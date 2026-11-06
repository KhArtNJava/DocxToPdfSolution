using System.Buffers.Binary;

namespace DocxToPdf.Fonts;

/// <summary>Ось вариативного шрифта (fvar).</summary>
public sealed record VarAxis(string Tag, double Min, double Default, double Max);

/// <summary>Именованное начертание вариативного шрифта (fvar): имя, координаты осей, PostScript-имя.</summary>
public sealed record VarInstance(string SubfamilyName, Dictionary<string, double> Coords, string? PostScriptName);

/// <summary>
/// Вариации OpenType: оси (fvar), нормализация координат с avar, хранилище дельт (ItemVariationStore, для HVAR и CFF2)
/// и дельты точек глифов TrueType (gvar) с интерполяцией нетронутых точек (IUP).
/// </summary>
static class Variations
{
    static ushort U16(byte[] d, int o) => BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o));
    static short S16(byte[] d, int o) => BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(o));
    static uint U32(byte[] d, int o) => BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o));
    static double Fixed(byte[] d, int o) => BinaryPrimitives.ReadInt32BigEndian(d.AsSpan(o)) / 65536.0;
    static double F2Dot14(byte[] d, int o) => S16(d, o) / 16384.0;

    // ---------------- fvar / avar ----------------

    public static List<VarAxis> ReadAxes(byte[] d, int fvar)
    {
        var axes = new List<VarAxis>();
        int arr = fvar + U16(d, fvar + 4), count = U16(d, fvar + 8), size = U16(d, fvar + 10);
        for (int i = 0; i < count; i++)
        {
            int a = arr + i * size;
            axes.Add(new VarAxis(System.Text.Encoding.ASCII.GetString(d, a, 4), Fixed(d, a + 4), Fixed(d, a + 8), Fixed(d, a + 12)));
        }
        return axes;
    }

    public static List<(int SubfamilyNameId, double[] Coords, int PsNameId)> ReadInstances(byte[] d, int fvar)
    {
        var list = new List<(int, double[], int)>();
        int arr = fvar + U16(d, fvar + 4), axisCount = U16(d, fvar + 8), axisSize = U16(d, fvar + 10);
        int instCount = U16(d, fvar + 12), instSize = U16(d, fvar + 14);
        int p = arr + axisCount * axisSize;
        for (int i = 0; i < instCount; i++, p += instSize)
        {
            if (p + instSize > d.Length) break;
            var c = new double[axisCount];
            for (int k = 0; k < axisCount; k++) c[k] = Fixed(d, p + 4 + k * 4);
            int ps = instSize >= axisCount * 4 + 6 ? U16(d, p + 4 + axisCount * 4) : 0xFFFF;
            list.Add((U16(d, p), c, ps));
        }
        return list;
    }

    /// <summary>Нормализованные координаты (−1…1) с поправкой avar.</summary>
    public static double[] Normalize(byte[] d, List<VarAxis> axes, IReadOnlyDictionary<string, double> coords, int avar)
    {
        var n = new double[axes.Count];
        for (int i = 0; i < axes.Count; i++)
        {
            var a = axes[i];
            double v = coords.TryGetValue(a.Tag, out var x) ? Math.Clamp(x, a.Min, a.Max) : a.Default;
            n[i] = v < a.Default ? (a.Default > a.Min ? (v - a.Default) / (a.Default - a.Min) : 0)
                 : v > a.Default ? (a.Max > a.Default ? (v - a.Default) / (a.Max - a.Default) : 0) : 0;
        }
        if (avar > 0)
        {
            int axisCount = U16(d, avar + 6), p = avar + 8;
            for (int i = 0; i < axisCount && i < n.Length; i++)
            {
                int pairs = U16(d, p);
                var map = new List<(double From, double To)>();
                for (int k = 0; k < pairs; k++) map.Add((F2Dot14(d, p + 2 + k * 4), F2Dot14(d, p + 4 + k * 4)));
                p += 2 + pairs * 4;
                if (map.Count < 2) continue;
                double v = n[i];
                for (int k = 1; k < map.Count; k++)
                {
                    if (v <= map[k].From)
                    {
                        var (f0, t0) = map[k - 1]; var (f1, t1) = map[k];
                        n[i] = f1 > f0 ? t0 + (t1 - t0) * (v - f0) / (f1 - f0) : t1;
                        break;
                    }
                }
            }
        }
        // Нормализованные координаты хранятся с точностью F2Dot14
        for (int i = 0; i < n.Length; i++) n[i] = Math.Round(n[i] * 16384) / 16384;
        return n;
    }

    /// <summary>Множитель региона (peak/start/end по каждой оси) для координат.</summary>
    static double RegionScalar(double[] coords, double[] start, double[] peak, double[] end)
    {
        double s = 1;
        for (int i = 0; i < peak.Length; i++)
        {
            double p = peak[i], c = i < coords.Length ? coords[i] : 0, st = start[i], en = end[i];
            if (p == 0) continue;                                     // ось не участвует
            if (st > p || p > en || (st < 0 && en > 0)) continue;     // недопустимый регион — ось игнорируется
            if (c == 0) return 0;
            if (c < st || c > en) return 0;
            if (c == p) continue;
            s *= c < p ? (c - st) / (p - st) : (en - c) / (en - p);
        }
        return s;
    }

    // ---------------- ItemVariationStore ----------------

    public sealed class ItemStore
    {
        readonly List<double[][]> _regions = [];                 // регион: [start, peak, end] по осям
        readonly List<(int[] Regions, List<int[]> Deltas)> _data = [];
        double[]? _scalars;

        public ItemStore(byte[] d, int off)
        {
            int rl = off + (int)U32(d, off + 2), count = U16(d, off + 6);
            int axisCount = U16(d, rl), regionCount = U16(d, rl + 2);
            for (int r = 0; r < regionCount; r++)
            {
                var st = new double[axisCount]; var pk = new double[axisCount]; var en = new double[axisCount];
                for (int a = 0; a < axisCount; a++)
                {
                    int q = rl + 4 + (r * axisCount + a) * 6;
                    st[a] = F2Dot14(d, q); pk[a] = F2Dot14(d, q + 2); en[a] = F2Dot14(d, q + 4);
                }
                _regions.Add([st, pk, en]);
            }
            for (int i = 0; i < count; i++)
            {
                int vd = off + (int)U32(d, off + 8 + i * 4);
                int itemCount = U16(d, vd), wordRaw = U16(d, vd + 2), regIdxCount = U16(d, vd + 4);
                bool longWords = (wordRaw & 0x8000) != 0;
                int wordCount = wordRaw & 0x7FFF;
                var regs = new int[regIdxCount];
                for (int k = 0; k < regIdxCount; k++) regs[k] = U16(d, vd + 6 + k * 2);
                int p = vd + 6 + regIdxCount * 2;
                var deltas = new List<int[]>(itemCount);
                for (int it = 0; it < itemCount; it++)
                {
                    var row = new int[regIdxCount];
                    for (int k = 0; k < regIdxCount; k++)
                    {
                        bool big = k < wordCount;
                        if (longWords) { row[k] = big ? BinaryPrimitives.ReadInt32BigEndian(d.AsSpan(p)) : S16(d, p); p += big ? 4 : 2; }
                        else { row[k] = big ? S16(d, p) : (sbyte)d[p]; p += big ? 2 : 1; }
                    }
                    deltas.Add(row);
                }
                _data.Add((regs, deltas));
            }
        }

        public int RegionCount(int outer) => outer < _data.Count ? _data[outer].Regions.Length : 0;

        public void SetCoords(double[] coords) => _scalars = _regions.Select(r => RegionScalar(coords, r[0], r[1], r[2])).ToArray();

        /// <summary>Итоговая дельта для пары (outer, inner).</summary>
        public double Delta(int outer, int inner)
        {
            if (_scalars == null || outer >= _data.Count) return 0;
            var (regs, deltas) = _data[outer];
            if (inner >= deltas.Count) return 0;
            double sum = 0;
            for (int k = 0; k < regs.Length; k++) sum += deltas[inner][k] * (regs[k] < _scalars.Length ? _scalars[regs[k]] : 0);
            return sum;
        }

        /// <summary>Смешение значений (blend в CFF2): для набора vsindex множители регионов по порядку.</summary>
        public double[] RegionScalars(int outer)
        {
            if (_scalars == null || outer >= _data.Count) return [];
            return _data[outer].Regions.Select(r => r < _scalars.Length ? _scalars[r] : 0).ToArray();
        }
    }

    /// <summary>DeltaSetIndexMap (HVAR): номер глифа → (outer, inner).</summary>
    public static (int Outer, int Inner) MapIndex(byte[] d, int map, int gid)
    {
        if (map <= 0) return (0, gid);
        int format = d[map], entryFormat = d[map + 1];
        int count = format == 0 ? U16(d, map + 2) : (int)U32(d, map + 2);
        int dataOff = map + (format == 0 ? 4 : 6);
        int size = ((entryFormat >> 4) & 3) + 1, innerBits = (entryFormat & 0xF) + 1;
        int idx = Math.Min(gid, count - 1);
        int v = 0;
        for (int k = 0; k < size; k++) v = v << 8 | d[dataOff + idx * size + k];
        return (v >> innerBits, v & ((1 << innerBits) - 1));
    }

    /// <summary>Дельты ширин из HVAR для всех глифов.</summary>
    public static double[]? HvarAdvanceDeltas(byte[] d, int hvar, double[] coords, int numGlyphs)
    {
        var store = new ItemStore(d, hvar + (int)U32(d, hvar + 4));
        store.SetCoords(coords);
        int map = (int)U32(d, hvar + 8);
        var res = new double[numGlyphs];
        for (int g = 0; g < numGlyphs; g++)
        {
            var (o, i) = MapIndex(d, map > 0 ? hvar + map : 0, g);
            res[g] = store.Delta(o, i);
        }
        return res;
    }

    // ---------------- gvar ----------------

    /// <summary>Дельты точек глифа (включая 4 фантомные) для нормализованных координат; null — у глифа нет вариаций.</summary>
    public static (double[] Dx, double[] Dy)? GlyphDeltas(byte[] d, int gvar, int gid, int pointCount, int[] endPts, double[] xs, double[] ys, double[] coords)
    {
        int axisCount = U16(d, gvar + 4), sharedCount = U16(d, gvar + 6);
        int sharedOff = gvar + (int)U32(d, gvar + 8);
        int glyphCount = U16(d, gvar + 12), flags = U16(d, gvar + 14);
        int dataArr = gvar + (int)U32(d, gvar + 16);
        if (gid >= glyphCount) return null;
        int Off(int i) => (flags & 1) != 0 ? (int)U32(d, gvar + 20 + i * 4) : U16(d, gvar + 20 + i * 2) * 2;
        int start = Off(gid), end = Off(gid + 1);
        if (end <= start) return null;
        int gv = dataArr + start;
        int tupleCount = U16(d, gv), dataOffset = U16(d, gv + 2);
        bool sharedPoints = (tupleCount & 0x8000) != 0;
        tupleCount &= 0x0FFF;
        int hdr = gv + 4, ser = gv + dataOffset;
        int[]? shared = null;
        if (sharedPoints) shared = ReadPoints(d, ref ser, pointCount);
        var dx = new double[pointCount]; var dy = new double[pointCount];
        for (int t = 0; t < tupleCount; t++)
        {
            int size = U16(d, hdr), tidx = U16(d, hdr + 2);
            hdr += 4;
            var peak = new double[axisCount];
            if ((tidx & 0x8000) != 0) { for (int a = 0; a < axisCount; a++) peak[a] = F2Dot14(d, hdr + a * 2); hdr += axisCount * 2; }
            else { int si = tidx & 0x0FFF; for (int a = 0; a < axisCount; a++) peak[a] = F2Dot14(d, sharedOff + (si * axisCount + a) * 2); }
            double[] st, en;
            if ((tidx & 0x4000) != 0)
            {
                st = new double[axisCount]; en = new double[axisCount];
                for (int a = 0; a < axisCount; a++) st[a] = F2Dot14(d, hdr + a * 2);
                hdr += axisCount * 2;
                for (int a = 0; a < axisCount; a++) en[a] = F2Dot14(d, hdr + a * 2);
                hdr += axisCount * 2;
            }
            else
            {
                st = peak.Select(p => Math.Min(p, 0)).ToArray();
                en = peak.Select(p => Math.Max(p, 0)).ToArray();
            }
            int p0 = ser;
            ser += size;
            double scalar = RegionScalar(coords, st, peak, en);
            if (scalar == 0) continue;
            int q = p0;
            int[]? pts = (tidx & 0x2000) != 0 ? ReadPoints(d, ref q, pointCount) : shared;
            int n = pts?.Length ?? pointCount;
            var ddx = ReadDeltas(d, ref q, n);
            var ddy = ReadDeltas(d, ref q, n);
            if (pts == null)
            {
                for (int i = 0; i < pointCount; i++) { dx[i] += ddx[i] * scalar; dy[i] += ddy[i] * scalar; }
            }
            else
            {
                // Нетронутые точки контура — интерполяция по соседним заданным (IUP)
                var tx = new double?[pointCount]; var ty = new double?[pointCount];
                for (int i = 0; i < pts.Length; i++) if (pts[i] < pointCount) { tx[pts[i]] = ddx[i]; ty[pts[i]] = ddy[i]; }
                Iup(tx, xs, endPts); Iup(ty, ys, endPts);
                for (int i = 0; i < pointCount; i++) { dx[i] += (tx[i] ?? 0) * scalar; dy[i] += (ty[i] ?? 0) * scalar; }
            }
        }
        return (dx, dy);
    }

    static int[] ReadPoints(byte[] d, ref int p, int pointCount)
    {
        int count = d[p++];
        if (count == 0) return Enumerable.Range(0, pointCount).ToArray();   // все точки
        if ((count & 0x80) != 0) count = (count & 0x7F) << 8 | d[p++];
        var res = new int[count];
        int i = 0, last = 0;
        while (i < count)
        {
            int ctl = d[p++];
            int run = (ctl & 0x7F) + 1;
            bool words = (ctl & 0x80) != 0;
            for (int k = 0; k < run && i < count; k++)
            {
                last += words ? U16(d, p) : d[p];
                p += words ? 2 : 1;
                res[i++] = last;
            }
        }
        return res;
    }

    static int[] ReadDeltas(byte[] d, ref int p, int count)
    {
        var res = new int[count];
        int i = 0;
        while (i < count)
        {
            int ctl = d[p++];
            int run = (ctl & 0x3F) + 1;
            for (int k = 0; k < run && i < count; k++)
            {
                if ((ctl & 0x80) != 0) res[i++] = 0;
                else if ((ctl & 0x40) != 0) { res[i++] = S16(d, p); p += 2; }
                else res[i++] = (sbyte)d[p++];
            }
        }
        return res;
    }

    /// <summary>IUP: дельты нетронутых точек каждого контура по ближайшим заданным с обеих сторон.</summary>
    static void Iup(double?[] delta, double[] orig, int[] endPts)
    {
        int start = 0;
        foreach (int end in endPts)
        {
            if (end >= delta.Length) break;
            var touched = Enumerable.Range(start, end - start + 1).Where(i => delta[i] != null).ToList();
            if (touched.Count == 0) { start = end + 1; continue; }
            if (touched.Count == 1) { for (int i = start; i <= end; i++) delta[i] ??= delta[touched[0]]; start = end + 1; continue; }
            int n = end - start + 1;
            for (int ti = 0; ti < touched.Count; ti++)
            {
                int a = touched[ti], b = touched[(ti + 1) % touched.Count];
                // Точки между a и b (по кругу)
                for (int k = (a - start + 1) % n; ; k = (k + 1) % n)
                {
                    int i = start + k;
                    if (i == b) break;
                    double oa = orig[a], ob = orig[b], oi = orig[i], da = delta[a]!.Value, db = delta[b]!.Value;
                    double lo = Math.Min(oa, ob), hi = Math.Max(oa, ob), dlo = oa <= ob ? da : db, dhi = oa <= ob ? db : da;
                    if (oi <= lo) delta[i] = dlo;
                    else if (oi >= hi) delta[i] = dhi;
                    else delta[i] = hi > lo ? dlo + (dhi - dlo) * (oi - lo) / (hi - lo) : dlo;
                }
            }
            start = end + 1;
        }
    }
}

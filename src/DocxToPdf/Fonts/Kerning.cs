using System.Buffers.Binary;

namespace DocxToPdf.Fonts;

/// <summary>
/// Парный кернинг шрифта: GPOS (функция kern, поиск PairPos форматов 1 и 2, в т.ч. через расширения — тип 9)
/// или, если в GPOS кернинга нет, старая таблица kern (формат 0). Значения — в единицах шрифта, с кешем пар.
/// </summary>
sealed class KerningData
{
    readonly byte[] _d;
    readonly List<int> _pairSubtables = [];          // смещения подтаблиц PairPos
    readonly Dictionary<uint, short>? _legacy;
    readonly Dictionary<uint, short> _cache = [];
    public bool Any => _pairSubtables.Count > 0 || _legacy is { Count: > 0 };

    public KerningData(byte[] data, Dictionary<string, (int Offset, int Length)> tables)
    {
        _d = data;
        try { if (tables.TryGetValue("GPOS", out var gpos)) ReadGpos(gpos.Offset); } catch { _pairSubtables.Clear(); }
        if (_pairSubtables.Count == 0 && tables.TryGetValue("kern", out var kern))
        {
            try { _legacy = ReadKern(kern.Offset, kern.Length); } catch { _legacy = null; }
        }
    }

    int U16(int o) => BinaryPrimitives.ReadUInt16BigEndian(_d.AsSpan(o));
    short S16(int o) => BinaryPrimitives.ReadInt16BigEndian(_d.AsSpan(o));
    uint U32(int o) => BinaryPrimitives.ReadUInt32BigEndian(_d.AsSpan(o));

    void ReadGpos(int g)
    {
        int featureList = g + U16(g + 6), lookupList = g + U16(g + 8);
        var lookups = new SortedSet<int>();
        int fc = U16(featureList);
        for (int i = 0; i < fc; i++)
        {
            int rec = featureList + 2 + i * 6;
            string tag = System.Text.Encoding.ASCII.GetString(_d, rec, 4);
            if (tag != "kern") continue;
            int f = featureList + U16(rec + 4);
            int n = U16(f + 2);
            for (int k = 0; k < n; k++) lookups.Add(U16(f + 4 + k * 2));
        }
        int lc = U16(lookupList);
        foreach (int li in lookups)
        {
            if (li >= lc) continue;
            int lk = lookupList + U16(lookupList + 2 + li * 2);
            int type = U16(lk), subCount = U16(lk + 4);
            for (int s = 0; s < subCount; s++)
            {
                int st = lk + U16(lk + 6 + s * 2);
                int t = type;
                if (t == 9)
                {
                    // Расширение: настоящий тип и 32-битное смещение
                    t = U16(st + 2);
                    st += (int)U32(st + 4);
                }
                if (t == 2 && U16(st) is 1 or 2) _pairSubtables.Add(st);
            }
        }
    }

    Dictionary<uint, short> ReadKern(int k, int len)
    {
        var res = new Dictionary<uint, short>();
        if (U16(k) != 0) return res;                 // формат Apple (версия 1.0) не поддерживается
        int n = U16(k + 2), p = k + 4;
        for (int i = 0; i < n && p + 6 <= k + len; i++)
        {
            int subLen = U16(p + 2), coverage = U16(p + 4);
            if (coverage >> 8 == 0 && (coverage & 1) == 1 && (coverage & 4) == 0)
            {
                int pairs = U16(p + 6);
                for (int j = 0; j < pairs; j++)
                {
                    int r = p + 14 + j * 6;
                    uint key = (uint)U16(r) << 16 | (uint)U16(r + 2);
                    res.TryAdd(key, S16(r + 4));
                }
            }
            p += Math.Max(6, subLen);
        }
        return res;
    }

    int Coverage(int cov, int gid)
    {
        int fmt = U16(cov);
        if (fmt == 1)
        {
            int n = U16(cov + 2), lo = 0, hi = n - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1, v = U16(cov + 4 + mid * 2);
                if (v == gid) return mid;
                if (v < gid) lo = mid + 1; else hi = mid - 1;
            }
            return -1;
        }
        if (fmt == 2)
        {
            int n = U16(cov + 2), lo = 0, hi = n - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1, r = cov + 4 + mid * 6;
                int start = U16(r), end = U16(r + 2);
                if (gid < start) hi = mid - 1;
                else if (gid > end) lo = mid + 1;
                else return U16(r + 4) + gid - start;
            }
        }
        return -1;
    }

    int ClassOf(int cd, int gid)
    {
        int fmt = U16(cd);
        if (fmt == 1)
        {
            int start = U16(cd + 2), n = U16(cd + 4);
            return gid >= start && gid < start + n ? U16(cd + 6 + (gid - start) * 2) : 0;
        }
        if (fmt == 2)
        {
            int n = U16(cd + 2), lo = 0, hi = n - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1, r = cd + 4 + mid * 6;
                int start = U16(r), end = U16(r + 2);
                if (gid < start) hi = mid - 1;
                else if (gid > end) lo = mid + 1;
                else return U16(r + 4);
            }
        }
        return 0;
    }

    static int Size(int vf) => System.Numerics.BitOperations.PopCount((uint)(vf & 0xFF)) * 2;

    /// <summary>Сдвиг в ValueRecord: XAdvance (0x4) или XPlacement (0x1); 0, если поля нет.</summary>
    int Field(int rec, int vf, int bit)
    {
        if ((vf & bit) == 0) return 0;
        int off = System.Numerics.BitOperations.PopCount((uint)(vf & (bit - 1))) * 2;
        return S16(rec + off);
    }

    /// <summary>Поправка к ширине левого глифа пары (единицы шрифта).</summary>
    public int Get(ushort left, ushort right)
    {
        uint key = (uint)left << 16 | right;
        if (_cache.TryGetValue(key, out var c)) return c;
        int v = 0;
        if (_legacy != null) v = _legacy.TryGetValue(key, out var lv) ? lv : 0;
        else
        {
            foreach (int st in _pairSubtables)
            {
                int ci = Coverage(st + U16(st + 2), left);
                if (ci < 0) continue;
                int vf1 = U16(st + 4), vf2 = U16(st + 6);
                int s1 = Size(vf1), s2 = Size(vf2);
                if (U16(st) == 1)
                {
                    int setCount = U16(st + 8);
                    if (ci >= setCount) continue;
                    int ps = st + U16(st + 10 + ci * 2);
                    int n = U16(ps), recSize = 2 + s1 + s2, lo = 0, hi = n - 1;
                    bool found = false;
                    while (lo <= hi)
                    {
                        int mid = (lo + hi) >> 1, r = ps + 2 + mid * recSize, g2 = U16(r);
                        if (g2 == right) { v = Field(r + 2, vf1, 4) + Field(r + 2 + s1, vf2, 1); found = true; break; }
                        if (g2 < right) lo = mid + 1; else hi = mid - 1;
                    }
                    if (found) break;
                }
                else
                {
                    int c1 = ClassOf(st + U16(st + 8), left), c2 = ClassOf(st + U16(st + 10), right);
                    int n1 = U16(st + 12), n2 = U16(st + 14);
                    if (c1 >= n1 || c2 >= n2) break;
                    int r = st + 16 + (c1 * n2 + c2) * (s1 + s2);
                    v = Field(r, vf1, 4) + Field(r + s1, vf2, 1);
                    break;   // класс пары найден — подтаблица применена
                }
            }
        }
        _cache[key] = (short)Math.Clamp(v, short.MinValue, short.MaxValue);
        return v;
    }
}

using System.Buffers.Binary;
using System.Globalization;
using System.Text;

namespace DocxToPdf.Fonts;

/// <summary>
/// Таблица CFF шрифта OpenType (контуры Type 2): разбор, подмножество для PDF в CID-форме
/// (CID = номер глифа исходного шрифта, поэтому подходит кодировка Identity-H) и контуры глифов.
/// </summary>
sealed class CffFont
{
    readonly byte[] _d;
    readonly int _start, _end;
    readonly List<(int S, int E)> _charStrings;
    readonly (int S, int E) _globalSubrsIndex;       // INDEX глобальных подпрограмм целиком (байты)
    readonly List<(int S, int E)> _globalSubrs;
    readonly List<Entry> _top;
    readonly List<FontDict> _fds = [];
    readonly byte[]? _fdSelect;                      // номер Font DICT для каждого глифа (CID-шрифты)
    public double Scale { get; }                     // FontMatrix[0] (обычно 0,001)

    /// <summary>Запись DICT: оператор (escape-операторы — 1200 + код), операнды как числа и исходные байты.</summary>
    sealed record Entry(int Op, List<double> Operands, byte[] Raw);

    sealed class FontDict
    {
        public List<Entry> Entries = [];          // Font DICT (для CID-шрифтов), без Private
        public List<Entry> Private = [];
        public (int S, int E)? LocalSubrsIndex;
        public List<(int S, int E)> LocalSubrs = [];
        public double DefaultWidthX, NominalWidthX;
        public int VsIndex;                       // набор регионов по умолчанию (CFF2)
    }

    readonly bool _cff2;                             // таблица CFF2 (вариативные контуры)
    readonly Variations.ItemStore? _store;           // VariationStore (CFF2)
    readonly bool _hasMatrix;

    public CffFont(byte[] data, int offset, int length, bool cff2 = false, int unitsPerEm = 1000)
    {
        _d = data;
        _start = offset;
        _end = Math.Min(data.Length, offset + length);
        _cff2 = cff2;
        int hdrSize = data[offset + 2];
        int p = offset + hdrSize;
        if (cff2)
        {
            // CFF2: Top DICT не в INDEX (длина в заголовке), строк и имён нет
            int topLen = U16(offset + 3);
            _top = ParseDict(p, p + topLen);
            p += topLen;
            int gsStart = p;
            _globalSubrs = ReadIndex(ref p);
            _globalSubrsIndex = (gsStart, p);
            if (Find(_top, 24) is { Operands.Count: >= 1 } vs && vs.Operands[0] > 0)
                _store = new Variations.ItemStore(_d, _start + (int)vs.Operands[0] + 2);
        }
        else
        {
            var names = ReadIndex(ref p);
            var topDicts = ReadIndex(ref p);
            ReadIndex(ref p);                           // строки (в подмножестве не нужны)
            int gsStart = p;
            _globalSubrs = ReadIndex(ref p);
            _globalSubrsIndex = (gsStart, p);
            if (topDicts.Count == 0) throw new InvalidDataException("CFF без Top DICT");
            _top = ParseDict(topDicts[0].S, topDicts[0].E);
        }

        int csOff = (int)Op(_top, 17, 0);
        if (csOff <= 0) throw new InvalidDataException("CFF без CharStrings");
        int q = _start + csOff;
        _charStrings = ReadIndex(ref q);
        _hasMatrix = Find(_top, 1207) is { Operands.Count: >= 1 } fm0 && fm0.Operands[0] > 0;
        Scale = _hasMatrix ? Find(_top, 1207)!.Operands[0] : cff2 ? 1.0 / Math.Max(1, unitsPerEm) : 0.001;

        if (Find(_top, 1230) != null || cff2 && Find(_top, 1236) != null)
        {
            // CID-шрифт (или CFF2): массив Font DICT и выбор FD для каждого глифа
            int fdArr = _start + (int)Op(_top, 1236, 0);
            foreach (var (s, e) in ReadIndex(ref fdArr))
                _fds.Add(LoadFontDict(ParseDict(s, e)));
            if (Find(_top, 1237) != null) _fdSelect = ReadFdSelect(_start + (int)Op(_top, 1237, 0), _charStrings.Count);
        }
        else
        {
            // Обычный (не CID) шрифт: один FD; FontMatrix остаётся только в Top DICT
            var fd = LoadFontDict(_top);
            fd.Entries.RemoveAll(e => e.Op == 1207);
            _fds.Add(fd);
        }
        if (_fds.Count == 0) _fds.Add(new FontDict());
    }

    /// <summary>Нормализованные координаты осей для смешения (blend) контуров CFF2.</summary>
    public void SetVariation(double[] norm) => _store?.SetCoords(norm);

    FontDict LoadFontDict(List<Entry> dict)
    {
        var fd = new FontDict { Entries = dict.Where(e => e.Op is not (18 or 1238) && !IsTopOnly(e.Op)).ToList() };
        if (Find(dict, 18) is { Operands.Count: >= 2 } pr)
        {
            int size = (int)pr.Operands[0], off = _start + (int)pr.Operands[1];
            if (off >= _start && off + size <= _end)
            {
                fd.Private = ParseDict(off, off + size);
                fd.DefaultWidthX = Op(fd.Private, 20, 0);
                fd.NominalWidthX = Op(fd.Private, 21, 0);
                fd.VsIndex = (int)Op(fd.Private, 22, 0);
                if (Find(fd.Private, 19) is { Operands.Count: >= 1 } sub)
                {
                    int ls = off + (int)sub.Operands[0], le = ls;
                    fd.LocalSubrs = ReadIndex(ref le);
                    fd.LocalSubrsIndex = (ls, le);
                }
            }
        }
        return fd;
    }

    // Операторы Top DICT, которые не переносятся в Font DICT подмножества
    static bool IsTopOnly(int op) => op is 0 or 1 or 2 or 3 or 4 or 13 or 14 or 15 or 16 or 17 or 1200 or 1201 or 1202 or 1203 or 1204 or 1205
        or 1206 or 1208 or 1220 or 1221 or 1222 or 1223 or 1230 or 1231 or 1232 or 1233 or 1234 or 1235 or 1236 or 1237 or 1238 or 5;

    byte[] ReadFdSelect(int p, int nGlyphs)
    {
        var r = new byte[nGlyphs];
        int format = _d[p];
        if (format == 0) { for (int i = 0; i < nGlyphs; i++) r[i] = _d[p + 1 + i]; }
        else if (format == 3)
        {
            int n = U16(p + 1);
            for (int i = 0; i < n; i++)
            {
                int first = U16(p + 3 + i * 3), fd = _d[p + 5 + i * 3], next = U16(p + 6 + i * 3);
                for (int g = first; g < next && g < nGlyphs; g++) r[g] = (byte)fd;
            }
        }
        else if (format == 4)
        {
            int n = (int)U32(p + 1);
            for (int i = 0; i < n; i++)
            {
                int o = p + 5 + i * 6;
                int first = (int)U32(o), fd = U16(o + 4), next = (int)U32(o + 6);
                for (int g = first; g < next && g < nGlyphs; g++) r[g] = (byte)fd;
            }
        }
        return r;
    }

    uint U32(int o) => BinaryPrimitives.ReadUInt32BigEndian(_d.AsSpan(o));
    int U16(int o) => BinaryPrimitives.ReadUInt16BigEndian(_d.AsSpan(o));

    List<(int S, int E)> ReadIndex(ref int p)
    {
        var list = new List<(int, int)>();
        int cw = _cff2 ? 4 : 2;
        int count = _cff2 ? (int)U32(p) : U16(p);
        if (count == 0) { p += cw; return list; }
        int offSize = _d[p + cw];
        int offBase = p + cw + 1;
        int dataStart = offBase + (count + 1) * offSize - 1;
        int Off(int i)
        {
            int v = 0, o = offBase + i * offSize;
            for (int k = 0; k < offSize; k++) v = v << 8 | _d[o + k];
            return v;
        }
        for (int i = 0; i < count; i++) list.Add((dataStart + Off(i), dataStart + Off(i + 1)));
        p = dataStart + Off(count);
        return list;
    }

    List<Entry> ParseDict(int p, int end)
    {
        var list = new List<Entry>();
        var ops = new List<double>();
        int rawStart = p;
        while (p < end)
        {
            int b0 = _d[p];
            if (b0 <= 27)
            {
                int op = b0;
                p++;
                if (b0 == 12) { op = 1200 + _d[p]; p++; }
                if (op == 23 && ops.Count > 0)
                {
                    // blend в DICT (CFF2): оставляем базовые значения, дельты отбрасываем
                    int n = (int)ops[^1];
                    ops.RemoveAt(ops.Count - 1);
                    int vsi = list.LastOrDefault(e => e.Op == 22) is { Operands.Count: > 0 } vx ? (int)vx.Operands[0] : 0;
                    int baseStart = ops.Count - n * (1 + (_store?.RegionCount(vsi) ?? 0));
                    if (baseStart < 0) baseStart = Math.Max(0, ops.Count - n);
                    if (n > 0) { var b = ops.GetRange(baseStart, n); ops.RemoveRange(baseStart, ops.Count - baseStart); ops.AddRange(b); }
                    continue;
                }
                // Исходные байты операндов (без оператора)
                int opLen = op >= 1200 ? 2 : 1;
                list.Add(new Entry(op, ops, _d.AsSpan(rawStart, p - opLen - rawStart).ToArray()));
                ops = [];
                rawStart = p;
            }
            else if (b0 == 28) { ops.Add((short)(_d[p + 1] << 8 | _d[p + 2])); p += 3; }
            else if (b0 == 29) { ops.Add(BinaryPrimitives.ReadInt32BigEndian(_d.AsSpan(p + 1))); p += 5; }
            else if (b0 == 30)
            {
                // Вещественное число: полубайты
                var sb = new StringBuilder();
                p++;
                bool done = false;
                while (!done && p < end)
                {
                    int b = _d[p++];
                    foreach (int nib in new[] { b >> 4, b & 15 })
                    {
                        if (nib <= 9) sb.Append((char)('0' + nib));
                        else if (nib == 10) sb.Append('.');
                        else if (nib == 11) sb.Append('E');
                        else if (nib == 12) sb.Append("E-");
                        else if (nib == 14) sb.Append('-');
                        else if (nib == 15) { done = true; break; }
                    }
                }
                ops.Add(double.TryParse(sb.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0);
            }
            else if (b0 >= 32 && b0 <= 246) { ops.Add(b0 - 139); p++; }
            else if (b0 >= 247 && b0 <= 250) { ops.Add((b0 - 247) * 256 + _d[p + 1] + 108); p += 2; }
            else if (b0 >= 251 && b0 <= 254) { ops.Add(-(b0 - 251) * 256 - _d[p + 1] - 108); p += 2; }
            else p++;
        }
        return list;
    }

    static Entry? Find(List<Entry> dict, int op) => dict.FirstOrDefault(e => e.Op == op);
    static double Op(List<Entry> dict, int op, double def) => Find(dict, op) is { Operands.Count: > 0 } e ? e.Operands[^1] : def;

    int FdOf(int gid) => _fdSelect != null && gid < _fdSelect.Length && _fdSelect[gid] < _fds.Count ? _fdSelect[gid] : 0;

    // ================================================================
    //  Подмножество: CID-keyed CFF, CID = исходный номер глифа
    // ================================================================

    static void Card16(MemoryStream m, int v) { m.WriteByte((byte)(v >> 8)); m.WriteByte((byte)v); }
    static void Int5(MemoryStream m, int v) { m.WriteByte(29); m.WriteByte((byte)(v >> 24)); m.WriteByte((byte)(v >> 16)); m.WriteByte((byte)(v >> 8)); m.WriteByte((byte)v); }
    static void Operator(MemoryStream m, int op) { if (op >= 1200) { m.WriteByte(12); m.WriteByte((byte)(op - 1200)); } else m.WriteByte((byte)op); }

    /// <summary>Вещественное число DICT (оператор 30, полубайты).</summary>
    static void Real(MemoryStream m, double v)
    {
        string t = v.ToString("0.##########E+0", CultureInfo.InvariantCulture);
        if (v == 0) t = "0";
        var nibs = new List<int>();
        for (int i = 0; i < t.Length; i++)
        {
            char c = t[i];
            if (c is >= '0' and <= '9') nibs.Add(c - '0');
            else if (c == '.') nibs.Add(10);
            else if (c == '-' && i == 0) nibs.Add(14);
            else if (c == 'E') { if (t[i + 1] == '-') { nibs.Add(12); i++; } else { nibs.Add(11); if (t[i + 1] == '+') i++; } }
        }
        nibs.Add(15);
        if (nibs.Count % 2 == 1) nibs.Add(15);
        m.WriteByte(30);
        for (int i = 0; i < nibs.Count; i += 2) m.WriteByte((byte)(nibs[i] << 4 | nibs[i + 1]));
    }

    static byte[] Index(IReadOnlyList<byte[]> items)
    {
        var m = new MemoryStream();
        Card16(m, items.Count);
        if (items.Count == 0) return m.ToArray();
        long total = items.Sum(i => (long)i.Length) + 1;
        int offSize = total < 0x100 ? 1 : total < 0x10000 ? 2 : total < 0x1000000 ? 3 : 4;
        m.WriteByte((byte)offSize);
        int off = 1;
        void Off(int v) { for (int k = offSize - 1; k >= 0; k--) m.WriteByte((byte)(v >> (8 * k))); }
        Off(off);
        foreach (var it in items) { off += it.Length; Off(off); }
        foreach (var it in items) m.Write(it);
        return m.ToArray();
    }

    byte[] Bytes((int S, int E) r) => _d.AsSpan(r.S, r.E - r.S).ToArray();

    public byte[] Subset(IEnumerable<ushort> used, string fontName)
    {
        int nGlyphs = _charStrings.Count;
        var glyphs = new SortedSet<int> { 0 };
        foreach (var g in used) if (g < nGlyphs) glyphs.Add(g);
        var list = glyphs.ToList();
        int cidCount = list[^1] + 1;

        // CFF2 переводится в обычный CFF: контуры смешиваются при текущих координатах осей
        var charStrings = Index(list.Select(g => _cff2 ? StaticCharString(g) : Bytes(_charStrings[g])).ToList());
        var gsubrs = _cff2 ? new byte[] { 0, 0 } : Bytes(_globalSubrsIndex);
        var strings = Index([Encoding.ASCII.GetBytes("Adobe"), Encoding.ASCII.GetBytes("Identity")]);
        var nameIdx = Index([Encoding.ASCII.GetBytes(fontName)]);

        // charset (формат 0): CID для каждого глифа, кроме .notdef
        var charset = new MemoryStream();
        charset.WriteByte(0);
        foreach (var g in list.Skip(1)) Card16(charset, g);
        // FDSelect (формат 0)
        var fdSelect = new MemoryStream();
        fdSelect.WriteByte(0);
        foreach (var g in list) fdSelect.WriteByte((byte)FdOf(g));

        // Private DICT каждого FD: исходные записи, Subrs — сразу после словаря
        var privates = new List<(byte[] Dict, byte[] Subrs)>();
        foreach (var fd in _fds)
        {
            var pm = new MemoryStream();
            if (_cff2) { privates.Add(([139, 20, 139, 21], [])); continue; }   // defaultWidthX 0, nominalWidthX 0
            var subrs = fd.LocalSubrsIndex is { } li ? Bytes(li) : null;
            foreach (var e in fd.Private.Where(e => e.Op != 19)) { pm.Write(e.Raw); Operator(pm, e.Op); }
            int dictLen = (int)pm.Length + (subrs != null ? 6 : 0);
            if (subrs != null) { Int5(pm, dictLen); Operator(pm, 19); }
            privates.Add((pm.ToArray(), subrs ?? []));
        }

        // Top DICT фиксированного размера (смещения — 5-байтовыми числами)
        byte[] TopDict(int charsetOff, int fdSelectOff, int csOff, int fdArrayOff)
        {
            var m = new MemoryStream();
            Int5(m, 391); Int5(m, 392); Int5(m, 0); Operator(m, 1230);          // ROS: Adobe-Identity-0
            foreach (var e in _top.Where(e => e.Op is 5 or 1207))   // FontBBox, FontMatrix
            { m.Write(e.Raw); Operator(m, e.Op); }
            if (!_hasMatrix && _cff2)
            {   // В CFF2 матрица по умолчанию задаётся unitsPerEm; в CFF — явно
                Real(m, Scale); Real(m, 0); Real(m, 0); Real(m, Scale); Real(m, 0); Real(m, 0); Operator(m, 1207);
            }
            Int5(m, cidCount); Operator(m, 1234);                                // CIDCount
            Int5(m, charsetOff); Operator(m, 15);
            Int5(m, csOff); Operator(m, 17);
            Int5(m, fdArrayOff); Operator(m, 1236);
            Int5(m, fdSelectOff); Operator(m, 1237);
            return m.ToArray();
        }
        int topLen = Index([TopDict(0, 0, 0, 0)]).Length;
        int pos = 4 + nameIdx.Length + topLen + strings.Length + gsubrs.Length;
        int charsetOffset = pos; pos += (int)charset.Length;
        int fdSelectOffset = pos; pos += (int)fdSelect.Length;
        int csOffset = pos; pos += charStrings.Length;

        // FDArray: Font DICT с Private (размер, смещение)
        int fdArrayOffset = pos;
        byte[] FdArray(List<int> privOffsets)
        {
            var items = new List<byte[]>();
            for (int i = 0; i < _fds.Count; i++)
            {
                var m = new MemoryStream();
                foreach (var e in _fds[i].Entries.Where(e => e.Op is 1207 or 1214 or 1209 or 1210 or 1211 or 1212 or 1213 or 1217 or 1218 or 1219))
                { m.Write(e.Raw); Operator(m, e.Op); }
                Int5(m, privates[i].Dict.Length); Int5(m, privOffsets[i]); Operator(m, 18);
                items.Add(m.ToArray());
            }
            return Index(items);
        }
        int fdArrayLen = FdArray([.. Enumerable.Repeat(0, _fds.Count)]).Length;
        pos += fdArrayLen;
        var privOffsets = new List<int>();
        foreach (var (dict, subrs) in privates) { privOffsets.Add(pos); pos += dict.Length + subrs.Length; }

        var o = new MemoryStream();
        o.Write([1, 0, 4, 4]);
        o.Write(nameIdx);
        o.Write(Index([TopDict(charsetOffset, fdSelectOffset, csOffset, fdArrayOffset)]));
        o.Write(strings);
        o.Write(gsubrs);
        o.Write(charset.ToArray());
        o.Write(fdSelect.ToArray());
        o.Write(charStrings);
        o.Write(FdArray(privOffsets));
        foreach (var (dict, subrs) in privates) { o.Write(dict); o.Write(subrs); }
        return o.ToArray();
    }

    // ================================================================
    //  Контуры глифов: интерпретатор Type 2 charstrings
    // ================================================================

    static int Bias(int n) => n < 1240 ? 107 : n < 33900 ? 1131 : 32768;

    /// <summary>Контур глифа (операторы пути PDF) в единицах 1/Scale (обычно = единицы шрифта).</summary>
    public string GlyphPath(int gid, double unitsPerEm)
    {
        double k = Scale * unitsPerEm;
        var sb = new StringBuilder();
        string N(double v) => Math.Round(v * k, 1).ToString("0.#", CultureInfo.InvariantCulture);
        bool open = false;
        Outline(gid,
            (x, y) => { if (open) sb.Append("h "); sb.Append(N(x)).Append(' ').Append(N(y)).Append(" m "); open = true; },
            (x, y) => sb.Append(N(x)).Append(' ').Append(N(y)).Append(" l "),
            (x1, y1, x2, y2, x, y) => sb.Append($"{N(x1)} {N(y1)} {N(x2)} {N(y2)} {N(x)} {N(y)} c "));
        if (open) sb.Append("h ");
        return sb.ToString();
    }

    /// <summary>
    /// Обход контура глифа (абсолютные координаты в единицах шрифта): интерпретатор Type 2
    /// и CFF2 charstrings (blend/vsindex смешиваются по текущим координатам вариативного шрифта).
    /// </summary>
    void Outline(int gid, Action<double, double> moveTo, Action<double, double> lineTo,
        Action<double, double, double, double, double, double> curveTo)
    {
        if (gid < 0 || gid >= _charStrings.Count) return;
        var fd = _fds[FdOf(gid)];
        var stack = new List<double>();
        double x = 0, y = 0;
        int nStems = 0;
        bool widthDone = _cff2;
        var trans = new double[32];
        int depth = 0;
        int vsIndex = fd.VsIndex;

        void Move(double dx, double dy) { x += dx; y += dy; moveTo(x, y); }
        void Line(double dx, double dy) { x += dx; y += dy; lineTo(x, y); }
        void Curve(double dx1, double dy1, double dx2, double dy2, double dx3, double dy3)
        {
            double x1 = x + dx1, y1 = y + dy1, x2 = x1 + dx2, y2 = y1 + dy2;
            x = x2 + dx3; y = y2 + dy3;
            curveTo(x1, y1, x2, y2, x, y);
        }
        // Первая очищающая стек команда может нести ширину глифа (в CFF2 ширины нет)
        void Width(bool hasWidth) { if (!widthDone && hasWidth && stack.Count > 0) stack.RemoveAt(0); widthDone = true; }

        bool Run(int s, int e)
        {
            if (++depth > 12) return false;
            int p = s;
            while (p < e)
            {
                int b0 = _d[p];
                if (b0 >= 32 || b0 == 28)
                {
                    if (b0 == 28) { stack.Add((short)(_d[p + 1] << 8 | _d[p + 2])); p += 3; }
                    else if (b0 <= 246) { stack.Add(b0 - 139); p++; }
                    else if (b0 <= 250) { stack.Add((b0 - 247) * 256 + _d[p + 1] + 108); p += 2; }
                    else if (b0 <= 254) { stack.Add(-(b0 - 251) * 256 - _d[p + 1] - 108); p += 2; }
                    else { stack.Add(BinaryPrimitives.ReadInt32BigEndian(_d.AsSpan(p + 1)) / 65536.0); p += 5; }
                    continue;
                }
                p++;
                switch (b0)
                {
                    case 1 or 3 or 18 or 23:   // hstem, vstem, hstemhm, vstemhm
                        Width(stack.Count % 2 == 1);
                        nStems += stack.Count / 2;
                        stack.Clear();
                        break;
                    case 19 or 20:             // hintmask, cntrmask (неявный vstem)
                        Width(stack.Count % 2 == 1);
                        nStems += stack.Count / 2;
                        stack.Clear();
                        p += (nStems + 7) / 8;
                        break;
                    case 21: Width(stack.Count > 2); if (stack.Count >= 2) Move(stack[0], stack[1]); stack.Clear(); break;
                    case 22: Width(stack.Count > 1); if (stack.Count >= 1) Move(stack[0], 0); stack.Clear(); break;
                    case 4: Width(stack.Count > 1); if (stack.Count >= 1) Move(0, stack[0]); stack.Clear(); break;
                    case 5: for (int i = 0; i + 1 < stack.Count; i += 2) Line(stack[i], stack[i + 1]); stack.Clear(); break;
                    case 6 or 7:
                    {
                        bool horiz = b0 == 6;
                        foreach (var v in stack) { if (horiz) Line(v, 0); else Line(0, v); horiz = !horiz; }
                        stack.Clear();
                        break;
                    }
                    case 8: for (int i = 0; i + 5 < stack.Count; i += 6) Curve(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]); stack.Clear(); break;
                    case 24:
                    {
                        int i = 0;
                        for (; i + 5 < stack.Count - 2; i += 6) Curve(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]);
                        if (i + 1 < stack.Count) Line(stack[i], stack[i + 1]);
                        stack.Clear();
                        break;
                    }
                    case 25:
                    {
                        int i = 0;
                        for (; i + 1 < stack.Count - 6; i += 2) Line(stack[i], stack[i + 1]);
                        if (i + 5 < stack.Count) Curve(stack[i], stack[i + 1], stack[i + 2], stack[i + 3], stack[i + 4], stack[i + 5]);
                        stack.Clear();
                        break;
                    }
                    case 26:   // vvcurveto
                    {
                        int i = 0; double dx1 = 0;
                        if (stack.Count % 4 == 1) dx1 = stack[i++];
                        for (; i + 3 < stack.Count; i += 4) { Curve(dx1, stack[i], stack[i + 1], stack[i + 2], 0, stack[i + 3]); dx1 = 0; }
                        stack.Clear();
                        break;
                    }
                    case 27:   // hhcurveto
                    {
                        int i = 0; double dy1 = 0;
                        if (stack.Count % 4 == 1) dy1 = stack[i++];
                        for (; i + 3 < stack.Count; i += 4) { Curve(stack[i], dy1, stack[i + 1], stack[i + 2], stack[i + 3], 0); dy1 = 0; }
                        stack.Clear();
                        break;
                    }
                    case 30 or 31:   // vhcurveto, hvcurveto
                    {
                        bool horiz = b0 == 31;
                        int i = 0;
                        while (i + 3 < stack.Count)
                        {
                            double last = stack.Count - i == 5 ? stack[i + 4] : 0;
                            if (horiz) Curve(stack[i], 0, stack[i + 1], stack[i + 2], last, stack[i + 3]);
                            else Curve(0, stack[i], stack[i + 1], stack[i + 2], stack[i + 3], last);
                            i += 4;
                            horiz = !horiz;
                        }
                        stack.Clear();
                        break;
                    }
                    case 10 or 29:   // callsubr, callgsubr
                    {
                        if (stack.Count == 0) return false;
                        var subrs = b0 == 10 ? fd.LocalSubrs : _globalSubrs;
                        int idx = (int)stack[^1] + Bias(subrs.Count);
                        stack.RemoveAt(stack.Count - 1);
                        if (idx < 0 || idx >= subrs.Count) return false;
                        if (!Run(subrs[idx].S, subrs[idx].E)) return false;
                        break;
                    }
                    case 11: depth--; return true;   // return
                    case 15 when _cff2:              // vsindex
                        if (stack.Count > 0) vsIndex = (int)stack[^1];
                        stack.Clear();
                        break;
                    case 16 when _cff2:              // blend: n базовых значений + n·k дельт, затем n
                    {
                        if (stack.Count == 0) break;
                        int n = (int)stack[^1];
                        stack.RemoveAt(stack.Count - 1);
                        var sc = _store?.RegionScalars(vsIndex) ?? [];
                        int k = _store?.RegionCount(vsIndex) ?? 0;
                        int b = stack.Count - n * (k + 1);
                        if (n < 0 || b < 0) { stack.Clear(); break; }
                        for (int v = 0; v < n; v++)
                        {
                            double sum = stack[b + v];
                            for (int r = 0; r < k; r++) sum += stack[b + n + v * k + r] * (r < sc.Length ? sc[r] : 0);
                            stack[b + v] = sum;
                        }
                        stack.RemoveRange(b + n, n * k);
                        break;
                    }
                    case 14:   // endchar
                        Width(stack.Count is 1 or 5);
                        depth--;
                        return false;
                    case 12:
                    {
                        int op = _d[p++];
                        var a = stack;
                        switch (op)
                        {
                            case 35 when a.Count >= 13: Curve(a[0], a[1], a[2], a[3], a[4], a[5]); Curve(a[6], a[7], a[8], a[9], a[10], a[11]); break;
                            case 34 when a.Count >= 7: { double y0 = y; Curve(a[0], 0, a[1], a[2], a[3], 0); Curve(a[4], 0, a[5], y0 - y, a[6], 0); break; }
                            case 36 when a.Count >= 9: { double y0 = y; Curve(a[0], a[1], a[2], a[3], a[4], 0); Curve(a[5], 0, a[6], a[7], a[8], y0 - y - a[7]); break; }
                            case 37 when a.Count >= 11:
                            {
                                double sx = a[0] + a[2] + a[4] + a[6] + a[8], sy = a[1] + a[3] + a[5] + a[7] + a[9];
                                Curve(a[0], a[1], a[2], a[3], a[4], a[5]);
                                if (Math.Abs(sx) > Math.Abs(sy)) Curve(a[6], a[7], a[8], a[9], a[10], -sy);
                                else Curve(a[6], a[7], a[8], a[9], -sx, a[10]);
                                break;
                            }
                            // Арифметика и работа со стеком
                            case 9 when a.Count >= 1: a[^1] = Math.Abs(a[^1]); continue;
                            case 10 when a.Count >= 2: a[^2] += a[^1]; a.RemoveAt(a.Count - 1); continue;
                            case 11 when a.Count >= 2: a[^2] -= a[^1]; a.RemoveAt(a.Count - 1); continue;
                            case 12 when a.Count >= 2: a[^2] = a[^1] != 0 ? a[^2] / a[^1] : 0; a.RemoveAt(a.Count - 1); continue;
                            case 14 when a.Count >= 1: a[^1] = -a[^1]; continue;
                            case 18 when a.Count >= 1: a.RemoveAt(a.Count - 1); continue;
                            case 24 when a.Count >= 2: a[^2] *= a[^1]; a.RemoveAt(a.Count - 1); continue;
                            case 26 when a.Count >= 1: a[^1] = Math.Sqrt(Math.Max(0, a[^1])); continue;
                            case 27 when a.Count >= 1: a.Add(a[^1]); continue;
                            case 28 when a.Count >= 2: (a[^1], a[^2]) = (a[^2], a[^1]); continue;
                            case 20 when a.Count >= 2: { int i = (int)a[^1]; if (i is >= 0 and < 32) trans[i] = a[^2]; a.RemoveRange(a.Count - 2, 2); continue; }
                            case 21 when a.Count >= 1: { int i = (int)a[^1]; a[^1] = i is >= 0 and < 32 ? trans[i] : 0; continue; }
                            case 22 when a.Count >= 4: { double v = a[^2] <= a[^1] ? a[^4] : a[^3]; a.RemoveRange(a.Count - 4, 4); a.Add(v); continue; }
                            case 23: a.Add(0.5); continue;
                        }
                        stack.Clear();
                        break;
                    }
                    default:
                        stack.Clear();
                        break;
                }
            }
            depth--;
            return true;
        }

        var (cs, ce) = _charStrings[gid];
        Run(cs, ce);
    }

    // ================================================================
    //  CFF2 → Type 2 (CFF): контур при текущих координатах осей без подпрограмм и подсказок
    // ================================================================

    static void Num(MemoryStream m, int v)
    {
        if (v is >= -107 and <= 107) m.WriteByte((byte)(v + 139));
        else if (v is >= 108 and <= 1131) { v -= 108; m.WriteByte((byte)((v >> 8) + 247)); m.WriteByte((byte)v); }
        else if (v is >= -1131 and <= -108) { v = -v - 108; m.WriteByte((byte)((v >> 8) + 251)); m.WriteByte((byte)v); }
        else { v = Math.Clamp(v, short.MinValue, short.MaxValue); m.WriteByte(28); m.WriteByte((byte)(v >> 8)); m.WriteByte((byte)v); }
    }

    byte[] StaticCharString(int gid)
    {
        var m = new MemoryStream();
        int cx = 0, cy = 0;
        void Pt(double x, double y) { int ix = (int)Math.Round(x), iy = (int)Math.Round(y); Num(m, ix - cx); Num(m, iy - cy); cx = ix; cy = iy; }
        Outline(gid,
            (x, y) => { Pt(x, y); m.WriteByte(21); },
            (x, y) => { Pt(x, y); m.WriteByte(5); },
            (x1, y1, x2, y2, x, y) => { Pt(x1, y1); Pt(x2, y2); Pt(x, y); m.WriteByte(8); });
        m.WriteByte(14);
        return m.ToArray();
    }
}

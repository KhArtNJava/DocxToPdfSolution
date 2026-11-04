using System.Buffers.Binary;
using System.Text;

namespace DocxToPdf.Fonts;

/// <summary>Краткие сведения о шрифте для индекса (без загрузки глифов).</summary>
public sealed record FontFaceInfo(string Path, int Index, List<string> Families, bool Bold, bool Italic, string PostScriptName);

/// <summary>
/// Минимальный парсер TrueType (glyf-контуры) с поддержкой TTC
/// и созданием подмножества шрифта для встраивания в PDF.
/// </summary>
public sealed class TrueTypeFont
{
    readonly byte[] _data;
    readonly Dictionary<string, (int Offset, int Length)> _tables = new();
    readonly Dictionary<int, ushort> _cmap = new();
    ushort[] _advances = [];
    int[] _loca = [];

    public string Family { get; private set; } = "";
    public string PostScriptName { get; private set; } = "Font";
    public int UnitsPerEm { get; private set; } = 1000;
    public int NumGlyphs { get; private set; }
    public short XMin, YMin, XMax, YMax;
    public int Ascent { get; private set; }      // в единицах шрифта
    public int Descent { get; private set; }     // положительное
    public int LineHeight { get; private set; }  // ascent + descent + lineGap
    public int CapHeight { get; private set; }
    public double ItalicAngle { get; private set; }
    public int UnderlinePosition { get; private set; }
    public int UnderlineThickness { get; private set; }
    public bool IsFixedPitch { get; private set; }
    public bool IsBold { get; private set; }
    public bool IsItalic { get; private set; }
    public bool IsSymbol { get; private set; }
    public bool IsSerif { get; private set; }

    TrueTypeFont(byte[] data) => _data = data;

    // ---------------------------------------------------------------
    //  Чтение
    // ---------------------------------------------------------------

    static ushort U16(byte[] d, int o) => BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(o));
    static short S16(byte[] d, int o) => BinaryPrimitives.ReadInt16BigEndian(d.AsSpan(o));
    static uint U32(byte[] d, int o) => BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o));

    static int FontOffset(byte[] data, int index)
    {
        if (data.Length >= 12 && Encoding.ASCII.GetString(data, 0, 4) == "ttcf")
        {
            int n = (int)U32(data, 8);
            if (index < 0 || index >= n) throw new InvalidDataException("Неверный индекс шрифта в TTC");
            return (int)U32(data, 12 + 4 * index);
        }
        return 0;
    }

    static Dictionary<string, (int, int)> ReadDirectory(byte[] data, int offset)
    {
        var t = new Dictionary<string, (int, int)>();
        int num = U16(data, offset + 4);
        for (int i = 0; i < num; i++)
        {
            int rec = offset + 12 + 16 * i;
            string tag = Encoding.ASCII.GetString(data, rec, 4);
            t[tag] = ((int)U32(data, rec + 8), (int)U32(data, rec + 12));
        }
        return t;
    }

    public static TrueTypeFont Load(string path, int index)
    {
        var data = File.ReadAllBytes(path);
        var f = new TrueTypeFont(data);
        f.Parse(FontOffset(data, index));
        return f;
    }

    /// <summary>Быстро читает имена и стиль всех шрифтов файла (TTF/TTC).</summary>
    public static List<FontFaceInfo> ReadInfo(string path)
    {
        var result = new List<FontFaceInfo>();
        byte[] data;
        try { data = File.ReadAllBytes(path); } catch { return result; }
        if (data.Length < 12) return result;
        int count = 1;
        if (Encoding.ASCII.GetString(data, 0, 4) == "ttcf") count = (int)U32(data, 8);
        for (int i = 0; i < count; i++)
        {
            try
            {
                int off = FontOffset(data, i);
                var tables = ReadDirectory(data, off);
                if (!tables.ContainsKey("glyf") || !tables.ContainsKey("cmap") || !tables.ContainsKey("head")) continue;
                var names = ReadNames(data, tables);
                var families = new List<string>();
                void Add(string? s) { if (!string.IsNullOrWhiteSpace(s) && !families.Contains(s, StringComparer.OrdinalIgnoreCase)) families.Add(s.Trim()); }
                foreach (var n in names.Where(n => n.Id == 1)) Add(n.Value);
                string sub = names.FirstOrDefault(n => n.Id == 17).Value ?? names.FirstOrDefault(n => n.Id == 2).Value ?? "";
                if (IsRibbi(sub)) foreach (var n in names.Where(n => n.Id == 16)) Add(n.Value);
                var (bold, italic) = ReadStyle(data, tables);
                string ps = names.FirstOrDefault(n => n.Id == 6).Value ?? families.FirstOrDefault() ?? "Font";
                if (families.Count > 0) result.Add(new FontFaceInfo(path, i, families, bold, italic, ps));
            }
            catch { /* повреждённый шрифт — пропускаем */ }
        }
        return result;
    }

    static bool IsRibbi(string sub)
    {
        var s = sub.ToLowerInvariant().Replace(" ", "");
        return s is "regular" or "normal" or "roman" or "book" or "bold" or "italic" or "oblique" or "bolditalic" or "boldoblique" or "";
    }

    static (bool Bold, bool Italic) ReadStyle(byte[] d, Dictionary<string, (int, int)> tables)
    {
        bool bold = false, italic = false;
        if (tables.TryGetValue("head", out var head))
        {
            int mac = U16(d, head.Item1 + 44);
            bold = (mac & 1) != 0; italic = (mac & 2) != 0;
        }
        if (tables.TryGetValue("OS/2", out var os2) && os2.Item2 >= 64)
        {
            int sel = U16(d, os2.Item1 + 62);
            int weight = U16(d, os2.Item1 + 4);
            bold = (sel & 0x20) != 0 || weight >= 600;
            italic = (sel & 0x01) != 0 || (sel & 0x200) != 0;
        }
        return (bold, italic);
    }

    static List<(int Id, string Value)> ReadNames(byte[] d, Dictionary<string, (int, int)> tables)
    {
        var list = new List<(int, string)>();
        if (!tables.TryGetValue("name", out var t)) return list;
        int o = t.Item1;
        int count = U16(d, o + 2), strOff = o + U16(d, o + 4);
        var preferred = new List<(int, string)>();
        for (int i = 0; i < count; i++)
        {
            int r = o + 6 + 12 * i;
            int platform = U16(d, r), encoding = U16(d, r + 2), lang = U16(d, r + 4), id = U16(d, r + 6);
            int len = U16(d, r + 8), off = U16(d, r + 10);
            if (id is not (1 or 2 or 6 or 16 or 17)) continue;
            if (strOff + off + len > d.Length) continue;
            string? s = null;
            if (platform == 3 || platform == 0) s = Encoding.BigEndianUnicode.GetString(d, strOff + off, len);
            else if (platform == 1 && encoding == 0) s = Encoding.Latin1.GetString(d, strOff + off, len);
            if (s == null) continue;
            s = s.TrimEnd('\0');
            if ((platform == 3 && lang == 0x409) || (platform == 1 && lang == 0)) preferred.Add((id, s));
            else list.Add((id, s));
        }
        preferred.AddRange(list);
        return preferred;
    }

    void Parse(int offset)
    {
        foreach (var (k, v) in ReadDirectory(_data, offset)) _tables[k] = v;
        if (!_tables.ContainsKey("glyf"))
            throw new NotSupportedException("Поддерживаются только TrueType-шрифты с таблицей glyf");

        var head = _tables["head"].Offset;
        UnitsPerEm = U16(_data, head + 18);
        XMin = S16(_data, head + 36); YMin = S16(_data, head + 38);
        XMax = S16(_data, head + 40); YMax = S16(_data, head + 42);
        int indexToLoc = S16(_data, head + 50);

        NumGlyphs = U16(_data, _tables["maxp"].Offset + 4);

        var hhea = _tables["hhea"].Offset;
        int hAsc = S16(_data, hhea + 4), hDesc = S16(_data, hhea + 6), hGap = S16(_data, hhea + 8);
        int numHMetrics = U16(_data, hhea + 34);

        var hmtx = _tables["hmtx"].Offset;
        _advances = new ushort[NumGlyphs];
        ushort last = 0;
        for (int g = 0; g < NumGlyphs; g++)
        {
            if (g < numHMetrics) last = U16(_data, hmtx + 4 * g);
            _advances[g] = last;
        }

        int winAsc = hAsc, winDesc = -hDesc;
        CapHeight = (int)(hAsc * 0.7);
        if (_tables.TryGetValue("OS/2", out var os2) && os2.Length >= 78)
        {
            int o = os2.Offset;
            winAsc = U16(_data, o + 74);
            winDesc = U16(_data, o + 76);
            int version = U16(_data, o);
            if (version >= 2 && os2.Length >= 90) CapHeight = S16(_data, o + 88);
            int family = _data[o + 30]; // panose bFamilyType
            int serif = _data[o + 31];
            IsSerif = family == 2 && serif is >= 2 and <= 10;
        }
        Ascent = Math.Max(winAsc, hAsc);
        int natural = Math.Max(winAsc + winDesc, hAsc - hDesc + hGap);
        if (natural <= 0) natural = UnitsPerEm * 115 / 100;
        LineHeight = natural;
        Descent = Math.Max(0, natural - Ascent);
        if (Descent == 0) Descent = Math.Max(1, -hDesc);

        if (_tables.TryGetValue("post", out var post))
        {
            ItalicAngle = S16(_data, post.Offset + 4) + U16(_data, post.Offset + 6) / 65536.0;
            UnderlinePosition = S16(_data, post.Offset + 8);
            UnderlineThickness = S16(_data, post.Offset + 10);
            IsFixedPitch = U32(_data, post.Offset + 12) != 0;
        }
        if (UnderlineThickness <= 0) UnderlineThickness = UnitsPerEm / 20;
        if (UnderlinePosition == 0) UnderlinePosition = -UnitsPerEm / 10;

        var loca = _tables["loca"].Offset;
        _loca = new int[NumGlyphs + 1];
        for (int g = 0; g <= NumGlyphs; g++)
            _loca[g] = indexToLoc == 0 ? U16(_data, loca + 2 * g) * 2 : (int)U32(_data, loca + 4 * g);

        var names = ReadNames(_data, _tables.ToDictionary(k => k.Key, v => (v.Value.Offset, v.Value.Length)));
        Family = names.FirstOrDefault(n => n.Id == 1).Value ?? "Font";
        string ps = names.FirstOrDefault(n => n.Id == 6).Value ?? Family;
        PostScriptName = new string(ps.Where(c => c > 32 && c < 127 && "[](){}<>/%#".IndexOf(c) < 0).ToArray());
        if (PostScriptName.Length == 0) PostScriptName = "Font";
        (IsBold, IsItalic) = ReadStyle(_data, _tables.ToDictionary(k => k.Key, v => (v.Value.Offset, v.Value.Length)));

        ParseCmap();
    }

    void ParseCmap()
    {
        int cmap = _tables["cmap"].Offset;
        int num = U16(_data, cmap + 2);
        int best = -1, bestScore = -1;
        for (int i = 0; i < num; i++)
        {
            int rec = cmap + 4 + 8 * i;
            int pid = U16(_data, rec), eid = U16(_data, rec + 2);
            int off = cmap + (int)U32(_data, rec + 4);
            int format = U16(_data, off);
            int score = (pid, eid, format) switch
            {
                (3, 10, 12) => 5,
                (0, _, 12) => 4,
                (3, 1, 4) => 3,
                (0, _, 4) => 2,
                (3, 0, 4) => 1,
                _ => -1
            };
            if (score > bestScore) { bestScore = score; best = off; IsSymbol = pid == 3 && eid == 0; }
        }
        if (best < 0) return;
        int fmt = U16(_data, best);
        if (fmt == 4)
        {
            int segX2 = U16(_data, best + 6);
            int ends = best + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ranges = deltas + segX2;
            for (int s = 0; s < segX2 / 2; s++)
            {
                int end = U16(_data, ends + 2 * s), start = U16(_data, starts + 2 * s);
                short delta = S16(_data, deltas + 2 * s);
                int rangeOff = U16(_data, ranges + 2 * s);
                if (start == 0xFFFF) continue;
                for (int c = start; c <= end; c++)
                {
                    int g;
                    if (rangeOff == 0) g = (c + delta) & 0xFFFF;
                    else
                    {
                        int addr = ranges + 2 * s + rangeOff + 2 * (c - start);
                        if (addr + 1 >= _data.Length) continue;
                        g = U16(_data, addr);
                        if (g != 0) g = (g + delta) & 0xFFFF;
                    }
                    if (g != 0 && g < NumGlyphs) _cmap.TryAdd(c, (ushort)g);
                }
            }
        }
        else if (fmt == 12)
        {
            int groups = (int)U32(_data, best + 12);
            for (int i = 0; i < groups; i++)
            {
                int o = best + 16 + 12 * i;
                uint start = U32(_data, o), end = U32(_data, o + 4), gid = U32(_data, o + 8);
                if (end - start > 0x10FFFF) continue;
                for (uint c = start; c <= end; c++, gid++)
                    if (gid < NumGlyphs) _cmap.TryAdd((int)c, (ushort)gid);
            }
        }
    }

    // ---------------------------------------------------------------
    //  Метрики
    // ---------------------------------------------------------------

    public bool TryGetGlyph(int codepoint, out ushort gid)
    {
        if (_cmap.TryGetValue(codepoint, out gid)) return true;
        if (IsSymbol)
        {
            if (codepoint < 0x100 && _cmap.TryGetValue(codepoint + 0xF000, out gid)) return true;
            if (codepoint is >= 0xF000 and <= 0xF0FF && _cmap.TryGetValue(codepoint - 0xF000, out gid)) return true;
        }
        gid = 0;
        return false;
    }

    /// <summary>Ширина глифа в долях em.</summary>
    public double Advance(ushort gid) => gid < _advances.Length ? _advances[gid] / (double)UnitsPerEm : 0;

    // ---------------------------------------------------------------
    //  Подмножество
    // ---------------------------------------------------------------

    /// <summary>
    /// Создаёт TrueType-файл, содержащий только нужные глифы. Номера глифов сохраняются
    /// (неиспользуемые глифы становятся пустыми), поэтому в PDF можно использовать CIDToGIDMap /Identity.
    /// </summary>
    public byte[] Subset(IEnumerable<ushort> used)
    {
        var keep = new HashSet<int> { 0 };
        var stack = new Stack<int>(used.Select(u => (int)u));
        while (stack.Count > 0)
        {
            int g = stack.Pop();
            if (g >= NumGlyphs || !keep.Add(g)) continue;
            foreach (var c in CompositeComponents(g)) stack.Push(c);
        }

        int glyfOff = _tables["glyf"].Offset;
        using var glyf = new MemoryStream();
        var newLoca = new int[NumGlyphs + 1];
        for (int g = 0; g < NumGlyphs; g++)
        {
            newLoca[g] = (int)glyf.Length;
            if (!keep.Contains(g)) continue;
            int len = _loca[g + 1] - _loca[g];
            if (len <= 0) continue;
            glyf.Write(_data, glyfOff + _loca[g], len);
            while (glyf.Length % 4 != 0) glyf.WriteByte(0);
        }
        newLoca[NumGlyphs] = (int)glyf.Length;

        var loca = new byte[(NumGlyphs + 1) * 4];
        for (int g = 0; g <= NumGlyphs; g++) BinaryPrimitives.WriteInt32BigEndian(loca.AsSpan(4 * g), newLoca[g]);

        var tables = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
        var headT = Table("head");
        BinaryPrimitives.WriteUInt32BigEndian(headT.AsSpan(8), 0);
        BinaryPrimitives.WriteInt16BigEndian(headT.AsSpan(50), 1);
        tables["head"] = headT;
        tables["hhea"] = Table("hhea");
        tables["maxp"] = Table("maxp");
        tables["loca"] = loca;
        tables["glyf"] = glyf.ToArray();

        // hmtx: обнуляем неиспользуемые метрики (лучше сжимается)
        var hmtx = Table("hmtx");
        int numH = U16(tables["hhea"], 34);
        for (int g = 0; g < numH && 4 * g + 3 < hmtx.Length; g++)
            if (!keep.Contains(g)) { hmtx[4 * g] = hmtx[4 * g + 1] = hmtx[4 * g + 2] = hmtx[4 * g + 3] = 0; }
        tables["hmtx"] = hmtx;
        foreach (var t in new[] { "cvt ", "fpgm", "prep", "OS/2" })
            if (_tables.ContainsKey(t)) tables[t] = Table(t);

        return BuildFont(tables);
    }

    byte[] Table(string tag)
    {
        var (o, l) = _tables[tag];
        l = Math.Min(l, _data.Length - o);
        return _data.AsSpan(o, l).ToArray();
    }

    IEnumerable<int> CompositeComponents(int g)
    {
        if (g >= NumGlyphs) yield break;
        int len = _loca[g + 1] - _loca[g];
        if (len < 10) yield break;
        int p = _tables["glyf"].Offset + _loca[g];
        if (S16(_data, p) >= 0) yield break;
        p += 10;
        while (true)
        {
            int flags = U16(_data, p);
            int comp = U16(_data, p + 2);
            yield return comp;
            p += 4;
            p += (flags & 0x0001) != 0 ? 4 : 2;
            if ((flags & 0x0008) != 0) p += 2;
            else if ((flags & 0x0040) != 0) p += 4;
            else if ((flags & 0x0080) != 0) p += 8;
            if ((flags & 0x0020) == 0) break;
        }
    }

    static uint Checksum(byte[] d)
    {
        uint sum = 0;
        int n = (d.Length + 3) / 4;
        for (int i = 0; i < n; i++)
        {
            uint v = 0;
            for (int k = 0; k < 4; k++)
            {
                int idx = 4 * i + k;
                v = (v << 8) | (idx < d.Length ? d[idx] : (byte)0);
            }
            sum += v;
        }
        return sum;
    }

    static byte[] BuildFont(SortedDictionary<string, byte[]> tables)
    {
        int n = tables.Count;
        int entrySelector = (int)Math.Floor(Math.Log2(n));
        int searchRange = (1 << entrySelector) * 16;
        using var ms = new MemoryStream();
        var hdr = new byte[12 + 16 * n];
        BinaryPrimitives.WriteUInt32BigEndian(hdr, 0x00010000);
        BinaryPrimitives.WriteUInt16BigEndian(hdr.AsSpan(4), (ushort)n);
        BinaryPrimitives.WriteUInt16BigEndian(hdr.AsSpan(6), (ushort)searchRange);
        BinaryPrimitives.WriteUInt16BigEndian(hdr.AsSpan(8), (ushort)entrySelector);
        BinaryPrimitives.WriteUInt16BigEndian(hdr.AsSpan(10), (ushort)(n * 16 - searchRange));
        int offset = hdr.Length;
        int i = 0;
        foreach (var (tag, data) in tables)
        {
            int rec = 12 + 16 * i++;
            Encoding.ASCII.GetBytes(tag).CopyTo(hdr, rec);
            BinaryPrimitives.WriteUInt32BigEndian(hdr.AsSpan(rec + 4), Checksum(data));
            BinaryPrimitives.WriteUInt32BigEndian(hdr.AsSpan(rec + 8), (uint)offset);
            BinaryPrimitives.WriteUInt32BigEndian(hdr.AsSpan(rec + 12), (uint)data.Length);
            offset += (data.Length + 3) & ~3;
        }
        ms.Write(hdr);
        int headPos = -1;
        foreach (var (tag, data) in tables)
        {
            if (tag == "head") headPos = (int)ms.Length;
            ms.Write(data);
            while (ms.Length % 4 != 0) ms.WriteByte(0);
        }
        var result = ms.ToArray();
        if (headPos >= 0)
        {
            uint adj = 0xB1B0AFBA - Checksum(result);
            BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(headPos + 8), adj);
        }
        return result;
    }
}

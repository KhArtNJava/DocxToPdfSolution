using System.Buffers.Binary;
using System.Text;

namespace DocxToPdf.Fonts;

/// <summary>Глиф эмодзи для вывода: шрифт, номер глифа и текст кластера для копирования.</summary>
public sealed record EmojiGlyph(ColorEmojiFont Font, ushort Gid, string Text);

/// <summary>
/// Цветной эмодзи-шрифт (CBDT/CBLC): PNG-растры глифов, cmap и GSUB-лигатуры
/// (ZWJ-последовательности, флаги, тона кожи, кейкапы). Контуры в таких шрифтах
/// отсутствуют, поэтому глифы выводятся растром, а текстовый слой — шрифтом Type 3.
/// </summary>
public sealed class ColorEmojiFont
{
    readonly byte[] _d;
    readonly Dictionary<string, (int Off, int Len)> _t;
    readonly List<(int First, int Last, int Hdr)> _index = [];
    Dictionary<int, ushort>? _cmap;
    Dictionary<ushort, List<(ushort[] Seq, ushort Lig)>>? _ligs;
    int _numH;

    public readonly string FilePath = "";
    public readonly List<string> Families = [];
    public int UnitsPerEm = 2048;
    /// <summary>Пикселей на em у растра strike (у Noto Color Emoji — 109).</summary>
    public int PpemY = 109;
    /// <summary>Аскент/дескент в долях em (положительные).</summary>
    public double AscentEm = 1900.0 / 2048, DescentEm = 500.0 / 2048;

    static readonly Dictionary<string, ColorEmojiFont?> _cache = new(StringComparer.OrdinalIgnoreCase);

    ColorEmojiFont(byte[] data, Dictionary<string, (int, int)> tables, string path)
    {
        _d = data;
        _t = tables;
        FilePath = path;
    }

    ushort U16(int o) => BinaryPrimitives.ReadUInt16BigEndian(_d.AsSpan(o));

    uint U32(int o) => BinaryPrimitives.ReadUInt32BigEndian(_d.AsSpan(o));

    // ================================================================
    //  Загрузка
    // ================================================================

    /// <summary>Быстрая проверка: есть ли в файле таблица CBLC (цветной битмап-шрифт).</summary>
    public static bool HasCbdtTable(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            Span<byte> hdr = stackalloc byte[12];
            if (fs.Read(hdr) < 12) return false;
            if (Encoding.ASCII.GetString(hdr[..4]) == "ttcf") return false;
            int n = BinaryPrimitives.ReadUInt16BigEndian(hdr[4..]);
            if (n <= 0 || 12 + 16L * n > fs.Length) return false;
            Span<byte> dir = stackalloc byte[16 * n];
            fs.ReadExactly(dir);
            for (int i = 0; i < n; i++)
                if (Encoding.ASCII.GetString(dir.Slice(16 * i, 4)) == "CBLC")
                    return true;
            return false;
        }
        catch { return false; }
    }

    /// <summary>Разбирает шрифт; null — файл не является CBDT-шрифтом. Результат кэшируется.</summary>
    public static ColorEmojiFont? TryLoad(string path)
    {
        if (_cache.TryGetValue(path, out var cached)) return cached;
        ColorEmojiFont? f = null;
        try
        {
            byte[] d = File.ReadAllBytes(path);
            var t = ReadTables(d);
            if (t != null && t.ContainsKey("CBLC") && t.ContainsKey("CBDT") && t.ContainsKey("cmap") && t.ContainsKey("head"))
            {
                var font = new ColorEmojiFont(d, t, path);
                font.ParseHead();
                font.ParseHhea();
                font.ReadName();
                font.ParseIndex();
                font.ParseCmap();
                if (font._index.Count > 0 && font._cmap is { Count: >100 }) f = font;
            }
        }
        catch { f = null; }
        _cache[path] = f;
        return f;
    }

    static Dictionary<string, (int, int)>? ReadTables(byte[] d)
    {
        if (d.Length < 12) return null;
        if (Encoding.ASCII.GetString(d, 0, 4) == "ttcf") return null; // коллекции с CBDT не встречаются
        uint sc = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(0));
        if (sc != 0x00010000) return null;
        int n = BinaryPrimitives.ReadUInt16BigEndian(d.AsSpan(4));
        if (n <= 0 || 12 + 16L * n > d.Length) return null;
        var t = new Dictionary<string, (int, int)>(12, StringComparer.Ordinal);
        for (int i = 0; i < n; i++)
        {
            int o = 12 + 16 * i;
            string tag = Encoding.ASCII.GetString(d, o, 4);
            uint off = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o + 8));
            uint len = BinaryPrimitives.ReadUInt32BigEndian(d.AsSpan(o + 12));
            if (off >= d.Length || len > d.Length - off) continue;
            t[tag] = ((int)off, (int)len);
        }
        return t;
    }

    void ParseHead()
    {
        int o = _t["head"].Off;
        UnitsPerEm = U16(o + 18);
        if (UnitsPerEm <= 0) UnitsPerEm = 2048;
    }

    void ParseHhea()
    {
        int o = _t["hhea"].Off;
        short asc = (short)U16(o + 4), desc = (short)U16(o + 6);
        AscentEm = asc / (double)UnitsPerEm;
        DescentEm = -desc / (double)UnitsPerEm;
        _numH = U16(o + 34);
    }

    void ReadName()
    {
        if (!_t.TryGetValue("name", out var nt)) return;
        int no = nt.Off;
        if (no + 6 > _d.Length) return;
        int count = U16(no + 2);
        for (int i = 0; i < count; i++)
        {
            int rec = no + 6 + 12 * i;
            if (rec + 12 > _d.Length) break;
            int plat = U16(rec);
            int nid = U16(rec + 6);
            int len = U16(rec + 8), off = U16(rec + 10);
            if (nid != 1) continue;
            string s;
            try { s = plat == 3 || plat == 0 ? Encoding.BigEndianUnicode.GetString(_d, no + off, len) : Encoding.Latin1.GetString(_d, no + off, len); }
            catch { continue; }
            if (!string.IsNullOrWhiteSpace(s) && !Families.Contains(s, StringComparer.OrdinalIgnoreCase)) Families.Add(s.Trim());
        }
    }

    // ================================================================
    //  CBLC/CBDT: индекс растров
    // ================================================================

    void ParseIndex()
    {
        int bo = _t["CBLC"].Off;
        uint numSizes = U32(bo + 4);
        if (numSizes == 0 || numSizes > 64) return;
        // CBLC: заголовок (8 байт) + bitmapSizeTable[numOfSizes] (по 48 байт);
        // выбирается strike с максимальным ppem (у Segoe UI Emoji их несколько)
        int bestBst = -1, bestPpem = 0;
        for (int s = 0; s < (int)numSizes; s++)
        {
            int bst = bo + 8 + 48 * s;
            if (bst + 48 > _d.Length) continue;
            int ppem = _d[bst + 45];
            if (ppem > bestPpem) { bestPpem = ppem; bestBst = bst; }
        }
        if (bestBst < 0) return;
        int arrOff = (int)U32(bestBst);
        int cnt = (int)U32(bestBst + 8);
        PpemY = bestPpem;
        if (PpemY <= 0) PpemY = 109;
        int baseArr = bo + arrOff;
        if (cnt > 4096) cnt = 4096;
        for (int k = 0; k < cnt; k++)
        {
            int rec = baseArr + 8 * k;
            if (rec + 8 > _d.Length) break;
            int first = U16(rec), last = U16(rec + 2);
            int add = (int)U32(rec + 4);
            if (first > last || last > 200000) continue;
            _index.Add((first, last, baseArr + add));
        }
    }

    /// <summary>Запись CBDT глифа: (смещение PNG, длина PNG, ширина, высота, сдвиг от базовой линии в пикселях).</summary>
    (int PngOff, int PngLen, int W, int H, int By)? BitmapRecord(ushort gid)
    {
        foreach (var (first, last, hdr) in _index)
        {
            if (gid < first || gid > last) continue;
            int fmt = U16(hdr);
            int imgOff = (int)U32(hdr + 4);
            int n = last - first + 1;
            uint start = 0, end = 0;
            int arr = hdr + 8;
            switch (fmt)
            {
                case 1:
                case 2:
                case 5:
                    if (arr + 4 * (n + 1) > _d.Length) continue;
                    start = U32(arr + 4 * (gid - first));
                    end = U32(arr + 4 * (gid - first + 1));
                    break;
                case 17:
                case 18:
                {
                    if (arr + 4 > _d.Length) continue;
                    int ng = (int)U32(arr);
                    if (gid - first >= ng) return null;
                    int a2 = arr + 4;
                    if (fmt == 18)
                    {
                        if (a2 + 4 * (ng + 1) > _d.Length) continue;
                        start = U32(a2 + 4 * (gid - first));
                        end = U32(a2 + 4 * (gid - first + 1));
                    }
                    else
                    {
                        if (a2 + 2 * (ng + 1) > _d.Length) continue;
                        start = U16(a2 + 2 * (gid - first));
                        end = U16(a2 + 2 * (gid - first + 1));
                    }
                    break;
                }
                default:
                    continue;
            }
            int rec = _t["CBDT"].Off + imgOff + (int)start;
            int len = (int)(end - start);
            if (len <= 0 || rec + 5 > _d.Length) return null;
            // Формат 17/18: метрики 4 байта + dataFormat + (данные); подпись PNG ищем в начале записи
            int w = _d[rec + 1], h = _d[rec];
            int by = (sbyte)_d[rec + 3];
            int sig = -1;
            for (int q = 4; q < Math.Min(17, len); q++)
                if (_d[rec + q] == 0x89 && _d.AsSpan(rec + q, 4).SequenceEqual([(byte)0x89, (byte)0x50, (byte)0x4E, (byte)0x47])) { sig = q; break; }
            if (sig < 0) return null;
            return (rec + sig, len - sig, w, h, by);
        }
        return null;
    }

    /// <summary>PNG-растр глифа (null — глифа нет).</summary>
    public byte[]? PngOf(ushort gid)
    {
        var r = BitmapRecord(gid);
        return r == null ? null : _d[r.Value.PngOff..(r.Value.PngOff + r.Value.PngLen)];
    }

    /// <summary>Метрики растра глифа: ширина, высота, сдвиг верха растра от базовой линии (пиксели).</summary>
    public (int W, int H, int By)? BitmapMetrics(ushort gid)
    {
        var r = BitmapRecord(gid);
        return r == null ? null : (r.Value.W, r.Value.H, r.Value.By);
    }

    /// <summary>Ширина глифа в em (hmtx).</summary>
    public double AdvanceEm(ushort gid)
    {
        if (!_t.TryGetValue("hmtx", out var hm)) return 1.245;
        int g = _numH > 0 ? Math.Min(gid, _numH - 1) : gid;
        return U16(hm.Off + 4 * g) / (double)UnitsPerEm;
    }

    // ================================================================
    //  cmap
    // ================================================================

    void ParseCmap()
    {
        var map = new Dictionary<int, ushort>();
        int o = _t["cmap"].Off;
        int nTab = U16(o + 2);
        for (int i = 0; i < nTab; i++)
        {
            int rec = o + 4 + 8 * i;
            int subOff = o + (int)U32(rec + 4);
            if (subOff + 2 > _d.Length) continue;
            int fmt = U16(subOff);
            switch (fmt)
            {
                case 4:
                {
                    int segX2 = U16(subOff + 6);
                    if (segX2 % 2 != 0) break;
                    int segs = segX2 / 2;
                    int endsOff = subOff + 14;
                    int startsOff = endsOff + segX2;
                    int deltasOff = startsOff + segX2;
                    int rngOffBase = deltasOff + segX2;
                    if (rngOffBase + segX2 > _d.Length) break;
                    for (int s = 0; s < segs; s++)
                    {
                        int end = U16(endsOff + 2 * s), start = U16(startsOff + 2 * s);
                        if (start == 0xFFFF) continue;
                        short delta = (short)U16(deltasOff + 2 * s);
                        int ro = U16(rngOffBase + 2 * s);
                        for (int cp = start; cp <= Math.Min(end, 0xFFFD); cp++)
                        {
                            ushort g;
                            if (ro == 0) g = (ushort)((cp + delta) & 0xFFFF);
                            else
                            {
                                int gi = rngOffBase + 2 * s + ro + 2 * (cp - start);
                                if (gi + 2 > _d.Length) break;
                                ushort g2 = U16(gi);
                                if (g2 == 0) continue;
                                g = (ushort)((g2 + delta) & 0xFFFF);
                            }
                            if (g != 0) map.TryAdd(cp, g);
                        }
                    }
                    break;
                }
                case 12:
                {
                    int ng = (int)U32(subOff + 12);
                    if (16 + 12L * ng > _d.Length - subOff) break;
                    for (int gi = 0; gi < ng; gi++)
                    {
                        uint s = U32(subOff + 16 + 12 * gi), e = U32(subOff + 20 + 12 * gi), first = U32(subOff + 24 + 12 * gi);
                        if (e - s > 100000) continue;
                        for (uint cp = s; cp <= e && cp <= 0x10FFFF; cp++)
                            map.TryAdd((int)cp, (ushort)(first + cp - s));
                    }
                    break;
                }
            }
        }
        _cmap = map;
    }

    /// <summary>Глиф кода или 0.</summary>
    public ushort GlyphOf(int cp) => _cmap != null && _cmap.TryGetValue(cp, out var g) ? g : (ushort)0;

    /// <summary>Есть ли код в шрифте.</summary>
    public bool HasCode(int cp) => GlyphOf(cp) != 0;

    // ================================================================
    //  GSUB: лигатуры (ZWJ-последовательности, флаги, кейкапы, тона кожи)
    // ================================================================

    void ParseLigatures()
    {
        var byFirst = new Dictionary<ushort, List<(ushort[] Seq, ushort Lig)>>();
        if (_cmap == null || !_t.TryGetValue("GSUB", out var gt)) { _ligs = byFirst; return; }
        int go = gt.Off;
        int lookOff = U16(go + 8);
        int ll = go + lookOff;
        if (ll + 2 > _d.Length) { _ligs = byFirst; return; }
        int lkCnt = U16(ll);
        for (int i = 0; i < lkCnt; i++)
        {
            if (ll + 4 + 2 * i + 2 > _d.Length) break;
            int lk = ll + U16(ll + 2 + 2 * i);
            if (lk + 6 > _d.Length) continue;
            int typ = U16(lk);
            if (typ != 4) continue;
            int subCnt = U16(lk + 4);
            for (int si = 0; si < subCnt; si++)
            {
                if (lk + 8 + 2 * si > _d.Length) break;
                int st = lk + U16(lk + 6 + 2 * si);
                if (st + 6 > _d.Length) continue;
                int fmt = U16(st), covOff = U16(st + 2), setCnt = U16(st + 4);
                if (fmt != 1 || setCnt > 200000) continue;
                int cov = st + covOff;
                if (cov + 4 > _d.Length) continue;
                var covg = new List<ushort>();
                int cf = U16(cov);
                if (cf == 1)
                {
                    int cc = U16(cov + 2);
                    if (cov + 4 + 2 * cc > _d.Length) continue;
                    for (int q = 0; q < cc; q++) covg.Add(U16(cov + 4 + 2 * q));
                }
                else if (cf == 2)
                {
                    int rc = U16(cov + 2);
                    if (cov + 4 + 6 * rc > _d.Length) continue;
                    for (int r = 0; r < rc; r++)
                    {
                        int s3 = U16(cov + 4 + 6 * r), e3 = U16(cov + 6 + 6 * r), s3g = U16(cov + 8 + 6 * r);
                        if (e3 - s3 > 20000) continue;
                        for (int q = 0; q <= e3 - s3; q++) covg.Add((ushort)(s3g + q));
                    }
                }
                else continue;
                for (int ls = 0; ls < setCnt && ls < covg.Count; ls++)
                {
                    if (st + 8 + 2 * ls > _d.Length) break;
                    int lset = st + U16(st + 6 + 2 * ls);
                    if (lset + 2 > _d.Length) continue;
                    int lcnt = U16(lset);
                    for (int li = 0; li < lcnt; li++)
                    {
                        if (lset + 4 + 2 * li > _d.Length) break;
                        int lg = lset + U16(lset + 2 + 2 * li);
                        if (lg + 4 > _d.Length) continue;
                        ushort lig = U16(lg);
                        int compCnt = U16(lg + 2);
                        if (compCnt is < 2 or > 12) continue;
                        var seq = new ushort[compCnt];
                        seq[0] = covg[ls];
                        bool ok = true;
                        for (int q = 1; q < compCnt; q++)
                        {
                            if (lg + 4 + 2 * q > _d.Length) { ok = false; break; }
                            seq[q] = U16(lg + 2 + 2 * q);
                        }
                        if (!ok || seq.Any(g => g == 0)) continue;
                        if (!byFirst.TryGetValue(seq[0], out var list)) byFirst[seq[0]] = list = [];
                        list.Add((seq, lig));
                    }
                }
            }
        }
        _ligs = byFirst;
    }

    /// <summary>Ищет лигатуру для последовательности кодов (FE0F пропускается); результат — глиф или 0.</summary>
    public ushort LigatureOf(ReadOnlySpan<int> cps)
    {
        _ligs ??= ParseLigaturesCached();
        var gids = new List<ushort>(cps.Length);
        foreach (var cp in cps)
        {
            if (cp == 0xFE0F) continue;
            var g = GlyphOf(cp);
            if (g == 0) return 0;
            gids.Add(g);
        }
        if (gids.Count == 0 || !_ligs!.TryGetValue(gids[0], out var list)) return 0;
        var arr = gids.ToArray();
        foreach (var (seq, lig) in list)
            if (seq.Length == arr.Length && seq.AsSpan().SequenceEqual(arr))
                return lig;
        return 0;
    }

    Dictionary<ushort, List<(ushort[] Seq, ushort Lig)>> ParseLigaturesCached()
    {
        ParseLigatures();
        return _ligs!;
    }

    // ================================================================
    //  Кластеризация эмодзи
    // ================================================================

    static bool IsRegional(int cp) => cp is >= 0x1F1E6 and <= 0x1F1FF;
    static bool IsSkin(int cp) => cp is >= 0x1F3FB and <= 0x1F3FF;
    static bool IsTag(int cp) => cp is >= 0xE0020 and <= 0xE007F;
    static bool IsKeycapBase(int cp) => cp is '#' or '*' or >= '0' and <= '9';

    /// <summary>Код может начинать эмодзи-кластер (не буквы/цифры/пробелы из cmap).</summary>
    public static bool IsEmojiStart(int cp) => cp is
        '#' or '*' or (>= '0' and <= '9') or 0xA9 or 0xAE or 0x203C or 0x2049 or 0x2122 or 0x2139 or
        (>= 0x2194 and <= 0x21AA) or (>= 0x231A and <= 0x231B) or 0x2328 or (>= 0x23CF and <= 0x23FA) or 0x24C2 or
        (>= 0x25AA and <= 0x25FE) or (>= 0x2600 and <= 0x27BF) or 0x2934 or 0x2935 or (>= 0x2B00 and <= 0x2BFF) or
        0x3030 or 0x303D or 0x3297 or 0x3299 or (>= 0x1F000 and <= 0x1FBFF);

    static int CodeAt(string t, int i)
    {
        if (i >= t.Length) return -1;
        return char.IsHighSurrogate(t[i]) && i + 1 < t.Length && char.IsLowSurrogate(t[i + 1])
            ? char.ConvertToUtf32(t[i], t[i + 1])
            : t[i];
    }

    static int CodeLen(int cp) => cp > 0xFFFF ? 2 : 1;

    /// <summary>
    /// Разбирает эмодзи-кластер с позиции start: ZWJ-последовательность, флаг (пара
    /// regional indicators), кейкап, тэг-последовательность или одиночный эмодзи.
    /// Возвращает список глифов (каждый со своим текстом) и длину в единицах UTF-16,
    /// либо null, если здесь не начинается эмодзи.
    /// </summary>
    public (List<EmojiGlyph> Glyphs, int Len)? ClusterAt(string text, int start)
    {
        int cp = CodeAt(text, start);
        if (cp < 0) return null;
        List<EmojiGlyph> Result(ushort gid, string txt) => [new EmojiGlyph(this, gid, txt)];

        // Кейкап: цифра/#/* + [FE0F] + 20E3
        if (IsKeycapBase(cp))
        {
            int j = start + CodeLen(cp);
            int c1 = CodeAt(text, j);
            if (c1 == 0xFE0F) { j += 1; c1 = CodeAt(text, j); }
            if (c1 == 0x20E3 && GlyphOf(cp) != 0)
            {
                var lig = LigatureOf([cp, 0xFE0F, 0x20E3]);
                if (lig == 0) lig = LigatureOf([cp, 0x20E3]);
                var txt = text.Substring(start, j + CodeLen(c1) - start);
                if (lig != 0) return (Result(lig, txt), j + CodeLen(c1) - start);
            }
            return null;
        }

        // Флаг: пара regional indicators
        if (IsRegional(cp))
        {
            int c1 = CodeAt(text, start + CodeLen(cp));
            if (IsRegional(c1))
            {
                var lig = LigatureOf([cp, c1]);
                int len = CodeLen(cp) + CodeLen(c1);
                var txt = text.Substring(start, len);
                if (lig != 0) return (Result(lig, txt), len);
                // Пары нет: каждый RI сам по себе
                var list = new List<EmojiGlyph>();
                if (GlyphOf(cp) != 0) list.Add(new(this, GlyphOf(cp), char.ConvertFromUtf32(cp)));
                if (GlyphOf(c1) != 0) list.Add(new(this, GlyphOf(c1), char.ConvertFromUtf32(c1)));
                return list.Count > 0 ? (list, len) : null;
            }
            if (GlyphOf(cp) != 0)
                return (Result(GlyphOf(cp), char.ConvertFromUtf32(cp)), CodeLen(cp));
            return null;
        }

        if (!HasCode(cp)) return null;

        // Тэг-последовательность: 1F3F4 + тэги + E007F
        if (cp == 0x1F3F4)
        {
            int j = start + CodeLen(cp);
            var seq = new List<int> { cp };
            while (IsTag(CodeAt(text, j)))
            {
                seq.Add(CodeAt(text, j));
                j += CodeLen(CodeAt(text, j));
            }
            if (seq.Count > 1)
            {
                var lig = LigatureOf(seq.ToArray());
                var txt = text.Substring(start, j - start);
                if (lig != 0) return (Result(lig, txt), j - start);
                return (Result(GlyphOf(cp), char.ConvertFromUtf32(cp)), CodeLen(cp));
            }
        }

        // База + жадное расширение: FE0F, ZWJ+код, тон кожи
        var parts = new List<int> { cp };
        int p = start + CodeLen(cp);
        bool extended;
        do
        {
            extended = false;
            int c1 = CodeAt(text, p);
            if (c1 == 0xFE0F || IsSkin(c1))
            {
                parts.Add(c1);
                p += CodeLen(c1);
                extended = true;
            }
            else if (c1 == 0x200D)
            {
                int c2 = CodeAt(text, p + 1);
                if (c2 > 0 && (HasCode(c2) || c2 == 0xFE0F || IsTag(c2)))
                {
                    parts.Add(0x200D);
                    parts.Add(c2);
                    p += 1 + CodeLen(c2);
                    extended = true;
                    // за тэгами может следовать E007F и другие тэги
                    while (IsTag(CodeAt(text, p)))
                    {
                        parts.Add(CodeAt(text, p));
                        p += CodeLen(CodeAt(text, p));
                        extended = true;
                    }
                }
            }
        } while (extended);

        // Самая длинная лигатура
        for (int len = parts.Count; len >= 1; len--)
        {
            var lig = LigatureOf(parts.Take(len).ToArray());
            if (lig != 0)
            {
                var txt = text.Substring(start, p - start);
                if (len == parts.Count) return (Result(lig, txt), p - start);
                // Частичное совпадение: остальное разберётся следующими вызовами — вернём лигатуру с её текстом
                int cut = start;
                for (int q = 0; q < len; q++) cut += CodeLen(parts[q]);
                return (Result(lig, text.Substring(start, cut - start)), cut - start);
            }
        }
        // Без лигатуры: базовый глиф; FE0F и тон кожи поглощаются (на вывод не влияют)
        return (Result(GlyphOf(cp), text.Substring(start, p - start)), p - start);
    }

    /// <summary>Код имеет эмодзи-презентацию по умолчанию (Unicode emoji-data).</summary>
    public static bool IsEmojiPresentation(int cp) => cp is
        (>= 0x231A and <= 0x231B) or (>= 0x23E9 and <= 0x23EC) or 0x23F0 or 0x23F3 or (>= 0x25FD and <= 0x25FE) or
        (>= 0x2614 and <= 0x2615) or (>= 0x2648 and <= 0x2653) or 0x267F or 0x2693 or 0x26A1 or (>= 0x26AA and <= 0x26AB) or
        (>= 0x26BD and <= 0x26BE) or (>= 0x26C4 and <= 0x26C5) or 0x26CE or 0x26D4 or 0x26EA or (>= 0x26F2 and <= 0x26F3) or
        0x26F5 or 0x26FA or 0x26FD or 0x2705 or (>= 0x270A and <= 0x270B) or 0x2728 or 0x274C or 0x274E or (>= 0x2753 and <= 0x2755) or
        0x2757 or (>= 0x2795 and <= 0x2797) or 0x27B0 or 0x27BF or (>= 0x2B1B and <= 0x2B1C) or 0x2B50 or 0x2B55 or
        0x1F004 or 0x1F0CF or 0x1F18E or (>= 0x1F191 and <= 0x1F19A) or (>= 0x1F1E6 and <= 0x1F1FF) or
        (>= 0x1F201 and <= 0x1F202) or 0x1F21A or 0x1F22F or (>= 0x1F232 and <= 0x1F23A) or (>= 0x1F250 and <= 0x1F251) or
        (>= 0x1F300 and <= 0x1F320) or (>= 0x1F32D and <= 0x1F335) or (>= 0x1F337 and <= 0x1F37C) or (>= 0x1F37E and <= 0x1F393) or
        (>= 0x1F3A0 and <= 0x1F3CA) or (>= 0x1F3CF and <= 0x1F3D3) or (>= 0x1F3E0 and <= 0x1F3F0) or 0x1F3F4 or
        (>= 0x1F3F8 and <= 0x1F43E) or 0x1F440 or (>= 0x1F442 and <= 0x1F4FC) or (>= 0x1F4FF and <= 0x1F53D) or
        (>= 0x1F54B and <= 0x1F54E) or (>= 0x1F550 and <= 0x1F567) or 0x1F57A or (>= 0x1F595 and <= 0x1F596) or 0x1F5A4 or
        (>= 0x1F5FB and <= 0x1F64F) or (>= 0x1F680 and <= 0x1F6C5) or 0x1F6CC or (>= 0x1F6D0 and <= 0x1F6D2) or
        (>= 0x1F6D5 and <= 0x1F6D7) or (>= 0x1F6DC and <= 0x1F6DF) or (>= 0x1F6EB and <= 0x1F6EC) or (>= 0x1F6F4 and <= 0x1F6FC) or
        (>= 0x1F7E0 and <= 0x1F7EB) or 0x1F7F0 or (>= 0x1F90C and <= 0x1F93A) or (>= 0x1F93C and <= 0x1F945) or
        (>= 0x1F947 and <= 0x1F9FF) or (>= 0x1FA70 and <= 0x1FA7C) or (>= 0x1FA80 and <= 0x1FA88) or (>= 0x1FA90 and <= 0x1FABD) or
        (>= 0x1FABF and <= 0x1FAC5) or (>= 0x1FACE and <= 0x1FADB) or (>= 0x1FAE0 and <= 0x1FAE8) or (>= 0x1FAF0 and <= 0x1FAF8);
}

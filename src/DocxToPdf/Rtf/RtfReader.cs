using System.Globalization;
using System.Text;
using DocxToPdf.Docx;

namespace DocxToPdf.Rtf;

/// <summary>
/// Чтение документов RTF (Rich Text Format 1.x: Word, WordPad, LibreOffice и др.) в общую модель документа:
/// таблицы шрифтов, цветов и стилей, кодовые страницы и Юникод, оформление символов и абзацев, табуляция, рамки и заливка,
/// списки (\listtext/\pntext и таблица списков), таблицы (объединения, вложенные), разделы, колонтитулы, сноски,
/// поля, закладки, исправления, картинки (\pict) и фигуры (\shp).
/// </summary>
public sealed partial class RtfReader
{
    static RtfReader() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    // ================================================================
    //  Токены и дерево групп
    // ================================================================

    internal abstract class Node { }
    internal sealed class Group : Node { public readonly List<Node> Items = []; }
    internal sealed class Word(string name, int? param) : Node { public readonly string Name = name; public readonly int? Param = param; }
    internal sealed class Sym(char c) : Node { public readonly char C = c; }
    internal sealed class Bytes(byte[] data, bool hex) : Node { public readonly byte[] Data = data; public readonly bool Hex = hex; }
    internal sealed class Bin(byte[] data) : Node { public readonly byte[] Data = data; }

    /// <summary>Разбор потока RTF в дерево групп (с \binN — двоичными данными).</summary>
    static Group Parse(byte[] d)
    {
        var root = new Group();
        var stack = new Stack<Group>();
        var cur = root;
        int i = 0, n = d.Length;
        var text = new List<byte>();
        void FlushText() { if (text.Count > 0) { cur.Items.Add(new Bytes(text.ToArray(), false)); text.Clear(); } }
        while (i < n)
        {
            byte b = d[i];
            if (b == '{') { FlushText(); var g = new Group(); cur.Items.Add(g); stack.Push(cur); cur = g; i++; continue; }
            if (b == '}') { FlushText(); if (stack.Count > 0) cur = stack.Pop(); i++; continue; }
            if (b == '\r' || b == '\n') { i++; continue; }
            if (b != '\\') { text.Add(b); i++; continue; }
            // Управляющее слово или символ
            FlushText();
            i++;
            if (i >= n) break;
            byte c = d[i];
            if (c is (byte)'\\' or (byte)'{' or (byte)'}') { text.Add(c); i++; continue; }
            if (c == '\'')
            {
                if (i + 2 < n + 0 && Hex(d[i + 1]) >= 0 && i + 2 < n && Hex(d[i + 2]) >= 0)
                {
                    cur.Items.Add(new Bytes([(byte)(Hex(d[i + 1]) << 4 | Hex(d[i + 2]))], true));
                    i += 3;
                }
                else i++;
                continue;
            }
            if (c == '\r' || c == '\n') { cur.Items.Add(new Word("par", null)); i++; continue; }   // «\» + перевод строки = \par
            if (!IsAsciiLetter(c)) { cur.Items.Add(new Sym((char)c)); i++; continue; }
            int s = i;
            while (i < n && IsAsciiLetter(d[i]) && i - s < 32) i++;
            string name = Encoding.ASCII.GetString(d, s, i - s);
            int? param = null;
            if (i < n && (d[i] == '-' || d[i] >= '0' && d[i] <= '9'))
            {
                int ps = i;
                if (d[i] == '-') i++;
                while (i < n && d[i] >= '0' && d[i] <= '9' && i - ps < 11) i++;
                if (long.TryParse(Encoding.ASCII.GetString(d, ps, i - ps), out var lv)) param = (int)Math.Clamp(lv, int.MinValue, int.MaxValue);
            }
            if (i < n && d[i] == ' ') i++;   // пробел-разделитель
            if (name == "bin" && param is > 0)
            {
                int len = Math.Min(param.Value, n - i);
                cur.Items.Add(new Bin(d.AsSpan(i, len).ToArray()));
                i += len;
                continue;
            }
            cur.Items.Add(new Word(name, param));
        }
        FlushText();
        return root;
    }

    static bool IsAsciiLetter(byte c) => c is >= (byte)'a' and <= (byte)'z' or >= (byte)'A' and <= (byte)'Z';
    static int Hex(byte c) => c switch { >= (byte)'0' and <= (byte)'9' => c - '0', >= (byte)'a' and <= (byte)'f' => c - 'a' + 10, >= (byte)'A' and <= (byte)'F' => c - 'A' + 10, _ => -1 };

    /// <summary>Похож ли файл на RTF.</summary>
    public static bool IsRtf(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var buf = new byte[8];
            int k = fs.Read(buf, 0, buf.Length);
            int o = k >= 3 && buf[0] == 0xEF && buf[1] == 0xBB && buf[2] == 0xBF ? 3 : 0;
            while (o < k && buf[o] is (byte)' ' or (byte)'\r' or (byte)'\n' or (byte)'\t') o++;
            return k - o >= 5 && buf[o] == '{' && buf[o + 1] == '\\' && buf[o + 2] == 'r' && buf[o + 3] == 't' && buf[o + 4] == 'f';
        }
        catch { return false; }
    }

    public static DocxDocument Read(string path, DocxOptions? options = null)
    {
        var data = File.ReadAllBytes(path);
        var root = Parse(data);
        var top = root.Items.OfType<Group>().FirstOrDefault() ?? throw new InvalidDataException("Файл не является документом RTF");
        var r = new RtfReader(options ?? new DocxOptions(), Path.GetDirectoryName(Path.GetFullPath(path)));
        return r.ReadAll(top);
    }

    // ================================================================
    //  Состояние
    // ================================================================

    internal sealed class CharProps
    {
        public int Font = -1, Size = 24, Color, Back, Highlight, UlColor;
        public bool Bold, Italic, Strike, DStrike, Caps, SmallCaps, Hidden, Outline, Shadow, Emboss, Imprint;
        public string? Underline;
        public int Vert;                              // 1 — верхний индекс, −1 — нижний
        public int Up;                                // \up/\dn, полупункты
        public int Expand;                            // \expndtw, твипы
        public int Scale = 100;
        public bool Revised, Deleted;
        public int RevAuthor;
        public int Codepage;                          // \cpg у шрифта (0 — по документу)
        public int Kerning;                           // \kerning, полупункты (0 — нет)
        public int Lang;                              // \lang (LCID)
        public CharProps Clone() => (CharProps)MemberwiseClone();
    }

    internal sealed class Tab { public int Pos; public string Align = "left", Leader = "none"; }

    internal sealed class ParaProps
    {
        public string Align = "left";
        public int Li, Ri, Fi, Sb, Sa, Sl, SlMult;
        public bool KeepN, Keep, PageBB, Contextual, Rtl, InTable;
        public bool WidCtl = true;
        public int Itap;
        public int Ls = -1, Ilvl;
        public int Style;
        public List<Tab> Tabs = [];
        public string? PendingTabAlign, PendingLeader;
        public BorderDef? BTop, BBottom, BLeft, BRight, BBetween;
        public int Shading = -1, ShadePct, ShadeFore;
        public int OutlineLevel = -1;
        public int SbAuto, SaAuto;
        public bool NoHyph;
        public ParaProps Clone()
        {
            var p = (ParaProps)MemberwiseClone();
            p.Tabs = [.. Tabs.Select(t => new Tab { Pos = t.Pos, Align = t.Align, Leader = t.Leader })];
            p.BTop = Copy(BTop); p.BBottom = Copy(BBottom); p.BLeft = Copy(BLeft); p.BRight = Copy(BRight); p.BBetween = Copy(BBetween);
            return p;
        }
        static BorderDef? Copy(BorderDef? b) => b == null ? null : new BorderDef { Val = b.Val, Width = b.Width, Space = b.Space, Color = b.Color };
    }

    sealed class State
    {
        public CharProps C = new();
        public ParaProps P = new();
        public int Uc = 1;
        public State Clone() => new() { C = C.Clone(), P = P.Clone(), Uc = Uc };
    }

    // ================================================================
    //  Таблицы документа
    // ================================================================

    readonly DocxOptions _options;
    readonly string? _baseDir;
    readonly DocxDocument _doc = new();
    readonly Dictionary<int, (string Name, int Charset, int Cpg)> _fonts = [];
    readonly List<Rgb?> _colors = [];
    readonly Dictionary<int, List<Node>> _paraStyles = [], _charStyles = [];
    readonly Dictionary<int, int> _styleBase = [];
    readonly Dictionary<int, string> _styleNames = [];
    readonly List<string> _revAuthors = [];
    int _docCodepage = 1252, _defFont, _defLang;
    bool _hyphAuto, _hyphNoCaps;
    double _hyphZone = 18;
    int _hyphConsec;

    /// <summary>Код языка Windows (LCID) → BCP 47 для словаря переносов.</summary>
    static string? LcidName(int lcid) => lcid switch
    {
        1049 => "ru-RU", 1058 => "uk-UA", 1059 => "be-BY", 1026 => "bg-BG", 1033 => "en-US", 2057 => "en-GB", 3081 => "en-AU",
        4105 => "en-CA", 1031 => "de-DE", 3079 => "de-AT", 2055 => "de-CH", 1036 => "fr-FR", 2060 => "fr-BE", 3084 => "fr-CA",
        1034 or 3082 => "es-ES", 2058 => "es-MX", 1040 => "it-IT", 1045 => "pl-PL", 2070 => "pt-PT", 1046 => "pt-BR",
        1043 => "nl-NL", 1029 => "cs-CZ", 1051 => "sk-SK", 1053 => "sv-SE", 1030 => "da-DK", 1044 => "nb-NO", 2068 => "nn-NO",
        1038 => "hu-HU", 1032 => "el-GR", 1050 => "hr-HR", 1060 => "sl-SI", 1063 => "lt-LT", 1062 => "lv-LV", 1061 => "et-EE",
        1048 => "ro-RO", 2074 or 3098 => "sr", _ => null,
    };
    CharProps _defChar = new();
    ParaProps _defPara = new();

    RtfReader(DocxOptions options, string? baseDir) { _options = options; _baseDir = baseDir; }

    Encoding Enc(int cp)
    {
        try { return Encoding.GetEncoding(cp); } catch { return Encoding.GetEncoding(1252); }
    }

    static int CharsetCodepage(int charset) => charset switch
    {
        0 => 1252, 1 => 0, 2 => 42, 77 => 10000, 78 => 10001, 79 => 10003, 80 => 10008, 81 => 10002, 83 => 10005, 84 => 10004,
        85 => 10006, 86 => 10081, 87 => 10021, 88 => 10029, 89 => 10007, 128 => 932, 129 => 949, 130 => 1361, 134 => 936,
        136 => 950, 161 => 1253, 162 => 1254, 163 => 1258, 177 => 1255, 178 => 1256, 186 => 1257, 204 => 1251, 222 => 874,
        238 => 1250, 254 => 437, 255 => 850, _ => 0,
    };

    /// <summary>Кодовая страница текущего шрифта (42 — символьный шрифт: байты как есть).</summary>
    int CodepageOf(CharProps c)
    {
        if (c.Codepage > 0) return c.Codepage;
        if (_fonts.TryGetValue(c.Font >= 0 ? c.Font : _defFont, out var f))
        {
            if (f.Cpg > 0) return f.Cpg;
            int cp = CharsetCodepage(f.Charset);
            if (cp > 0) return cp;
        }
        return _docCodepage;
    }

    Rgb? ColorAt(int i) => i > 0 && i < _colors.Count ? _colors[i] : null;

    // ================================================================
    //  Построение документа
    // ================================================================

    /// <summary>Приёмник блоков: тело, колонтитул, сноска, надпись (со своим стеком таблиц).</summary>
    sealed class Sink
    {
        public List<Block> Blocks = [];
        public List<Inline> Inlines = [];
        public readonly List<TableBuilder> Tables = [];
        public List<string>? Bookmarks;
        public List<Block> Pending = [];             // плавающие объекты текущего абзаца
    }

    Sink _sink = new();
    readonly List<Section> _sections = [];
    SectionProps _sect = new();
    SectionProps _docSect = new();
    string? _link;                                    // активная гиперссылка (результат поля)
    int _footCount, _endCount;
    readonly List<List<Block>> _endnotes = [];
    bool _markup;
    string _footFmt = "decimal", _endFmt = "lowerRoman";
    int _footStart = 1, _endStart = 1;

    DocxDocument ReadAll(Group top)
    {
        _markup = _options.Revisions != "final";
        _docSect = new SectionProps
        {
            PageWidth = 612, PageHeight = 792, MarginLeft = 90, MarginRight = 90, MarginTop = 72, MarginBottom = 72,
            HeaderDistance = 36, FooterDistance = 36,
        };
        _sect = CloneSect(_docSect);
        var st = new State();
        Walk(top, st, first: true);
        FlushParagraph(st, force: false);
        CloseTables(0, st);
        AppendEndnotes();
        _sections.Add(new Section { Props = _sect, Blocks = _sink.Blocks });
        _doc.Sections.AddRange(_sections);
        // Колонтитулы наследуются следующими разделами
        SectionProps? prev = null;
        foreach (var s in _doc.Sections)
        {
            var p = s.Props;
            if (prev != null)
            {
                p.HeaderDefault ??= prev.HeaderDefault; p.HeaderFirst ??= prev.HeaderFirst; p.HeaderEven ??= prev.HeaderEven;
                p.FooterDefault ??= prev.FooterDefault; p.FooterFirst ??= prev.FooterFirst; p.FooterEven ??= prev.FooterEven;
            }
            prev = p;
            if (s.Blocks.Count == 0) s.Blocks.Add(new Paragraph { MarkStyle = new TextStyle { Size = 1 }, LineRule = "exact", LineValue = 1 });
        }
        return _doc;
    }

    static SectionProps CloneSect(SectionProps s)
    {
        var c = new SectionProps
        {
            PageWidth = s.PageWidth, PageHeight = s.PageHeight, MarginLeft = s.MarginLeft, MarginRight = s.MarginRight,
            MarginTop = s.MarginTop, MarginBottom = s.MarginBottom, HeaderDistance = s.HeaderDistance, FooterDistance = s.FooterDistance,
            Columns = s.Columns, ColumnSpace = s.ColumnSpace, TitlePage = s.TitlePage, Type = s.Type,
            PageBorderTop = s.PageBorderTop, PageBorderBottom = s.PageBorderBottom, PageBorderLeft = s.PageBorderLeft, PageBorderRight = s.PageBorderRight,
            PageBorderFromText = s.PageBorderFromText, PageNumberFormat = s.PageNumberFormat, VAlign = s.VAlign,
        };
        return c;
    }

    static double Tw(int twips) => twips / 20.0;

    // ---------------- Обход ----------------

    static readonly HashSet<string> SkipDest =
    [
        "fonttbl", "colortbl", "stylesheet", "info", "listtable", "listoverridetable", "revtbl", "rsidtbl", "generator", "xmlnstbl",
        "themedata", "colorschememapping", "latentstyles", "datastore", "pgdsctbl", "mmathPr", "filetbl", "fchars", "lchars",
        "nonshppict", "bkmkend", "atnid", "atnauthor", "atndate", "atnref", "annotation", "atrfstart", "atrfend", "atnparent",
        "nonesttables", "listpicture", "pnseclvl", "panose", "falt", "userprops", "docvar", "xform", "wgrffmtfilter", "aftncn",
        "aftnsep", "aftnsepc", "ftncn", "ftnsep", "ftnsepc", "template", "operator", "company", "picprop",
        "defchp", "defpap", "pgptbl", "protusertbl", "ffdeftext", "ffname", "ffentrymcr", "ffexitmcr", "ffstattext", "ffhelptext",
        "fldtype", "datafield", "objdata", "objclass", "objname", "objtime", "oleclsid", "passwordhash", "leveltext", "levelnumbers",
        "mhtmltag", "htmltag", "blipuid", "shpgrp_", "tc", "tcn", "txe", "xe", "rxe", "pn", "listname", "styrsid", "shprslt_",
        "dptxbxtext", "do", "macpict", "wbitmap", "sp", "sn", "sv", "hl", "hlfr", "hlloc", "hlsrc", "factoidname", "mmathPict",
    ];

    void Walk(Group g, State st, bool first = false)
    {
        // Назначение группы
        string? dest = null;
        bool star = false;
        int k = 0;
        if (g.Items.Count > 0 && g.Items[0] is Sym { C: '*' }) { star = true; k = 1; }
        if (k < g.Items.Count && g.Items[k] is Word w0) dest = w0.Name;
        if (dest != null && HandleDestination(dest, g, st, star)) return;
        if (star && dest != null && !KnownStarDest(dest)) return;   // неизвестное \* — пропускаем

        for (int i = 0; i < g.Items.Count; i++)
        {
            var it = g.Items[i];
            switch (it)
            {
                case Group sub:
                    Walk(sub, st.Clone());
                    break;
                case Word w:
                    if (w.Name == "u") { EmitUnicode(w.Param ?? 0, st); i = SkipFallback(g, i, st.Uc); break; }
                    if (w.Name == "upr")
                    {
                        // Пара ANSI/Юникод: \upr{ANSI}{\*\ud{Юникод}} — выводится только вариант \ud
                        int j = i + 1, groups = 0;
                        Group? ud = null;
                        while (j < g.Items.Count && groups < 2)
                        {
                            if (g.Items[j] is Group gg)
                            {
                                groups++;
                                if (gg.Items.Count > 1 && gg.Items[0] is Sym { C: '*' } && gg.Items[1] is Word { Name: "ud" }) ud = gg;
                            }
                            j++;
                        }
                        if (ud != null) { var inner = new Group(); inner.Items.AddRange(ud.Items.Skip(2)); Walk(inner, st.Clone()); }
                        i = j - 1;
                        break;
                    }
                    ControlWord(w, st);
                    break;
                case Sym s:
                    ControlSymbol(s.C, st);
                    break;
                case Bytes b:
                {
                    // Байты текста подряд (включая \'hh) декодируются вместе (двухбайтовые кодировки)
                    var buf = new List<byte>(b.Data);
                    while (i + 1 < g.Items.Count && g.Items[i + 1] is Bytes nb) { buf.AddRange(nb.Data); i++; }
                    EmitBytes(buf.ToArray(), st);
                    break;
                }
            }
        }
    }

    static bool KnownStarDest(string d) => d is "shppict" or "fldinst" or "bkmkstart" or "footnote" or "listtext" or "shpinst" or "shptxt"
        or "nesttableprops" or "pntext" or "ud" or "chdate" or "pntxta" or "pntxtb";

    /// <summary>Пропуск замещающих символов после \u (по \uc), с учётом байтов \'hh.</summary>
    static int SkipFallback(Group g, int i, int uc)
    {
        int left = uc;
        int j = i + 1;
        while (left > 0 && j < g.Items.Count)
        {
            if (g.Items[j] is Bytes b)
            {
                if (b.Hex) { left--; j++; continue; }
                // Обычный текст: отрезаем первые символы
                int take = Math.Min(left, b.Data.Length);
                if (take == b.Data.Length) { left -= take; j++; continue; }
                g.Items[j] = new Bytes(b.Data[take..], false);
                left = 0;
                break;
            }
            if (g.Items[j] is Sym) { left--; j++; continue; }
            break;
        }
        return j - 1;
    }

    // ---------------- Текст ----------------

    void EmitUnicode(int code, State st)
    {
        if (code < 0) code += 65536;
        if (code is < 0 or > 0x10FFFF) code = 0xFFFD;
        if (st.C.Hidden) return;
        if (st.C.Deleted && !_markup) return;
        // Символы шрифтов Symbol/Wingdings (F0xx) — в Юникод
        string t = char.ConvertFromUtf32(code is >= 0xD800 and <= 0xDFFF ? 0xFFFD : code);
        if (code is >= 0xF020 and <= 0xF0FF || CodepageOf(st.C) == 42) t = SymbolText(t, st.C);
        AddText(t, st);
    }

    /// <summary>Символы символьного шрифта (Symbol, Wingdings) → Юникод.</summary>
    string SymbolText(string s, CharProps c)
    {
        string font = _fonts.TryGetValue(c.Font >= 0 ? c.Font : _defFont, out var f) ? f.Name : "";
        if (!font.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase) && !font.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase)) return s;
        var sb = new StringBuilder(s.Length);
        foreach (char ch in s) sb.Append(char.ConvertFromUtf32(DocxToPdf.Fonts.FontManager.MapSymbolChar(font, ch)));
        return sb.ToString();
    }

    void EmitBytes(byte[] data, State st)
    {
        if (st.C.Hidden) return;
        if (st.C.Deleted && !_markup) return;
        int cp = CodepageOf(st.C);
        string s = cp == 42 ? SymbolText(Encoding.Latin1.GetString(data), st.C) : Enc(cp).GetString(data);
        s = s.Replace('\t', '\u0009');
        // Табуляция буквально — как \tab
        int start = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '\t') continue;
            if (i > start) AddText(s[start..i], st);
            _sink.Inlines.Add(new TabInline(Style(st.C)));
            start = i + 1;
        }
        if (start < s.Length) AddText(s[start..], st);
    }

    void AddText(string s, State st)
    {
        if (s.Length == 0) return;
        var ts = Style(st.C);
        var inl = _sink.Inlines;
        if (inl.Count > 0 && inl[^1] is TextInline prev && prev.Style == ts) prev.Text += s;
        else inl.Add(new TextInline(s, ts));
    }

    readonly Dictionary<string, int> _authorIdx = [];
    static readonly Rgb[] AuthorColors =
    [
        new(0.71, 0.05, 0.06), new(0.12, 0.3, 0.75), new(0.0, 0.5, 0.2), new(0.5, 0.15, 0.6),
        new(0.75, 0.4, 0.0), new(0.0, 0.5, 0.55), new(0.55, 0.35, 0.15), new(0.4, 0.4, 0.4),
    ];

    internal TextStyle Style(CharProps c)
    {
        string font = _fonts.TryGetValue(c.Font >= 0 ? c.Font : _defFont, out var f) && f.Name.Length > 0 ? f.Name : "Times New Roman";
        Rgb? bg = c.Highlight > 0 ? ColorAt(c.Highlight) : c.Back > 0 ? ColorAt(c.Back) : null;
        var ts = new TextStyle
        {
            FontAscii = font, FontHAnsi = font, FontEastAsia = font,
            Size = Math.Max(1, c.Size / 2.0),
            Bold = c.Bold, Italic = c.Italic,
            Underline = c.Underline != null, UnderlineStyle = c.Underline, DoubleUnderline = c.Underline is "double" or "wavyDouble",
            UnderlineColor = c.UlColor > 0 ? ColorAt(c.UlColor) : null,
            Strike = c.Strike || c.DStrike, DoubleStrike = c.DStrike,
            Caps = c.Caps, SmallCaps = c.SmallCaps,
            Color = ColorAt(c.Color) ?? Rgb.Black,
            Background = bg,
            VertAlign = c.Vert,
            Position = c.Up / 2.0,
            LetterSpacing = c.Expand / 20.0,
            HScale = c.Scale / 100.0,
            Outline = c.Outline,
            ShadowColor = c.Shadow || c.Emboss || c.Imprint ? new Rgb(0.55, 0.55, 0.55) : null,
            ShadowOffset = c.Imprint ? -0.04 : c.Emboss ? 0.04 : c.Shadow ? 0.06 : 0,
            Link = _link,
            Kerning = c.Kerning > 0 && c.Size >= c.Kerning,
            Hyphens = _hyphAuto ? 2 : 1,
            Lang = LcidName(c.Lang > 0 ? c.Lang : _defLang),
        };
        if (c.Emboss || c.Imprint) ts = ts with { Color = new Rgb(0.93, 0.93, 0.93) };
        if (_markup && (c.Revised || c.Deleted))
        {
            string author = c.RevAuthor >= 0 && c.RevAuthor < _revAuthors.Count ? _revAuthors[c.RevAuthor] : "";
            if (!_authorIdx.TryGetValue(author, out int ai)) _authorIdx[author] = ai = _authorIdx.Count;
            var col = AuthorColors[ai % AuthorColors.Length];
            ts = c.Deleted
                ? ts with { Color = col, Strike = true, RevisionMark = true }
                : ts with { Color = col, Underline = true, UnderlineStyle = "single", UnderlineColor = col, RevisionMark = true };
        }
        return ts;
    }

    // ---------------- Абзацы ----------------

    Paragraph MakeParagraph(ParaProps p, CharProps markChar)
    {
        var para = new Paragraph
        {
            StyleId = _styleNames.GetValueOrDefault(p.Style) ?? "",
            Align = p.Align,
            IndLeft = Tw(p.Li), IndRight = Tw(p.Ri), IndFirst = Tw(p.Fi),
            SpaceBefore = p.SbAuto == 1 ? 14 : Math.Max(0, Tw(p.Sb)),
            SpaceAfter = p.SaAuto == 1 ? 14 : Math.Max(0, Tw(p.Sa)),
            KeepNext = p.KeepN, KeepLines = p.Keep, PageBreakBefore = p.PageBB, WidowControl = p.WidCtl,
            ContextualSpacing = p.Contextual, Rtl = p.Rtl,
            Tabs = p.Tabs.OrderBy(t => t.Pos).Select(t => new TabStop(Tw(t.Pos), t.Align, t.Leader)).ToList(),
            BorderTop = Vis(p.BTop), BorderBottom = Vis(p.BBottom), BorderLeft = Vis(p.BLeft), BorderRight = Vis(p.BRight), BorderBetween = Vis(p.BBetween),
            MarkStyle = Style(markChar),
        };
        para.NoHyphenation = p.NoHyph;
        para.HyphenZone = _hyphZone;
        para.HyphenLimit = _hyphConsec;
        para.HyphenNoCaps = _hyphNoCaps;
        if (p.Rtl && para.Align == "left") para.Align = "right";
        else if (p.Rtl && para.Align == "right") para.Align = "left";
        if (p.Sl != 0)
        {
            if (p.SlMult == 1) { para.LineRule = "auto"; para.LineValue = Math.Abs(p.Sl); }
            else if (p.Sl < 0) { para.LineRule = "exact"; para.LineValue = Tw(-p.Sl); }
            else { para.LineRule = "atLeast"; para.LineValue = Tw(p.Sl); }
        }
        if (p.Shading > 0 && ColorAt(p.Shading) is Rgb sh) para.Shading = sh;
        else if (p.ShadePct > 0)
        {
            var fg = ColorAt(p.ShadeFore) ?? Rgb.Black;
            double a = p.ShadePct / 10000.0;
            para.Shading = new Rgb(1 - (1 - fg.R) * a, 1 - (1 - fg.G) * a, 1 - (1 - fg.B) * a);
        }
        return para;
    }

    static BorderDef? Vis(BorderDef? b) => b is { IsNone: false } ? b : null;

    /// <summary>Конец абзаца: \par (или \cell / \row в таблице, конец документа).</summary>
    void FlushParagraph(State st, bool force = true)
    {
        var inl = _sink.Inlines;
        if (!force && inl.Count == 0 && _sink.Pending.Count == 0) return;
        var para = MakeParagraph(st.P, st.C);
        // Метка списка: \listtext/\pntext или по таблице списков
        if (_pendingLabel != null) { para.Inlines.AddRange(_pendingLabel); _pendingLabel = null; }
        else if (st.P.Ls > 0 && ListLabel(st.P, st.C) is { } lbl) para.Inlines.AddRange(lbl);
        para.Inlines.AddRange(inl);
        para.Bookmarks = _sink.Bookmarks;
        _sink.Bookmarks = null;
        _sink.Inlines = [];
        var pending = _sink.Pending;
        _sink.Pending = [];
        int level = st.P.InTable ? Math.Max(1, st.P.Itap) : 0;
        if (level == 0) CloseTables(0, st);
        var target = Target(level, st);
        target.AddRange(pending);
        target.Add(para);
    }

    List<Inline>? _pendingLabel;

    /// <summary>Куда добавлять блоки уровня таблицы L (0 — вне таблиц).</summary>
    List<Block> Target(int level, State st)
    {
        if (level == 0) return _sink.Blocks;
        while (_sink.Tables.Count < level) _sink.Tables.Add(new TableBuilder());
        // Таблицы глубже текущего уровня закрываются (их строки уже завершены)
        if (_sink.Tables.Count > level) CloseTables(level, st);
        return _sink.Tables[level - 1].CurCell;
    }

    // ---------------- Управляющие слова ----------------

    BorderDef? _curBorder;                            // последняя упомянутая рамка (для \brdrs, \brdrw, \brdrcf, \brsp)

    void ControlSymbol(char c, State st)
    {
        switch (c)
        {
            case '~': AddText(" ", st); break;
            case '_': AddText("‑", st); break;
            case '-': AddText("\u00AD", st); break;  // мягкий перенос
            case '|': break;
            case ':': break;
            case '\t': EmitBytes([(byte)'\t'], st); break;
        }
    }

    void ControlWord(Word w, State st)
    {
        var c = st.C; var p = st.P;
        int v = w.Param ?? 1;
        bool on = w.Param != 0;
        switch (w.Name)
        {
            // ----- документ -----
            case "ansi": _docCodepage = 1252; break;
            case "mac": _docCodepage = 10000; break;
            case "pc": _docCodepage = 437; break;
            case "pca": _docCodepage = 850; break;
            case "ansicpg": if (v > 0) _docCodepage = v; break;
            case "deff": _defFont = v; if (c.Font < 0) c.Font = v; _defChar.Font = v; break;
            case "deflang": _defLang = v; break;
            case "lang": c.Lang = v; break;
            case "kerning": c.Kerning = v; break;
            case "hyphauto": _hyphAuto = on; break;
            case "hyphhotz": _hyphZone = Tw(v); break;
            case "hyphconsec": _hyphConsec = v; break;
            case "hyphcaps": _hyphNoCaps = !on; break;
            case "hyphpar": p.NoHyph = !on; break;
            case "deftab": if (v > 0) _doc.DefaultTabStop = Tw(v); break;
            case "paperw": _docSect.PageWidth = _sect.PageWidth = Tw(v); break;
            case "paperh": _docSect.PageHeight = _sect.PageHeight = Tw(v); break;
            case "margl": _docSect.MarginLeft = _sect.MarginLeft = Tw(v); break;
            case "margr": _docSect.MarginRight = _sect.MarginRight = Tw(v); break;
            case "margt": _docSect.MarginTop = _sect.MarginTop = Tw(v); break;
            case "margb": _docSect.MarginBottom = _sect.MarginBottom = Tw(v); break;
            case "gutter": _docSect.MarginLeft += Tw(v); _sect.MarginLeft += Tw(v); break;
            case "landscape": if (_docSect.PageWidth < _docSect.PageHeight) { (_docSect.PageWidth, _docSect.PageHeight) = (_docSect.PageHeight, _docSect.PageWidth); _sect.PageWidth = _docSect.PageWidth; _sect.PageHeight = _docSect.PageHeight; } break;
            case "facingp": _doc.EvenAndOddHeaders = true; break;
            case "margmirror": _doc.MirrorMargins = true; break;
            case "titlepg": _sect.TitlePage = true; break;
            case "ftnnar": _footFmt = "decimal"; break;
            case "ftnnalc": _footFmt = "lowerLetter"; break;
            case "ftnnauc": _footFmt = "upperLetter"; break;
            case "ftnnrlc": _footFmt = "lowerRoman"; break;
            case "ftnnruc": _footFmt = "upperRoman"; break;
            case "aftnnar": _endFmt = "decimal"; break;
            case "aftnnalc": _endFmt = "lowerLetter"; break;
            case "aftnnauc": _endFmt = "upperLetter"; break;
            case "aftnnrlc": _endFmt = "lowerRoman"; break;
            case "aftnnruc": _endFmt = "upperRoman"; break;
            case "ftnstart": _footStart = v; break;
            case "aftnstart": _endStart = v; break;
            case "pgbrdrt": _curBorder = _sect.PageBorderTop = new BorderDef(); break;
            case "pgbrdrb": _curBorder = _sect.PageBorderBottom = new BorderDef(); break;
            case "pgbrdrl": _curBorder = _sect.PageBorderLeft = new BorderDef(); break;
            case "pgbrdrr": _curBorder = _sect.PageBorderRight = new BorderDef(); break;
            case "pgbrdropt": _sect.PageBorderFromText = (v & 32) == 0 ? _sect.PageBorderFromText : false; if ((v & 32) == 0) _sect.PageBorderFromText = true; break;

            // ----- раздел -----
            case "sectd":
                _sect = CloneSect(_docSect);
                _sect.HeaderDefault = _sect.HeaderFirst = _sect.HeaderEven = _sect.FooterDefault = _sect.FooterFirst = _sect.FooterEven = null;
                _sect.TitlePage = false;
                break;
            case "sect": EndSection(st); break;
            case "sbknone": _sect.Type = "continuous"; break;
            case "sbkpage": _sect.Type = "nextPage"; break;
            case "sbkcol": _sect.Type = "nextColumn"; break;
            case "sbkeven": _sect.Type = "evenPage"; break;
            case "sbkodd": _sect.Type = "oddPage"; break;
            case "cols": _sect.Columns = Math.Max(1, v); break;
            case "colsx": _sect.ColumnSpace = Tw(v); break;
            case "linebetcol": _sect.ColumnSeparator = true; break;
            case "pgwsxn": _sect.PageWidth = Tw(v); break;
            case "pghsxn": _sect.PageHeight = Tw(v); break;
            case "marglsxn": _sect.MarginLeft = Tw(v); break;
            case "margrsxn": _sect.MarginRight = Tw(v); break;
            case "margtsxn": _sect.MarginTop = Tw(v); break;
            case "margbsxn": _sect.MarginBottom = Tw(v); break;
            case "guttersxn": _sect.MarginLeft += Tw(v); break;
            case "lndscpsxn": if (_sect.PageWidth < _sect.PageHeight) (_sect.PageWidth, _sect.PageHeight) = (_sect.PageHeight, _sect.PageWidth); break;
            case "headery": _sect.HeaderDistance = Tw(v); if (_sections.Count == 0 && _sink.Blocks.Count == 0) _docSect.HeaderDistance = Tw(v); break;
            case "footery": _sect.FooterDistance = Tw(v); if (_sections.Count == 0 && _sink.Blocks.Count == 0) _docSect.FooterDistance = Tw(v); break;
            case "pgnstarts": _sect.PageNumberStart = v; break;
            case "pgnrestart": _sect.PageNumberStart ??= 1; break;
            case "pgnlcrm": _sect.PageNumberFormat = "lowerRoman"; break;
            case "pgnucrm": _sect.PageNumberFormat = "upperRoman"; break;
            case "pgnlcltr": _sect.PageNumberFormat = "lowerLetter"; break;
            case "pgnucltr": _sect.PageNumberFormat = "upperLetter"; break;
            case "pgndec": _sect.PageNumberFormat = "decimal"; break;
            case "vertalc": _sect.VAlign = "center"; break;
            case "vertalb": _sect.VAlign = "bottom"; break;
            case "vertalj": _sect.VAlign = "both"; break;
            case "linemod": _sect.LineNumberCountBy = v; break;
            case "linex": _sect.LineNumberDistance = Tw(v); break;
            case "linestarts": _sect.LineNumberStart = v; break;
            case "linecont": _sect.LineNumberRestart = "continuous"; break;
            case "lineppage": _sect.LineNumberRestart = "newPage"; break;

            // ----- абзац -----
            case "pard":
                st.P = _defPara.Clone();
                break;
            case "par":
                FlushParagraph(st);
                break;
            case "s": p.Style = v; ApplyStyle(v, st, paragraph: true); break;
            case "ql": p.Align = "left"; break;
            case "qr": p.Align = "right"; break;
            case "qc": p.Align = "center"; break;
            case "qj": p.Align = "both"; break;
            case "qd": p.Align = "distribute"; break;
            case "li": case "lin": p.Li = v; break;
            case "ri": case "rin": p.Ri = v; break;
            case "fi": p.Fi = v; break;
            case "sb": p.Sb = v; break;
            case "sa": p.Sa = v; break;
            case "sbauto": p.SbAuto = v; break;
            case "saauto": p.SaAuto = v; break;
            case "sl": p.Sl = v; if (w.Param == null) p.Sl = 0; break;
            case "slmult": p.SlMult = v; break;
            case "keepn": p.KeepN = on; break;
            case "keep": p.Keep = on; break;
            case "pagebb": p.PageBB = on; break;
            case "widctlpar": p.WidCtl = true; break;
            case "nowidctlpar": p.WidCtl = false; break;
            case "contextualspace": p.Contextual = true; break;
            case "rtlpar": p.Rtl = true; break;
            case "ltrpar": p.Rtl = false; break;
            case "intbl": p.InTable = true; if (p.Itap == 0) p.Itap = 1; break;
            case "itap": p.Itap = v; p.InTable = v > 0; break;
            case "ls": p.Ls = v; break;
            case "ilvl": p.Ilvl = v; break;
            case "outlinelevel": p.OutlineLevel = v; break;
            case "tqr": p.PendingTabAlign = "right"; break;
            case "tqc": p.PendingTabAlign = "center"; break;
            case "tqdec": p.PendingTabAlign = "decimal"; break;
            case "tldot": p.PendingLeader = "dot"; break;
            case "tlmdot": p.PendingLeader = "middleDot"; break;
            case "tlhyph": p.PendingLeader = "hyphen"; break;
            case "tlul": case "tlth": p.PendingLeader = "underscore"; break;
            case "tleq": p.PendingLeader = "hyphen"; break;
            case "tx":
                p.Tabs.RemoveAll(t => t.Pos == v);
                p.Tabs.Add(new Tab { Pos = v, Align = p.PendingTabAlign ?? "left", Leader = p.PendingLeader ?? "none" });
                p.PendingTabAlign = null; p.PendingLeader = null;
                break;
            case "tb": p.PendingTabAlign = null; p.PendingLeader = null; break;
            case "brdrt": _curBorder = p.BTop = new BorderDef(); break;
            case "brdrb": _curBorder = p.BBottom = new BorderDef(); break;
            case "brdrl": _curBorder = p.BLeft = new BorderDef(); break;
            case "brdrr": _curBorder = p.BRight = new BorderDef(); break;
            case "brdrbtw": _curBorder = p.BBetween = new BorderDef(); break;
            case "box":
                _curBorder = new BorderDef();
                p.BTop = p.BBottom = p.BLeft = p.BRight = _curBorder;
                break;
            case "cbpat": p.Shading = v; break;
            case "shading": p.ShadePct = v; break;
            case "cfpat": p.ShadeFore = v; break;

            // ----- рамки (общие модификаторы) -----
            case "brdrs" or "brdrsh" or "brdroutset" or "brdrinset" or "brdrframe" or "brdremboss" or "brdrengrave": SetBorder("single"); break;
            case "brdrth": SetBorder("single"); if (_curBorder != null && _curBorder.Width < 1) _curBorder.Width = 1.5; break;
            case "brdrdb" or "brdrtriple" or "brdrtnthsg" or "brdrthtnsg" or "brdrtnthtnsg" or "brdrtnthmg" or "brdrthtnmg" or "brdrtnthlg" or "brdrthtnlg": SetBorder("double"); break;
            case "brdrdot": SetBorder("dotted"); break;
            case "brdrdash" or "brdrdashsm" or "brdrdashd" or "brdrdashdd" or "brdrdashdotstr": SetBorder("dashed"); break;
            case "brdrhair": SetBorder("single"); if (_curBorder != null) _curBorder.Width = 0.25; break;
            case "brdrwavy" or "brdrwavydb": SetBorder("single"); break;
            case "brdrnone" or "brdrnil" or "brdrtbl": SetBorder("nil"); break;
            case "brdrw": if (_curBorder != null) _curBorder.Width = Math.Max(0.25, Tw(v) * (_curBorder.Val == "double" ? 3 : 1)); break;
            case "brdrcf": if (_curBorder != null) _curBorder.Color = ColorAt(v) ?? Rgb.Black; break;
            case "brsp": if (_curBorder != null) _curBorder.Space = Tw(v); break;

            // ----- символы -----
            case "plain":
            {
                var keepCp = c.Codepage;
                st.C = _defChar.Clone();
                st.C.Codepage = 0;
                _ = keepCp;
                break;
            }
            case "b": c.Bold = on; break;
            case "i": c.Italic = on; break;
            case "ul": c.Underline = on ? "single" : null; break;
            case "uld": c.Underline = on ? "dotted" : null; break;
            case "uldash": c.Underline = on ? "dash" : null; break;
            case "uldashd": c.Underline = on ? "dotDash" : null; break;
            case "uldashdd": c.Underline = on ? "dotDotDash" : null; break;
            case "uldb": c.Underline = on ? "double" : null; break;
            case "ulth": case "ulthd": case "ulthdash": c.Underline = on ? "thick" : null; break;
            case "ulw": c.Underline = on ? "words" : null; break;
            case "ulwave": c.Underline = on ? "wave" : null; break;
            case "ululdbwave": c.Underline = on ? "wavyDouble" : null; break;
            case "ulldash": c.Underline = on ? "dashLong" : null; break;
            case "ulnone": c.Underline = null; break;
            case "ulc": c.UlColor = v; break;
            case "strike": c.Strike = on; break;
            case "striked": c.DStrike = on; break;
            case "caps": c.Caps = on; break;
            case "scaps": c.SmallCaps = on; break;
            case "v": c.Hidden = on; break;
            case "outl": c.Outline = on; break;
            case "shad": c.Shadow = on; break;
            case "embo": c.Emboss = on; break;
            case "impr": c.Imprint = on; break;
            case "super": c.Vert = 1; break;
            case "sub": c.Vert = -1; break;
            case "nosupersub": c.Vert = 0; break;
            case "up": c.Up = w.Param ?? 6; break;
            case "dn": c.Up = -(w.Param ?? 6); break;
            case "fs": c.Size = w.Param ?? 24; break;
            case "f": c.Font = v; c.Codepage = 0; break;
            case "cf": c.Color = v; break;
            case "cb": c.Back = v; break;
            case "chcbpat": c.Back = v; break;
            case "highlight": c.Highlight = v; break;
            case "expnd": c.Expand = v * 5; break;
            case "expndtw": c.Expand = v; break;
            case "charscalex": c.Scale = v; break;
            case "cs": ApplyStyle(v, st, paragraph: false); break;
            case "revised": c.Revised = on; break;
            case "deleted": c.Deleted = on; break;
            case "revauth": case "revauthdel": c.RevAuthor = v; break;
            case "uc": st.Uc = Math.Max(0, v); break;
            case "cpg": c.Codepage = v; break;

            // ----- специальные символы -----
            case "tab": _sink.Inlines.Add(new TabInline(Style(c))); break;
            case "line": _sink.Inlines.Add(new BreakInline(BreakKind.Line, Style(c))); break;
            case "page": if (!p.InTable) _sink.Inlines.Add(new BreakInline(BreakKind.Page, Style(c))); break;
            case "column": _sink.Inlines.Add(new BreakInline(BreakKind.Column, Style(c))); break;
            case "emdash": AddText("—", st); break;
            case "endash": AddText("–", st); break;
            case "bullet": AddText("•", st); break;
            case "lquote": AddText("‘", st); break;
            case "rquote": AddText("’", st); break;
            case "ldblquote": AddText("“", st); break;
            case "rdblquote": AddText("”", st); break;
            case "emspace": AddText(" ", st); break;
            case "enspace": AddText(" ", st); break;
            case "qmspace": AddText(" ", st); break;
            case "zwj": AddText("‍", st); break;
            case "zwnj": AddText("‌", st); break;
            case "ltrmark": AddText("‎", st); break;
            case "rtlmark": AddText("‏", st); break;
            case "chpgn": _sink.Inlines.Add(new FieldInline(FieldKind.Page, Style(c))); break;
            case "chftn":
                if (_noteMark != null) AddText(_noteMark, st);
                else
                {
                    // Знак сноски в тексте: заполнитель, его заменит следующая группа \footnote
                    _chftn = new TextInline(NoteMark(_footCount + _footStart, _footFmt), Style(c));
                    _sink.Inlines.Add(_chftn);
                }
                break;

            // ----- таблицы -----
            case "cell": EndCell(1, st); break;
            case "nestcell": EndCell(Math.Max(2, p.Itap), st); break;
            case "row": EndRow(1, st); break;
            case "nestrow": EndRow(Math.Max(2, _rowDefLevel), st); break;
            default:
                TableWord(w, st);
                break;
        }
    }

    void SetBorder(string val)
    {
        if (_curBorder == null) return;
        _curBorder.Val = val;
        if (val == "double" && _curBorder.Width < 1.5) _curBorder.Width = Math.Max(_curBorder.Width * 3, 1.5);
    }

    void EndSection(State st)
    {
        FlushParagraph(st, force: false);
        CloseTables(0, st);
        _sections.Add(new Section { Props = _sect, Blocks = _sink.Blocks });
        _sink.Blocks = [];
        // Новый раздел: свойства продолжаются до \sectd
        var next = CloneSect(_sect);
        next.Type = "nextPage";
        next.PageNumberStart = null;
        _sect = next;
    }

    // ---------------- Стили ----------------

    void ApplyStyle(int idx, State st, bool paragraph)
    {
        var dict = paragraph ? _paraStyles : _charStyles;
        var chain = new List<List<Node>>();
        int cur = idx;
        for (int guard = 0; guard < 16 && dict.TryGetValue(cur, out var toks); guard++)
        {
            chain.Add(toks);
            if (!_styleBase.TryGetValue((paragraph ? 0 : 100000) + cur, out var b) || b == cur || b == 222) break;
            cur = b;
        }
        chain.Reverse();
        foreach (var toks in chain)
            foreach (var t in toks)
                if (t is Word w && w.Name is not ("s" or "cs" or "ts" or "sbasedon" or "snext" or "slink" or "additive" or "sautoupd" or "shidden" or "sqformat" or "spriority" or "ssemihidden" or "sunhideused" or "styrsid" or "par" or "tab" or "cell" or "row"))
                    ControlWord(w, st);
    }

    // ---------------- Назначения ----------------

    bool HandleDestination(string dest, Group g, State st, bool star)
    {
        switch (dest)
        {
            case "fonttbl": ReadFonts(g); return true;
            case "colortbl": ReadColors(g); return true;
            case "stylesheet": ReadStyles(g); return true;
            case "info": ReadInfo(g); return true;
            case "listtable": ReadListTable(g); return true;
            case "listoverridetable": ReadOverrides(g); return true;
            case "revtbl":
                foreach (var a in g.Items.OfType<Group>()) _revAuthors.Add(PlainText(a).TrimEnd(';').Trim());
                return true;
            case "defchp":
            {
                var s = new State { C = _defChar.Clone() };
                foreach (var w in g.Items.OfType<Word>().Skip(1)) ControlWord(w, s);
                _defChar = s.C;
                st.C = _defChar.Clone();
                return true;
            }
            case "defpap":
            {
                var s = new State { P = _defPara.Clone() };
                foreach (var w in g.Items.OfType<Word>().Skip(1)) ControlWord(w, s);
                _defPara = s.P;
                st.P = _defPara.Clone();
                return true;
            }
            case "pict": Picture(g, st); return true;
            case "shppict":
                foreach (var pg in g.Items.OfType<Group>()) Walk(pg, st.Clone());
                return true;
            case "nonshppict": return true;
            case "background":
            {
                // Фон страниц: заливка фигуры (fillColor — BGR)
                var fc = FindProp(g, "fillColor");
                if (fc != null && int.TryParse(fc, out var bgr)) _doc.Background = new Rgb((bgr & 0xFF) / 255.0, (bgr >> 8 & 0xFF) / 255.0, (bgr >> 16 & 0xFF) / 255.0);
                return true;
            }
            case "object": ObjectGroup(g, st); return true;
            case "field": Field(g, st); return true;
            case "footnote": Footnote(g, st); return true;
            case "header" or "headerl" or "headerr" or "headerf" or "footer" or "footerl" or "footerr" or "footerf":
                HeaderFooter(dest, g, st); return true;
            case "shp": Shape(g, st); return true;
            case "listtext" or "pntext":
            {
                // Готовая метка списка (номер или маркер) с оформлением
                var saved = _sink.Inlines;
                _sink.Inlines = [];
                var s = st.Clone();
                Walk(StripFirst(g, star), s);
                _pendingLabel = _sink.Inlines;
                _sink.Inlines = saved;
                return true;
            }
            case "bkmkstart":
            {
                string name = PlainText(StripFirst(g, star)).Trim();
                if (name.Length > 0) (_sink.Bookmarks ??= []).Add(name);
                return true;
            }
            case "nesttableprops":
            {
                // Свойства строки вложенной таблицы: \trowd … \nestrow
                int saved = _rowDefLevel;
                _rowDefLevel = Math.Max(2, st.P.Itap);
                foreach (var it in StripFirst(g, star).Items)
                {
                    if (it is Word w) ControlWord(w, st);
                    else if (it is Group sg) Walk(sg, st);
                }
                _rowDefLevel = saved;
                return true;
            }
            case "pn":
            {
                // Нумерация Word 6/95: метка — по \pntext; маркер, если текста нет
                if (_pendingLabel == null && g.Items.OfType<Word>().Any(x => x.Name == "pnlvlblt"))
                {
                    var ts = Style(st.C);
                    _pendingLabel = [new TextInline("•", ts), new TabInline(ts)];
                }
                return true;
            }
            case "shpinst" or "shptxt": return false;
        }
        if (SkipDest.Contains(dest) && (star || dest is not ("sp" or "sn" or "sv"))) return true;
        return false;
    }

    /// <summary>Значение свойства фигуры {\sp{\sn имя}{\sv значение}} в группе (на любой глубине).</summary>
    string? FindProp(Group g, string name)
    {
        foreach (var sg in g.Items.OfType<Group>())
        {
            if (sg.Items.FirstOrDefault() is Word { Name: "sp" })
            {
                var sn = sg.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "sn" });
                var sv = sg.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "sv" });
                if (sn != null && sv != null && PlainText(StripFirst(sn, false)).Trim() == name) return PlainText(StripFirst(sv, false)).Trim();
            }
            else if (FindProp(sg, name) is { } r) return r;
        }
        return null;
    }

    static Group StripFirst(Group g, bool star)
    {
        var r = new Group();
        r.Items.AddRange(g.Items.Skip(star ? 2 : 1));
        return r;
    }

    /// <summary>Текст группы (без управляющих слов), для таблиц шрифтов, стилей, сведений.</summary>
    string PlainText(Group g, int cp = 0)
    {
        var sb = new StringBuilder();
        var bytes = new List<byte>();
        void Flush() { if (bytes.Count > 0) { sb.Append(Enc(cp > 0 ? cp : _docCodepage).GetString(bytes.ToArray())); bytes.Clear(); } }
        for (int i = 0; i < g.Items.Count; i++)
        {
            var it = g.Items[i];
            if (it is Bytes b) bytes.AddRange(b.Data);
            else if (it is Word { Name: "u" } u)
            {
                Flush();
                int code = u.Param ?? 0; if (code < 0) code += 65536;
                sb.Append((char)code);
                i = SkipFallback(g, i, 1);
            }
            else if (it is Group sub && !(sub.Items.Count > 0 && sub.Items[0] is Sym { C: '*' })) { Flush(); sb.Append(PlainText(sub, cp)); }
            else if (it is Sym { C: '~' }) { Flush(); sb.Append(' '); }
        }
        Flush();
        return sb.ToString();
    }

    void ReadFonts(Group g)
    {
        void One(Group fg)
        {
            int num = -1, charset = 0, cpg = 0;
            foreach (var w in fg.Items.OfType<Word>())
            {
                if (w.Name == "f") num = w.Param ?? 0;
                else if (w.Name == "fcharset") charset = w.Param ?? 0;
                else if (w.Name == "cpg") cpg = w.Param ?? 0;
            }
            if (num < 0) return;
            int cp = cpg > 0 ? cpg : CharsetCodepage(charset);
            string name = PlainText(fg, cp is > 0 and not 42 ? cp : 0).Trim().TrimEnd(';').Trim();
            _fonts[num] = (name, charset, cpg);
        }
        // Шрифты бывают как группами {\f0 …;}, так и подряд без групп
        var direct = new Group();
        foreach (var it in g.Items.Skip(1))
        {
            if (it is Group fg && fg.Items.OfType<Word>().Any(w => w.Name == "f")) One(fg);
            else
            {
                direct.Items.Add(it);
                if (it is Bytes b && b.Data.Contains((byte)';')) { One(direct); direct = new Group(); }
            }
        }
    }

    void ReadColors(Group g)
    {
        int r = -1, gr = -1, bl = -1;
        bool any = false;
        foreach (var it in g.Items.Skip(1))
        {
            if (it is Word w)
            {
                if (w.Name == "red") { r = w.Param ?? 0; any = true; }
                else if (w.Name == "green") { gr = w.Param ?? 0; any = true; }
                else if (w.Name == "blue") { bl = w.Param ?? 0; any = true; }
            }
            else if (it is Bytes b)
            {
                foreach (var ch in b.Data)
                    if (ch == ';')
                    {
                        _colors.Add(any ? new Rgb(Math.Max(0, r) / 255.0, Math.Max(0, gr) / 255.0, Math.Max(0, bl) / 255.0) : null);
                        r = gr = bl = -1; any = false;
                    }
            }
        }
    }

    void ReadStyles(Group g)
    {
        foreach (var sg in g.Items.OfType<Group>())
        {
            var words = new List<Node>();
            int idx = 0; bool character = false, table = false; int? based = null;
            foreach (var it in sg.Items)
            {
                if (it is Word w)
                {
                    if (w.Name == "s") idx = w.Param ?? 0;
                    else if (w.Name == "cs") { idx = w.Param ?? 0; character = true; }
                    else if (w.Name == "ts") { idx = w.Param ?? 0; table = true; }
                    else if (w.Name == "sbasedon") based = w.Param;
                    words.Add(w);
                }
                else if (it is Sym) words.Add(it);
            }
            if (table) continue;
            string name = PlainText(sg).Trim().TrimEnd(';').Trim();
            if (character) { _charStyles[idx] = words; if (based != null) _styleBase[100000 + idx] = based.Value; }
            else { _paraStyles[idx] = words; _styleNames[idx] = name; if (based != null) _styleBase[idx] = based.Value; }
        }
    }

    void ReadInfo(Group g)
    {
        foreach (var sg in g.Items.OfType<Group>())
        {
            string? n = sg.Items.FirstOrDefault() as Word is { } w ? w.Name : null;
            string text = PlainText(sg).Trim();
            if (text.Length == 0) continue;
            if (n == "title") _doc.Title = text;
            else if (n == "author") _doc.Author = text;
            else if (n == "subject") _doc.Subject = text;
        }
    }

    // ---------------- Колонтитулы, сноски, поля ----------------

    void HeaderFooter(string dest, Group g, State st)
    {
        var saved = _sink;
        var savedLabel = _pendingLabel; _pendingLabel = null;
        _sink = new Sink();
        var s = new State { C = _defChar.Clone(), P = _defPara.Clone(), Uc = st.Uc };
        Walk(StripFirst(g, false), s);
        FlushParagraph(s, force: false);
        CloseTables(0, s);
        var hf = new HeaderFooter { Blocks = _sink.Blocks };
        _sink = saved;
        _pendingLabel = savedLabel;
        switch (dest)
        {
            case "header": case "headerr": _sect.HeaderDefault = hf; if (dest == "header" && !_doc.EvenAndOddHeaders) _sect.HeaderEven ??= null; break;
            case "headerl": _sect.HeaderEven = hf; break;
            case "headerf": _sect.HeaderFirst = hf; break;
            case "footer": case "footerr": _sect.FooterDefault = hf; break;
            case "footerl": _sect.FooterEven = hf; break;
            case "footerf": _sect.FooterFirst = hf; break;
        }
    }

    string? _noteMark;
    TextInline? _chftn;

    internal static string NoteMark(int n, string fmt) => fmt switch
    {
        "lowerRoman" => Roman(n).ToLowerInvariant(),
        "upperRoman" => Roman(n),
        "lowerLetter" => ((char)('a' + (n - 1) % 26)).ToString(),
        "upperLetter" => ((char)('A' + (n - 1) % 26)).ToString(),
        _ => n.ToString(CultureInfo.InvariantCulture),
    };

    internal static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString();
        var sb = new StringBuilder();
        (int V, string S)[] t = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        foreach (var (v, s) in t) while (n >= v) { sb.Append(s); n -= v; }
        return sb.ToString();
    }

    void Footnote(Group g, State st)
    {
        bool end = g.Items.OfType<Word>().Any(w => w.Name == "ftnalt");
        int num = end ? ++_endCount + _endStart - 1 : ++_footCount + _footStart - 1;
        string mark = NoteMark(num, end ? _endFmt : _footFmt);
        // Знак сноски в тексте — заполнитель \chftn перед группой (его оформление — у знака)
        var inl = _sink.Inlines;
        TextStyle markStyle = Style(st.C);
        bool custom = _chftn == null || !inl.Contains(_chftn);
        if (!custom) { markStyle = _chftn!.Style; inl.Remove(_chftn); }
        else if (inl.Count > 0 && inl[^1] is TextInline cm && cm.Text.Length <= 3 && cm.Style.VertAlign == 1)
        {
            // Собственный знак (\chftn нет): последний короткий фрагмент
            mark = cm.Text; markStyle = cm.Style; inl.RemoveAt(inl.Count - 1);
        }
        _chftn = null;
        var saved = _sink;
        var savedLabel = _pendingLabel; _pendingLabel = null;
        var savedMark = _noteMark;
        _noteMark = mark;
        _sink = new Sink();
        var s = new State { C = _defChar.Clone(), P = _defPara.Clone(), Uc = st.Uc };
        Walk(StripFirst(g, true), s);
        FlushParagraph(s, force: false);
        CloseTables(0, s);
        var blocks = _sink.Blocks;
        _sink = saved;
        _pendingLabel = savedLabel;
        _noteMark = savedMark;
        if (markStyle.VertAlign == 0 && markStyle.Position == 0) markStyle = markStyle with { VertAlign = 1 };
        if (end)
        {
            _endnotes.Add(blocks);
            _sink.Inlines.Add(new TextInline(mark, markStyle));
        }
        else _sink.Inlines.Add(new FootnoteInline(mark, markStyle, blocks));
    }

    void AppendEndnotes()
    {
        if (_endnotes.Count == 0) return;
        _sink.Blocks.Add(new Paragraph
        {
            SpaceBefore = 12, IndRight = 360, LineRule = "exact", LineValue = 2,
            BorderBottom = new BorderDef { Width = 0.5, Color = Rgb.Black }, MarkStyle = new TextStyle { Size = 1 },
        });
        foreach (var b in _endnotes) _sink.Blocks.AddRange(b);
    }

    void Field(Group g, State st)
    {
        var inst = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.Count > 1 && x.Items[0] is Sym { C: '*' } && x.Items[1] is Word { Name: "fldinst" });
        var rslt = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.Count > 0 && x.Items[0] is Word { Name: "fldrslt" });
        string instr = inst != null ? PlainText(StripFirst(inst, true)).Trim() : "";
        var tokens = FieldTokens(instr);
        string kind = tokens.Count > 0 ? tokens[0].ToUpperInvariant() : "";
        var style = Style(st.C);
        if (rslt != null && rslt.Items.OfType<Group>().FirstOrDefault() is { } rg)
        {
            // Оформление поля — по первому фрагменту результата
            var tmp = st.Clone();
            foreach (var w in rg.Items.OfType<Word>()) if (w.Name is not ("par" or "tab" or "cell")) ControlWord(w, tmp);
            style = Style(tmp.C);
        }
        switch (kind)
        {
            case "PAGE": _sink.Inlines.Add(new FieldInline(FieldKind.Page, style)); return;
            case "NUMPAGES" or "SECTIONPAGES": _sink.Inlines.Add(new FieldInline(FieldKind.NumPages, style)); return;
            case "PAGEREF" when tokens.Count > 1:
                _sink.Inlines.Add(new FieldInline(FieldKind.PageRef, tokens.Any(t => t.Equals("\\h", StringComparison.OrdinalIgnoreCase)) ? style with { Link = "#" + tokens[1] } : style)
                    { Arg = tokens[1], Cached = rslt != null ? PlainText(StripFirst(rslt, false)).Trim() : null });
                return;
            case "SYMBOL" when tokens.Count > 1:
            {
                // SYMBOL код \f "шрифт" \s размер
                int code = tokens[1].StartsWith("0x", StringComparison.OrdinalIgnoreCase) ? int.Parse(tokens[1][2..], NumberStyles.HexNumber) : int.TryParse(tokens[1], out var sc) ? sc : 32;
                int fi = tokens.FindIndex(t => t == "\\f");
                var ss = style;
                if (fi >= 0 && fi + 1 < tokens.Count) ss = ss with { FontAscii = tokens[fi + 1], FontHAnsi = tokens[fi + 1], FontEastAsia = tokens[fi + 1] };
                int si = tokens.FindIndex(t => t == "\\s");
                if (si >= 0 && si + 1 < tokens.Count && double.TryParse(tokens[si + 1], NumberStyles.Float, CultureInfo.InvariantCulture, out var sz)) ss = ss with { Size = sz };
                _sink.Inlines.Add(new TextInline(char.ConvertFromUtf32(code), ss));
                return;
            }
        }
        if (rslt == null) return;
        string? saved = _link;
        if (kind == "HYPERLINK") _link = HyperlinkTarget(tokens) ?? _link;
        else if (kind is "REF" or "NOTEREF" && tokens.Count > 1 && tokens.Any(t => t.Equals("\\h", StringComparison.OrdinalIgnoreCase))) _link = "#" + tokens[1];
        Walk(StripFirst(rslt, false), st.Clone());
        _link = saved;
    }

    static List<string> FieldTokens(string instr)
    {
        var list = new List<string>();
        int i = 0;
        while (i < instr.Length)
        {
            if (char.IsWhiteSpace(instr[i])) { i++; continue; }
            if (instr[i] == '"')
            {
                int j = instr.IndexOf('"', i + 1);
                if (j < 0) j = instr.Length;
                list.Add(instr[(i + 1)..j]);
                i = j + 1;
                continue;
            }
            int k = i;
            while (k < instr.Length && !char.IsWhiteSpace(instr[k])) k++;
            list.Add(instr[i..k]);
            i = k;
        }
        return list;
    }

    static string? HyperlinkTarget(List<string> t)
    {
        string? url = null, anchor = null;
        for (int i = 1; i < t.Count; i++)
        {
            if (t[i].Equals("\\l", StringComparison.OrdinalIgnoreCase) && i + 1 < t.Count) anchor = t[++i];
            else if (t[i].StartsWith('\\')) { if (t[i] is "\\o" or "\\O" or "\\t" or "\\T") i++; }
            else url ??= t[i].Replace("\\\\", "\\");
        }
        if (url != null) return anchor != null ? url + "#" + anchor : url;
        return anchor != null ? "#" + anchor : null;
    }
}

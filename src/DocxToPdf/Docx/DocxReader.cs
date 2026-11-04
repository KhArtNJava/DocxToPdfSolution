using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>
/// Читает пакет DOCX (ZIP + WordprocessingML) и строит вычисленную модель документа:
/// стили разрешены, нумерация посчитана, размеры переведены в пункты.
/// </summary>
public sealed class DocxReader
{
    static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    static readonly XNamespace PR = "http://schemas.openxmlformats.org/package/2006/relationships";
    static readonly XNamespace DC = "http://purl.org/dc/elements/1.1/";

    sealed class StyleDef
    {
        public string Id = "";
        public string Type = "paragraph";
        public string? BasedOn;
        public ParaProps? PPr;
        public RunProps? RPr;
        public CellBorders? TblBorders;
        public Margins? TblCellMar;
    }

    sealed class NumLevel
    {
        public int Start = 1;
        public string NumFmt = "decimal";
        public string LvlText = "";
        public string Suff = "tab";
        public ParaProps? PPr;
        public RunProps? RPr;
    }

    sealed class NumInstance
    {
        public int AbstractId;
        public Dictionary<int, int> StartOverrides = [];
        public Dictionary<int, NumLevel> LevelOverrides = [];
        public bool Used;
    }

    /// <summary>Контекст разбора конкретной части пакета (document, header, footer).</summary>
    sealed class PartContext
    {
        public Dictionary<string, (string Target, bool External)> Rels = [];
        public string BaseDir = "word/";
        public bool HeaderFooter;
        public ParaProps? TablePPr;
        public RunProps? TableRPr;
    }

    readonly ZipArchive _zip;
    readonly DocxDocument _doc = new();
    readonly Dictionary<string, StyleDef> _styles = [];
    readonly Dictionary<string, (ParaProps P, RunProps R)> _styleChainCache = [];
    readonly Dictionary<int, Dictionary<int, NumLevel>> _abstractNums = [];
    readonly Dictionary<int, NumInstance> _nums = [];
    readonly Dictionary<int, int[]> _counters = [];
    readonly Dictionary<int, int> _lastLevel = [];
    ParaProps _defaultPPr = new();
    RunProps _defaultRPr = new();
    string? _defaultParaStyle, _defaultTableStyle;
    string _majorFont = "Calibri Light", _minorFont = "Calibri";
    string? _majorEa, _minorEa;

    DocxReader(ZipArchive zip) => _zip = zip;

    public static DocxDocument Read(string path)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        return new DocxReader(zip).ReadAll();
    }

    // ================================================================
    //  Пакет
    // ================================================================

    ZipArchiveEntry? Entry(string name)
    {
        name = name.TrimStart('/');
        return _zip.GetEntry(name) ?? _zip.Entries.FirstOrDefault(e => e.FullName.Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    XDocument? LoadXml(string name)
    {
        var e = Entry(name);
        if (e == null) return null;
        using var s = e.Open();
        return XDocument.Load(s, LoadOptions.PreserveWhitespace);
    }

    byte[]? LoadBytes(string name)
    {
        var e = Entry(name);
        if (e == null) return null;
        using var s = e.Open();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    static string CombinePath(string baseDir, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = new List<string>(baseDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var seg in target.Split('/'))
        {
            if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
            else if (seg != "." && seg.Length > 0) parts.Add(seg);
        }
        return string.Join('/', parts);
    }

    Dictionary<string, (string, bool)> LoadRels(string partName)
    {
        var result = new Dictionary<string, (string, bool)>();
        int slash = partName.LastIndexOf('/');
        string dir = slash >= 0 ? partName[..(slash + 1)] : "";
        string file = slash >= 0 ? partName[(slash + 1)..] : partName;
        var x = LoadXml(dir + "_rels/" + file + ".rels");
        if (x?.Root == null) return result;
        foreach (var rel in x.Root.Elements(PR + "Relationship"))
        {
            string? id = (string?)rel.Attribute("Id");
            string? target = (string?)rel.Attribute("Target");
            if (id == null || target == null) continue;
            bool ext = string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            result[id] = (ext ? target : CombinePath(dir, target), ext);
        }
        return result;
    }

    string? FindMainDocumentPart()
    {
        var x = LoadXml("_rels/.rels");
        if (x?.Root != null)
        {
            foreach (var rel in x.Root.Elements(PR + "Relationship"))
            {
                string type = (string?)rel.Attribute("Type") ?? "";
                if (type.EndsWith("/officeDocument"))
                    return CombinePath("", (string?)rel.Attribute("Target") ?? "word/document.xml");
            }
        }
        return Entry("word/document.xml") != null ? "word/document.xml" : null;
    }

    DocxDocument ReadAll()
    {
        string mainPart = FindMainDocumentPart() ?? throw new InvalidDataException("В пакете не найден word/document.xml — это не DOCX?");
        var mainRels = LoadRels(mainPart);
        string baseDir = mainPart.Contains('/') ? mainPart[..(mainPart.LastIndexOf('/') + 1)] : "";

        string? RelTarget(string typeSuffix) =>
            LoadXml(baseDir + "_rels/" + mainPart[baseDir.Length..] + ".rels")?.Root?
                .Elements(PR + "Relationship")
                .Where(r => ((string?)r.Attribute("Type") ?? "").EndsWith(typeSuffix))
                .Select(r => CombinePath(baseDir, (string)r.Attribute("Target")!))
                .FirstOrDefault();

        LoadTheme(RelTarget("/theme") ?? baseDir + "theme/theme1.xml");
        LoadStyles(RelTarget("/styles") ?? baseDir + "styles.xml");
        LoadNumbering(RelTarget("/numbering") ?? baseDir + "numbering.xml");
        LoadSettings(RelTarget("/settings") ?? baseDir + "settings.xml");
        LoadCoreProps();

        var xdoc = LoadXml(mainPart) ?? throw new InvalidDataException("Не удалось прочитать " + mainPart);
        var body = xdoc.Root?.Element(W + "body") ?? throw new InvalidDataException("В документе нет w:body");

        var ctx = new PartContext { Rels = mainRels, BaseDir = baseDir };
        var blocks = new List<Block>();
        ParseBlockContainer(body, ctx, blocks, isBody: true);

        // Финальный sectPr
        var finalSect = body.Element(W + "sectPr");
        var finalProps = finalSect != null ? ParseSectPr(finalSect, ctx) : new SectionProps();
        _doc.Sections.Add(new Section { Props = finalProps, Blocks = blocks });

        // Разделы, завершённые абзацами с sectPr: разрезаем общий список
        SplitSections();
        InheritHeaders();
        return _doc;
    }

    void SplitSections()
    {
        var all = _doc.Sections[0];
        var result = new List<Section>();
        var cur = new List<Block>();
        foreach (var b in all.Blocks)
        {
            cur.Add(b);
            if (b is Paragraph { Section: not null } p)
            {
                result.Add(new Section { Props = p.Section, Blocks = cur });
                cur = [];
            }
        }
        result.Add(new Section { Props = all.Props, Blocks = cur });
        _doc.Sections.Clear();
        _doc.Sections.AddRange(result);
    }

    void InheritHeaders()
    {
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
        }
    }

    // ================================================================
    //  Тема, стили, нумерация, настройки
    // ================================================================

    void LoadTheme(string part)
    {
        var x = LoadXml(part);
        var fs = x?.Descendants(A + "fontScheme").FirstOrDefault();
        if (fs == null) return;
        string? Latin(string n) => (string?)fs.Element(A + n)?.Element(A + "latin")?.Attribute("typeface");
        string? Ea(string n) => (string?)fs.Element(A + n)?.Element(A + "ea")?.Attribute("typeface");
        _majorFont = NonEmpty(Latin("majorFont")) ?? _majorFont;
        _minorFont = NonEmpty(Latin("minorFont")) ?? _minorFont;
        _majorEa = NonEmpty(Ea("majorFont"));
        _minorEa = NonEmpty(Ea("minorFont"));
    }

    static string? NonEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;

    void LoadStyles(string part)
    {
        var x = LoadXml(part);
        if (x?.Root == null) return;
        var dd = x.Root.Element(W + "docDefaults");
        if (dd != null)
        {
            _defaultRPr = ParseRPr(dd.Element(W + "rPrDefault")?.Element(W + "rPr")) ?? new RunProps();
            _defaultPPr = ParsePPr(dd.Element(W + "pPrDefault")?.Element(W + "pPr")) ?? new ParaProps();
        }
        foreach (var s in x.Root.Elements(W + "style"))
        {
            var def = new StyleDef
            {
                Id = (string?)s.Attribute(W + "styleId") ?? "",
                Type = (string?)s.Attribute(W + "type") ?? "paragraph",
                BasedOn = (string?)s.Element(W + "basedOn")?.Attribute(W + "val"),
                PPr = ParsePPr(s.Element(W + "pPr")),
                RPr = ParseRPr(s.Element(W + "rPr")),
            };
            var tblPr = s.Element(W + "tblPr");
            if (tblPr != null)
            {
                def.TblBorders = ParseBorders(tblPr.Element(W + "tblBorders"));
                def.TblCellMar = ParseMargins(tblPr.Element(W + "tblCellMar"));
            }
            _styles[def.Id] = def;
            if (IsOn(s.Attribute(W + "default")))
            {
                if (def.Type == "paragraph") _defaultParaStyle = def.Id;
                else if (def.Type == "table") _defaultTableStyle = def.Id;
            }
        }
    }

    (ParaProps P, RunProps R) StyleChain(string? id, int depth = 0)
    {
        if (id == null || depth > 20 || !_styles.TryGetValue(id, out var st))
            return (new ParaProps(), new RunProps());
        if (_styleChainCache.TryGetValue(id, out var cached)) return cached;
        var (bp, br) = StyleChain(st.BasedOn, depth + 1);
        var p = bp.Clone(); p.Overlay(st.PPr);
        var r = br.Clone(); r.Overlay(st.RPr);
        p.StyleId = id;
        _styleChainCache[id] = (p, r);
        return (p, r);
    }

    (CellBorders? B, Margins? M) TableStyleChain(string? id, int depth = 0)
    {
        if (id == null || depth > 20 || !_styles.TryGetValue(id, out var st)) return (null, null);
        var (bb, bm) = TableStyleChain(st.BasedOn, depth + 1);
        CellBorders? b = null;
        if (bb != null || st.TblBorders != null) { b = new CellBorders(); b.Overlay(bb); b.Overlay(st.TblBorders); }
        Margins? m = null;
        if (bm != null || st.TblCellMar != null) { m = new Margins(); m.Overlay(bm); m.Overlay(st.TblCellMar); }
        return (b, m);
    }

    void LoadNumbering(string part)
    {
        var x = LoadXml(part);
        if (x?.Root == null) return;
        foreach (var an in x.Root.Elements(W + "abstractNum"))
        {
            int id = ParseInt((string?)an.Attribute(W + "abstractNumId")) ?? -1;
            var levels = new Dictionary<int, NumLevel>();
            foreach (var lvl in an.Elements(W + "lvl"))
            {
                int ilvl = ParseInt((string?)lvl.Attribute(W + "ilvl")) ?? 0;
                levels[ilvl] = ParseNumLevel(lvl);
            }
            // numStyleLink / styleLink — упрощённо игнорируем
            _abstractNums[id] = levels;
        }
        foreach (var num in x.Root.Elements(W + "num"))
        {
            int id = ParseInt((string?)num.Attribute(W + "numId")) ?? -1;
            var inst = new NumInstance { AbstractId = ParseInt(Val(num.Element(W + "abstractNumId"))) ?? -1 };
            foreach (var ov in num.Elements(W + "lvlOverride"))
            {
                int ilvl = ParseInt((string?)ov.Attribute(W + "ilvl")) ?? 0;
                var so = ParseInt(Val(ov.Element(W + "startOverride")));
                if (so != null) inst.StartOverrides[ilvl] = so.Value;
                var lvl = ov.Element(W + "lvl");
                if (lvl != null) inst.LevelOverrides[ilvl] = ParseNumLevel(lvl);
            }
            _nums[id] = inst;
        }
    }

    NumLevel ParseNumLevel(XElement lvl) => new()
    {
        Start = ParseInt(Val(lvl.Element(W + "start"))) ?? 1,
        NumFmt = Val(lvl.Element(W + "numFmt")) ?? "decimal",
        LvlText = Val(lvl.Element(W + "lvlText")) ?? "",
        Suff = Val(lvl.Element(W + "suff")) ?? "tab",
        PPr = ParsePPr(lvl.Element(W + "pPr")),
        RPr = ParseRPr(lvl.Element(W + "rPr")),
    };

    void LoadSettings(string part)
    {
        var x = LoadXml(part);
        if (x?.Root == null) return;
        var dts = Twips(Val(x.Root.Element(W + "defaultTabStop")));
        if (dts is > 1) _doc.DefaultTabStop = dts.Value;
        _doc.EvenAndOddHeaders = x.Root.Element(W + "evenAndOddHeaders") is { } e && IsOnVal(e);
    }

    void LoadCoreProps()
    {
        var x = LoadXml("docProps/core.xml");
        if (x?.Root == null) return;
        _doc.Title = NonEmpty(x.Root.Element(DC + "title")?.Value);
        _doc.Author = NonEmpty(x.Root.Element(DC + "creator")?.Value);
        _doc.Subject = NonEmpty(x.Root.Element(DC + "subject")?.Value);
    }

    // ================================================================
    //  Разбор свойств
    // ================================================================

    static string? Val(XElement? e) => (string?)e?.Attribute(W + "val");

    static bool IsOn(XAttribute? a) => a != null && a.Value is "1" or "true" or "on";

    static bool IsOnVal(XElement e)
    {
        var v = (string?)e.Attribute(W + "val");
        return v == null || !(v is "0" or "false" or "off" or "none");
    }

    static bool? OnOff(XElement? e) => e == null ? null : IsOnVal(e);

    static int? ParseInt(string? s)
    {
        if (s == null) return null;
        if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v)) return v;
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return (int)d;
        return null;
    }

    /// <summary>Значение в твипах (или с единицами измерения) → пункты.</summary>
    static double? Twips(string? s)
    {
        if (string.IsNullOrEmpty(s)) return null;
        s = s.Trim();
        double factor = 1 / 20.0;
        string num = s;
        if (s.EndsWith("pt")) { factor = 1; num = s[..^2]; }
        else if (s.EndsWith("in")) { factor = 72; num = s[..^2]; }
        else if (s.EndsWith("cm")) { factor = 72 / 2.54; num = s[..^2]; }
        else if (s.EndsWith("mm")) { factor = 72 / 25.4; num = s[..^2]; }
        else if (s.EndsWith("pc") || s.EndsWith("pi")) { factor = 12; num = s[..^2]; }
        return double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v * factor : null;
    }

    static double? HalfPoints(string? s) =>
        double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 2 : null;

    RunProps? ParseRPr(XElement? rPr)
    {
        if (rPr == null) return null;
        var r = new RunProps();
        var fonts = rPr.Element(W + "rFonts");
        if (fonts != null)
        {
            r.FontAscii = (string?)fonts.Attribute(W + "ascii");
            r.FontHAnsi = (string?)fonts.Attribute(W + "hAnsi");
            r.FontEastAsia = (string?)fonts.Attribute(W + "eastAsia");
            r.ThemeAscii = (string?)fonts.Attribute(W + "asciiTheme");
            r.ThemeHAnsi = (string?)fonts.Attribute(W + "hAnsiTheme");
            r.ThemeEastAsia = (string?)fonts.Attribute(W + "eastAsiaTheme");
        }
        r.Bold = OnOff(rPr.Element(W + "b"));
        r.Italic = OnOff(rPr.Element(W + "i"));
        r.Strike = OnOff(rPr.Element(W + "strike"));
        r.DStrike = OnOff(rPr.Element(W + "dstrike"));
        r.Caps = OnOff(rPr.Element(W + "caps"));
        r.SmallCaps = OnOff(rPr.Element(W + "smallCaps"));
        r.Vanish = OnOff(rPr.Element(W + "vanish"));
        var u = rPr.Element(W + "u");
        if (u != null) r.Underline = Val(u) ?? "single";
        r.Size = HalfPoints(Val(rPr.Element(W + "sz")));
        r.Color = Val(rPr.Element(W + "color"));
        r.Highlight = Val(rPr.Element(W + "highlight"));
        var shd = rPr.Element(W + "shd");
        if (shd != null) r.Shading = (string?)shd.Attribute(W + "fill");
        r.VertAlign = Val(rPr.Element(W + "vertAlign"));
        r.StyleId = Val(rPr.Element(W + "rStyle"));
        r.Position = HalfPoints(Val(rPr.Element(W + "position")));
        r.Spacing = Twips(Val(rPr.Element(W + "spacing")));
        return r;
    }

    ParaProps? ParsePPr(XElement? pPr)
    {
        if (pPr == null) return null;
        var p = new ParaProps
        {
            StyleId = Val(pPr.Element(W + "pStyle")),
            Jc = Val(pPr.Element(W + "jc")),
        };
        var ind = pPr.Element(W + "ind");
        if (ind != null)
        {
            p.IndLeft = Twips((string?)ind.Attribute(W + "left") ?? (string?)ind.Attribute(W + "start"));
            p.IndRight = Twips((string?)ind.Attribute(W + "right") ?? (string?)ind.Attribute(W + "end"));
            p.IndFirstLine = Twips((string?)ind.Attribute(W + "firstLine"));
            p.IndHanging = Twips((string?)ind.Attribute(W + "hanging"));
            if (p.IndHanging != null) p.IndFirstLine = null;
        }
        var sp = pPr.Element(W + "spacing");
        if (sp != null)
        {
            p.SpBefore = Twips((string?)sp.Attribute(W + "before"));
            p.SpAfter = Twips((string?)sp.Attribute(W + "after"));
            if (sp.Attribute(W + "beforeAutospacing") is { } ba) p.BeforeAuto = IsOn(ba);
            if (sp.Attribute(W + "afterAutospacing") is { } aa) p.AfterAuto = IsOn(aa);
            string? line = (string?)sp.Attribute(W + "line");
            if (line != null)
            {
                p.LineRule = (string?)sp.Attribute(W + "lineRule") ?? "auto";
                p.Line = p.LineRule == "auto"
                    ? (double.TryParse(line, NumberStyles.Float, CultureInfo.InvariantCulture, out var lv) ? lv : 240)
                    : Twips(line);
            }
        }
        var numPr = pPr.Element(W + "numPr");
        if (numPr != null)
        {
            p.NumId = ParseInt(Val(numPr.Element(W + "numId")));
            p.Ilvl = ParseInt(Val(numPr.Element(W + "ilvl")));
        }
        p.PageBreakBefore = OnOff(pPr.Element(W + "pageBreakBefore"));
        p.KeepNext = OnOff(pPr.Element(W + "keepNext"));
        p.KeepLines = OnOff(pPr.Element(W + "keepLines"));
        p.ContextualSpacing = OnOff(pPr.Element(W + "contextualSpacing"));
        p.WidowControl = OnOff(pPr.Element(W + "widowControl"));
        var tabs = pPr.Element(W + "tabs");
        if (tabs != null)
        {
            p.Tabs = [];
            foreach (var t in tabs.Elements(W + "tab"))
            {
                double? pos = Twips((string?)t.Attribute(W + "pos"));
                if (pos == null) continue;
                p.Tabs.Add(new TabStop(pos.Value, Val(t) ?? "left", (string?)t.Attribute(W + "leader") ?? "none"));
            }
        }
        var shd = pPr.Element(W + "shd");
        if (shd != null) p.Shading = (string?)shd.Attribute(W + "fill");
        var bdr = pPr.Element(W + "pBdr");
        if (bdr != null)
        {
            p.BorderTop = ParseBorder(bdr.Element(W + "top"));
            p.BorderBottom = ParseBorder(bdr.Element(W + "bottom"));
        }
        p.MarkRunProps = ParseRPr(pPr.Element(W + "rPr"));
        return p;
    }

    static BorderDef? ParseBorder(XElement? e)
    {
        if (e == null) return null;
        var b = new BorderDef { Val = Val(e) ?? "single" };
        if (double.TryParse((string?)e.Attribute(W + "sz"), NumberStyles.Float, CultureInfo.InvariantCulture, out var sz))
            b.Width = Math.Max(0.25, sz / 8.0);
        if (double.TryParse((string?)e.Attribute(W + "space"), NumberStyles.Float, CultureInfo.InvariantCulture, out var sp))
            b.Space = sp;
        b.Color = Rgb.Parse((string?)e.Attribute(W + "color")) ?? Rgb.Black;
        if (b.Val is "double") b.Width = Math.Max(b.Width, 1.5);
        return b;
    }

    static CellBorders? ParseBorders(XElement? e)
    {
        if (e == null) return null;
        return new CellBorders
        {
            Top = ParseBorder(e.Element(W + "top")),
            Bottom = ParseBorder(e.Element(W + "bottom")),
            Left = ParseBorder(e.Element(W + "left") ?? e.Element(W + "start")),
            Right = ParseBorder(e.Element(W + "right") ?? e.Element(W + "end")),
            InsideH = ParseBorder(e.Element(W + "insideH")),
            InsideV = ParseBorder(e.Element(W + "insideV")),
        };
    }

    static Margins? ParseMargins(XElement? e)
    {
        if (e == null) return null;
        double? M(params string[] names)
        {
            foreach (var n in names)
            {
                var el = e.Element(W + n);
                if (el != null)
                {
                    string? type = (string?)el.Attribute(W + "type");
                    if (type is "nil") return 0;
                    return Twips((string?)el.Attribute(W + "w"));
                }
            }
            return null;
        }
        return new Margins { Top = M("top"), Bottom = M("bottom"), Left = M("left", "start"), Right = M("right", "end") };
    }

    // ================================================================
    //  Вычисление итоговых стилей
    // ================================================================

    string ThemeFont(string? theme, string? fallback)
    {
        if (theme == null) return fallback ?? _minorFont;
        if (theme.StartsWith("major"))
            return theme.Contains("EastAsia") ? (_majorEa ?? _majorFont) : _majorFont;
        return theme.Contains("EastAsia") ? (_minorEa ?? _minorFont) : _minorFont;
    }

    TextStyle ToTextStyle(RunProps r, string? link = null)
    {
        string ascii = r.ThemeAscii != null ? ThemeFont(r.ThemeAscii, null) : r.FontAscii ?? r.FontHAnsi ?? "Times New Roman";
        string hansi = r.ThemeHAnsi != null ? ThemeFont(r.ThemeHAnsi, null) : r.FontHAnsi ?? ascii;
        string ea = r.ThemeEastAsia != null ? ThemeFont(r.ThemeEastAsia, null) : r.FontEastAsia ?? hansi;
        Rgb? bg = Rgb.FromHighlight(r.Highlight) ?? Rgb.Parse(r.Shading);
        return new TextStyle
        {
            FontAscii = ascii,
            FontHAnsi = hansi,
            FontEastAsia = ea,
            Bold = r.Bold ?? false,
            Italic = r.Italic ?? false,
            Underline = r.Underline != null && r.Underline != "none",
            DoubleUnderline = r.Underline == "double",
            Strike = (r.Strike ?? false) || (r.DStrike ?? false),
            Caps = r.Caps ?? false,
            SmallCaps = r.SmallCaps ?? false,
            LetterSpacing = r.Spacing ?? 0,
            Size = r.Size ?? 10,
            Color = Rgb.Parse(r.Color) ?? (link != null && r.Color == null ? new Rgb(0.02, 0.39, 0.76) : Rgb.Black),
            Background = bg,
            VertAlign = r.VertAlign switch { "superscript" => 1, "subscript" => -1, _ => 0 },
            Position = r.Position ?? 0,
            Link = link,
        };
    }

    /// <summary>Базовые свойства абзаца: умолчания → стиль таблицы → стиль абзаца.</summary>
    (ParaProps P, RunProps R) BaseProps(string? styleId, PartContext ctx)
    {
        var p = _defaultPPr.Clone();
        var r = _defaultRPr.Clone();
        p.Overlay(ctx.TablePPr);
        r.Overlay(ctx.TableRPr);
        var (sp, sr) = StyleChain(styleId ?? _defaultParaStyle);
        p.Overlay(sp);
        r.Overlay(sr);
        return (p, r);
    }

    RunProps ResolveRunProps(RunProps paraBase, RunProps? direct)
    {
        var r = paraBase.Clone();
        string? cs = direct?.StyleId;
        if (cs != null) r.Overlay(StyleChain(cs).R);
        r.Overlay(direct);
        return r;
    }

    // ================================================================
    //  Блоки
    // ================================================================

    void ParseBlockContainer(XElement container, PartContext ctx, List<Block> output, bool isBody = false)
    {
        foreach (var el in container.Elements())
        {
            if (el.Name == W + "p") output.Add(ParseParagraph(el, ctx));
            else if (el.Name == W + "tbl") output.Add(ParseTable(el, ctx));
            else if (el.Name == W + "sdt") { var c = el.Element(W + "sdtContent"); if (c != null) ParseBlockContainer(c, ctx, output); }
            else if (el.Name == W + "customXml" || el.Name == W + "ins" || el.Name == W + "moveTo") ParseBlockContainer(el, ctx, output);
            else if (el.Name == MC + "AlternateContent")
            {
                var alt = el.Element(MC + "Choice") ?? el.Element(MC + "Fallback");
                if (alt != null) ParseBlockContainer(alt, ctx, output);
            }
        }
    }

    Paragraph ParseParagraph(XElement p, PartContext ctx)
    {
        var pPrEl = p.Element(W + "pPr");
        var direct = ParsePPr(pPrEl);
        string? styleId = direct?.StyleId;
        if (styleId != null && !_styles.ContainsKey(styleId)) styleId = null;
        var (pp, baseR) = BaseProps(styleId, ctx);

        // Нумерация (уровень вставляется между стилем и прямыми свойствами)
        int? numId = direct?.NumId ?? pp.NumId;
        int ilvl = direct?.Ilvl ?? pp.Ilvl ?? 0;
        NumLevel? lvl = null;
        if (numId is > 0) lvl = GetLevel(numId.Value, ilvl);
        if (lvl?.PPr != null) pp.Overlay(StripNum(lvl.PPr));
        pp.Overlay(direct);

        var para = new Paragraph
        {
            StyleId = styleId ?? _defaultParaStyle ?? "",
            Align = pp.Jc switch
            {
                "center" => "center",
                "right" or "end" => "right",
                "both" or "distribute" or "lowKashida" or "mediumKashida" or "highKashida" or "thaiDistribute" => "both",
                _ => "left"
            },
            IndLeft = pp.IndLeft ?? 0,
            IndRight = pp.IndRight ?? 0,
            IndFirst = pp.IndHanging != null ? -pp.IndHanging.Value : pp.IndFirstLine ?? 0,
            SpaceBefore = pp.BeforeAuto == true ? 14 : pp.SpBefore ?? 0,
            SpaceAfter = pp.AfterAuto == true ? 14 : pp.SpAfter ?? 0,
            LineValue = pp.Line ?? 240,
            LineRule = pp.LineRule ?? "auto",
            PageBreakBefore = pp.PageBreakBefore ?? false,
            KeepNext = pp.KeepNext ?? false,
            KeepLines = pp.KeepLines ?? false,
            ContextualSpacing = pp.ContextualSpacing ?? false,
            WidowControl = pp.WidowControl ?? true,
            Tabs = (pp.Tabs ?? []).Where(t => t.Align != "clear").OrderBy(t => t.Pos).ToList(),
            Shading = Rgb.Parse(pp.Shading),
            BorderTop = pp.BorderTop is { IsNone: false } ? pp.BorderTop : null,
            BorderBottom = pp.BorderBottom is { IsNone: false } ? pp.BorderBottom : null,
        };

        var markR = baseR.Clone();
        markR.Overlay(pp.MarkRunProps);
        para.MarkStyle = ToTextStyle(markR);

        // Метка списка
        if (lvl != null && numId != null)
        {
            string label = NextNumberLabel(numId.Value, ilvl);
            if (label.Length > 0 || lvl.NumFmt == "none")
            {
                var lr = markR.Clone();
                lr.Underline = null;
                lr.Overlay(lvl.RPr);
                var ls = ToTextStyle(lr);
                if (label.Length > 0) para.Inlines.Add(new TextInline(label, ls));
                if (lvl.Suff == "tab") para.Inlines.Add(new TabInline(ls));
                else if (lvl.Suff == "space") para.Inlines.Add(new TextInline(" ", ls));
            }
        }

        var state = new RunState { BaseR = baseR, Ctx = ctx };
        ParseInlineContainer(p, state, para.Inlines);

        if (pPrEl?.Element(W + "sectPr") is { } sp && !ctx.HeaderFooter)
            para.Section = ParseSectPr(sp, ctx);
        return para;
    }

    static ParaProps StripNum(ParaProps p)
    {
        var c = p.Clone();
        c.NumId = null; c.Ilvl = null; c.StyleId = null;
        return c;
    }

    // ---------------- Нумерация ----------------

    NumLevel? GetLevel(int numId, int ilvl)
    {
        if (!_nums.TryGetValue(numId, out var inst)) return null;
        if (inst.LevelOverrides.TryGetValue(ilvl, out var ov)) return ov;
        if (!_abstractNums.TryGetValue(inst.AbstractId, out var levels)) return null;
        return levels.GetValueOrDefault(ilvl);
    }

    string NextNumberLabel(int numId, int ilvl)
    {
        var inst = _nums[numId];
        int key = inst.AbstractId;
        if (!_counters.TryGetValue(key, out var counters))
        {
            counters = new int[10];
            for (int i = 0; i < 10; i++) counters[i] = (GetLevel(numId, i)?.Start ?? 1) - 1;
            _counters[key] = counters;
        }
        if (!inst.Used)
        {
            inst.Used = true;
            foreach (var (l, start) in inst.StartOverrides)
                if (l is >= 0 and < 10) counters[l] = start - 1;
        }
        ilvl = Math.Clamp(ilvl, 0, 9);
        counters[ilvl]++;
        for (int i = ilvl + 1; i < 10; i++) counters[i] = (GetLevel(numId, i)?.Start ?? 1) - 1;
        _lastLevel[key] = ilvl;

        var lvl = GetLevel(numId, ilvl)!;
        if (lvl.NumFmt == "none") return "";
        if (lvl.NumFmt == "bullet") return lvl.LvlText;
        var sb = new StringBuilder();
        string t = lvl.LvlText;
        for (int i = 0; i < t.Length; i++)
        {
            if (t[i] == '%' && i + 1 < t.Length && char.IsDigit(t[i + 1]))
            {
                int l = t[i + 1] - '1';
                i++;
                if (l < 0 || l > 9) continue;
                var ll = GetLevel(numId, l);
                int value = Math.Max(counters[l], l == ilvl ? 1 : (ll?.Start ?? 1));
                string fmt = ll?.NumFmt ?? "decimal";
                // isLgl и "legal" формат — всегда арабские цифры
                sb.Append(FormatNumber(value, l == ilvl ? fmt : (fmt == "bullet" ? "decimal" : fmt)));
            }
            else sb.Append(t[i]);
        }
        return sb.ToString();
    }

    static string FormatNumber(int n, string fmt) => fmt switch
    {
        "lowerLetter" => Letters(n, false),
        "upperLetter" => Letters(n, true),
        "lowerRoman" => Roman(n).ToLowerInvariant(),
        "upperRoman" => Roman(n),
        "decimalZero" => n < 10 ? "0" + n : n.ToString(CultureInfo.InvariantCulture),
        "russianLower" => RussianLetters(n, false),
        "russianUpper" => RussianLetters(n, true),
        "ordinal" => n + ".",
        _ => n.ToString(CultureInfo.InvariantCulture)
    };

    static string Letters(int n, bool upper)
    {
        if (n <= 0) return "";
        char c = (char)((upper ? 'A' : 'a') + (n - 1) % 26);
        return new string(c, (n - 1) / 26 + 1);
    }

    static string RussianLetters(int n, bool upper)
    {
        const string abc = "абвгдежзиклмнопрстуфхцчшщэюя";
        if (n <= 0) return "";
        char c = abc[(n - 1) % abc.Length];
        if (upper) c = char.ToUpperInvariant(c);
        return new string(c, (n - 1) / abc.Length + 1);
    }

    static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString(CultureInfo.InvariantCulture);
        int[] v = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] s = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        var sb = new StringBuilder();
        for (int i = 0; i < v.Length; i++)
            while (n >= v[i]) { sb.Append(s[i]); n -= v[i]; }
        return sb.ToString();
    }

    // ---------------- Инлайны ----------------

    sealed class FieldFrame
    {
        public StringBuilder Instr = new();
        public bool InResult;
        public bool SkipResult;
        public bool Emitted;
        public TextStyle? Style;
    }

    sealed class RunState
    {
        public required RunProps BaseR;
        public required PartContext Ctx;
        public string? Link;
        public Stack<FieldFrame> Fields = new();

        public bool Suppressed => Fields.Any(f => !f.InResult || f.SkipResult);
    }

    void ParseInlineContainer(XElement container, RunState st, List<Inline> output)
    {
        foreach (var el in container.Elements())
        {
            var n = el.Name;
            if (n == W + "r") ParseRun(el, st, output);
            else if (n == W + "hyperlink")
            {
                string? prev = st.Link;
                string? rid = (string?)el.Attribute(R + "id");
                if (rid != null && st.Ctx.Rels.TryGetValue(rid, out var rel) && rel.External) st.Link = rel.Target;
                ParseInlineContainer(el, st, output);
                st.Link = prev;
            }
            else if (n == W + "fldSimple")
            {
                string instr = ((string?)el.Attribute(W + "instr") ?? "").Trim();
                var kind = FieldKindOf(instr);
                if (kind != null && st.Ctx.HeaderFooter)
                {
                    var r = el.Descendants(W + "rPr").FirstOrDefault();
                    output.Add(new FieldInline(kind.Value, ToTextStyle(ResolveRunProps(st.BaseR, ParseRPr(r)))));
                }
                else ParseInlineContainer(el, st, output);
            }
            else if (n == W + "ins" || n == W + "moveTo" || n == W + "smartTag" || n == W + "customXml" || n == W + "dir" || n == W + "bdo")
                ParseInlineContainer(el, st, output);
            else if (n == W + "sdt")
            {
                var c = el.Element(W + "sdtContent");
                if (c != null) ParseInlineContainer(c, st, output);
            }
            else if (n == MC + "AlternateContent")
            {
                var alt = el.Element(MC + "Choice") ?? el.Element(MC + "Fallback");
                if (alt != null) ParseInlineContainer(alt, st, output);
            }
            else if (n == M + "oMathPara" || n == M + "oMath")
            {
                // Формулы: выводим только текст
                string text = string.Concat(el.Descendants(M + "t").Select(t => t.Value));
                if (text.Length > 0 && !st.Suppressed)
                {
                    var rr = ResolveRunProps(st.BaseR, null);
                    rr.Italic = true;
                    output.Add(new TextInline(text, ToTextStyle(rr, st.Link)));
                }
            }
        }
    }

    static FieldKind? FieldKindOf(string instr)
    {
        var word = instr.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.ToUpperInvariant();
        return word switch { "PAGE" => FieldKind.Page, "NUMPAGES" or "SECTIONPAGES" => FieldKind.NumPages, _ => null };
    }

    void ParseRun(XElement r, RunState st, List<Inline> output)
    {
        var rp = ResolveRunProps(st.BaseR, ParseRPr(r.Element(W + "rPr")));
        var style = ToTextStyle(rp, st.Link);
        bool hidden = rp.Vanish == true;

        foreach (var el in r.Elements())
        {
            var n = el.Name;
            if (n == W + "fldChar")
            {
                string type = (string?)el.Attribute(W + "fldCharType") ?? "";
                if (type == "begin") st.Fields.Push(new FieldFrame { Style = style });
                else if (type == "separate" && st.Fields.Count > 0)
                {
                    var f = st.Fields.Peek();
                    f.InResult = true;
                    var kind = FieldKindOf(f.Instr.ToString().Trim());
                    if (kind != null && st.Ctx.HeaderFooter && !OuterSuppressed(st))
                    {
                        output.Add(new FieldInline(kind.Value, f.Style ?? style));
                        f.SkipResult = true;
                        f.Emitted = true;
                    }
                }
                else if (type == "end" && st.Fields.Count > 0)
                {
                    var f = st.Fields.Pop();
                    if (!f.InResult && !f.Emitted && st.Ctx.HeaderFooter && !st.Suppressed)
                    {
                        var kind = FieldKindOf(f.Instr.ToString().Trim());
                        if (kind != null) output.Add(new FieldInline(kind.Value, f.Style ?? style));
                    }
                }
                continue;
            }
            if (n == W + "instrText")
            {
                if (st.Fields.Count > 0 && !st.Fields.Peek().InResult) st.Fields.Peek().Instr.Append(el.Value);
                continue;
            }
            if (st.Suppressed || hidden) continue;

            if (n == W + "t") { if (el.Value.Length > 0) output.Add(new TextInline(el.Value, style)); }
            else if (n == W + "tab") output.Add(new TabInline(style));
            else if (n == W + "br")
            {
                string? type = (string?)el.Attribute(W + "type");
                output.Add(new BreakInline(type switch { "page" => BreakKind.Page, "column" => BreakKind.Column, _ => BreakKind.Line }, style));
            }
            else if (n == W + "cr") output.Add(new BreakInline(BreakKind.Line, style));
            else if (n == W + "noBreakHyphen") output.Add(new TextInline("-", style));
            else if (n == W + "sym")
            {
                string? font = (string?)el.Attribute(W + "font");
                string? ch = (string?)el.Attribute(W + "char");
                if (ch != null && int.TryParse(ch, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code))
                {
                    if (code < 0x100 && font != null) code += 0xF000;
                    var s2 = font != null ? style with { FontAscii = font, FontHAnsi = font, FontEastAsia = font } : style;
                    output.Add(new TextInline(char.ConvertFromUtf32(code), s2));
                }
            }
            else if (n == W + "drawing") ParseDrawing(el, st, style, output);
            else if (n == W + "pict" || n == W + "object") ParseVml(el, st, style, output);
            else if (n == MC + "AlternateContent")
            {
                var alt = el.Element(MC + "Choice") ?? el.Element(MC + "Fallback");
                if (alt == null) continue;
                foreach (var d in alt.Elements())
                {
                    if (d.Name == W + "drawing") ParseDrawing(d, st, style, output);
                    else if (d.Name == W + "pict") ParseVml(d, st, style, output);
                }
            }
        }
    }

    static bool OuterSuppressed(RunState st) => st.Fields.Skip(1).Any(f => !f.InResult || f.SkipResult);

    void ParseDrawing(XElement drawing, RunState st, TextStyle style, List<Inline> output)
    {
        foreach (var container in drawing.Elements())
        {
            if (container.Name != WP + "inline" && container.Name != WP + "anchor") continue;
            var extent = container.Element(WP + "extent");
            double cx = (double?)extent?.Attribute("cx") ?? 0;
            double cy = (double?)extent?.Attribute("cy") ?? 0;
            var blip = container.Descendants(A + "blip").FirstOrDefault();
            string? rid = (string?)blip?.Attribute(R + "embed") ?? (string?)blip?.Attribute(R + "link");
            if (rid == null) continue;
            AddImage(rid, cx / 12700.0, cy / 12700.0, st, style, output);
        }
    }

    void ParseVml(XElement pict, RunState st, TextStyle style, List<Inline> output)
    {
        foreach (var img in pict.Descendants(V + "imagedata"))
        {
            string? rid = (string?)img.Attribute(R + "id");
            if (rid == null) continue;
            var shape = img.Parent;
            string css = (string?)shape?.Attribute("style") ?? "";
            double w = CssLength(css, "width"), h = CssLength(css, "height");
            AddImage(rid, w, h, st, style, output);
        }
    }

    static double CssLength(string css, string prop)
    {
        foreach (var part in css.Split(';'))
        {
            var kv = part.Split(':', 2);
            if (kv.Length != 2 || !kv[0].Trim().Equals(prop, StringComparison.OrdinalIgnoreCase)) continue;
            string v = kv[1].Trim();
            if (v.EndsWith("px") && double.TryParse(v[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var px))
                return px * 0.75;
            if (v.Any(char.IsLetter)) return Twips(v) ?? 0;
            return double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n * 0.75 : 0;
        }
        return 0;
    }

    void AddImage(string rid, double w, double h, RunState st, TextStyle style, List<Inline> output)
    {
        if (!st.Ctx.Rels.TryGetValue(rid, out var rel) || rel.External) return;
        if (!_doc.Media.ContainsKey(rel.Target))
        {
            var bytes = LoadBytes(rel.Target);
            if (bytes == null) return;
            _doc.Media[rel.Target] = bytes;
        }
        output.Add(new ImageInline(rel.Target, w, h, style));
    }

    // ---------------- Таблицы ----------------

    Table ParseTable(XElement tbl, PartContext ctx)
    {
        var t = new Table();
        var tblPr = tbl.Element(W + "tblPr");
        string? styleId = Val(tblPr?.Element(W + "tblStyle")) ?? _defaultTableStyle;

        var (sb, sm) = TableStyleChain(_defaultTableStyle);
        t.CellMargins.Overlay(sm);
        t.Borders.Overlay(sb);
        if (styleId != _defaultTableStyle)
        {
            (sb, sm) = TableStyleChain(styleId);
            t.Borders.Overlay(sb);
            t.CellMargins.Overlay(sm);
        }
        if (tblPr != null)
        {
            t.Borders.Overlay(ParseBorders(tblPr.Element(W + "tblBorders")));
            t.CellMargins.Overlay(ParseMargins(tblPr.Element(W + "tblCellMar")));
            var ind = tblPr.Element(W + "tblInd");
            if (ind != null && (string?)ind.Attribute(W + "type") != "pct") t.Indent = Twips((string?)ind.Attribute(W + "w")) ?? 0;
            t.Align = Val(tblPr.Element(W + "jc")) switch { "center" => "center", "right" or "end" => "right", _ => "left" };
            var tw = tblPr.Element(W + "tblW");
            if (tw != null && (string?)tw.Attribute(W + "type") == "dxa") t.Width = Twips((string?)tw.Attribute(W + "w"));
        }
        foreach (var gc in tbl.Element(W + "tblGrid")?.Elements(W + "gridCol") ?? [])
            t.Grid.Add(Twips((string?)gc.Attribute(W + "w")) ?? 0);

        // Свойства абзацев внутри таблицы из стиля таблицы
        var cellCtx = ctx;
        if (styleId != null && _styles.ContainsKey(styleId))
        {
            var (tp, tr) = StyleChain(styleId);
            cellCtx = new PartContext
            {
                Rels = ctx.Rels, BaseDir = ctx.BaseDir, HeaderFooter = ctx.HeaderFooter,
                TablePPr = tp, TableRPr = tr
            };
        }

        foreach (var trEl in RowsOf(tbl))
        {
            var row = new TableRow();
            var trPr = trEl.Element(W + "trPr");
            if (trPr != null)
            {
                var h = trPr.Element(W + "trHeight");
                if (h != null)
                {
                    row.Height = Twips(Val(h)) ?? 0;
                    row.HeightRule = (string?)h.Attribute(W + "hRule") ?? "atLeast";
                }
                row.IsHeader = trPr.Element(W + "tblHeader") is { } th && IsOnVal(th);
                row.CantSplit = trPr.Element(W + "cantSplit") is { } cs && IsOnVal(cs);
                int before = ParseInt(Val(trPr.Element(W + "gridBefore"))) ?? 0;
                if (before > 0)
                    row.Cells.Add(new TableCell
                    {
                        GridSpan = before,
                        Borders = new CellBorders { Top = new BorderDef { Val = "nil" }, Bottom = new BorderDef { Val = "nil" }, Left = new BorderDef { Val = "nil" } }
                    });
            }
            foreach (var tcEl in CellsOf(trEl))
            {
                var cell = new TableCell();
                var tcPr = tcEl.Element(W + "tcPr");
                if (tcPr != null)
                {
                    cell.GridSpan = Math.Max(1, ParseInt(Val(tcPr.Element(W + "gridSpan"))) ?? 1);
                    var vm = tcPr.Element(W + "vMerge");
                    if (vm != null) cell.VMerge = Val(vm) == "restart" ? VMerge.Restart : VMerge.Continue;
                    if (tcPr.Element(W + "hMerge") is { } hm && Val(hm) != "restart") cell.GridSpan = 0; // склеено с предыдущей
                    cell.Borders = ParseBorders(tcPr.Element(W + "tcBorders"));
                    cell.Shading = Rgb.Parse((string?)tcPr.Element(W + "shd")?.Attribute(W + "fill"));
                    cell.Margins = ParseMargins(tcPr.Element(W + "tcMar"));
                    cell.VAlign = Val(tcPr.Element(W + "vAlign")) ?? "top";
                    var tcW = tcPr.Element(W + "tcW");
                    if (tcW != null && (string?)tcW.Attribute(W + "type") is "dxa" or null) cell.Width = Twips((string?)tcW.Attribute(W + "w"));
                }
                ParseBlockContainer(tcEl, cellCtx, cell.Blocks);
                if (cell.GridSpan == 0 && row.Cells.Count > 0) { row.Cells[^1].GridSpan++; continue; }
                cell.GridSpan = Math.Max(1, cell.GridSpan);
                row.Cells.Add(cell);
            }
            t.Rows.Add(row);
        }

        // Если сетки нет — строим её по ширинам ячеек первой строки
        int cols = t.Rows.Count == 0 ? 0 : t.Rows.Max(r => r.Cells.Sum(c => c.GridSpan));
        if (t.Grid.Count < cols || t.Grid.All(g => g <= 0))
        {
            t.Grid.Clear();
            var widest = t.Rows.OrderByDescending(r => r.Cells.Count).FirstOrDefault();
            for (int i = 0; i < cols; i++) t.Grid.Add(0);
            if (widest != null)
            {
                int gi = 0;
                foreach (var c in widest.Cells)
                {
                    double w = (c.Width ?? 72) / c.GridSpan;
                    for (int k = 0; k < c.GridSpan && gi < cols; k++) t.Grid[gi++] = w;
                }
            }
            for (int i = 0; i < cols; i++) if (t.Grid[i] <= 0) t.Grid[i] = 72;
        }
        return t;
    }

    static IEnumerable<XElement> RowsOf(XElement tbl)
    {
        foreach (var el in tbl.Elements())
        {
            if (el.Name == W + "tr") yield return el;
            else if (el.Name == W + "sdt" || el.Name == W + "customXml")
            {
                var c = el.Name == W + "sdt" ? el.Element(W + "sdtContent") : el;
                if (c != null) foreach (var r in RowsOf(c)) yield return r;
            }
        }
    }

    static IEnumerable<XElement> CellsOf(XElement tr)
    {
        foreach (var el in tr.Elements())
        {
            if (el.Name == W + "tc") yield return el;
            else if (el.Name == W + "sdt" || el.Name == W + "customXml")
            {
                var c = el.Name == W + "sdt" ? el.Element(W + "sdtContent") : el;
                if (c != null) foreach (var r in CellsOf(c)) yield return r;
            }
        }
    }

    // ---------------- Разделы, колонтитулы ----------------

    SectionProps ParseSectPr(XElement sp, PartContext ctx)
    {
        var s = new SectionProps();
        var pgSz = sp.Element(W + "pgSz");
        if (pgSz != null)
        {
            s.PageWidth = Twips((string?)pgSz.Attribute(W + "w")) ?? s.PageWidth;
            s.PageHeight = Twips((string?)pgSz.Attribute(W + "h")) ?? s.PageHeight;
            if ((string?)pgSz.Attribute(W + "orient") == "landscape" && s.PageWidth < s.PageHeight)
                (s.PageWidth, s.PageHeight) = (s.PageHeight, s.PageWidth);
        }
        var m = sp.Element(W + "pgMar");
        if (m != null)
        {
            s.MarginTop = Math.Abs(Twips((string?)m.Attribute(W + "top")) ?? s.MarginTop);
            s.MarginBottom = Math.Abs(Twips((string?)m.Attribute(W + "bottom")) ?? s.MarginBottom);
            s.MarginLeft = Twips((string?)m.Attribute(W + "left") ?? (string?)m.Attribute(W + "start")) ?? s.MarginLeft;
            s.MarginRight = Twips((string?)m.Attribute(W + "right") ?? (string?)m.Attribute(W + "end")) ?? s.MarginRight;
            s.HeaderDistance = Twips((string?)m.Attribute(W + "header")) ?? s.HeaderDistance;
            s.FooterDistance = Twips((string?)m.Attribute(W + "footer")) ?? s.FooterDistance;
            s.MarginLeft += Twips((string?)m.Attribute(W + "gutter")) ?? 0;
        }
        s.Type = Val(sp.Element(W + "type")) ?? "nextPage";
        s.TitlePage = sp.Element(W + "titlePg") is { } tp && IsOnVal(tp);
        s.PageNumberStart = ParseInt((string?)sp.Element(W + "pgNumType")?.Attribute(W + "start"));

        foreach (var hr in sp.Elements(W + "headerReference"))
        {
            var hf = LoadHeaderFooter((string?)hr.Attribute(R + "id"), ctx);
            switch ((string?)hr.Attribute(W + "type") ?? "default")
            {
                case "first": s.HeaderFirst = hf; break;
                case "even": s.HeaderEven = hf; break;
                default: s.HeaderDefault = hf; break;
            }
        }
        foreach (var fr in sp.Elements(W + "footerReference"))
        {
            var hf = LoadHeaderFooter((string?)fr.Attribute(R + "id"), ctx);
            switch ((string?)fr.Attribute(W + "type") ?? "default")
            {
                case "first": s.FooterFirst = hf; break;
                case "even": s.FooterEven = hf; break;
                default: s.FooterDefault = hf; break;
            }
        }
        return s;
    }

    readonly Dictionary<string, HeaderFooter> _hfCache = [];

    HeaderFooter? LoadHeaderFooter(string? rid, PartContext ctx)
    {
        if (rid == null || !ctx.Rels.TryGetValue(rid, out var rel) || rel.External) return null;
        if (_hfCache.TryGetValue(rel.Target, out var cached)) return cached;
        var x = LoadXml(rel.Target);
        if (x?.Root == null) return null;
        var hctx = new PartContext { Rels = LoadRels(rel.Target), HeaderFooter = true };
        var hf = new HeaderFooter();
        ParseBlockContainer(x.Root, hctx, hf.Blocks);
        _hfCache[rel.Target] = hf;
        return hf;
    }
}

using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>
/// Читает пакет DOCX (ZIP + WordprocessingML) и строит вычисленную модель документа:
/// стили разрешены, нумерация посчитана, размеры переведены в пункты.
/// </summary>
public sealed partial class DocxReader
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
        public Dictionary<string, CondFormat> Cond = [];
        public int RowBand = 1, ColBand = 1;
    }

    /// <summary>Условное форматирование табличного стиля (tblStylePr): шапка, полосы, первая колонка…</summary>
    sealed class CondFormat
    {
        public ParaProps? PPr;
        public RunProps? RPr;
        public string? Shading;
        public CellBorders? Borders;

        public CondFormat Merge(CondFormat? over)
        {
            if (over == null) return this;
            var r = new CondFormat { PPr = PPr?.Clone() ?? new ParaProps(), RPr = RPr?.Clone() ?? new RunProps(), Shading = over.Shading ?? Shading };
            r.PPr.Overlay(over.PPr); r.RPr.Overlay(over.RPr);
            r.Borders = new CellBorders(); r.Borders.Overlay(Borders); r.Borders.Overlay(over.Borders);
            return r;
        }
    }

    sealed class NumLevel
    {
        public int Start = 1;
        public string NumFmt = "decimal";
        public string LvlText = "";
        public string Jc = "left";
        public string Suff = "tab";
        /// <summary>w:isLgl — номера всех уровней в тексте этого уровня арабскими цифрами.</summary>
        public bool IsLgl;
        /// <summary>w:lvlRestart: перезапуск после уровня N (1-based); 0 — не перезапускать; null — после любого вышестоящего.</summary>
        public int? Restart;
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

    readonly string _docDir;

    DocxReader(ZipArchive zip, string docDir) { _zip = zip; _docDir = docDir; }

    public static DocxDocument Read(string path)
    {
        using var fs = File.OpenRead(path);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read);
        return new DocxReader(zip, Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".").ReadAll();
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
        LoadNotes(RelTarget("/footnotes"), _footnotes);
        LoadNotes(RelTarget("/endnotes"), _endnotes);

        var xdoc = LoadXml(mainPart) ?? throw new InvalidDataException("Не удалось прочитать " + mainPart);
        var body = xdoc.Root?.Element(W + "body") ?? throw new InvalidDataException("В документе нет w:body");
        _doc.Background = Rgb.Parse((string?)xdoc.Root!.Element(W + "background")?.Attribute(W + "color"));

        var ctx = new PartContext { Rels = mainRels, BaseDir = baseDir };
        var blocks = new List<Block>();
        // Левое поле (для рамок и плавающих таблиц, привязанных к странице)
        _bodyMarginLeft = Twips((string?)body.Element(W + "sectPr")?.Element(W + "pgMar")?.Attribute(W + "left")) ?? 85.05;
        // Свойства нумерации сносок по разделам: settings → sectPr каждого раздела
        _noteSections = body.Descendants(W + "sectPr").ToList();
        _noteSection = -1;
        NextNoteSection();
        ParseBlockContainer(body, ctx, blocks, isBody: true);

        // Финальный sectPr
        var finalSect = body.Element(W + "sectPr");
        var finalProps = finalSect != null ? ParseSectPr(finalSect, ctx) : new SectionProps();
        _doc.Sections.Add(new Section { Props = finalProps, Blocks = blocks });

        AppendEndnotes(blocks);

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
                def.RowBand = Math.Max(1, ParseInt(Val(tblPr.Element(W + "tblStyleRowBandSize"))) ?? 1);
                def.ColBand = Math.Max(1, ParseInt(Val(tblPr.Element(W + "tblStyleColBandSize"))) ?? 1);
            }
            foreach (var cp in s.Elements(W + "tblStylePr"))
            {
                string? type = (string?)cp.Attribute(W + "type");
                if (type == null) continue;
                var tcPr = cp.Element(W + "tcPr");
                def.Cond[type] = new CondFormat
                {
                    PPr = ParsePPr(cp.Element(W + "pPr")),
                    RPr = ParseRPr(cp.Element(W + "rPr")),
                    Shading = ShdHex(tcPr?.Element(W + "shd")),
                    Borders = ParseBorders(tcPr?.Element(W + "tcBorders")),
                };
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

    /// <summary>Условные форматы стиля с учётом наследования (basedOn).</summary>
    Dictionary<string, CondFormat> CondChain(string? id, int depth = 0)
    {
        if (id == null || depth > 20 || !_styles.TryGetValue(id, out var st)) return [];
        var baseCond = CondChain(st.BasedOn, depth + 1);
        var r = new Dictionary<string, CondFormat>(baseCond);
        foreach (var (k, v) in st.Cond) r[k] = r.TryGetValue(k, out var b) ? b.Merge(v) : v;
        return r;
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
        Jc = Val(lvl.Element(W + "lvlJc")) ?? "left",
        IsLgl = OnOff(lvl.Element(W + "isLgl")) == true,
        Restart = ParseInt(Val(lvl.Element(W + "lvlRestart"))),
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
        _doc.MirrorMargins = x.Root.Element(W + "mirrorMargins") is { } mm && IsOnVal(mm);
        _footPrDefault = x.Root.Element(W + "footnotePr");
        _endPrDefault = x.Root.Element(W + "endnotePr");
    }

    // Нумерация сносок: формат, начальный номер, перезапуск (w:footnotePr / w:endnotePr)
    XElement? _footPrDefault, _endPrDefault;
    List<XElement> _noteSections = [];
    int _noteSection;
    string _footFmt = "decimal", _endFmt = "lowerRoman";
    int _footStart = 1, _endStart = 1;

    /// <summary>Переход к следующему разделу: свойства сносок раздела и перезапуск нумерации (numRestart="eachSect").</summary>
    void NextNoteSection()
    {
        _noteSection++;
        var sect = _noteSection < _noteSections.Count ? _noteSections[_noteSection] : null;
        string? Get(XElement? sectPr, XElement? def, string name, string child) =>
            Val(sectPr?.Element(W + name)?.Element(W + child)) ?? Val(def?.Element(W + child));
        bool first = _noteSection == 0;
        _footFmt = Get(sect, _footPrDefault, "footnotePr", "numFmt") ?? "decimal";
        _endFmt = Get(sect, _endPrDefault, "endnotePr", "numFmt") ?? "lowerRoman";
        _footStart = ParseInt(Get(sect, _footPrDefault, "footnotePr", "numStart")) ?? 1;
        _endStart = ParseInt(Get(sect, _endPrDefault, "endnotePr", "numStart")) ?? 1;
        if (first || Get(sect, _footPrDefault, "footnotePr", "numRestart") == "eachSect") _footCount = 0;
        if (first || Get(sect, _endPrDefault, "endnotePr", "numRestart") == "eachSect") _endCount = 0;
    }

    /// <summary>Знак сноски в формате раздела (chicago — *, †, ‡, §, затем удвоенные).</summary>
    static string NoteMark(int n, string fmt)
    {
        if (fmt == "chicago")
        {
            string[] sym = ["*", "\u2020", "\u2021", "\u00A7"];
            return string.Concat(Enumerable.Repeat(sym[(n - 1) % 4], (n - 1) / 4 + 1));
        }
        return FormatNumber(n, fmt);
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

    /// <summary>
    /// Итоговый цвет заливки w:shd: фон (fill) смешивается с цветом узора (color) в доле узора
    /// (pct25 — 25 %, полосы — 50 %/25 %, solid — цвет узора). Возвращает "RRGGBB" или "auto".
    /// </summary>
    static string? ShdHex(XElement? shd)
    {
        if (shd == null) return null;
        string val = (string?)shd.Attribute(W + "val") ?? "clear";
        string? fill = (string?)shd.Attribute(W + "fill");
        if (val is "clear" or "nil") return fill;
        double k = val switch
        {
            "solid" => 1,
            _ when val.StartsWith("pct") && double.TryParse(val[3..], NumberStyles.Float, CultureInfo.InvariantCulture, out var pc) => pc / 100,
            _ when val.StartsWith("thin") => 0.25,       // thinHorzStripe, thinDiagCross …
            "horzStripe" or "vertStripe" or "diagStripe" or "reverseDiagStripe" => 0.5,
            "horzCross" or "diagCross" => 0.75,
            _ => 0
        };
        if (k <= 0) return fill;
        var bg = Rgb.Parse(fill) ?? new Rgb(1, 1, 1);
        var fg = Rgb.Parse((string?)shd.Attribute(W + "color")) ?? new Rgb(0, 0, 0);
        int C(double a, double b) => (int)Math.Round((a * (1 - k) + b * k) * 255);
        return $"{C(bg.R, fg.R):X2}{C(bg.G, fg.G):X2}{C(bg.B, fg.B):X2}";
    }

    static int? ParseHex(string? s) =>
        s != null && int.TryParse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var v) ? v : null;

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
        if (u != null)
        {
            r.Underline = Val(u) ?? "single";
            var uc = (string?)u.Attribute(W + "color");
            r.UnderlineColor = uc is null or "auto" ? null : uc;
        }
        r.Size = HalfPoints(Val(rPr.Element(W + "sz")));
        r.Color = Val(rPr.Element(W + "color"));
        r.Highlight = Val(rPr.Element(W + "highlight"));
        var shd = rPr.Element(W + "shd");
        if (shd != null) r.Shading = ShdHex(shd);
        r.VertAlign = Val(rPr.Element(W + "vertAlign"));
        r.StyleId = Val(rPr.Element(W + "rStyle"));
        r.Position = HalfPoints(Val(rPr.Element(W + "position")));
        if (ParseInt(Val(rPr.Element(W + "w"))) is int sc && sc > 0) r.Scale = sc;
        r.Outline = OnOff(rPr.Element(W + "outline"));
        r.Emboss = OnOff(rPr.Element(W + "emboss"));
        r.Imprint = OnOff(rPr.Element(W + "imprint"));
        r.Shadow = OnOff(rPr.Element(W + "shadow"));
        if (rPr.Element(W + "bdr") is { } bdrEl) r.Border = ParseBorder(bdrEl);
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
            // Интервал в сотых долях строки (beforeLines/afterLines) имеет приоритет; строка = 12 pt, как в Word
            if (ParseInt((string?)sp.Attribute(W + "beforeLines")) is { } bl and > 0) p.SpBefore = bl / 100.0 * 12;
            if (ParseInt((string?)sp.Attribute(W + "afterLines")) is { } al and > 0) p.SpAfter = al / 100.0 * 12;
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
        if (shd != null) p.Shading = ShdHex(shd);
        var bdr = pPr.Element(W + "pBdr");
        if (bdr != null)
        {
            p.BorderTop = ParseBorder(bdr.Element(W + "top"));
            p.BorderBottom = ParseBorder(bdr.Element(W + "bottom"));
            p.BorderLeft = ParseBorder(bdr.Element(W + "left") ?? bdr.Element(W + "start"));
            p.BorderRight = ParseBorder(bdr.Element(W + "right") ?? bdr.Element(W + "end"));
            p.BorderBetween = ParseBorder(bdr.Element(W + "between"));
        }
        p.MarkRunProps = ParseRPr(pPr.Element(W + "rPr"));
        p.SnapToGrid = OnOff(pPr.Element(W + "snapToGrid"));
        p.SuppressLineNumbers = OnOff(pPr.Element(W + "suppressLineNumbers"));
        p.Bidi = OnOff(pPr.Element(W + "bidi"));
        if (pPr.Element(W + "framePr") is { } fp && (string?)fp.Attribute(W + "dropCap") is "drop" or "margin")
        {
            p.DropLines = Math.Max(1, ParseInt((string?)fp.Attribute(W + "lines")) ?? 3);
            p.DropGap = Twips((string?)fp.Attribute(W + "hSpace")) ?? 0;
        }
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
            DoubleUnderline = r.Underline is "double" or "wavyDouble",
            UnderlineStyle = r.Underline,
            UnderlineColor = Rgb.Parse(r.UnderlineColor),
            DoubleStrike = r.DStrike ?? false,
            Strike = (r.Strike ?? false) || (r.DStrike ?? false),
            Caps = r.Caps ?? false,
            SmallCaps = r.SmallCaps ?? false,
            LetterSpacing = r.Spacing ?? 0,
            HScale = (r.Scale ?? 100) / 100.0,
            Outline = r.Outline ?? false,
            RunBorder = r.Border is { IsNone: false } ? r.Border : null,
            // Эффекты Word: «приподнятый» — светлый текст с тёмной тенью снизу-справа, «утопленный» — тень сверху-слева, «тень» — серая тень
            ShadowColor = r.Emboss == true || r.Imprint == true || r.Shadow == true ? new Rgb(0.55, 0.55, 0.55) : null,
            ShadowOffset = r.Imprint == true ? -0.04 : r.Emboss == true ? 0.04 : r.Shadow == true ? 0.06 : 0,
            Size = r.Size ?? 10,
            Color = (r.Emboss == true || r.Imprint == true)
                ? new Rgb(0.93, 0.93, 0.93)
                : Rgb.Parse(r.Color) ?? (link != null && r.Color == null ? new Rgb(0.02, 0.39, 0.76) : Rgb.Black),
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
            if (el.Name == W + "p")
            {
                // Плавающие объекты абзаца (картинки, надписи) — блоки перед ним
                var saved = _pending;
                var pending = new List<Block>();
                _pending = pending;
                var para = ParseParagraph(el, ctx);
                _pending = saved;
                output.AddRange(pending);
                if (para.DropLines > 0)
                    output.Add(new FloatBlock { Side = "left", Width = 0, DropLines = para.DropLines, MarginRight = para.DropGap, Blocks = [para] });
                else if (FrameFloat(el.Element(W + "pPr")?.Element(W + "framePr")) is { } ff)
                {
                    // Соседние абзацы с одинаковой рамкой (framePr) — одна рамка
                    string key = el.Element(W + "pPr")!.Element(W + "framePr")!.ToString(SaveOptions.DisableFormatting);
                    if (output.Count > 0 && output[^1] is FloatBlock prevF && prevF.FrameKey == key) prevF.Blocks.Add(para);
                    else { ff.FrameKey = key; ff.Blocks.Add(para); output.Add(ff); }
                }
                else output.Add(para);
            }
            else if (el.Name == W + "tbl")
            {
                var table = ParseTable(el, ctx);
                if (TableFloat(el.Element(W + "tblPr")?.Element(W + "tblpPr"), table) is { } tf) { tf.Blocks.Add(table); output.Add(tf); }
                else output.Add(table);
            }
            else if (el.Name == W + "sdt") { var c = el.Element(W + "sdtContent"); if (c != null) ParseBlockContainer(c, ctx, output); }
            else if (el.Name == W + "customXml" || el.Name == W + "ins" || el.Name == W + "moveTo") ParseBlockContainer(el, ctx, output);
            else if (el.Name == MC + "AlternateContent")
            {
                var alt = el.Element(MC + "Choice") ?? el.Element(MC + "Fallback");
                if (alt != null) ParseBlockContainer(alt, ctx, output);
            }
        }
    }

    /// <summary>
    /// Рамка-абзац w:framePr (не буквица) с обтеканием — плавающий блок у левого или правого края.
    /// Без обтекания (wrap="none"/"notBeside") абзац остаётся в потоке.
    /// </summary>
    double _bodyMarginLeft = 85.05;

    FloatBlock? FrameFloat(XElement? fp)
    {
        if (fp == null || fp.Attribute(W + "dropCap") is { } dc && dc.Value != "none") return null;
        string wrap = (string?)fp.Attribute(W + "wrap") ?? "auto";
        if (wrap is "none" or "notBeside") return null;
        double w = Twips((string?)fp.Attribute(W + "w")) ?? 0;
        double hSpace = Twips((string?)fp.Attribute(W + "hSpace")) ?? 0, vSpace = Twips((string?)fp.Attribute(W + "vSpace")) ?? 0;
        string hAnchor = (string?)fp.Attribute(W + "hAnchor") ?? "page", vAnchor = (string?)fp.Attribute(W + "vAnchor") ?? "margin";
        string? xAlign = (string?)fp.Attribute(W + "xAlign");
        double x = Twips((string?)fp.Attribute(W + "x")) ?? 0, y = Twips((string?)fp.Attribute(W + "y")) ?? 0;
        return MakeSideFloat(xAlign, hAnchor, x, w, vAnchor == "text" ? y : 0, hSpace, hSpace, vSpace);
    }

    /// <summary>Плавающая таблица (w:tblpPr): обтекается текстом следующих абзацев.</summary>
    FloatBlock? TableFloat(XElement? tp, Table t)
    {
        if (tp == null) return null;
        double L(string n) => Twips((string?)tp.Attribute(W + n)) ?? 0;
        string hAnchor = (string?)tp.Attribute(W + "horzAnchor") ?? "column", vAnchor = (string?)tp.Attribute(W + "vertAnchor") ?? "margin";
        string? xSpec = (string?)tp.Attribute(W + "tblpXSpec");
        double w = t.Width ?? t.Grid.Sum();
        t.Indent = 0;
        t.Align = "left";
        return MakeSideFloat(xSpec, hAnchor, L("tblpX"), w, vAnchor == "text" ? L("tblpY") : 0, L("leftFromText"), L("rightFromText"), L("bottomFromText"));
    }

    FloatBlock MakeSideFloat(string? align, string hAnchor, double x, double w, double top, double spaceL, double spaceR, double spaceB)
    {
        // Смещение от края страницы приводим к смещению от поля
        double ml = Math.Max(0, hAnchor == "page" ? x - _bodyMarginLeft : x);
        string side = align switch
        {
            "right" or "outside" => "right",
            "left" or "inside" or "center" => "left",
            _ => "left"
        };
        return new FloatBlock
        {
            Side = side, Width = w,
            MarginLeft = side == "right" ? spaceL : align == null ? ml : 0,
            MarginRight = side == "left" ? spaceR : 0,
            MarginTop = Math.Max(0, top), MarginBottom = spaceB,
        };
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
            BorderLeft = pp.BorderLeft is { IsNone: false } ? pp.BorderLeft : null,
            BorderRight = pp.BorderRight is { IsNone: false } ? pp.BorderRight : null,
            BorderBetween = pp.BorderBetween is { IsNone: false } ? pp.BorderBetween : null,
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

        if (lvl != null && para.Inlines.Count > 0 && lvl.Jc is "right" or "end" or "center") para.LabelAlign = lvl.Jc == "center" ? "center" : "right";
        para.SnapToGrid = pp.SnapToGrid ?? true;
        para.SuppressLineNumbers = pp.SuppressLineNumbers ?? false;
        para.DropLines = direct?.DropLines ?? 0;
        para.DropGap = direct?.DropGap ?? 0;
        para.Rtl = pp.Bidi ?? false;

        var state = new RunState { BaseR = baseR, Ctx = ctx };
        ParseInlineContainer(p, state, para.Inlines);
        ShareRunBorders(para.Inlines);
        para.Bookmarks = state.Bookmarks;
        // Отдельная формула (m:oMathPara) — по центру, как в Word
        if (p.Elements(M + "oMathPara").Any() && !p.Elements(W + "r").Any(r => r.Elements(W + "t").Any(t => t.Value.Trim().Length > 0)))
            para.Align = "center";

        if (pPrEl?.Element(W + "sectPr") is { } sp && !ctx.HeaderFooter)
        {
            para.Section = ParseSectPr(sp, ctx);
            if (ctx.BaseDir != null && _noteMark == null) NextNoteSection();
        }
        return para;
    }

    /// <summary>Соседние фрагменты с одинаковой рамкой w:bdr объединяются в один строчный блок.</summary>
    static void ShareRunBorders(List<Inline> inlines)
    {
        InlineBox? box = null;
        BorderDef? last = null;
        foreach (var inl in inlines)
        {
            TextStyle? st = inl switch { TextInline t => t.Style, TabInline tb => tb.Style, FieldInline fi => fi.Style, ImageInline ii => ii.Style, _ => null };
            if (st == null) continue;
            var b = st.RunBorder;
            if (b == null) { box = null; last = null; continue; }
            if (box == null || last == null || last.Val != b.Val || Math.Abs(last.Width - b.Width) > 0.01 || last.Color != b.Color)
            {
                double pad = b.Space;
                box = new InlineBox { Top = b, Bottom = b, Left = b, Right = b, PadLeft = pad, PadRight = pad, PadTop = 0.5, PadBottom = 0.5, Background = st.Background };
                last = b;
            }
            var ns = st with { Box = box };
            switch (inl)
            {
                case TextInline t: t.Style = ns; break;
                case TabInline tb: tb.Style = ns; break;
                case FieldInline fi: fi.Style = ns; break;
                case ImageInline ii: ii.Style = ns; break;
            }
        }
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
        for (int i = ilvl + 1; i < 10; i++)
        {
            var li = GetLevel(numId, i);
            int restartAfter = li?.Restart ?? i;   // по умолчанию — после любого вышестоящего уровня
            if (ilvl < restartAfter) counters[i] = (li?.Start ?? 1) - 1;
        }
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
                if (lvl.IsLgl && fmt != "decimalZero") fmt = "decimal";
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
        public List<string>? Bookmarks;
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
                // Ссылка на закладку внутри документа
                else if ((string?)el.Attribute(W + "anchor") is { Length: > 0 } anchorName) st.Link = "#" + anchorName;
                ParseInlineContainer(el, st, output);
                st.Link = prev;
            }
            else if (n == W + "fldSimple")
            {
                string instr = ((string?)el.Attribute(W + "instr") ?? "").Trim();
                var kind = FieldKindOf(instr);
                if (kind != null)
                {
                    var r = el.Descendants(W + "rPr").FirstOrDefault();
                    output.Add(new FieldInline(kind.Value, ToTextStyle(ResolveRunProps(st.BaseR, ParseRPr(r)))));
                }
                else ParseInlineContainer(el, st, output);
            }
            else if (n == W + "bookmarkStart")
            {
                if ((string?)el.Attribute(W + "name") is { Length: > 0 } bm) (st.Bookmarks ??= []).Add(bm);
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
                // Формулы OMML — линейная запись с индексами
                if (!st.Suppressed)
                {
                    var rr = ResolveRunProps(st.BaseR, null);
                    var baseStyle = ToTextStyle(rr, st.Link) with { FontAscii = "Cambria Math", FontHAnsi = "Cambria Math", FontEastAsia = "Cambria Math" };
                    MathToInlines(el, baseStyle, output);
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
                    if (kind != null && !OuterSuppressed(st))
                    {
                        output.Add(new FieldInline(kind.Value, f.Style ?? style));
                        f.SkipResult = true;
                        f.Emitted = true;
                    }
                }
                else if (type == "end" && st.Fields.Count > 0)
                {
                    var f = st.Fields.Pop();
                    if (!f.InResult && !f.Emitted && !st.Suppressed)
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
            if (ParseNoteElement(el, style, output)) continue;

            if (n == W + "t") { if (el.Value.Length > 0) output.Add(new TextInline(el.Value, style)); }
            else if (n == W + "tab") output.Add(new TabInline(style));
            else if (n == W + "ptab")
                output.Add(new TabInline(style)
                {
                    PtabAlign = (string?)el.Attribute(W + "alignment") ?? "left",
                    PtabRelativeTo = (string?)el.Attribute(W + "relativeTo") ?? "margin",
                    PtabLeader = (string?)el.Attribute(W + "leader") ?? "none",
                });
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
        if (!st.Ctx.Rels.TryGetValue(rid, out var rel)) return;
        if (!_doc.Media.ContainsKey(rel.Target))
        {
            var bytes = rel.External ? LoadExternal(rel.Target) : LoadBytes(rel.Target);
            if (bytes == null) return;
            _doc.Media[rel.Target] = bytes;
        }
        output.Add(new ImageInline(rel.Target, w, h, style));
    }

    /// <summary>Картинка по внешней ссылке (r:link): file:// или путь относительно документа.</summary>
    byte[]? LoadExternal(string target)
    {
        try
        {
            string path = target.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(target).LocalPath : Uri.UnescapeDataString(target);
            if (!Path.IsPathRooted(path)) path = Path.Combine(_docDir, path.Replace('/', Path.DirectorySeparatorChar));
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }
        catch { return null; }
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
            if (tblPr.Element(W + "tblCellSpacing") is { } tcs && (string?)tcs.Attribute(W + "type") is "dxa" or null)
                t.CellSpacing = Twips((string?)tcs.Attribute(W + "w")) ?? 0;
            var tw = tblPr.Element(W + "tblW");
            if (tw != null && (string?)tw.Attribute(W + "type") == "dxa") t.Width = Twips((string?)tw.Attribute(W + "w"));
            else if (tw != null && (string?)tw.Attribute(W + "type") == "pct" && (string?)tw.Attribute(W + "w") is { } pw)
            {
                // Пятидесятые доли процента ("5000" = 100 %) или явные проценты ("100%")
                if (pw.EndsWith('%') && double.TryParse(pw.TrimEnd('%'), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pp)) t.WidthPct = pp / 100;
                else if (double.TryParse(pw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var pf)) t.WidthPct = pf / 5000;
                if (t.WidthPct is <= 0) t.WidthPct = null;
            }
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

        // Условное форматирование стиля: какие области включены (tblLook) и размеры полос
        var cond = styleId != null ? CondChain(styleId) : [];
        var look = tblPr?.Element(W + "tblLook");
        int lookVal = ParseHex((string?)look?.Attribute(W + "val")) ?? 0x04A0;
        bool LookOn(string attr, int bit) => look?.Attribute(W + attr) is { } a ? IsOn(a) : (lookVal & bit) != 0;
        bool firstRowOn = LookOn("firstRow", 0x20), lastRowOn = LookOn("lastRow", 0x40);
        bool firstColOn = LookOn("firstColumn", 0x80), lastColOn = LookOn("lastColumn", 0x100);
        bool hBandOn = !LookOn("noHBand", 0x200), vBandOn = !LookOn("noVBand", 0x400);
        StyleDef? sdef = styleId != null ? _styles.GetValueOrDefault(styleId) : null;
        int rowBand = sdef?.RowBand ?? 1, colBand = sdef?.ColBand ?? 1;
        var allRows = RowsOf(tbl).ToList();
        int totalCols = Math.Max(t.Grid.Count, allRows.Count == 0 ? 0 : allRows.Max(r => CellsOf(r).Sum(c => Math.Max(1, ParseInt(Val(c.Element(W + "tcPr")?.Element(W + "gridSpan"))) ?? 1))));

        CondFormat? CondFor(int r, int c, int span)
        {
            if (cond.Count == 0) return null;
            bool firstRow = firstRowOn && r == 0, lastRow = lastRowOn && r == allRows.Count - 1;
            bool firstCol = firstColOn && c == 0, lastCol = lastColOn && c + span >= totalCols;
            var parts = new List<string> { "wholeTable" };
            if (vBandOn && !firstCol && !lastCol)
            {
                int ci = c - (firstColOn ? 1 : 0);
                parts.Add((ci / colBand) % 2 == 0 ? "band1Vert" : "band2Vert");
            }
            if (hBandOn && !firstRow && !lastRow)
            {
                int ri = r - (firstRowOn ? 1 : 0);
                parts.Add((ri / rowBand) % 2 == 0 ? "band1Horz" : "band2Horz");
            }
            if (firstCol) parts.Add("firstCol");
            if (lastCol) parts.Add("lastCol");
            if (firstRow) parts.Add("firstRow");
            if (lastRow) parts.Add("lastRow");
            if (firstRow && firstCol) parts.Add("nwCell");
            if (firstRow && lastCol) parts.Add("neCell");
            if (lastRow && firstCol) parts.Add("swCell");
            if (lastRow && lastCol) parts.Add("seCell");
            CondFormat? acc = null;
            foreach (var p in parts)
                if (cond.TryGetValue(p, out var cf)) acc = acc == null ? new CondFormat().Merge(cf) : acc.Merge(cf);
            return acc;
        }

        int rowIndex = -1;
        foreach (var trEl in allRows)
        {
            rowIndex++;
            int gridCol = 0;
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
            gridCol = row.Cells.Sum(c => c.GridSpan);
            foreach (var tcEl in CellsOf(trEl))
            {
                var cell = new TableCell();
                var tcPr = tcEl.Element(W + "tcPr");
                int spanHere = Math.Max(1, ParseInt(Val(tcPr?.Element(W + "gridSpan"))) ?? 1);
                var cf = CondFor(rowIndex, gridCol, spanHere);
                gridCol += spanHere;
                if (tcPr != null)
                {
                    cell.GridSpan = Math.Max(1, ParseInt(Val(tcPr.Element(W + "gridSpan"))) ?? 1);
                    var vm = tcPr.Element(W + "vMerge");
                    if (vm != null) cell.VMerge = Val(vm) == "restart" ? VMerge.Restart : VMerge.Continue;
                    if (tcPr.Element(W + "hMerge") is { } hm && Val(hm) != "restart") cell.GridSpan = 0; // склеено с предыдущей
                    cell.Borders = ParseBorders(tcPr.Element(W + "tcBorders"));
                    cell.Shading = Rgb.Parse(ShdHex(tcPr.Element(W + "shd")));
                    cell.Margins = ParseMargins(tcPr.Element(W + "tcMar"));
                    cell.VAlign = Val(tcPr.Element(W + "vAlign")) ?? "top";
                    cell.TextDirection = Val(tcPr.Element(W + "textDirection")) switch
                    {
                        "btLr" => "btLr",
                        "tbRl" or "tbRlV" or "tbLrV" => "tbRl",
                        _ => null
                    };
                    var tcW = tcPr.Element(W + "tcW");
                    if (tcW != null && (string?)tcW.Attribute(W + "type") is "dxa" or null) cell.Width = Twips((string?)tcW.Attribute(W + "w"));
                }
                var useCtx = cellCtx;
                if (cf != null)
                {
                    // Условный формат: заливка и рамки ячейки (если не заданы явно), свойства текста
                    if (cell.Shading == null) cell.Shading = Rgb.Parse(cf.Shading);
                    if (cf.Borders != null)
                    {
                        var cb = new CellBorders
                        {
                            Top = cf.Borders.Top ?? cf.Borders.InsideH, Bottom = cf.Borders.Bottom ?? cf.Borders.InsideH,
                            Left = cf.Borders.Left ?? cf.Borders.InsideV, Right = cf.Borders.Right ?? cf.Borders.InsideV,
                        };
                        cb.Overlay(cell.Borders);
                        cell.Borders = cb;
                    }
                    var tp2 = (cellCtx.TablePPr ?? new ParaProps()).Clone(); tp2.Overlay(cf.PPr);
                    var tr2 = (cellCtx.TableRPr ?? new RunProps()).Clone(); tr2.Overlay(cf.RPr);
                    useCtx = new PartContext { Rels = ctx.Rels, BaseDir = ctx.BaseDir, HeaderFooter = ctx.HeaderFooter, TablePPr = tp2, TableRPr = tr2 };
                }
                ParseBlockContainer(tcEl, useCtx, cell.Blocks);
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
        s.PageNumberFormat = (string?)sp.Element(W + "pgNumType")?.Attribute(W + "fmt") ?? "decimal";
        s.VAlign = Val(sp.Element(W + "vAlign")) ?? "top";
        if (sp.Element(W + "lnNumType") is { } ln)
        {
            s.LineNumberCountBy = ParseInt((string?)ln.Attribute(W + "countBy")) ?? 0;
            // w:start — номер первой строки минус один (отсчёт с нуля)
            s.LineNumberStart = (ParseInt((string?)ln.Attribute(W + "start")) ?? 0) + 1;
            s.LineNumberDistance = Twips((string?)ln.Attribute(W + "distance"));
            s.LineNumberRestart = (string?)ln.Attribute(W + "restart") ?? "newPage";
        }
        if (sp.Element(W + "pgBorders") is { } pb)
        {
            BorderDef? B(string n) => ParseBorder(pb.Element(W + n)) is { IsNone: false } b ? b : null;
            s.PageBorderTop = B("top"); s.PageBorderBottom = B("bottom"); s.PageBorderLeft = B("left"); s.PageBorderRight = B("right");
            s.PageBorderFromText = (string?)pb.Attribute(W + "offsetFrom") == "text";
            s.PageBorderDisplay = (string?)pb.Attribute(W + "display") ?? "allPages";
        }
        if (sp.Elements(W + "docGrid").LastOrDefault() is { } grid && (string?)grid.Attribute(W + "type") is "lines" or "linesAndChars" or "snapToChars")
            s.LinePitch = Twips((string?)grid.Attribute(W + "linePitch")) ?? 0;
        if (sp.Elements(W + "cols").LastOrDefault() is { } cols)
        {
            s.Columns = Math.Max(1, ParseInt((string?)cols.Attribute(W + "num")) ?? 1);
            s.ColumnSpace = Twips((string?)cols.Attribute(W + "space")) ?? 36;
            s.ColumnSeparator = cols.Attribute(W + "sep") is { } sepA && IsOn(sepA);
            var explicitCols = cols.Elements(W + "col").ToList();
            if (explicitCols.Count > 0 && (string?)cols.Attribute(W + "equalWidth") is "0" or "false")
            {
                s.ColumnWidths = explicitCols.Select(c => (Twips((string?)c.Attribute(W + "w")) ?? 0, Twips((string?)c.Attribute(W + "space")) ?? 0)).ToList();
                s.Columns = s.ColumnWidths.Count;
            }
        }

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

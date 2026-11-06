using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using DocxToPdf.Docx;

namespace DocxToPdf.Odt;

/// <summary>
/// Чтение документов OpenDocument Text (.odt, .ott, плоский .fodt) в общую модель документа:
/// стили (наследование, стили по умолчанию, автоматические стили), абзацы и заголовки, списки и нумерация
/// заголовков, таблицы (объединения, повторы), страницы (макеты, стили страниц, колонтитулы), разделы с колонками,
/// сноски, поля, закладки и ссылки, рамки с картинками и надписями, фигуры, формулы (MathML) и диаграммы.
/// </summary>
public sealed partial class OdtReader
{
    internal static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    internal static readonly XNamespace St = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    internal static readonly XNamespace Tx = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    internal static readonly XNamespace Tb = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    internal static readonly XNamespace Dr = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    internal static readonly XNamespace Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    internal static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
    internal static readonly XNamespace XLink = "http://www.w3.org/1999/xlink";
    internal static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    internal static readonly XNamespace Meta = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
    internal static readonly XNamespace Ch = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
    internal static readonly XNamespace Mml = "http://www.w3.org/1998/Math/MathML";
    internal static readonly XNamespace Cfg = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";
    internal static readonly XNamespace Loext = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    internal static readonly XNamespace Xml = XNamespace.Xml;

    // ---------------- Пакет ----------------

    readonly ZipArchive? _zip;
    readonly XDocument? _flat;                       // плоский документ (.fodt)
    readonly string? _baseDir;                       // каталог файла (для внешних ссылок)
    readonly DocxDocument _doc = new();

    readonly DocxOptions _options;

    OdtReader(ZipArchive? zip, XDocument? flat, string? baseDir, DocxOptions? options) { _zip = zip; _flat = flat; _baseDir = baseDir; _options = options ?? new DocxOptions(); }

    /// <summary>Является ли файл документом OpenDocument Text (ZIP с content.xml или плоский XML office:document).</summary>
    public static bool IsOdt(string path)
    {
        try
        {
            using var fs = File.OpenRead(path);
            var head = new byte[2];
            if (fs.Read(head, 0, 2) == 2 && head[0] == 'P' && head[1] == 'K')
            {
                fs.Position = 0;
                using var z = new ZipArchive(fs, ZipArchiveMode.Read);
                if (z.GetEntry("word/document.xml") != null) return false;
                if (z.GetEntry("mimetype") is { } mt)
                {
                    using var r = new StreamReader(mt.Open());
                    string m = r.ReadToEnd().Trim();
                    if (m.StartsWith("application/vnd.oasis.opendocument.text", StringComparison.Ordinal)) return true;
                    if (m.StartsWith("application/vnd.oasis.opendocument.", StringComparison.Ordinal)) return false;
                }
                return z.GetEntry("content.xml") != null;
            }
            fs.Position = 0;
            var buf = new byte[4096];
            int n = fs.Read(buf, 0, buf.Length);
            string s = Encoding.UTF8.GetString(buf, 0, n);
            return s.Contains("urn:oasis:names:tc:opendocument:xmlns:office:1.0") && s.Contains("office:document")
                && !s.Contains("office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet")
                && !s.Contains("office:mimetype=\"application/vnd.oasis.opendocument.presentation");
        }
        catch { return false; }
    }

    /// <summary>Тип документа ODF, если это не текст (электронная таблица, презентация, рисунок…), иначе null.</summary>
    public static string? OtherOdfKind(string path)
    {
        try
        {
            using var z = ZipFile.OpenRead(path);
            if (z.GetEntry("mimetype") is not { } mt) return null;
            using var r = new StreamReader(mt.Open());
            string m = r.ReadToEnd().Trim();
            if (!m.StartsWith("application/vnd.oasis.opendocument.", StringComparison.Ordinal) || m.Contains(".text")) return null;
            return m["application/vnd.oasis.opendocument.".Length..];
        }
        catch { return null; }
    }

    public static DocxDocument Read(string path, DocxOptions? options = null)
    {
        string? dir = Path.GetDirectoryName(Path.GetFullPath(path));
        using var fs = File.OpenRead(path);
        bool zip = fs.ReadByte() == 'P' && fs.ReadByte() == 'K';
        fs.Position = 0;
        if (zip)
        {
            using var z = new ZipArchive(fs, ZipArchiveMode.Read);
            if (LoadXmlFrom(z, "META-INF/manifest.xml")?.Descendants().Any(e => e.Name.LocalName == "encryption-data") == true)
                throw new InvalidDataException("Документ ODT зашифрован (защищён паролем) — конвертация невозможна");
            return new OdtReader(z, null, dir, options).ReadAll();
        }
        var x = XDocument.Load(fs, LoadOptions.PreserveWhitespace);
        return new OdtReader(null, x, dir, options).ReadAll();
    }

    static XDocument? LoadXmlFrom(ZipArchive z, string name)
    {
        var e = z.GetEntry(name);
        if (e == null) return null;
        try { using var s = e.Open(); return XDocument.Load(s, LoadOptions.PreserveWhitespace); }
        catch { return null; }
    }

    XDocument? LoadXml(string name) => _zip != null ? LoadXmlFrom(_zip, name.TrimStart('.', '/')) : null;

    internal byte[]? LoadBytes(string name)
    {
        if (_zip == null) return null;
        name = Uri.UnescapeDataString(name);
        while (name.StartsWith("./")) name = name[2..];
        var e = _zip.GetEntry(name.TrimStart('/'));
        if (e == null) return null;
        using var s = e.Open();
        var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    // ================================================================
    //  Стили
    // ================================================================

    /// <summary>Стиль ODF: свойства по группам (атрибуты *-properties как «префикс:имя» → значение).</summary>
    internal sealed class OStyle
    {
        public string Name = "", Family = "";
        public string? Parent, ListStyle, MasterPage, DisplayName;
        public int? OutlineLevel;
        public readonly Dictionary<string, string> Text = [], Para = [], Table = [], Cell = [], Column = [], Row = [], Graphic = [], Section = [];
        public XElement? TabStops, DropCap, Columns, BackgroundImage;
        public XElement El = null!;
    }

    readonly Dictionary<(string Family, string Name), OStyle> _common = [];
    readonly Dictionary<(string Family, string Name), OStyle> _autoContent = [];
    readonly Dictionary<(string Family, string Name), OStyle> _autoStyles = [];
    readonly Dictionary<string, OStyle> _defaults = [];
    readonly Dictionary<string, XElement> _listStyles = [];          // общие и автоматические стили списков
    readonly Dictionary<string, XElement> _listStylesAutoStyles = [];
    readonly Dictionary<string, XElement> _pageLayouts = [];
    readonly Dictionary<string, XElement> _masterPages = [];
    readonly Dictionary<string, string> _fontFaces = [];             // style:font-face name → svg:font-family
    readonly Dictionary<string, XElement> _drawStyles = [];          // draw:gradient, draw:hatch, draw:marker, draw:fill-image
    XElement? _outlineStyle;
    XElement? _footConfig, _endConfig;
    string? _firstMaster;

    static string Key(XName n) => n.Namespace == Fo ? "fo:" + n.LocalName : n.Namespace == St ? "style:" + n.LocalName
        : n.Namespace == Svg ? "svg:" + n.LocalName : n.Namespace == Tx ? "text:" + n.LocalName : n.Namespace == Tb ? "table:" + n.LocalName
        : n.Namespace == Dr ? "draw:" + n.LocalName : n.Namespace == Loext ? "loext:" + n.LocalName : n.LocalName;

    static void ReadProps(XElement? props, Dictionary<string, string> into)
    {
        if (props == null) return;
        foreach (var a in props.Attributes()) if (!a.IsNamespaceDeclaration) into[Key(a.Name)] = a.Value;
    }

    static OStyle ParseStyle(XElement s)
    {
        var st = new OStyle
        {
            El = s,
            Name = (string?)s.Attribute(St + "name") ?? "",
            Family = (string?)s.Attribute(St + "family") ?? "",
            Parent = (string?)s.Attribute(St + "parent-style-name"),
            ListStyle = (string?)s.Attribute(St + "list-style-name"),
            MasterPage = (string?)s.Attribute(St + "master-page-name"),
            DisplayName = (string?)s.Attribute(St + "display-name"),
            OutlineLevel = int.TryParse((string?)s.Attribute(St + "default-outline-level"), out var ol) ? ol : null,
        };
        ReadProps(s.Element(St + "text-properties"), st.Text);
        var pp = s.Element(St + "paragraph-properties");
        ReadProps(pp, st.Para);
        st.TabStops = pp?.Element(St + "tab-stops");
        st.DropCap = pp?.Element(St + "drop-cap");
        st.BackgroundImage = pp?.Element(St + "background-image");
        ReadProps(s.Element(St + "table-properties"), st.Table);
        ReadProps(s.Element(St + "table-cell-properties"), st.Cell);
        ReadProps(s.Element(St + "table-column-properties"), st.Column);
        ReadProps(s.Element(St + "table-row-properties"), st.Row);
        var gp = s.Element(St + "graphic-properties");
        ReadProps(gp, st.Graphic);
        var sp = s.Element(St + "section-properties");
        ReadProps(sp, st.Section);
        st.Columns = sp?.Element(St + "columns") ?? gp?.Element(St + "columns");
        return st;
    }

    void LoadStyleContainer(XElement? c, Dictionary<(string, string), OStyle> into, Dictionary<string, XElement> lists, bool common)
    {
        if (c == null) return;
        foreach (var e in c.Elements())
        {
            if (e.Name == St + "style") { var s = ParseStyle(e); into[(s.Family, s.Name)] = s; }
            else if (e.Name == St + "default-style" && common) { var s = ParseStyle(e); _defaults[s.Family] = s; }
            else if (e.Name == Tx + "list-style" && (string?)e.Attribute(St + "name") is { } ln) lists[ln] = e;
            else if (e.Name == Tx + "outline-style") _outlineStyle = e;
            else if (e.Name == Tx + "notes-configuration")
            {
                if ((string?)e.Attribute(Tx + "note-class") == "endnote") _endConfig = e; else _footConfig = e;
            }
            else if (e.Name == St + "page-layout" && (string?)e.Attribute(St + "name") is { } pl) _pageLayouts[pl] = e;
            else if (e.Name.Namespace == Dr && (string?)e.Attribute(Dr + "name") is { } dn) _drawStyles[e.Name.LocalName + ":" + dn] = e;
            else if (e.Name == Tx + "linenumbering-configuration") _lineNumbering = e;
        }
    }

    XElement? _lineNumbering;

    void LoadFontFaces(XElement? decls)
    {
        if (decls == null) return;
        foreach (var f in decls.Elements(St + "font-face"))
        {
            string? name = (string?)f.Attribute(St + "name");
            string? fam = (string?)f.Attribute(Svg + "font-family");
            if (name != null) _fontFaces[name] = Unquote(fam ?? name);
        }
    }

    static string Unquote(string s)
    {
        s = s.Trim();
        if (s.Length >= 2 && (s[0] == '\'' && s[^1] == '\'' || s[0] == '"' && s[^1] == '"')) s = s[1..^1];
        int comma = s.IndexOf(',');
        if (comma > 0) s = Unquote(s[..comma]);
        return s;
    }

    /// <summary>Контекст разбора: тело документа или колонтитулы (свои автоматические стили).</summary>
    internal sealed class Ctx
    {
        public bool HeaderFooter;
        public bool InNote;
        public int TableDepth;
        /// <summary>Свойства текста фигуры (draw:text-style-name): поверх стиля абзаца.</summary>
        public OStyle? ShapeText;
    }

    internal OStyle? FindStyle(string? name, string family, Ctx ctx)
    {
        if (string.IsNullOrEmpty(name)) return null;
        var auto = ctx.HeaderFooter ? _autoStyles : _autoContent;
        if (auto.TryGetValue((family, name), out var s)) return s;
        if (_common.TryGetValue((family, name), out s)) return s;
        // Автоматические стили колонтитулов иногда ссылаются на стили содержимого и наоборот
        if ((ctx.HeaderFooter ? _autoContent : _autoStyles).TryGetValue((family, name), out s)) return s;
        return null;
    }

    /// <summary>Цепочка стилей от корня к самому стилю (без стиля по умолчанию).</summary>
    List<OStyle> Chain(OStyle? s, Ctx ctx)
    {
        var list = new List<OStyle>();
        for (int i = 0; s != null && i < 32; i++)
        {
            list.Add(s);
            s = s.Parent != null ? FindStyle(s.Parent, s.Family, ctx) : null;
            if (s != null && list.Contains(s)) break;
        }
        list.Reverse();
        return list;
    }

    /// <summary>Свойства текста: значения из цепочки накладываются; % у размера шрифта — от предыдущего значения.</summary>
    static void OverlayText(Dictionary<string, string> acc, Dictionary<string, string> src)
    {
        foreach (var (k, v) in src)
        {
            if (k is "fo:font-size" or "style:font-size-asian" or "style:font-size-complex" && v.EndsWith('%'))
            {
                double baseSize = acc.TryGetValue(k, out var prev) && Length(prev) is double p ? p : 12;
                if (double.TryParse(v[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)) acc[k] = (baseSize * pct / 100).ToString(CultureInfo.InvariantCulture) + "pt";
                continue;
            }
            acc[k] = v;
        }
    }

    static void OverlayPara(Dictionary<string, string> acc, Dictionary<string, string> src)
    {
        foreach (var (k, v) in src)
        {
            // Отступы в % — от значения родителя
            if (v.EndsWith('%') && k is "fo:margin-left" or "fo:margin-right" or "fo:margin-top" or "fo:margin-bottom" or "fo:text-indent")
            {
                double baseV = acc.TryGetValue(k, out var prev) && Length(prev) is double p ? p : 0;
                if (double.TryParse(v[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)) acc[k] = (baseV * pct / 100).ToString(CultureInfo.InvariantCulture) + "pt";
                continue;
            }
            acc[k] = v;
        }
    }

    internal Dictionary<string, string> TextPropsOf(OStyle? style, Ctx ctx, Dictionary<string, string>? baseProps = null)
    {
        var acc = baseProps != null ? new Dictionary<string, string>(baseProps) : new Dictionary<string, string>();
        if (baseProps == null && _defaults.TryGetValue(style?.Family ?? "paragraph", out var d)) OverlayText(acc, d.Text);
        if (baseProps == null && style?.Family == "text" && _defaults.TryGetValue("paragraph", out var dp)) OverlayText(acc, dp.Text);
        foreach (var s in Chain(style, ctx)) OverlayText(acc, s.Text);
        return acc;
    }

    (Dictionary<string, string> Para, Dictionary<string, string> Text, XElement? Tabs, XElement? Drop, string? ListStyle, int? Outline) ParaPropsOf(OStyle? style, Ctx ctx)
    {
        var p = new Dictionary<string, string>();
        var t = new Dictionary<string, string>();
        XElement? tabs = null, drop = null;
        string? list = null;
        int? outline = null;
        if (_defaults.TryGetValue("paragraph", out var d)) { OverlayPara(p, d.Para); OverlayText(t, d.Text); tabs = d.TabStops ?? tabs; }
        foreach (var s in Chain(style, ctx))
        {
            OverlayPara(p, s.Para);
            OverlayText(t, s.Text);
            if (s.TabStops != null) tabs = s.TabStops;
            if (s.DropCap != null) drop = s.DropCap;
            if (s.ListStyle != null) list = s.ListStyle.Length > 0 ? s.ListStyle : null;
            if (s.OutlineLevel != null) outline = s.OutlineLevel;
        }
        return (p, t, tabs, drop, list, outline);
    }

    // ================================================================
    //  Единицы, цвета, рамки
    // ================================================================

    /// <summary>Длина ODF (cm, mm, in, pt, pc, px) в пунктах.</summary>
    internal static double? Length(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        v = v.Trim();
        int i = 0;
        while (i < v.Length && (char.IsDigit(v[i]) || v[i] is '.' or '-' or '+' or 'e' or 'E') && !(v[i] is 'e' or 'E' && (i + 1 >= v.Length || !char.IsDigit(v[i + 1]) && v[i + 1] != '-'))) i++;
        if (!double.TryParse(v[..i], NumberStyles.Float, CultureInfo.InvariantCulture, out var n)) return null;
        return v[i..].Trim().ToLowerInvariant() switch
        {
            "cm" => n * 72 / 2.54,
            "mm" => n * 72 / 25.4,
            "in" or "inch" => n * 72,
            "pt" or "" => n,
            "pc" => n * 12,
            "px" => n * 0.75,
            "em" => n * 12,
            _ => null,
        };
    }

    internal static double Len(Dictionary<string, string> d, string key, double def = 0) => d.TryGetValue(key, out var v) && Length(v) is double x ? x : def;

    internal static Rgb? Color(string? v) => v is null or "transparent" or "none" ? null : Rgb.Parse(v.TrimStart('#'));

    /// <summary>Рамка: «0.06pt solid #000000» (style:border-line-width — толщины двойной линии).</summary>
    internal static BorderDef? Border(string? v, string? lineWidth = null, string? padding = null)
    {
        if (string.IsNullOrWhiteSpace(v) || v.Trim() is "none" or "hidden") return null;
        double w = 0.75;
        string style = "solid";
        Rgb col = Rgb.Black;
        foreach (var tok in v.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (tok.StartsWith('#')) col = Rgb.Parse(tok[1..]) ?? Rgb.Black;
            else if (Length(tok) is double lw && char.IsDigit(tok[0]) | tok[0] == '.') w = lw;
            else if (tok is "solid" or "double" or "dotted" or "dashed" or "dot-dash" or "dot-dot-dash" or "groove" or "ridge" or "inset" or "outset" or "fine-dashed" or "dash-dot" or "dash-dot-dot" or "double-thin") style = tok;
            else if (tok is "thin") w = 0.5; else if (tok is "medium") w = 1.5; else if (tok is "thick") w = 3;
            else if (tok is "none" or "hidden") return null;
        }
        if (style == "double" && lineWidth != null)
        {
            double tot = lineWidth.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(t => Length(t) ?? 0).Sum();
            if (tot > 0) w = tot;
        }
        return new BorderDef
        {
            Val = style switch { "double" or "double-thin" => "double", "dotted" => "dotted", "dashed" or "fine-dashed" => "dashed", "dot-dash" or "dash-dot" => "dotDash", "dot-dot-dash" or "dash-dot-dot" => "dotDotDash", _ => "single" },
            Width = Math.Max(0.35, w), Color = col, Space = Length(padding) ?? 0,
        };
    }

    /// <summary>Рамка стороны: fo:border-{side} или общая fo:border; отступ — fo:padding-{side}/fo:padding.</summary>
    internal static BorderDef? SideBorder(Dictionary<string, string> d, string side)
    {
        string? v = d.GetValueOrDefault("fo:border-" + side) ?? d.GetValueOrDefault("fo:border");
        string? lw = d.GetValueOrDefault("style:border-line-width-" + side) ?? d.GetValueOrDefault("style:border-line-width");
        string? pad = d.GetValueOrDefault("fo:padding-" + side) ?? d.GetValueOrDefault("fo:padding");
        return Border(v, lw, pad);
    }

    internal static double Padding(Dictionary<string, string> d, string side) =>
        Length(d.GetValueOrDefault("fo:padding-" + side) ?? d.GetValueOrDefault("fo:padding")) ?? 0;

    // ================================================================
    //  Чтение документа
    // ================================================================

    XElement? _body;
    readonly List<Section> _sections = [];
    List<Block> _blocks = [];
    SectionProps _curProps = new();
    string? _curMaster;
    bool _tabsRelativeToIndent = true;

    DocxDocument ReadAll()
    {
        XElement? stylesRoot, contentRoot, metaRoot, settingsRoot;
        if (_flat != null)
        {
            var r = _flat.Root!;
            stylesRoot = r; contentRoot = r; metaRoot = r.Element(Office + "meta") != null ? r : null; settingsRoot = r.Element(Office + "settings") != null ? r : null;
        }
        else
        {
            if (_zip!.GetEntry("content.xml") is not { } ce) throw new InvalidDataException("В пакете нет content.xml — это не ODT?");
            try { using var cs = ce.Open(); contentRoot = XDocument.Load(cs, LoadOptions.PreserveWhitespace).Root!; }
            catch (System.Xml.XmlException xe) { throw new InvalidDataException("content.xml повреждён: " + xe.Message); }
            stylesRoot = LoadXml("styles.xml")?.Root;
            metaRoot = LoadXml("meta.xml")?.Root;
            settingsRoot = LoadXml("settings.xml")?.Root;
        }

        LoadFontFaces(stylesRoot?.Element(Office + "font-face-decls"));
        LoadFontFaces(contentRoot.Element(Office + "font-face-decls"));
        LoadStyleContainer(stylesRoot?.Element(Office + "styles"), _common, _listStyles, common: true);
        LoadStyleContainer(stylesRoot?.Element(Office + "automatic-styles"), _flat != null ? _autoContent : _autoStyles, _flat != null ? _listStyles : _listStylesAutoStyles, common: false);
        if (_flat == null) LoadStyleContainer(contentRoot.Element(Office + "automatic-styles"), _autoContent, _listStyles, common: false);
        if (_flat != null) foreach (var (k, v) in _autoContent) _autoStyles[k] = v;
        foreach (var mp in stylesRoot?.Element(Office + "master-styles")?.Elements(St + "master-page") ?? [])
            if ((string?)mp.Attribute(St + "name") is { } mn) { _masterPages[mn] = mp; _firstMaster ??= mn; }

        if (metaRoot?.Element(Office + "meta") is { } meta)
        {
            _doc.Title = NonEmpty(meta.Element(Dc + "title")?.Value);
            _doc.Author = NonEmpty(meta.Element(Dc + "creator")?.Value ?? meta.Element(Meta + "initial-creator")?.Value);
            _doc.Subject = NonEmpty(meta.Element(Dc + "subject")?.Value);
        }
        foreach (var item in settingsRoot?.Descendants(Cfg + "config-item") ?? [])
            if ((string?)item.Attribute(Cfg + "name") == "TabsRelativeToIndent") _tabsRelativeToIndent = item.Value.Trim() != "false";
        if (_defaults.TryGetValue("paragraph", out var dps) && Length(dps.Para.GetValueOrDefault("style:tab-stop-distance")) is double tsd && tsd > 0)
            _doc.DefaultTabStop = tsd;

        _body = contentRoot.Element(Office + "body")?.Element(Office + "text") ?? throw new InvalidDataException("В документе нет office:text — это не текстовый документ ODF");

        // Первый стиль страницы — у первого абзаца/таблицы (через стиль), иначе «Standard»/первый
        string startMaster = FirstMasterOf(_body) ?? (_masterPages.ContainsKey("Standard") ? "Standard" : _masterPages.ContainsKey("Default") ? "Default" : _firstMaster ?? "");
        StartSection(startMaster, null, first: true);

        // Исправления: вставки и удаления (текст удалений хранится в text:tracked-changes)
        foreach (var cr in _body.Element(Tx + "tracked-changes")?.Elements(Tx + "changed-region") ?? [])
        {
            string? id = (string?)cr.Attribute(Tx + "id") ?? (string?)cr.Attribute(Xml + "id");
            var ch = cr.Elements().FirstOrDefault();
            if (id == null || ch == null) continue;
            string author = ch.Element(Office + "change-info")?.Element(Dc + "creator")?.Value ?? "";
            _changes[id] = (ch.Name.LocalName, author, ch);
        }
        _markup = _changes.Count > 0 && _options.Revisions != "final";

        var ctx = new Ctx();
        ParseBlocks(_body, ctx, _blocks, top: true);
        AppendEndnotes();
        CloseSection();
        _doc.Sections.AddRange(_sections);
        if (_doc.Sections.Count == 0) _doc.Sections.Add(new Section { Props = _curProps, Blocks = [new Paragraph()] });
        foreach (var s in _doc.Sections) if (s.Blocks.Count == 0) s.Blocks.Add(new Paragraph { MarkStyle = new TextStyle { Size = 1 }, LineRule = "exact", LineValue = 1 });
        return _doc;
    }

    static string? NonEmpty(string? s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    string? FirstMasterOf(XElement body)
    {
        foreach (var e in body.Elements())
        {
            if (e.Name == Tx + "p" || e.Name == Tx + "h")
                return MasterOfStyle(FindStyle((string?)e.Attribute(Tx + "style-name"), "paragraph", new Ctx()));
            if (e.Name == Tb + "table")
                return MasterOfStyle(FindStyle((string?)e.Attribute(Tb + "style-name"), "table", new Ctx()));
            if (e.Name == Tx + "list")
            {
                var p = e.Descendants().FirstOrDefault(x => x.Name == Tx + "p" || x.Name == Tx + "h");
                return p != null ? MasterOfStyle(FindStyle((string?)p.Attribute(Tx + "style-name"), "paragraph", new Ctx())) : null;
            }
            if (e.Name == Tx + "section" || e.Name == Tx + "table-of-content" || e.Name.LocalName.EndsWith("-index"))
            {
                var r = FirstMasterOf(e.Name == Tx + "section" ? e : e.Element(Tx + "index-body") ?? e);
                if (r != null) return r;
                continue;
            }
            if (e.Name.Namespace == Tx && e.Name.LocalName is "sequence-decls" or "variable-decls" or "user-field-decls" or "tracked-changes" or "soft-page-break") continue;
            if (e.Name.Namespace == Office && e.Name.LocalName == "forms") continue;
            if (e.Name.Namespace == Dr) continue;
        }
        return null;
    }

    string? MasterOfStyle(OStyle? s)
    {
        for (int i = 0; s != null && i < 32; i++)
        {
            if (s.MasterPage is { Length: > 0 } m) return m;
            if (s.MasterPage == "") return null;
            s = s.Parent != null ? FindStyle(s.Parent, s.Family, new Ctx()) : null;
        }
        return null;
    }

    void StartSection(string master, int? pageNumber, bool first = false, string type = "nextPage", XElement? columns = null)
    {
        _curMaster = master;
        _curProps = PageProps(master);
        _curProps.Type = type;
        if (pageNumber != null) _curProps.PageNumberStart = pageNumber;
        if (columns != null) ApplyColumns(_curProps, columns);
        _blocks = [];
    }

    void CloseSection()
    {
        _sections.Add(new Section { Props = _curProps, Blocks = _blocks });
    }

    /// <summary>Новый раздел с тем же стилем страницы (смена колонок) — продолжается на той же странице.</summary>
    void ContinuousSection(XElement? columns)
    {
        if (_blocks.Count > 0) CloseSection();
        else if (_sections.Count == 0 && _curProps.Type != "continuous") { /* ещё ничего не выведено — меняем текущий */ }
        var keepType = _blocks.Count == 0 && _sections.Count == 0 ? _curProps.Type : "continuous";
        var pn = _blocks.Count == 0 && _sections.Count == 0 ? _curProps.PageNumberStart : null;
        StartSection(_curMaster ?? "", pn, type: keepType, columns: columns);
    }

    static void ApplyColumns(SectionProps p, XElement cols)
    {
        int n = int.TryParse((string?)cols.Attribute(Fo + "column-count"), out var c) ? c : 1;
        p.Columns = Math.Max(1, n);
        p.ColumnSpace = Length((string?)cols.Attribute(Fo + "column-gap")) ?? 0;
        p.ColumnSeparator = cols.Element(St + "column-sep") is { } sep && (string?)sep.Attribute(St + "style") != "none";
        var explicitCols = cols.Elements(St + "column").ToList();
        if (n > 1 && explicitCols.Count == n)
        {
            // Относительные ширины (rel-width «1234*») и поля колонок
            double total = explicitCols.Sum(e => RelWidth((string?)e.Attribute(St + "rel-width")));
            var starts = explicitCols.Select(e => Length((string?)e.Attribute(Fo + "start-indent")) ?? 0).ToList();
            var ends = explicitCols.Select(e => Length((string?)e.Attribute(Fo + "end-indent")) ?? 0).ToList();
            bool equal = explicitCols.Select(e => RelWidth((string?)e.Attribute(St + "rel-width"))).Distinct().Count() == 1;
            if (!equal && total > 0)
            {
                p.ColumnWidths = [];
                for (int i = 0; i < n; i++)
                {
                    double share = RelWidth((string?)explicitCols[i].Attribute(St + "rel-width")) / total;
                    p.ColumnWidths.Add((share, i < n - 1 ? ends[i] + starts[i + 1] : 0));
                }
            }
            else if (p.ColumnSpace == 0 && n > 1) p.ColumnSpace = ends[0] + starts[1];
        }
    }

    static double RelWidth(string? v) => v != null && double.TryParse(v.TrimEnd('*'), NumberStyles.Float, CultureInfo.InvariantCulture, out var r) ? r : 1;

    /// <summary>Свойства страницы по стилю страницы (макет, колонтитулы; следующий стиль — для первой страницы).</summary>
    SectionProps PageProps(string master)
    {
        var p = new SectionProps();
        if (!_masterPages.TryGetValue(master, out var mp)) return p;
        FillLayout(p, mp);
        // Стиль первой страницы со сменой на другой (style:next-style-name): первая страница — колонтитулы этого стиля
        string? next = (string?)mp.Attribute(St + "next-style-name");
        if (next != null && next != master && _masterPages.TryGetValue(next, out var nmp))
        {
            var np = new SectionProps();
            FillLayout(np, nmp);
            p.TitlePage = true;
            p.HeaderFirst = p.HeaderDefault ?? new HeaderFooter();
            p.FooterFirst = p.FooterDefault ?? new HeaderFooter();
            p.HeaderDefault = np.HeaderDefault; p.FooterDefault = np.FooterDefault;
            p.HeaderEven = np.HeaderEven; p.FooterEven = np.FooterEven;
        }
        return p;
    }

    void FillLayout(SectionProps p, XElement mp)
    {
        var ctx = new Ctx { HeaderFooter = true };
        if (_pageLayouts.TryGetValue((string?)mp.Attribute(St + "page-layout-name") ?? "", out var pl))
        {
            var lp = new Dictionary<string, string>();
            ReadProps(pl.Element(St + "page-layout-properties"), lp);
            p.PageWidth = Len(lp, "fo:page-width", p.PageWidth);
            p.PageHeight = Len(lp, "fo:page-height", p.PageHeight);
            if (lp.GetValueOrDefault("style:print-orientation") == "landscape" && p.PageWidth < p.PageHeight) (p.PageWidth, p.PageHeight) = (p.PageHeight, p.PageWidth);
            double mt = Len(lp, "fo:margin-top", 56.7), mb = Len(lp, "fo:margin-bottom", 56.7);
            p.MarginLeft = Len(lp, "fo:margin-left", 56.7);
            p.MarginRight = Len(lp, "fo:margin-right", 56.7);
            // Рамка страницы: отступы считаются от текста
            p.PageBorderTop = SideBorder(lp, "top"); p.PageBorderBottom = SideBorder(lp, "bottom");
            p.PageBorderLeft = SideBorder(lp, "left"); p.PageBorderRight = SideBorder(lp, "right");
            p.PageBorderFromText = true;
            // Рамка страницы стоит на поле; текст — внутри рамки и отступа (padding)
            p.MarginLeft += Padding(lp, "left") + (p.PageBorderLeft?.Width ?? 0);
            p.MarginRight += Padding(lp, "right") + (p.PageBorderRight?.Width ?? 0);
            mt += Padding(lp, "top") + (p.PageBorderTop?.Width ?? 0);
            mb += Padding(lp, "bottom") + (p.PageBorderBottom?.Width ?? 0);
            if (Color(lp.GetValueOrDefault("fo:background-color")) is Rgb bg) _doc.Background ??= bg;
            p.PageNumberFormat = lp.GetValueOrDefault("style:num-format") switch { "i" => "lowerRoman", "I" => "upperRoman", "a" => "lowerLetter", "A" => "upperLetter", _ => "decimal" };
            if (pl.Element(St + "page-layout-properties")?.Element(St + "columns") is { } cols) ApplyColumns(p, cols);
            // Сетка строк (азиатская вёрстка)
            if (lp.GetValueOrDefault("style:layout-grid-mode") is "line" or "both" && Length(lp.GetValueOrDefault("style:layout-grid-base-height")) is double gh)
                p.LinePitch = gh + (Length(lp.GetValueOrDefault("style:layout-grid-ruby-height")) ?? 0);

            // Колонтитулы: поле страницы — до колонтитула; тело начинается после высоты колонтитула и отступа
            var hp = new Dictionary<string, string>(); ReadProps(pl.Element(St + "header-style")?.Element(St + "header-footer-properties"), hp);
            var fp = new Dictionary<string, string>(); ReadProps(pl.Element(St + "footer-style")?.Element(St + "header-footer-properties"), fp);
            bool hasHeader = mp.Element(St + "header") is { } he && (string?)he.Attribute(St + "display") != "false";
            bool hasFooter = mp.Element(St + "footer") is { } fe && (string?)fe.Attribute(St + "display") != "false";
            p.HeaderDistance = mt; p.FooterDistance = mb;
            p.MarginTop = hasHeader ? mt + Len(hp, "fo:min-height", Len(hp, "svg:height", 0)) + (hp.ContainsKey("fo:min-height") || hp.ContainsKey("svg:height") ? 0 : Len(hp, "fo:margin-bottom", 0)) : mt;
            p.MarginBottom = hasFooter ? mb + Len(fp, "fo:min-height", Len(fp, "svg:height", 0)) + (fp.ContainsKey("fo:min-height") || fp.ContainsKey("svg:height") ? 0 : Len(fp, "fo:margin-top", 0)) : mb;
            if (hasHeader && Len(hp, "fo:min-height", 0) == 0) p.MarginTop = mt + 14 + Len(hp, "fo:margin-bottom", 0);
            if (hasFooter && Len(fp, "fo:min-height", 0) == 0) p.MarginBottom = mb + 14 + Len(fp, "fo:margin-top", 0);
            // Отступы колонтитула слева/справа
            _hfIndent = (Len(hp, "fo:margin-left", 0), Len(hp, "fo:margin-right", 0));
            // Рамка страницы охватывает и колонтитулы: верх/низ рамки — на поле страницы
            double rawT = Len(lp, "fo:margin-top", 56.7), rawB = Len(lp, "fo:margin-bottom", 56.7);
            if (p.PageBorderTop != null) p.PageBorderTop.Space = Math.Max(0, p.MarginTop - rawT - p.PageBorderTop.Width / 2);
            if (p.PageBorderBottom != null) p.PageBorderBottom.Space = Math.Max(0, p.MarginBottom - rawB - p.PageBorderBottom.Width / 2);
            if (p.PageBorderLeft != null) p.PageBorderLeft.Space += p.PageBorderLeft.Width / 2;
            if (p.PageBorderRight != null) p.PageBorderRight.Space += p.PageBorderRight.Width / 2;
        }
        HeaderFooter? HF(string name)
        {
            var e = mp.Element(St + name);
            if (e == null || (string?)e.Attribute(St + "display") == "false") return null;
            var hf = new HeaderFooter();
            ParseBlocks(e, ctx, hf.Blocks, top: false);
            if (_hfIndent.L > 0 || _hfIndent.R > 0)
                foreach (var b in hf.Blocks.OfType<Paragraph>()) { b.IndLeft += _hfIndent.L; b.IndRight += _hfIndent.R; }
            return hf;
        }
        p.HeaderDefault = HF("header");
        p.FooterDefault = HF("footer");
        var hl = HF("header-left"); var fl = HF("footer-left");
        if (hl != null || fl != null)
        {
            _doc.EvenAndOddHeaders = true;
            p.HeaderEven = hl ?? (mp.Element(St + "header-left") == null ? p.HeaderDefault : null);
            p.FooterEven = fl ?? (mp.Element(St + "footer-left") == null ? p.FooterDefault : null);
        }
        var hfirst = HF("header-first"); var ffirst = HF("footer-first");
        if (hfirst != null || ffirst != null)
        {
            p.TitlePage = true;
            p.HeaderFirst = hfirst ?? p.HeaderDefault;
            p.FooterFirst = ffirst ?? p.FooterDefault;
        }
        if (_lineNumbering is { } ln && (string?)ln.Attribute(Tx + "number-lines") != "false")
        {
            p.LineNumberCountBy = int.TryParse((string?)ln.Attribute(Tx + "increment"), out var inc) ? inc : 1;
            p.LineNumberDistance = Length((string?)ln.Attribute(Tx + "offset"));
            p.LineNumberRestart = (string?)ln.Attribute(Tx + "restart-on-page") == "true" ? "newPage" : "continuous";
        }
    }

    (double L, double R) _hfIndent;

    // ================================================================
    //  Блоки
    // ================================================================

    List<Block>? _pendingAnchors;                    // плавающие объекты текущего абзаца (вставляются перед ним)
    bool _pendingBreak;                              // fo:break-after предыдущего абзаца

    void ParseBlocks(XElement container, Ctx ctx, List<Block> output, bool top)
    {
        foreach (var el in container.Elements())
            ParseBlock(el, ctx, output, top);
    }

    /// <summary>Текущий список вывода: на верхнем уровне он меняется при смене раздела.</summary>
    List<Block> Out(List<Block> output, bool top) => top ? _blocks : output;

    void ParseBlock(XElement el, Ctx ctx, List<Block> output, bool top)
    {
        var n = el.Name;
        if (n == Tx + "p" || n == Tx + "h")
        {
            ParseParagraphEl(el, ctx, output, top, null);
        }
        else if (n == Tx + "list")
        {
            ParseList(el, ctx, output, top, 0, null);
        }
        else if (n == Tb + "table")
        {
            var style = FindStyle((string?)el.Attribute(Tb + "style-name"), "table", ctx);
            if (top) CheckMaster(style, ctx);
            var t = ParseTable(el, ctx);
            if (_pendingBreak) { t.PageBreakBefore = true; _pendingBreak = false; }
            Out(output, top).Add(t);
        }
        else if (n == Tx + "section")
        {
            var style = FindStyle((string?)el.Attribute(Tx + "style-name"), "section", ctx);
            if ((string?)el.Attribute(Tx + "display") == "none") return;
            bool cols = top && style?.Columns is { } c && int.TryParse((string?)c.Attribute(Fo + "column-count"), out var cc) && cc > 1;
            var sectionBlocks = new List<Block>();
            if (cols)
            {
                int prevCols = _curProps.Columns;
                var prevColsEl = new XElement(St + "columns", new XAttribute(Fo + "column-count", prevCols));
                ContinuousSection(style!.Columns);
                ParseBlocks(el, ctx, output, top);
                ContinuousSection(prevCols > 1 ? null : prevColsEl);
            }
            else
            {
                // Раздел без колонок: фон и отступы раздела — к абзацам
                int start = Out(output, top).Count;
                ParseBlocks(el, ctx, output, top);
                var list = Out(output, top);
                double ml = style != null ? Len(style.Section, "fo:margin-left") : 0, mr = style != null ? Len(style.Section, "fo:margin-right") : 0;
                var bg = style != null ? Color(style.Section.GetValueOrDefault("fo:background-color")) : null;
                for (int i = start; i < list.Count; i++)
                    if (list[i] is Paragraph p) { p.IndLeft += ml; p.IndRight += mr; if (bg != null) p.Shading ??= bg; }
                    else if (list[i] is Table t) t.Indent += ml;
            }
        }
        else if (n == Tx + "table-of-content" || n.Namespace == Tx && n.LocalName is "illustration-index" or "table-index" or "object-index" or "user-index" or "alphabetical-index" or "bibliography")
        {
            if (el.Element(Tx + "index-body") is { } ib) ParseBlocks(ib, ctx, output, top);
        }
        else if (n == Tx + "index-title")
        {
            ParseBlocks(el, ctx, output, top);
        }
        else if (n == Tx + "change-start" || n == Tx + "change-end" || n == Tx + "change")
        {
            ChangeMark(el);
            if (n == Tx + "change" && _markup && _changes.TryGetValue((string?)el.Attribute(Tx + "change-id") ?? "", out var del) && del.Kind == "deletion")
            {
                // Удалённые абзацы целиком — зачёркнутыми абзацами
                var saved = _revision; var savedA = _revisionAuthor;
                _revision = 2; _revisionAuthor = del.Author;
                foreach (var c in del.El.Elements().Where(e => e.Name != Office + "change-info")) ParseBlock(c, ctx, output, top);
                _revision = saved; _revisionAuthor = savedA;
            }
        }
        else if (n == Tx + "soft-page-break" || n == Tx + "sequence-decls" || n == Tx + "variable-decls" || n == Tx + "user-field-decls"
                 || n == Tx + "tracked-changes" || n == Office + "forms" || n == Tx + "bookmark" || n == Tx + "bookmark-start" || n == Tx + "bookmark-end")
        {
        }
        else if (n.Namespace == Dr)
        {
            // Объекты, привязанные к странице (или стоящие между абзацами): плавающие блоки перед следующим текстом
            var saved = _pendingAnchors;
            var pend = new List<Block>();
            _pendingAnchors = pend;
            var tmp = new List<Inline>();
            ParseDrawing(el, ctx, DefaultTextStyle(ctx), tmp, blockLevel: true);
            _pendingAnchors = saved;
            var o = Out(output, top);
            o.AddRange(pend);
            if (tmp.Count > 0)
            {
                var p = new Paragraph { MarkStyle = DefaultTextStyle(ctx) };
                p.Inlines.AddRange(tmp);
                o.Add(p);
            }
        }
        else if (n == Tx + "illustration-index-source" || n == Tx + "table-of-content-source") { }
        else if (n == Tx + "list-item" || n == Tx + "list-header") ParseBlocks(el, ctx, output, top);
        else if (n.Namespace == Tx && el.HasElements) ParseBlocks(el, ctx, output, top);
    }

    /// <summary>Смена стиля страницы по стилю абзаца/таблицы (style:master-page-name) — новый раздел.</summary>
    void CheckMaster(OStyle? style, Ctx ctx)
    {
        if (ctx.HeaderFooter || ctx.InNote || ctx.TableDepth > 0 || style == null) return;
        // Только у автоматического или прямо назначенного стиля (наследуемый master-page не действует)
        string? master = style.MasterPage;
        if (master == null && style.Parent != null)
        {
            // LibreOffice пишет master-page-name в автоматический стиль; у общих стилей (например, «Заголовок») — тоже действует
            master = MasterOfStyle(style);
        }
        if (string.IsNullOrEmpty(master)) return;
        int? pn = null;
        if (style.Para.TryGetValue("style:page-number", out var pnv) && int.TryParse(pnv, out var pni)) pn = pni;
        if (_blocks.Count == 0 && _sections.Count == 0)
        {
            // Первый абзац документа: стиль страницы уже выбран
            if (pn != null) _curProps.PageNumberStart = pn;
            return;
        }
        CloseSection();
        StartSection(master, pn);
        _pendingBreak = false;
    }

    TextStyle DefaultTextStyle(Ctx ctx) => ToTextStyle(TextPropsOf(null, ctx));

    // ---------------- Абзацы ----------------

    internal sealed class ListCtx
    {
        public required XElement Style;              // text:list-style
        public int Level;                             // 1..10
        public bool Header;                           // text:list-header — без метки
        public int? StartValue;                       // text:start-value элемента списка
        public required ListInstance Inst;
        public bool FirstParaInItem = true;
    }

    internal sealed class ListInstance
    {
        public string StyleName = "";
        public readonly int[] Counters = new int[11];
        public readonly bool[] Started = new bool[11];
        public string? Id;
    }

    readonly List<ListInstance> _lists = [];
    readonly ListInstance _outline = new() { StyleName = "#outline" };

    void ParseParagraphEl(XElement el, Ctx ctx, List<Block> output, bool top, ListCtx? list)
    {
        var style = FindStyle((string?)el.Attribute(Tx + "style-name"), "paragraph", ctx);
        if (top && list == null) CheckMaster(style, ctx);
        else if (top && list != null) CheckMaster(style, ctx);

        var saved = _pendingAnchors;
        var pend = new List<Block>();
        _pendingAnchors = pend;
        var para = BuildParagraph(el, style, ctx, list);
        _pendingAnchors = saved;

        var o = Out(output, top);
        if (_pendingBreak) { para.PageBreakBefore = true; _pendingBreak = false; }
        // Объекты абзаца, привязанные к нему, — перед ним (как в DOCX)
        if (pend.Count > 0 && para.PageBreakBefore)
        {
            // Разрыв страницы должен быть перед объектами
            if (pend[0] is Paragraph pp0) pp0.PageBreakBefore = true;
            else o.Add(new Paragraph { PageBreakBefore = true, MarkStyle = new TextStyle { Size = 1 }, LineRule = "exact", LineValue = 0.1 });
            para.PageBreakBefore = false;
        }
        o.AddRange(pend);

        if (para.DropLines > 0 && _dropText is { } dt)
        {
            o.Add(new FloatBlock { Side = "left", Width = 0, DropLines = para.DropLines, MarginRight = para.DropGap, Blocks = [dt] });
            _dropText = null;
        }
        o.Add(para);
        if (_breakAfter) { _pendingBreak = true; _breakAfter = false; }
    }

    Paragraph? _dropText;
    bool _breakAfter;

    Paragraph BuildParagraph(XElement el, OStyle? style, Ctx ctx, ListCtx? list)
    {
        var (pp, tp, tabsEl, dropEl, styleList, styleOutline) = ParaPropsOf(style, ctx);
        if (ctx.ShapeText is { } stx)
        {
            OverlayText(tp, stx.Text);
            foreach (var (k, v) in stx.Para) if (k is "fo:text-align" or "fo:line-height") pp[k] = v;
        }
        // Заголовок: text:h, либо абзац со стилем уровня структуры (LibreOffice нумерует и такие)
        bool isHeading = el.Name == Tx + "h" || styleOutline is > 0 && (string?)el.Attribute(Tx + "outline-level") != "0";
        int outlineLevel = isHeading ? (int.TryParse((string?)el.Attribute(Tx + "outline-level"), out var olv) ? olv : styleOutline ?? 1) : 0;

        var para = new Paragraph { StyleId = style?.DisplayName ?? style?.Name ?? "" };
        para.Align = (pp.GetValueOrDefault("fo:text-align") ?? "start") switch
        {
            "center" => "center",
            "end" or "right" => pp.GetValueOrDefault("style:writing-mode") is "rl-tb" or "rl" ? "left" : "right",
            "justify" => "both",
            "start" or "left" => pp.GetValueOrDefault("style:writing-mode") is "rl-tb" or "rl" ? "right" : "left",
            _ => "left",
        };
        para.Rtl = pp.GetValueOrDefault("style:writing-mode") is "rl-tb" or "rl";
        if (para.Rtl && para.Align == "right" && (pp.GetValueOrDefault("fo:text-align") ?? "start") is "start") para.Align = "left";
        para.IndLeft = Len(pp, "fo:margin-left");
        para.IndRight = Len(pp, "fo:margin-right");
        para.IndFirst = Len(pp, "fo:text-indent");
        para.SpaceBefore = Math.Max(0, Len(pp, "fo:margin-top"));
        para.SpaceAfter = Math.Max(0, Len(pp, "fo:margin-bottom"));
        para.ContextualSpacing = pp.GetValueOrDefault("style:contextual-spacing") == "true";
        var markProps = new Dictionary<string, string>(tp);
        double fontSize = Len(markProps, "fo:font-size", 12);
        string lh = pp.GetValueOrDefault("fo:line-height") ?? "normal";
        if (lh.EndsWith('%') && double.TryParse(lh[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var lpct)) { para.LineRule = "auto"; para.LineValue = 240 * lpct / 100; }
        else if (Length(lh) is double lfix) { para.LineRule = "exact"; para.LineValue = lfix; }
        if (Length(pp.GetValueOrDefault("style:line-height-at-least")) is double lal) { para.LineRule = "atLeast"; para.LineValue = lal; }
        if (Length(pp.GetValueOrDefault("style:line-spacing")) is double ls && ls != 0)
        {
            // Интерлиньяж: обычная высота строки плюс добавка
            para.LineRule = "auto"; para.LineValue = 240 * (1 + ls / (fontSize * 1.17));
        }
        para.KeepNext = pp.GetValueOrDefault("fo:keep-with-next") == "always";
        para.KeepLines = pp.GetValueOrDefault("fo:keep-together") == "always";
        para.WidowControl = !(pp.GetValueOrDefault("fo:widows") == "0" && pp.GetValueOrDefault("fo:orphans") == "0");
        if (pp.GetValueOrDefault("fo:break-before") == "page" && !ctx.HeaderFooter && !ctx.InNote && ctx.TableDepth == 0) para.PageBreakBefore = true;
        if (pp.GetValueOrDefault("fo:break-after") == "page" && !ctx.HeaderFooter && !ctx.InNote && ctx.TableDepth == 0) _breakAfter = true;
        para.Shading = Color(pp.GetValueOrDefault("fo:background-color"));
        para.BorderTop = SideBorder(pp, "top"); para.BorderBottom = SideBorder(pp, "bottom");
        para.BorderLeft = SideBorder(pp, "left"); para.BorderRight = SideBorder(pp, "right");
        if (para.BorderTop == null && para.BorderBottom == null && para.BorderLeft == null && para.BorderRight == null && para.Shading != null)
        {
            // Фон с отступами (padding) без рамки
            double pad = Padding(pp, "left");
            if (pad > 0) { para.BorderLeft = new BorderDef { Val = "nil", Width = 0, Space = pad }; para.BorderRight = new BorderDef { Val = "nil", Width = 0, Space = Padding(pp, "right") }; }
        }
        para.SnapToGrid = pp.GetValueOrDefault("style:snap-to-layout-grid") != "false";
        // Параметры переносов (свойства текста абзаца и абзаца LibreOffice)
        if (int.TryParse(tp.GetValueOrDefault("fo:hyphenation-remain-char-count"), out var hrc)) para.HyphenMinLeft = hrc;
        if (int.TryParse(tp.GetValueOrDefault("fo:hyphenation-push-char-count"), out var hpc)) para.HyphenMinRight = hpc;
        if (int.TryParse(pp.GetValueOrDefault("fo:hyphenation-ladder-count"), out var hlc)) para.HyphenLimit = hlc;
        if (int.TryParse(tp.GetValueOrDefault("loext:hyphenation-word-char-count") ?? pp.GetValueOrDefault("loext:hyphenation-word-char-count"), out var hmw) && hmw > 0) para.HyphenMinWord = hmw;
        if ((tp.GetValueOrDefault("loext:hyphenation-no-caps") ?? pp.GetValueOrDefault("loext:hyphenation-no-caps")) == "true") para.HyphenNoCaps = true;
        // Зона переноса LibreOffice: число без единиц — в твипах, «no-limit» — без зоны
        string? hzs = tp.GetValueOrDefault("loext:hyphenation-zone") ?? pp.GetValueOrDefault("loext:hyphenation-zone");
        if (hzs != null) para.HyphenZone = int.TryParse(hzs, out var hzt) ? hzt / 20.0 : Length(hzs) ?? 0;
        if (pp.GetValueOrDefault("text:number-lines") == "false") para.SuppressLineNumbers = true;

        // Метка списка / номер заголовка
        XElement? lvl = null;
        string label = "";
        ListInstance? inst = null;
        int level = 0;
        if (list != null)
        {
            var lstyle = list.Style;
            if (styleList != null && FindListStyle(styleList, ctx) is { } sl && (string?)list.Style.Attribute(St + "name") is null) lstyle = sl;
            level = list.Level;
            lvl = LevelStyle(lstyle, level);
            inst = list.Inst;
            if (list.Header || !list.FirstParaInItem) lvl = lvl == null ? null : lvl;
        }
        else if (isHeading && _outlineStyle != null && (string?)el.Attribute(Tx + "is-list-header") != "true")
        {
            level = Math.Clamp(outlineLevel, 1, 10);
            lvl = LevelStyle(_outlineStyle, level);
            inst = _outline;
            if (lvl != null && string.IsNullOrEmpty((string?)lvl.Attribute(St + "num-format")) && lvl.Name == Tx + "outline-level-style" && (string?)lvl.Attribute(St + "num-prefix") == null && (string?)lvl.Attribute(St + "num-suffix") == null)
            {
                // Без нумерации: счётчик всё равно идёт (для «1.1» у нижних уровней)
                Count(inst, level, null, lvl);
                lvl = null;
            }
        }

        TextStyle markStyle = ToTextStyle(markProps);
        para.MarkStyle = markStyle;
        bool hasLabel = false;
        if (lvl != null && inst != null)
        {
            bool numbered = lvl.Name == Tx + "list-level-style-number" || lvl.Name == Tx + "outline-level-style";
            bool showLabel = list == null || (!list.Header && list.FirstParaInItem);
            if (showLabel)
            {
                if (numbered)
                {
                    Count(inst, level, list?.StartValue, lvl);
                    label = NumberLabel(inst, level, lvl, list?.Style ?? _outlineStyle!);
                }
                else if (lvl.Name == Tx + "list-level-style-bullet") label = (string?)lvl.Attribute(Tx + "bullet-char") ?? "•";
                else if (lvl.Name == Tx + "list-level-style-image") label = "•";
                if (list != null) list.FirstParaInItem = false;
            }
            // Отступы уровня списка
            var lp = lvl.Element(St + "list-level-properties");
            var directPara = style != null && (style.El.Parent?.Name == Office + "automatic-styles") ? style.Para : null;
            string mode = (string?)lp?.Attribute(Tx + "list-level-position-and-space-mode") ?? "label-width-and-position";
            string follow = "listtab";
            double? tabPos = null;
            if (mode == "label-alignment" && lp?.Element(St + "list-level-label-alignment") is { } la)
            {
                double ml = Length((string?)la.Attribute(Fo + "margin-left")) ?? 0;
                double ti = Length((string?)la.Attribute(Fo + "text-indent")) ?? 0;
                // Прямые отступы абзаца (в автоматическом стиле) важнее отступов уровня
                para.IndLeft = directPara?.ContainsKey("fo:margin-left") == true ? Len(directPara, "fo:margin-left") : ml;
                para.IndFirst = directPara?.ContainsKey("fo:text-indent") == true ? Len(directPara, "fo:text-indent") : ti;
                follow = (string?)la.Attribute(Tx + "label-followed-by") ?? "listtab";
                tabPos = Length((string?)la.Attribute(Tx + "list-tab-stop-position"));
                // Следующие абзацы пункта (без метки) начинаются от позиции текста
                if (list != null && !showLabel && !list.Header && directPara?.ContainsKey("fo:text-indent") != true) para.IndFirst = 0;
            }
            else if (lp != null)
            {
                double sb = Length((string?)lp.Attribute(Tx + "space-before")) ?? 0;
                double mlw = Length((string?)lp.Attribute(Tx + "min-label-width")) ?? 0;
                double mld = Length((string?)lp.Attribute(Tx + "min-label-distance")) ?? 0;
                double baseInd = para.IndLeft;
                para.IndLeft = baseInd + sb + mlw;
                para.IndFirst = -mlw;
                if (mlw <= 0) { para.IndFirst = 0; follow = "space"; }
                tabPos = null;
                if (mld > 0 && mlw > 0) _labelMinDistance = mld;
                if ((string?)lp.Attribute(Fo + "text-align") is "end" or "right") para.LabelAlign = "right";
                else if ((string?)lp.Attribute(Fo + "text-align") is "center") para.LabelAlign = "center";
            }
            if (label.Length > 0 || numbered && showLabel)
            {
                // Оформление метки: стиль символов уровня + свойства текста уровня (шрифт маркера)
                var lprops = new Dictionary<string, string>(markProps);
                lprops.Remove("style:text-underline-style"); lprops.Remove("fo:background-color");
                if (FindStyle((string?)lvl.Attribute(Tx + "style-name"), "text", ctx) is { } cs)
                    foreach (var s in Chain(cs, ctx)) OverlayText(lprops, s.Text);
                var ltp = new Dictionary<string, string>();
                ReadProps(lvl.Element(St + "text-properties"), ltp);
                if (ltp.TryGetValue("fo:font-size", out var lfs) && lfs.EndsWith('%')) { OverlayText(lprops, new() { ["fo:font-size"] = lfs }); ltp.Remove("fo:font-size"); }
                OverlayText(lprops, ltp);
                if ((string?)lvl.Attribute(Tx + "bullet-relative-size") is { } brs && brs.EndsWith('%')) OverlayText(lprops, new() { ["fo:font-size"] = brs });
                var ls2 = ToTextStyle(lprops);
                if (lvl.Name == Tx + "list-level-style-bullet" && ls2.FontAscii is "OpenSymbol" or "StarSymbol" && label.Length == 1 && label[0] < 0x2000)
                    ls2 = ls2 with { FontAscii = markStyle.FontAscii, FontHAnsi = markStyle.FontHAnsi };
                if (label.Length > 0)
                {
                    para.Inlines.Add(new TextInline(label, ls2));
                    hasLabel = true;
                }
                if (hasLabel || numbered)
                {
                    if (follow == "listtab") para.Inlines.Add(new TabInline(ls2));
                    else if (follow == "space") para.Inlines.Add(new TextInline(" ", ls2));
                }
                if (follow == "listtab" && tabPos != null && hasLabel)
                {
                    // Позиция табуляции после метки (от поля; в режиме «относительно отступа» — от отступа абзаца)
                    double tpos = tabPos.Value;
                    if (tpos > para.IndLeft + para.IndFirst + 0.5) para.Tabs.Add(new TabStop(tpos, "left", "none"));
                }
            }
        }

        // Табуляция
        if (tabsEl != null)
        {
            double rel = _tabsRelativeToIndent ? para.IndLeft : 0;
            foreach (var t in tabsEl.Elements(St + "tab-stop"))
            {
                double pos = (Length((string?)t.Attribute(St + "position")) ?? 0) + rel;
                string align = (string?)t.Attribute(St + "type") switch { "right" => "right", "center" => "center", "char" => "decimal", _ => "left" };
                string? leaderText = (string?)t.Attribute(St + "leader-text");
                string leaderStyle = (string?)t.Attribute(St + "leader-style") ?? "none";
                string leader = leaderText switch
                {
                    "." => "dot", "-" => "hyphen", "_" => "underscore", "·" => "middleDot",
                    _ => leaderStyle switch { "dotted" => "dot", "dash" or "long-dash" => "hyphen", "solid" => "underscore", _ => "none" }
                };
                if (leaderText == null && leaderStyle == "none") leader = "none";
                para.Tabs.Add(new TabStop(pos, align, leader));
            }
        }
        para.Tabs = para.Tabs.OrderBy(t => t.Pos).ToList();

        // Содержимое
        var st = new InlineState { Ctx = ctx, ParaText = markProps, Para = para };
        ParseInlines(el, st, para.Inlines, markProps, null);
        para.Bookmarks = st.Bookmarks;
        TrimTrailingSpace(para.Inlines);

        // Буквица: первые символы абзаца — отдельный плавающий абзац крупным шрифтом
        if (dropEl != null && int.TryParse((string?)dropEl.Attribute(St + "lines"), out var dl) && dl > 1 && !ctx.HeaderFooter)
            MakeDropCap(para, dropEl, dl, markStyle, ctx);
        return para;
    }

    double _labelMinDistance;

    static void TrimTrailingSpace(List<Inline> inl)
    {
        for (int i = inl.Count - 1; i >= 0; i--)
        {
            if (inl[i] is TextInline t)
            {
                string s = t.Text.TrimEnd(' ');
                if (s.Length == t.Text.Length) return;
                t.Text = s;
                if (s.Length > 0) return;
                inl.RemoveAt(i);
                continue;
            }
            return;
        }
    }

    void MakeDropCap(Paragraph para, XElement dropEl, int lines, TextStyle markStyle, Ctx ctx)
    {
        string len = (string?)dropEl.Attribute(St + "length") ?? "1";
        double gap = Length((string?)dropEl.Attribute(St + "distance")) ?? 0;
        // Берём первые N символов (или первое слово) из начала текста
        int skip = 0;
        while (skip < para.Inlines.Count && para.Inlines[skip] is not TextInline) skip++;
        if (skip >= para.Inlines.Count || para.Inlines[skip] is not TextInline first) return;
        string text = first.Text;
        int count = len == "word" ? Math.Max(1, text.IndexOf(' ') is int sp && sp > 0 ? sp : text.Length) : int.TryParse(len, out var c) ? c : 1;
        count = Math.Min(count, text.Length);
        if (count <= 0) return;
        var fs = first.Style;
        if (FindStyle((string?)dropEl.Attribute(St + "style-name"), "text", ctx) is { } dcs)
        {
            var tp = TextPropsOf(dcs, ctx, new Dictionary<string, string>());
            var tmp = ToTextStyle(tp);
            fs = fs with { FontAscii = tp.ContainsKey("style:font-name") || tp.ContainsKey("fo:font-family") ? tmp.FontAscii : fs.FontAscii, FontHAnsi = tp.ContainsKey("style:font-name") || tp.ContainsKey("fo:font-family") ? tmp.FontHAnsi : fs.FontHAnsi, Bold = tp.ContainsKey("fo:font-weight") ? tmp.Bold : fs.Bold, Color = tp.ContainsKey("fo:color") ? tmp.Color : fs.Color };
        }
        // Высота прописной буквы ≈ высота N строк (до базовой линии последней)
        double lineH = markStyle.Size * 1.17;
        double size = ((lines - 1) * lineH + 0.7 * markStyle.Size) / 0.7;
        var dp = new Paragraph { MarkStyle = fs with { Size = size } };
        dp.Inlines.Add(new TextInline(text[..count], fs with { Size = size }));
        first.Text = text[count..];
        para.DropLines = lines;
        para.DropGap = gap;
        _dropText = dp;
    }

    // ---------------- Списки ----------------

    XElement? FindListStyle(string? name, Ctx ctx)
    {
        if (string.IsNullOrEmpty(name)) return null;
        if (ctx.HeaderFooter && _listStylesAutoStyles.TryGetValue(name, out var a)) return a;
        if (_listStyles.TryGetValue(name, out var l)) return l;
        if (_listStylesAutoStyles.TryGetValue(name, out a)) return a;
        return null;
    }

    static XElement? LevelStyle(XElement listStyle, int level)
    {
        foreach (var e in listStyle.Elements())
            if (int.TryParse((string?)e.Attribute(Tx + "level"), out var l) && l == level) return e;
        // Уровни выше заданных — последний заданный
        return listStyle.Elements().Where(e => e.Attribute(Tx + "level") != null).LastOrDefault();
    }

    static int StartOf(XElement? lvl) => int.TryParse((string?)lvl?.Attribute(Tx + "start-value"), out var s) ? s : 1;

    static void Count(ListInstance inst, int level, int? startValue, XElement? lvl)
    {
        if (startValue != null) { inst.Counters[level] = startValue.Value; inst.Started[level] = true; }
        else if (!inst.Started[level]) { inst.Counters[level] = StartOf(lvl); inst.Started[level] = true; }
        else inst.Counters[level]++;
        for (int l = level + 1; l < inst.Counters.Length; l++) inst.Started[l] = false;
    }

    static string NumberLabel(ListInstance inst, int level, XElement lvl, XElement listStyle)
    {
        string fmt = (string?)lvl.Attribute(St + "num-format") ?? "";
        bool legal = (string?)lvl.Attribute(Loext + "is-legal") == "true";
        // Шаблон LibreOffice 7+ («Глава %1%», «%1%.%2%.»): номера уровней подставляются в текст
        if ((string?)lvl.Attribute(Loext + "num-list-format") is { Length: > 0 } tmpl)
        {
            return System.Text.RegularExpressions.Regex.Replace(tmpl, @"%(\d+)%", m =>
            {
                int l = int.Parse(m.Groups[1].Value);
                if (l < 1 || l > level) return "";
                var pl = l == level ? lvl : LevelStyle(listStyle, l);
                string pf = (string?)pl?.Attribute(St + "num-format") ?? "1";
                if (l < level && (legal || pf.Length == 0)) pf = "1";
                int v = l == level || inst.Started[l] ? inst.Counters[l] : StartOf(pl);
                return FormatNum(v, l == level ? fmt : pf);
            });
        }
        string prefix = (string?)lvl.Attribute(St + "num-prefix") ?? "";
        string suffix = (string?)lvl.Attribute(St + "num-suffix") ?? "";
        int display = int.TryParse((string?)lvl.Attribute(Tx + "display-levels"), out var dl) ? Math.Clamp(dl, 1, level) : 1;
        var parts = new List<string>();
        for (int l = level - display + 1; l <= level; l++)
        {
            if (l == level) { if (fmt.Length > 0) parts.Add(FormatNum(inst.Counters[l], fmt)); continue; }
            var pl = LevelStyle(listStyle, l);
            string pf = (string?)pl?.Attribute(St + "num-format") ?? "1";
            int v = inst.Started[l] ? inst.Counters[l] : StartOf(pl);
            parts.Add(FormatNum(v, pf.Length > 0 && !legal ? pf : "1"));
        }
        string body = string.Join(".", parts.Where(p => p.Length > 0));
        if (body.Length == 0 && fmt.Length == 0) return prefix + suffix;
        return prefix + body + suffix;
    }

    internal static string FormatNum(int n, string fmt)
    {
        if (fmt.Length == 0) return "";
        if (fmt.StartsWith("1")) return fmt.StartsWith("01") && n < 10 ? "0" + n : n.ToString(CultureInfo.InvariantCulture);
        char f = fmt[0];
        return f switch
        {
            'a' => Letters(n, false), 'A' => Letters(n, true),
            'i' => Roman(n).ToLowerInvariant(), 'I' => Roman(n),
            'а' => RuLetters(n, false), 'А' => RuLetters(n, true),
            '①' => n is >= 1 and <= 20 ? ((char)(0x2460 + n - 1)).ToString() : n.ToString(),
            _ => n.ToString(CultureInfo.InvariantCulture),
        };
    }

    static string Letters(int n, bool upper)
    {
        if (n <= 0) return n.ToString();
        int rep = (n - 1) / 26 + 1;
        char c = (char)((upper ? 'A' : 'a') + (n - 1) % 26);
        return new string(c, rep);
    }

    static string RuLetters(int n, bool upper)
    {
        const string abc = "абвгдежзиклмнопрстуфхцчшщэюя";
        if (n <= 0) return n.ToString();
        int rep = (n - 1) / abc.Length + 1;
        char c = abc[(n - 1) % abc.Length];
        if (upper) c = char.ToUpperInvariant(c);
        return new string(c, rep);
    }

    static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString();
        var sb = new StringBuilder();
        (int V, string S)[] t = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        foreach (var (v, s) in t) while (n >= v) { sb.Append(s); n -= v; }
        return sb.ToString();
    }

    void ParseList(XElement el, Ctx ctx, List<Block> output, bool top, int depth, ListCtx? parent)
    {
        string? styleName = (string?)el.Attribute(Tx + "style-name");
        XElement? lstyle = FindListStyle(styleName, ctx) ?? parent?.Style;
        ListInstance inst;
        if (parent != null) inst = parent.Inst;
        else
        {
            string key = styleName ?? "";
            string? cont = (string?)el.Attribute(Tx + "continue-list");
            if (cont != null && _lists.LastOrDefault(l => l.Id == cont) is { } byId) inst = byId;
            else if ((string?)el.Attribute(Tx + "continue-numbering") == "true" && _lists.LastOrDefault(l => l.StyleName == key) is { } prev) inst = prev;
            else { inst = new ListInstance { StyleName = key }; _lists.Add(inst); }
            if ((string?)el.Attribute(Xml + "id") is { } id) inst.Id = id;
        }
        foreach (var item in el.Elements())
        {
            bool header = item.Name == Tx + "list-header";
            if (item.Name != Tx + "list-item" && !header) continue;
            int? start = int.TryParse((string?)item.Attribute(Tx + "start-value"), out var sv) ? sv : null;
            // Стиль списка может быть задан у стиля абзаца, если у списка его нет
            var effStyle = lstyle;
            if (effStyle == null)
            {
                var firstP = item.Elements().FirstOrDefault(e => e.Name == Tx + "p" || e.Name == Tx + "h");
                var ps = FindStyle((string?)firstP?.Attribute(Tx + "style-name"), "paragraph", ctx);
                effStyle = FindListStyle(ParaPropsOf(ps, ctx).ListStyle, ctx);
            }
            var lc = effStyle == null ? null : new ListCtx { Style = effStyle, Level = Math.Min(10, depth + 1), Header = header, StartValue = start, Inst = inst };
            foreach (var child in item.Elements())
            {
                if (child.Name == Tx + "p" || child.Name == Tx + "h")
                {
                    if (lc == null) { ParseParagraphEl(child, ctx, output, top, null); continue; }
                    ParseParagraphEl(child, ctx, output, top, lc);
                    lc.StartValue = null;
                }
                else if (child.Name == Tx + "list")
                {
                    if (lc != null && lc.FirstParaInItem && !header)
                    {
                        // Пункт без собственного абзаца: счётчик уровня всё равно увеличивается
                        var lvl = LevelStyle(lc.Style, lc.Level);
                        if (lvl?.Name == Tx + "list-level-style-number") Count(inst, lc.Level, lc.StartValue, lvl);
                        lc.FirstParaInItem = false;
                    }
                    ParseList(child, ctx, output, top, depth + 1, lc ?? (effStyle != null ? new ListCtx { Style = effStyle, Level = depth + 1, Inst = inst } : null));
                }
                else if (child.Name == Tx + "number") { }
                else ParseBlock(child, ctx, output, top);
            }
        }
    }

    // ================================================================
    //  Строчное содержимое
    // ================================================================

    internal sealed class InlineState
    {
        public required Ctx Ctx;
        public required Dictionary<string, string> ParaText;
        public required Paragraph Para;
        public List<string>? Bookmarks;
        public bool LastSpace = true;                 // для схлопывания пробелов (в начале абзаца пробелы опускаются)
    }

    void ParseInlines(XElement container, InlineState st, List<Inline> output, Dictionary<string, string> props, string? link)
    {
        foreach (var node in container.Nodes())
        {
            if (node is XText txt)
            {
                AddText(txt.Value, st, output, props, link);
                continue;
            }
            if (node is not XElement e) continue;
            var n = e.Name;
            if (n == Tx + "span")
            {
                var cs = FindStyle((string?)e.Attribute(Tx + "style-name"), "text", st.Ctx);
                var p2 = cs != null ? TextPropsOf(cs, st.Ctx, props) : props;
                if (p2.GetValueOrDefault("text:display") == "none") continue;
                ParseInlines(e, st, output, p2, link);
            }
            else if (n == Tx + "s")
            {
                int c = int.TryParse((string?)e.Attribute(Tx + "c"), out var cc) ? cc : 1;
                output.Add(new TextInline(new string(' ', Math.Max(1, c)), TS(props, link)));
                st.LastSpace = false;
            }
            else if (n == Tx + "tab") { output.Add(new TabInline(TS(props, link))); st.LastSpace = true; }
            else if (n == Tx + "line-break") { output.Add(new BreakInline(BreakKind.Line, TS(props, link))); st.LastSpace = true; }
            else if (n == Tx + "a")
            {
                string? href = (string?)e.Attribute(XLink + "href");
                var cs = FindStyle((string?)e.Attribute(Tx + "style-name"), "text", st.Ctx);
                var p2 = cs != null ? TextPropsOf(cs, st.Ctx, props) : props;
                string? lk = href;
                if (lk != null && lk.StartsWith('#'))
                {
                    lk = Uri.UnescapeDataString(lk);
                    int bar = lk.IndexOf('|');
                    if (bar > 0) lk = lk[..bar];   // «#Заголовок|outline» — ссылка на заголовок
                }
                ParseInlines(e, st, output, p2, lk ?? link);
            }
            else if (n == Tx + "note") AddNote(e, st, output, props);
            else if (n == Tx + "change-start" || n == Tx + "change-end") ChangeMark(e);
            else if (n == Tx + "change")
            {
                // Удалённый фрагмент: в режиме разметки — зачёркнутым цветом автора
                if (_markup && _changes.TryGetValue((string?)e.Attribute(Tx + "change-id") ?? "", out var del) && del.Kind == "deletion")
                {
                    var saved = _revision; var savedA = _revisionAuthor;
                    _revision = 2; _revisionAuthor = del.Author;
                    bool firstP = true;
                    foreach (var dp in del.El.Elements().Where(x => x.Name == Tx + "p" || x.Name == Tx + "h"))
                    {
                        if (!firstP) AddText(" ", st, output, props, link);
                        var ps = FindStyle((string?)dp.Attribute(Tx + "style-name"), "paragraph", st.Ctx);
                        ParseInlines(dp, st, output, ps != null ? TextPropsOf(ps, st.Ctx, props) : props, link);
                        firstP = false;
                    }
                    _revision = saved; _revisionAuthor = savedA;
                }
            }
            else if (n == Tx + "bookmark" || n == Tx + "bookmark-start")
            {
                if ((string?)e.Attribute(Tx + "name") is { } bn) (st.Bookmarks ??= []).Add(bn);
            }
            else if (n == Tx + "page-number")
            {
                string sel = (string?)e.Attribute(Tx + "select-page") ?? "current";
                if (sel == "current") output.Add(new FieldInline(FieldKind.Page, TS(props, link)));
                else AddText(e.Value, st, output, props, link);
                st.LastSpace = false;
            }
            else if (n == Tx + "page-count")
            {
                output.Add(new FieldInline(FieldKind.NumPages, TS(props, link)));
                st.LastSpace = false;
            }
            else if (n == Tx + "bookmark-ref" || n == Tx + "reference-ref" || n == Tx + "sequence-ref" || n == Tx + "note-ref")
            {
                string? refName = (string?)e.Attribute(Tx + "ref-name");
                string format = (string?)e.Attribute(Tx + "reference-format") ?? "text";
                if (format == "page" && refName != null)
                    output.Add(new FieldInline(FieldKind.PageRef, ToTextStyle(props, "#" + refName)) { Arg = refName, Cached = e.Value });
                else ParseInlines(e, st, output, props, refName != null ? "#" + refName : link);
            }
            else if (n == Office + "annotation" || n == Office + "annotation-end" || n == Tx + "bookmark-end" || n == Tx + "soft-page-break"
                     || n == Tx + "reference-mark-end"
                     || n == Tx + "toc-mark-end" || n == Tx + "alphabetical-index-mark-end" || n == Tx + "user-index-mark-end"
                     || n == Tx + "alphabetical-index-mark" || n == Tx + "toc-mark" || n == Tx + "user-index-mark" || n == Tx + "number")
            {
                if (n == Tx + "reference-mark-start" && (string?)e.Attribute(Tx + "name") is { } rm) (st.Bookmarks ??= []).Add(rm);
            }
            else if (n == Tx + "reference-mark-start" || n == Tx + "reference-mark")
            {
                if ((string?)e.Attribute(Tx + "name") is { } rm) (st.Bookmarks ??= []).Add(rm);
            }
            else if (n == Tx + "ruby")
            {
                if (e.Element(Tx + "ruby-base") is { } rb) ParseInlines(rb, st, output, props, link);
            }
            else if (n == Tx + "hidden-text" || n == Tx + "hidden-paragraph" || n == Tx + "conditional-text" && (string?)e.Attribute(Tx + "current-value") == "false") { }
            else if (n.Namespace == Dr)
            {
                ParseDrawing(e, st.Ctx, TS(props, link), output, blockLevel: false);
                st.LastSpace = false;
            }
            else if (n == Tx + "meta" || n == Tx + "meta-field" || n == Tx + "span" || n == Loext + "content-control")
                ParseInlines(e, st, output, props, link);
            else if (n.Namespace == Tx || n.Namespace == Office || n.Namespace == Loext)
            {
                // Прочие поля (дата, автор, заглавие, номер в последовательности, переменные): выводится сохранённое значение
                ParseInlines(e, st, output, props, link);
            }
        }
    }

    void AddText(string s, InlineState st, List<Inline> output, Dictionary<string, string> props, string? link)
    {
        if (s.Length == 0) return;
        // Пробельные символы схлопываются (ODF 1.2, 6.1.2); в начале абзаца опускаются
        var sb = new StringBuilder(s.Length);
        foreach (char ch in s)
        {
            if (ch is ' ' or '\t' or '\n' or '\r')
            {
                if (st.LastSpace) continue;
                sb.Append(' ');
                st.LastSpace = true;
            }
            else { sb.Append(ch); st.LastSpace = false; }
        }
        if (sb.Length == 0) return;
        var ts = TS(props, link);
        string text = sb.ToString();
        if (props.GetValueOrDefault("fo:text-transform") == "lowercase") text = text.ToLowerInvariant();
        else if (props.GetValueOrDefault("fo:text-transform") == "capitalize") text = Capitalize(text, output);
        if (output.Count > 0 && output[^1] is TextInline prev && prev.Style == ts) prev.Text += text;
        else output.Add(new TextInline(text, ts));
    }

    static string Capitalize(string text, List<Inline> before)
    {
        var sb = new StringBuilder(text.Length);
        bool start = before.Count == 0 || before[^1] is TextInline { Text: var t } && (t.Length == 0 || char.IsWhiteSpace(t[^1]));
        foreach (char c in text)
        {
            sb.Append(start && char.IsLetter(c) ? char.ToUpper(c) : c);
            start = char.IsWhiteSpace(c);
        }
        return sb.ToString();
    }

    // ---------------- Исправления ----------------

    readonly Dictionary<string, (string Kind, string Author, XElement El)> _changes = [];
    bool _markup;
    int _revision;                                   // 0 — нет, 1 — вставка, 2 — удаление
    string _revisionAuthor = "";
    readonly Dictionary<string, int> _authors = [];
    static readonly Rgb[] AuthorColors =
    [
        new(0.71, 0.05, 0.06), new(0.12, 0.3, 0.75), new(0.0, 0.5, 0.2), new(0.5, 0.15, 0.6),
        new(0.75, 0.4, 0.0), new(0.0, 0.5, 0.55), new(0.55, 0.35, 0.15), new(0.4, 0.4, 0.4),
    ];

    void ChangeMark(XElement e)
    {
        if (!_markup || !_changes.TryGetValue((string?)e.Attribute(Tx + "change-id") ?? "", out var ch) || ch.Kind != "insertion") return;
        if (e.Name == Tx + "change-start") { _revision = 1; _revisionAuthor = ch.Author; }
        else if (e.Name == Tx + "change-end" && _revision == 1) _revision = 0;
    }

    TextStyle TS(Dictionary<string, string> props, string? link)
    {
        var s = ToTextStyle(props, link);
        if (_revision == 0) return s;
        if (!_authors.TryGetValue(_revisionAuthor, out int i)) _authors[_revisionAuthor] = i = _authors.Count;
        var c = AuthorColors[i % AuthorColors.Length];
        return _revision == 1
            ? s with { Color = c, Underline = true, UnderlineStyle = "single", UnderlineColor = c, RevisionMark = true }
            : s with { Color = c, Strike = true, RevisionMark = true };
    }

    // ---------------- Сноски ----------------

    int _footCount, _endCount;
    readonly List<List<Block>> _endnotes = [];

    void AddNote(XElement e, InlineState st, List<Inline> output, Dictionary<string, string> props)
    {
        bool end = (string?)e.Attribute(Tx + "note-class") == "endnote";
        var cfg = end ? _endConfig : _footConfig;
        var citEl = e.Element(Tx + "note-citation");
        string fmt = (string?)cfg?.Attribute(St + "num-format") ?? (end ? "i" : "1");
        int startAt = int.TryParse((string?)cfg?.Attribute(Tx + "start-value"), out var sv) ? sv : 0;
        int num = (end ? ++_endCount : ++_footCount) + startAt;
        string mark = (string?)citEl?.Attribute(Tx + "label") ?? (citEl?.Value is { Length: > 0 } cv ? cv : FormatNum(num, fmt));
        // Знак сноски в тексте: стиль citation-body-style-name (по умолчанию — верхний индекс)
        var anchorStyle = FindStyle((string?)cfg?.Attribute(Tx + "citation-body-style-name"), "text", st.Ctx);
        var ap = anchorStyle != null ? TextPropsOf(anchorStyle, st.Ctx, props) : new Dictionary<string, string>(props) { ["style:text-position"] = "super 58%" };
        var ats = ToTextStyle(ap);
        // Текст сноски: знак в начале первого абзаца (стиль citation-style-name)
        var noteCtx = new Ctx { HeaderFooter = st.Ctx.HeaderFooter, InNote = true };
        var blocks = new List<Block>();
        var savedPend = _pendingAnchors;
        var savedBreak = (_pendingBreak, _breakAfter);
        if (e.Element(Tx + "note-body") is { } body) ParseBlocks(body, noteCtx, blocks, top: false);
        _pendingAnchors = savedPend;
        (_pendingBreak, _breakAfter) = savedBreak;
        var firstP = blocks.OfType<Paragraph>().FirstOrDefault();
        if (firstP != null)
        {
            var bodyStyle = FindStyle((string?)cfg?.Attribute(Tx + "citation-style-name"), "text", st.Ctx);
            var bp = bodyStyle != null ? TextPropsOf(bodyStyle, st.Ctx, TextPropsOfStyleOf(firstP, noteCtx)) : TextPropsOfStyleOf(firstP, noteCtx);
            var bts = ToTextStyle(bp);
            var ins = new List<Inline> { new TextInline(mark, bts) };
            // Висячая строка (стиль «Сноска»): знак на выступе, текст — после табуляции
            if (firstP.IndFirst < 0) ins.Add(new TabInline(bts)); else ins.Add(new TextInline(" ", bts));
            firstP.Inlines.InsertRange(0, ins);
        }
        if (end)
        {
            _endnotes.Add(blocks);
            output.Add(new TextInline(mark, ats));
        }
        else output.Add(new FootnoteInline(mark, ats, blocks));
        st.LastSpace = false;
    }

    Dictionary<string, string> TextPropsOfStyleOf(Paragraph p, Ctx ctx) => new()
    {
        ["fo:font-size"] = p.MarkStyle.Size.ToString(CultureInfo.InvariantCulture) + "pt",
        ["style:font-name"] = p.MarkStyle.FontAscii,
    };

    void AppendEndnotes()
    {
        if (_endnotes.Count == 0) return;
        // LibreOffice выводит концевые сноски на отдельной странице в конце документа
        bool first = true;
        foreach (var b in _endnotes)
        {
            if (first && b.FirstOrDefault() is Paragraph p) p.PageBreakBefore = true;
            else if (first) _blocks.Add(new Paragraph { PageBreakBefore = true });
            first = false;
            _blocks.AddRange(b);
        }
    }

    // ================================================================
    //  Стиль текста
    // ================================================================

    string FontName(Dictionary<string, string> p, string nameKey, string famKey)
    {
        if (p.TryGetValue(nameKey, out var n)) return _fontFaces.TryGetValue(n, out var f) ? f : Unquote(n);
        if (p.TryGetValue(famKey, out var fam)) return Unquote(fam);
        return "";
    }

    internal TextStyle ToTextStyle(Dictionary<string, string> p, string? link = null)
    {
        string font = FontName(p, "style:font-name", "fo:font-family");
        if (font.Length == 0) font = "Liberation Serif";
        string asian = FontName(p, "style:font-name-asian", "style:font-family-asian");
        if (asian.Length == 0) asian = font;
        string weight = p.GetValueOrDefault("fo:font-weight") ?? "normal";
        int numWeight = weight switch { "bold" => 700, "normal" => 400, _ => int.TryParse(weight, out var w) ? w : 400 };
        string ul = p.GetValueOrDefault("style:text-underline-style") ?? "none";
        string ulType = p.GetValueOrDefault("style:text-underline-type") ?? "single";
        string ulWidth = p.GetValueOrDefault("style:text-underline-width") ?? "auto";
        string? underline = ul == "none" || ulType == "none" ? null : ul switch
        {
            "dotted" => ulWidth == "bold" ? "dottedHeavy" : "dotted",
            "dash" => ulWidth == "bold" ? "dashedHeavy" : "dash",
            "long-dash" => "dashLong",
            "dot-dash" => "dotDash",
            "dot-dot-dash" => "dotDotDash",
            "wave" => ulType == "double" ? "wavyDouble" : "wave",
            _ => ulType == "double" ? "double" : ulWidth is "bold" or "thick" ? "thick" : "single",
        };
        string? ulColor = p.GetValueOrDefault("style:text-underline-color");
        string strike = p.GetValueOrDefault("style:text-line-through-style") ?? "none";
        bool doubleStrike = p.GetValueOrDefault("style:text-line-through-type") == "double";
        double size = Len(p, "fo:font-size", 12);
        int vert = 0;
        double position = 0;
        if (p.GetValueOrDefault("style:text-position") is { } tpos)
        {
            var toks = tpos.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            double? relSize = toks.Length > 1 && toks[1].EndsWith('%') && double.TryParse(toks[1][..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var rs) ? rs / 100 : null;
            if (toks.Length > 0 && toks[0] == "super") vert = 1;
            else if (toks.Length > 0 && toks[0] == "sub") vert = -1;
            else if (toks.Length > 0 && toks[0].EndsWith('%') && double.TryParse(toks[0][..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct) && pct != 0)
            {
                if (relSize is double r0 && r0 < 0.99) { vert = pct > 0 ? 1 : -1; }
                else position = size * pct / 100;
            }
        }
        string relief = p.GetValueOrDefault("style:font-relief") ?? "none";
        bool shadow = p.TryGetValue("fo:text-shadow", out var sh) && sh != "none";
        var color = Color(p.GetValueOrDefault("fo:color"));
        if (p.GetValueOrDefault("style:use-window-font-color") == "true" && color == null) color = Rgb.Black;
        double scale = p.TryGetValue("style:text-scale", out var sc) && sc.EndsWith('%') && double.TryParse(sc[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var scv) ? scv / 100 : 1;
        string? emph = p.GetValueOrDefault("style:text-emphasize") is { } em && !em.StartsWith("none") ? "dot" : null;
        BorderDef? runBorder = null;
        if (p.ContainsKey("fo:border") || p.ContainsKey("fo:border-top")) runBorder = SideBorder(p, "top") ?? SideBorder(p, "left");
        return new TextStyle
        {
            FontAscii = font, FontHAnsi = font, FontEastAsia = asian,
            Bold = numWeight >= 600,
            Weight = numWeight is 400 or 700 ? 0 : numWeight,
            Italic = p.GetValueOrDefault("fo:font-style") is "italic" or "oblique",
            Underline = underline != null,
            DoubleUnderline = underline is "double" or "wavyDouble",
            UnderlineStyle = underline,
            UnderlineColor = ulColor is null or "font-color" ? null : Color(ulColor),
            Strike = strike != "none" && (p.GetValueOrDefault("style:text-line-through-type") ?? "single") != "none",
            DoubleStrike = doubleStrike && strike != "none",
            Caps = p.GetValueOrDefault("fo:text-transform") == "uppercase",
            SmallCaps = p.GetValueOrDefault("fo:font-variant") == "small-caps",
            LetterSpacing = p.TryGetValue("fo:letter-spacing", out var lsp) && lsp != "normal" ? Length(lsp) ?? 0 : 0,
            HScale = scale,
            Outline = p.GetValueOrDefault("style:text-outline") == "true",
            ShadowColor = shadow || relief != "none" ? new Rgb(0.55, 0.55, 0.55) : null,
            ShadowOffset = relief == "engraved" ? -0.04 : relief == "embossed" ? 0.04 : shadow ? 0.06 : 0,
            Size = size,
            Color = relief != "none" ? new Rgb(0.93, 0.93, 0.93) : color ?? Rgb.Black,
            Background = Color(p.GetValueOrDefault("fo:background-color")) ?? Color(p.GetValueOrDefault("loext:char-shading-value") is null ? null : null),
            VertAlign = vert,
            Position = position,
            Link = link,
            Emphasis = emph,
            RunBorder = runBorder,
            Kerning = p.GetValueOrDefault("style:letter-kerning") == "true",
            Hyphens = p.GetValueOrDefault("fo:hyphenate") == "true" ? 2 : 1,
            Lang = p.TryGetValue("fo:language", out var lng) && lng is not ("zxx" or "none")
                ? lng + (p.TryGetValue("fo:country", out var cty) && cty != "none" ? "-" + cty : "") : null,
        };
    }
}

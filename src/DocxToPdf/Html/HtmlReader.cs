using System.Globalization;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Images;

namespace DocxToPdf.Html;

/// <summary>Настройки конвертации HTML.</summary>
public sealed class HtmlOptions
{
    public double PageWidth = 595.3, PageHeight = 841.9;   // A4
    public double MarginTop = 56.7, MarginBottom = 56.7, MarginLeft = 56.7, MarginRight = 56.7;
    public bool Landscape;
    public bool PageNumbers;
    public bool AllowRemote = true;
}

/// <summary>
/// Читает HTML (+ CSS) и строит ту же модель документа, что и DocxReader,
/// после чего используется общий движок вёрстки и PDF.
/// </summary>
public sealed partial class HtmlReader
{
    const string UserAgentCss = """
        p { margin: 1em 0 }
        h1 { font-size: 2em; margin: .67em 0; font-weight: bold }
        h2 { font-size: 1.5em; margin: .83em 0; font-weight: bold }
        h3 { font-size: 1.17em; margin: 1em 0; font-weight: bold }
        h4 { margin: 1.33em 0; font-weight: bold }
        h5 { font-size: .83em; margin: 1.67em 0; font-weight: bold }
        h6 { font-size: .67em; margin: 2.33em 0; font-weight: bold }
        b, strong, th { font-weight: bold }
        i, em, cite, var, dfn, address { font-style: italic }
        u, ins { text-decoration: underline }
        s, strike, del { text-decoration: line-through }
        a[href] { color: #0645ad; text-decoration: underline }
        pre, code, kbd, samp, tt, listing, xmp { font-family: monospace }
        pre, listing, xmp { white-space: pre; margin: 1em 0 }
        blockquote, figure { margin: 1em 40px }
        ul, ol, menu, dir { margin: 1em 0; padding-left: 40px }
        ul ul, ul ol, ol ul, ol ol, li ul, li ol { margin-top: 0; margin-bottom: 0 }
        ol { list-style-type: decimal }
        ul ul, ol ul { list-style-type: circle }
        ul ul ul, ol ul ul, ul ol ul, ol ol ul { list-style-type: square }
        dl { margin: 1em 0 }
        dd { margin-left: 40px }
        th { text-align: center }
        td, th { padding: 1px 2px; vertical-align: middle }
        caption { text-align: center }
        sup { vertical-align: super }
        sub { vertical-align: sub }
        small { font-size: smaller }
        big { font-size: larger }
        mark { background-color: yellow }
        q::before { content: open-quote }
        q::after { content: close-quote }
        center { text-align: center }
        hr { margin: .5em 0 }
        fieldset { margin: 0 2px; padding: .35em .75em .625em }
        """;

    static readonly HashSet<string> BlockTags =
    [
        "html", "body", "div", "p", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "menu", "dir", "dl", "dt", "dd", "blockquote",
        "pre", "address", "article", "aside", "footer", "header", "nav", "section", "main", "figure", "figcaption", "form",
        "fieldset", "legend", "center", "details", "summary", "hgroup", "listing", "xmp", "hr", "caption", "dialog", "search"
    ];
    static readonly HashSet<string> HiddenTags =
        ["head", "script", "style", "title", "meta", "link", "template", "noscript", "base", "datalist", "param", "source", "track", "area", "map", "audio", "video", "iframe", "object", "embed", "svg", "math", "canvas", "button", "input", "select", "textarea", "option"];

    static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(10) };

    static HtmlReader()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    readonly FontManager _fonts;
    readonly HtmlOptions _opt;
    readonly string _baseDir;
    MhtmlArchive? _mht;                       // веб-архив MHTML: ресурсы берутся из него
    string? _baseUrl;                         // адрес страницы в архиве (Content-Location или <base href>)
    readonly DocxDocument _doc = new() { DefaultTabStop = 36 };
    readonly StyleSheet _ua = new();
    readonly StyleSheet _author = new();
    readonly Dictionary<string, string?> _fontPick = new(StringComparer.OrdinalIgnoreCase);
    double _contentWidth;
    int _imageCounter;

    HtmlReader(FontManager fonts, HtmlOptions opt, string baseDir)
    {
        _fonts = fonts;
        _opt = opt;
        _baseDir = baseDir;
        _ua.Parse(UserAgentCss);
    }

    public static DocxDocument Read(string path, FontManager fonts, HtmlOptions? options = null)
    {
        var bytes = File.ReadAllBytes(path);
        var reader = new HtmlReader(fonts, options ?? new HtmlOptions(), Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        if (MhtmlArchive.Looks(bytes))
        {
            var mht = MhtmlArchive.Parse(bytes);
            if (mht.Main != null)
            {
                reader._mht = mht;
                reader._baseUrl = mht.Main.Location;
                return reader.Convert(DecodePart(mht.Main));
            }
        }
        return reader.Convert(DecodeBytes(bytes));
    }

    /// <summary>Текст части MHTML: кодировка из Content-Type, иначе — как у обычного файла.</summary>
    static string DecodePart(MhtmlArchive.Part part)
    {
        var b = part.Data;
        bool bom = b.Length >= 2 && (b[0] == 0xEF && b[1] == 0xBB || b[0] == 0xFF && b[1] == 0xFE || b[0] == 0xFE && b[1] == 0xFF);
        if (part.Charset is { Length: > 0 } cs && !bom)
        {
            try { return Encoding.GetEncoding(cs.Trim()).GetString(b); } catch { /* неизвестная кодировка */ }
        }
        return DecodeBytes(b);
    }

    public static DocxDocument ReadString(string html, FontManager fonts, HtmlOptions? options = null, string? baseDir = null) =>
        new HtmlReader(fonts, options ?? new HtmlOptions(), baseDir ?? Directory.GetCurrentDirectory()).Convert(html);

    /// <summary>Определение кодировки: BOM → meta charset → UTF-8.</summary>
    static string DecodeBytes(byte[] b)
    {
        if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
        if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
        if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
        string head = Encoding.Latin1.GetString(b, 0, Math.Min(b.Length, 4096)).ToLowerInvariant();
        int ci = head.IndexOf("charset=", StringComparison.Ordinal);
        if (ci >= 0)
        {
            int s = ci + 8;
            while (s < head.Length && head[s] is '"' or '\'' or ' ') s++;
            int e = s;
            while (e < head.Length && (char.IsLetterOrDigit(head[e]) || head[e] is '-' or '_')) e++;
            string cs = head[s..e];
            if (cs.Length > 0 && cs != "utf-8" && cs != "utf8")
            {
                try { return Encoding.GetEncoding(cs).GetString(b); } catch { /* неизвестная кодировка */ }
            }
        }
        // Если UTF-8 некорректен — пробуем windows-1251
        try { return new UTF8Encoding(false, true).GetString(b); }
        catch { return Encoding.GetEncoding(1251).GetString(b); }
    }

    // =================================================================
    //  Общий проход
    // =================================================================

    DocxDocument Convert(string html)
    {
        var root = HtmlParser.Parse(html);
        if (_mht != null && root.Descendants().FirstOrDefault(d => d.Tag == "base" && d.Attr("href") != null) is { } baseEl)
        {
            string bh = baseEl.Attr("href")!.Trim();
            if (Uri.TryCreate(bh, UriKind.Absolute, out var bu)) _baseUrl = bu.AbsoluteUri;
            else if (_baseUrl != null && Uri.TryCreate(new Uri(_baseUrl), bh, out var rel)) _baseUrl = rel.AbsoluteUri;
        }

        // CSS документа действует и внутри встроенных SVG (в т.ч. стили из других SVG страницы, как в браузерах)
        var svgCss = new StringBuilder();
        foreach (var d in root.Descendants())
        {
            if (d.Tag == "style") svgCss.Append(string.Concat(d.Children.Where(c => c.IsText).Select(c => c.Text))).Append('\n');
            else if (d.Tag == "img" && d.Attr("data-tag") == "svg" && d.Attr("src") is { } isrc && isrc.IndexOf(',') is int ic and > 0)
            {
                try
                {
                    string raw = Encoding.UTF8.GetString(System.Convert.FromBase64String(isrc[(ic + 1)..]));
                    foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(raw, @"<style[^>]*>(.*?)</style>", System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        svgCss.Append(m.Groups[1].Value).Append('\n');
                }
                catch { }
            }
        }
        foreach (var st in root.Descendants().Where(d => d.Tag == "style"))
        {
            string media = (st.Attr("media") ?? "all").ToLowerInvariant();
            if (media.Contains("all") || media.Contains("print") || media.Length == 0)
                _author.Parse(string.Concat(st.Children.Where(c => c.IsText).Select(c => c.Text)));
        }
        foreach (var link in root.Descendants().Where(d => d.Tag == "link"))
        {
            if (!(link.Attr("rel") ?? "").Contains("stylesheet", StringComparison.OrdinalIgnoreCase)) continue;
            string media = (link.Attr("media") ?? "all").ToLowerInvariant();
            if (!(media.Contains("all") || media.Contains("print"))) continue;
            var css = LoadResource(link.Attr("href"));
            if (css != null) { _author.Parse(DecodeBytes(css)); svgCss.Insert(0, DecodeBytes(css) + "\n"); }
        }
        _svgCss = svgCss.ToString();
        _doc.Title = root.Descendants().FirstOrDefault(d => d.Tag == "title")?.Children.FirstOrDefault()?.Text.Trim();

        var section = new Section { Props = PageSetup() };
        _contentWidth = section.Props.PageWidth - section.Props.MarginLeft - section.Props.MarginRight;

        string lang = (root.Descendants().FirstOrDefault(d => d.Tag == "html")?.Attr("lang") ?? "").ToLowerInvariant();
        if (lang.StartsWith("ru") || lang.StartsWith("uk") || lang.StartsWith("be")) _defaultQuotes = ["\u00AB", "\u00BB", "\u201E", "\u201C"];
        else if (lang.StartsWith("de")) _defaultQuotes = ["\u201E", "\u201C", "\u201A", "\u2018"];
        else if (lang.StartsWith("fr")) _defaultQuotes = ["\u00AB\u00A0", "\u00A0\u00BB", "\u2039", "\u203A"];
        ComputeStyles(root, new CssStyle { Display = "block" });

        var flow = new Flow(section.Blocks);
        var ctx = new BlockCtx { IndL = 0, IndR = 0, Owner = root.Style!, Width = _contentWidth };
        WalkChildren(root, flow, ctx);
        Flush(flow);
        if (section.Blocks.Count == 0) section.Blocks.Add(new Paragraph());

        if (_opt.PageNumbers)
        {
            var st = new TextStyle { FontAscii = "Arial", FontHAnsi = "Arial", FontEastAsia = "Arial", Size = 9, Color = new Rgb(0.3, 0.3, 0.3) };
            var p = new Paragraph { Align = "center", MarkStyle = st };
            p.Inlines.Add(new FieldInline(FieldKind.Page, st));
            p.Inlines.Add(new TextInline(" / ", st));
            p.Inlines.Add(new FieldInline(FieldKind.NumPages, st));
            section.Props.FooterDefault = new HeaderFooter { Blocks = [p] };
        }
        section.Props.FixedBlocks.AddRange(_fixedBlocks);
        _doc.Sections.Add(section);
        return _doc;
    }

    SectionProps PageSetup()
    {
        var sp = new SectionProps
        {
            PageWidth = _opt.PageWidth, PageHeight = _opt.PageHeight,
            MarginTop = _opt.MarginTop, MarginBottom = _opt.MarginBottom, MarginLeft = _opt.MarginLeft, MarginRight = _opt.MarginRight,
            HeaderDistance = 28, FooterDistance = 28
        };
        bool landscape = _opt.Landscape;
        foreach (var d in _author.PageDecls)
        {
            if (d.Property == "size")
            {
                foreach (var tok in CssValues.Tokens(d.Value.ToLowerInvariant()))
                {
                    switch (tok)
                    {
                        case "a3": sp.PageWidth = 841.9; sp.PageHeight = 1190.6; break;
                        case "a4": sp.PageWidth = 595.3; sp.PageHeight = 841.9; break;
                        case "a5": sp.PageWidth = 419.5; sp.PageHeight = 595.3; break;
                        case "b5": sp.PageWidth = 498.9; sp.PageHeight = 708.7; break;
                        case "letter": sp.PageWidth = 612; sp.PageHeight = 792; break;
                        case "legal": sp.PageWidth = 612; sp.PageHeight = 1008; break;
                        case "landscape": landscape = true; break;
                        case "portrait": landscape = false; break;
                    }
                }
                var lens = CssValues.Tokens(d.Value).Select(t => CssValues.Length(t, 12)).Where(l => l is > 0).ToList();
                if (lens.Count >= 1) { sp.PageWidth = lens[0]!.Value; sp.PageHeight = lens.Count > 1 ? lens[1]!.Value : lens[0]!.Value; }
            }
            else if (d.Property == "margin")
            {
                var (t, r, b, l) = CssValues.Box(d.Value);
                sp.MarginTop = CssValues.Length(t, 12) ?? sp.MarginTop;
                sp.MarginRight = CssValues.Length(r, 12) ?? sp.MarginRight;
                sp.MarginBottom = CssValues.Length(b, 12) ?? sp.MarginBottom;
                sp.MarginLeft = CssValues.Length(l, 12) ?? sp.MarginLeft;
            }
            else if (d.Property.StartsWith("margin-"))
            {
                double? v = CssValues.Length(d.Value, 12);
                if (v == null) continue;
                switch (d.Property)
                {
                    case "margin-top": sp.MarginTop = v.Value; break;
                    case "margin-bottom": sp.MarginBottom = v.Value; break;
                    case "margin-left": sp.MarginLeft = v.Value; break;
                    case "margin-right": sp.MarginRight = v.Value; break;
                }
            }
        }
        if (landscape && sp.PageWidth < sp.PageHeight) (sp.PageWidth, sp.PageHeight) = (sp.PageHeight, sp.PageWidth);
        return sp;
    }

    // =================================================================
    //  Каскад стилей
    // =================================================================

    static readonly HashSet<string> NoPseudo = ["img", "br", "hr", "input", "textarea", "select", "iframe", "video", "audio", "canvas", "wbr", "meta", "link", "col", "area", "base", "embed", "param", "source", "track"];

    void ComputeStyles(Node n, CssStyle parent)
    {
        if (n.Tag == "#root")
        {
            n.Style = parent;
            foreach (var c in n.Children.ToList()) ComputeStyles(c, parent);
            return;
        }
        if (!n.IsElement || n.IsPseudo) return;

        var s = parent.InheritFrom();
        s.Display = HiddenTags.Contains(n.Tag) ? "none"
            : n.Tag == "li" ? "list-item"
            : n.Tag switch
            {
                "table" => "table", "tr" => "table-row", "td" or "th" => "table-cell",
                "thead" => "table-header-group", "tbody" => "table-row-group", "tfoot" => "table-footer-group",
                "col" or "colgroup" => "table-column", "caption" => "table-caption",
                _ => BlockTags.Contains(n.Tag) ? "block" : "inline"
            };
        if (n.Tag is "body" or "html") s.Display = "block";

        ApplyAll(s, parent, CollectDecls(n, null));

        if (n.Tag == "a" && n.Attr("href") is { } href && IsExternalLink(href)) s.Link = href;
        if ((n.Attr("lang") ?? n.Attr("xml:lang")) is { Length: > 0 } lang) s.Lang = lang;
        Blockify(s, parent);
        FixTextBackground(s);
        SetupInlineBox(s);
        s.ParentPerspective = parent.Perspective;
        s.ParentPerspectiveOrigin = parent.PerspectiveOrigin;
        s.ParentPreserve3D = parent.Preserve3D;
        n.Style = s;
        if (s.Display == "none") return;

        // Счётчики: сначала элемента, затем ::before, потомков и ::after (порядок документа)
        ApplyCounters(s, n.Parent ?? n);
        bool pseudo = !NoPseudo.Contains(n.Tag) && !HiddenTags.Contains(n.Tag);
        if (pseudo && MakePseudo(n, s, "before") is { } before) n.Children.Insert(0, before);
        foreach (var c in n.Children.ToList())
            if (!c.IsPseudo) ComputeStyles(c, s);
        if (pseudo && MakePseudo(n, s, "after") is { } after) n.Children.Add(after);
        PopCounters(n);
        if (s.Display is not ("inline" or "none")) ApplyFirstLetter(n, s);
    }

    // =================================================================
    //  ::first-letter / ::first-line
    // =================================================================

    static bool IsPunct(char c) => char.GetUnicodeCategory(c) is UnicodeCategory.OpenPunctuation or UnicodeCategory.ClosePunctuation
        or UnicodeCategory.InitialQuotePunctuation or UnicodeCategory.FinalQuotePunctuation or UnicodeCategory.OtherPunctuation
        or UnicodeCategory.DashPunctuation;

    /// <summary>Первый текстовый узел первой строки блока (через строчные элементы и первый вложенный блок).</summary>
    static Node? FirstTextNode(Node n)
    {
        foreach (var c in n.Children)
        {
            if (c.IsText)
            {
                if (c.Text.Trim().Length > 0) return c;
                continue;
            }
            if (!c.IsElement || c.Style == null || c.Style.Display == "none") continue;
            if (c.Style.Float != null || c.Style.Position is "absolute" or "fixed") continue;
            if (c.Tag is "img" or "br" or "hr" or "table" or "input" or "textarea" or "select" || c.Style.Display is "table" or "flex" or "grid") return null;
            var r = FirstTextNode(c);
            if (r != null) return r;
        }
        return null;
    }

    /// <summary>Выделяет первую букву (с окружающей пунктуацией) в узел ::first-letter.</summary>
    void ApplyFirstLetter(Node n, CssStyle s)
    {
        var decls = CollectDecls(n, "first-letter");
        if (decls.Count == 0) return;
        // ::first-letter наследует от ::first-line
        decls = CollectDecls(n, "first-line").Concat(decls).ToList();
        var t = FirstTextNode(n);
        if (t?.Parent?.Style == null || t.Parent.Tag == "::first-letter") return;
        string text = t.Text;
        int i = 0;
        while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
        int start = i;
        while (i < text.Length && IsPunct(text[i])) i++;
        if (i < text.Length && !char.IsWhiteSpace(text[i])) i += char.IsSurrogatePair(text, i) ? 2 : 1;
        while (i < text.Length && IsPunct(text[i])) i++;
        if (i == start) return;

        var owner = t.Parent;
        var ps = owner.Style.InheritFrom();
        ps.Display = "inline";
        ApplyAll(ps, owner.Style, decls);
        if (ps.InitialLetter != null) SetupInitialLetter(ps, owner.Style, n.Style ?? owner.Style);
        Blockify(ps, owner.Style);
        FixTextBackground(ps);
        SetupInlineBox(ps);
        var pseudo = new Node { Tag = "::first-letter", Parent = owner, Style = ps };
        pseudo.Children.Add(new Node { Tag = "#text", Text = text[start..i], Parent = pseudo });
        t.Text = text[i..];
        owner.Children.Insert(owner.Children.IndexOf(t), pseudo);
    }

    /// <summary>
    /// initial-letter: N [погружение] — буквица высотой ровно в N строк. Верх заглавной буквы
    /// выравнивается по высоте прописных первой строки, базовая линия — по базовой линии строки «погружения».
    /// Реализуется как плавающий блок с точно рассчитанными размером шрифта и высотой.
    /// </summary>
    void SetupInitialLetter(CssStyle ps, CssStyle owner, CssStyle block)
    {
        var tok = CssValues.Tokens(ps.InitialLetter!);
        if (tok.Count == 0 || !double.TryParse(tok[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double size) || size < 1)
            return;
        int sink = tok.Count > 1 && int.TryParse(tok[1], out var sk) ? Math.Max(1, sk) : (int)Math.Round(size);
        if (tok.Count > 1 && tok[1] == "drop") sink = (int)Math.Round(size);
        if (tok.Count > 1 && tok[1] == "raise") sink = 1;

        // Метрики основного текста (как их считает вёрстка)
        var bodyFace = _fonts.Get(block.FontFamily, block.Bold, block.Italic, block.Weight).Font;
        double em = bodyFace.UnitsPerEm, bs = block.FontSize;
        double asc = bodyFace.Ascent / em * bs, desc = bodyFace.Descent / em * bs, natural = asc + desc;
        double factor = block.LineHeightFactor is double lf ? lf / 1.15 : 1;
        double lh = block.LineHeightPt is double lp ? lp : natural * factor;
        double baseline1 = block.LineHeightPt != null ? asc + (lh - natural) / 2 : factor >= 1 ? asc : asc * factor;
        double capBody = bodyFace.CapHeight / em * bs;

        // Размер буквицы: от высоты прописных первой строки до базовой линии строки N
        var letterFace = _fonts.Get(ps.FontFamily, ps.Bold, ps.Italic, ps.Weight).Font;
        double capRatio = Math.Max(0.3, letterFace.CapHeight / (double)letterFace.UnitsPerEm);
        double capLetter = (size - 1) * lh + capBody;
        double fsLetter = capLetter / capRatio;
        double descLetter = letterFace.Descent / (double)letterFace.UnitsPerEm * fsLetter;

        double baselineSink = (sink - 1) * lh + baseline1;
        // Приподнятая буквица (погружение меньше размера) выступает над первой строкой —
        // абзац сдвигается вниз, чтобы буква не наезжала на предыдущий текст
        double raise = Math.Max(0, capLetter - baselineSink + capBody * 0.1);
        ps.FontSize = fsLetter;
        ps.Float = "left";
        ps.ExactLineHeight = raise + baselineSink + descLetter;   // базовая линия буквы = базовая линия строки sink
        ps.LineHeightFactor = null; ps.LineHeightPt = null;
        ps.PaddingTop = ps.PaddingBottom = 0;
        ps.MarginTop = -raise;
        // Обтекание — ровно sink строк
        ps.MarginBottom = sink * lh - ps.ExactLineHeight.Value + raise;
        if (raise > 0) block.PaddingTop += raise;
        ps.MarginRight = Math.Max(ps.MarginRight, fsLetter * 0.1);
    }

    /// <summary>Радиусы скругления по исходным значениям CSS; проценты вычисляются при отрисовке.</summary>
    static RadiusSpec? RadiusOf(CssStyle s)
    {
        if (!s.HasRadius) return null;
        var r = new RadiusSpec();
        for (int i = 0; i < 8; i++)
        {
            string? v = s.Radius[i];
            if (v == null) continue;
            if (v.EndsWith('%') && double.TryParse(v[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct))
            {
                r.V[i] = Math.Max(0, pct / 100);
                r.Pct[i] = true;
            }
            else r.V[i] = Math.Max(0, CssValues.Length(v, s.FontSize, 0) ?? 0);
        }
        return r.Any ? r : null;
    }

    /// <summary>Назначает стиль ::first-line первому абзацу, созданному внутри элемента.</summary>
    void ApplyFirstLine(Node n, CssStyle s, IEnumerable<Block> created)
    {
        if (!n.IsElement) return;
        var decls = CollectDecls(n, "first-line");
        if (decls.Count == 0) return;
        var para = created.OfType<Paragraph>().FirstOrDefault();
        if (para == null || para.FirstLine != null) return;

        var fl = s.InheritFrom();
        ApplyAll(fl, s, decls);
        var props = decls.Select(d => d.Property).ToHashSet();
        bool font = props.Contains("font");
        bool weight = font || props.Contains("font-weight"), style = font || props.Contains("font-style");
        bool size = font || props.Contains("font-size") || props.Contains("font-variant") || props.Contains("font-variant-caps");
        bool family = font || props.Contains("font-family");
        bool color = props.Contains("color");
        bool caps = props.Contains("text-transform") || props.Contains("font-variant") || props.Contains("font-variant-caps");
        bool deco = props.Contains("text-decoration") || props.Contains("text-decoration-line");
        bool bg = props.Contains("background") || props.Contains("background-color");
        para.FirstLine = ts => ts with
        {
            Bold = weight ? fl.Bold : ts.Bold,
            Weight = weight ? fl.Weight : ts.Weight,
            Italic = style ? fl.Italic : ts.Italic,
            Size = size ? (ts.Size == s.FontSize ? fl.FontSize : ts.Size * fl.FontSize / Math.Max(1, s.FontSize)) : ts.Size,
            FontAscii = family ? fl.FontFamily : ts.FontAscii,
            FontHAnsi = family ? fl.FontFamily : ts.FontHAnsi,
            FontEastAsia = family ? fl.FontFamily : ts.FontEastAsia,
            Color = color ? fl.Color : ts.Color,
            Caps = caps ? fl.Upper : ts.Caps,
            SmallCaps = caps ? fl.SmallCaps : ts.SmallCaps,
            LetterSpacing = props.Contains("letter-spacing") ? fl.LetterSpacing : ts.LetterSpacing,
            Underline = deco ? fl.Underline : ts.Underline,
            Strike = deco ? fl.Strike : ts.Strike,
            Background = bg ? fl.Background : ts.Background,
        };
    }

    /// <summary>Элементы flex/grid, float и absolute становятся блочными (CSS blockification).</summary>
    static void Blockify(CssStyle s, CssStyle parent)
    {
        if (s.Display != "inline") return;
        if (parent.Display is "flex" or "grid" || s.Float != null || s.Position is "absolute" or "fixed") s.Display = "block";
    }

    /// <summary>Строчный элемент с рамкой, отступами или полями получает собственный строчный блок.</summary>
    static void SetupInlineBox(CssStyle s)
    {
        if (s.Display != "inline") { s.InlineBox = null; return; }
        var t = s.BorderTop?.ToBorderDef(); var b = s.BorderBottom?.ToBorderDef();
        var l = s.BorderLeft?.ToBorderDef(); var r = s.BorderRight?.ToBorderDef();
        bool pad = s.PaddingLeft > 0 || s.PaddingRight > 0 || s.PaddingTop > 0 || s.PaddingBottom > 0;
        bool margin = Math.Abs(s.MarginLeft) > 0.01 || Math.Abs(s.MarginRight) > 0.01;
        if (t == null && b == null && l == null && r == null && !pad && !margin) return;   // наследуем внешний блок
        var box = new InlineBox
        {
            Parent = s.InlineBox, Background = s.Background, Top = t, Bottom = b, Left = l, Right = r,
            PadTop = s.PaddingTop, PadBottom = s.PaddingBottom, PadLeft = s.PaddingLeft, PadRight = s.PaddingRight,
            MarginLeft = s.MarginLeft, MarginRight = s.MarginRight, Radius = RadiusOf(s),
        };
        if (box.Background != null) s.TextBackground = null;   // фон рисует сам блок (вместе с padding)
        s.InlineBox = box;
    }

    static void FixTextBackground(CssStyle s)
    {
        if (s.Display is "block" or "list-item" or "table" or "table-cell" or "flex" or "grid") s.TextBackground = null;
        if (s.Display == "inline" && s.Background is { } bg) s.TextBackground = bg;
    }

    /// <summary>Объявления, применимые к элементу (или его псевдоэлементу), в порядке каскада.</summary>
    List<CssDecl> CollectDecls(Node n, string? pseudo)
    {
        var decls = new List<(CssDecl D, int Prio)>();
        AddMatches(_ua, n, decls, 0, pseudo);
        if (pseudo == null) foreach (var d in PresentationalHints(n)) decls.Add((d, 1_000_000));
        AddMatches(_author, n, decls, 2_000_000, pseudo);
        if (pseudo == null && n.Attr("style") is { } inline)
            foreach (var d in StyleSheet.ParseDeclarations(inline)) decls.Add((d, 3_000_000));
        // !important поверх всего
        return decls.Select((x, i) => (x.D, Key: (x.D.Important ? 10_000_000 : 0) + x.Prio, i))
            .OrderBy(x => x.Key).ThenBy(x => x.i).Select(x => x.D).ToList();
    }

    /// <summary>Применяет объявления: пользовательские свойства, подстановка var(), затем остальное.</summary>
    void ApplyAll(CssStyle s, CssStyle parent, List<CssDecl> ordered)
    {
        bool copied = false;
        foreach (var d in ordered)
        {
            if (!d.Property.StartsWith("--")) continue;
            if (!copied) { s.Vars = s.Vars == null ? new() : new(s.Vars); copied = true; }
            s.Vars![d.Property] = d.Value;
        }
        var resolved = new List<CssDecl>();
        foreach (var d in ordered)
        {
            if (d.Property.StartsWith("--")) continue;
            var v = CssValues.SubstituteVars(d.Value, s.Vars);
            if (v == null) continue;   // var() без значения и запасного варианта — объявление недействительно
            resolved.Add(v == d.Value ? d : d with { Value = v });
        }
        // Сначала размер шрифта (от него считаются em)
        foreach (var d in resolved.Where(d => d.Property is "font-size" or "font"))
            Apply(s, parent, d);
        foreach (var d in resolved.Where(d => d.Property is not "font-size"))
            Apply(s, parent, d);
    }

    // =================================================================
    //  ::before / ::after
    // =================================================================

    Node? MakePseudo(Node n, CssStyle s, string which)
    {
        var decls = CollectDecls(n, which);
        if (decls.Count == 0) return null;
        var ps = s.InheritFrom();
        ps.Display = "inline";
        ApplyAll(ps, s, decls);
        if (ps.Content is null || ps.Display == "none") return null;
        string c = ps.Content.Trim();
        if (c is "none" or "normal" || c.Length == 0) return null;
        Blockify(ps, s);
        FixTextBackground(ps);
        SetupInlineBox(ps);
        ps.ParentPerspective = s.Perspective;
        ps.ParentPerspectiveOrigin = s.PerspectiveOrigin;
        ps.ParentPreserve3D = s.Preserve3D;

        ApplyCounters(ps, n);
        var node = new Node { Tag = "::" + which, Parent = n, Style = ps };
        var text = new System.Text.StringBuilder();
        void FlushText()
        {
            if (text.Length == 0) return;
            node.Children.Add(new Node { Tag = "#text", Text = text.ToString(), Parent = node });
            text.Clear();
        }
        foreach (var (kind, value) in ParseContent(c))
        {
            switch (kind)
            {
                case "string": text.Append(value); break;
                case "attr": text.Append(n.Attr(value.Trim()) ?? ""); break;
                case "counter":
                case "counters":
                {
                    var args = SplitArgs(value);
                    string name = args.ElementAtOrDefault(0)?.Trim() ?? "";
                    var stack = _counters.TryGetValue(name, out var st) && st.Count > 0 ? st : null;
                    if (kind == "counter")
                    {
                        string style = args.ElementAtOrDefault(1)?.Trim() ?? "decimal";
                        text.Append(CounterText(style, stack?[^1].Value ?? 0));
                    }
                    else
                    {
                        string sep = Unquote(args.ElementAtOrDefault(1)?.Trim() ?? "\".\"");
                        string style = args.ElementAtOrDefault(2)?.Trim() ?? "decimal";
                        text.Append(stack == null ? CounterText(style, 0) : string.Join(sep, stack.Select(x => CounterText(style, x.Value))));
                    }
                    break;
                }
                case "open-quote": text.Append(Quote(ps, _quoteDepth++, open: true)); break;
                case "close-quote": text.Append(Quote(ps, _quoteDepth = Math.Max(0, _quoteDepth - 1), open: false)); break;
                case "no-open-quote": _quoteDepth++; break;
                case "no-close-quote": _quoteDepth = Math.Max(0, _quoteDepth - 1); break;
                case "url":
                    FlushText();
                    var img = new Node { Tag = "img", Parent = node, Style = ps.InheritFrom() };
                    img.Style.Display = "inline";
                    img.Attrs["src"] = Unquote(value.Trim());
                    node.Children.Add(img);
                    break;
            }
        }
        FlushText();
        return node;
    }

    int _quoteDepth;
    string[] _defaultQuotes = ["\u201C", "\u201D", "\u2018", "\u2019"];

    string Quote(CssStyle s, int depth, bool open)
    {
        var q = _defaultQuotes;
        if (s.Quotes != null && s.Quotes != "auto")
        {
            var list = ParseContent(s.Quotes).Where(x => x.Kind == "string").Select(x => x.Value).ToArray();
            if (list.Length >= 2) q = list;
            else if (s.Quotes == "none") return "";
        }
        int pairs = q.Length / 2;
        int idx = Math.Min(depth, pairs - 1) * 2 + (open ? 0 : 1);
        return q[idx];
    }

    static string Unquote(string v) =>
        v.Length >= 2 && (v[0] == '"' || v[0] == '\'') && v[^1] == v[0] ? UnescapeCss(v[1..^1]) : v;

    static List<string> SplitArgs(string v)
    {
        var list = new List<string>();
        int depth = 0, start = 0;
        char quote = '\0';
        for (int i = 0; i < v.Length; i++)
        {
            char c = v[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == ',' && depth == 0) { list.Add(v[start..i]); start = i + 1; }
        }
        list.Add(v[start..]);
        return list;
    }

    /// <summary>Экранирование в строках CSS: \201C, \A, \" …</summary>
    static string UnescapeCss(string s)
    {
        if (s.IndexOf('\\') < 0) return s;
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] != '\\' || i + 1 >= s.Length) { sb.Append(s[i]); continue; }
            int j = i + 1, st = j;
            while (j < s.Length && j - st < 6 && Uri.IsHexDigit(s[j])) j++;
            if (j > st)
            {
                int cp = int.Parse(s[st..j], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                if (cp == 0xA) sb.Append('\n');
                else if (cp is > 0 and <= 0x10FFFF and not (>= 0xD800 and <= 0xDFFF)) sb.Append(char.ConvertFromUtf32(cp));
                if (j < s.Length && s[j] == ' ') j++;
                i = j - 1;
            }
            else { sb.Append(s[j]); i = j; }
        }
        return sb.ToString();
    }

    /// <summary>Разбор значения content: строки, attr(), counter(), counters(), кавычки, url().</summary>
    static List<(string Kind, string Value)> ParseContent(string v)
    {
        var parts = new List<(string, string)>();
        int i = 0;
        while (i < v.Length)
        {
            char c = v[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (c is '"' or '\'')
            {
                int j = i + 1;
                var sb = new System.Text.StringBuilder();
                while (j < v.Length && v[j] != c)
                {
                    if (v[j] == '\\' && j + 1 < v.Length) { sb.Append(v[j]).Append(v[j + 1]); j += 2; continue; }
                    sb.Append(v[j++]);
                }
                parts.Add(("string", UnescapeCss(sb.ToString())));
                i = j + 1;
                continue;
            }
            int st = i;
            while (i < v.Length && (char.IsLetterOrDigit(v[i]) || v[i] is '-' or '_')) i++;
            if (i == st) { i++; continue; }
            string ident = v[st..i].ToLowerInvariant();
            if (i < v.Length && v[i] == '(')
            {
                int depth = 0, a = i + 1;
                for (; i < v.Length; i++)
                {
                    if (v[i] == '(') depth++;
                    else if (v[i] == ')' && --depth == 0) break;
                }
                parts.Add((ident, v[a..Math.Min(i, v.Length)]));
                i++;
            }
            else parts.Add((ident, ""));
        }
        return parts;
    }

    // =================================================================
    //  Счётчики CSS (counter-reset / counter-set / counter-increment)
    // =================================================================

    readonly Dictionary<string, List<(Node Owner, int Value)>> _counters = new();

    static IEnumerable<(string Name, int Value)> CounterPairs(string? v, int def)
    {
        if (string.IsNullOrWhiteSpace(v) || v.Trim() == "none") yield break;
        var t = CssValues.Tokens(v);
        for (int i = 0; i < t.Count; i++)
        {
            string name = t[i];
            int val = def;
            if (i + 1 < t.Count && int.TryParse(t[i + 1], out var x)) { val = x; i++; }
            yield return (name, val);
        }
    }

    void ApplyCounters(CssStyle s, Node scopeOwner)
    {
        foreach (var (name, val) in CounterPairs(s.CounterReset, 0))
        {
            if (!_counters.TryGetValue(name, out var list)) _counters[name] = list = [];
            if (list.Count > 0 && list[^1].Owner == scopeOwner) list[^1] = (scopeOwner, val);
            else list.Add((scopeOwner, val));
        }
        foreach (var (name, val) in CounterPairs(s.CounterSet, 0))
        {
            if (!_counters.TryGetValue(name, out var list) || list.Count == 0) _counters[name] = list = [(scopeOwner, 0)];
            list[^1] = (list[^1].Owner, val);
        }
        foreach (var (name, val) in CounterPairs(s.CounterIncrement, 1))
        {
            if (!_counters.TryGetValue(name, out var list) || list.Count == 0) _counters[name] = list = [(scopeOwner, 0)];
            list[^1] = (list[^1].Owner, list[^1].Value + val);
        }
    }

    /// <summary>Закрывает области видимости счётчиков, созданных внутри элемента (в т.ч. его ::before/::after).</summary>
    void PopCounters(Node owner)
    {
        foreach (var list in _counters.Values)
            list.RemoveAll(x => x.Owner == owner);
    }

    static string CounterText(string style, int n)
    {
        string label = ListLabel(style.Trim().ToLowerInvariant(), n);
        if (style is "decimal" or "decimal-leading-zero" or "lower-alpha" or "lower-latin" or "upper-alpha" or "upper-latin"
            or "lower-roman" or "upper-roman" or "lower-greek" or "cyrillic-lower" or "cyrillic-upper")
            label = label.TrimEnd('.', ')');
        return label;
    }

    static bool IsExternalLink(string href) =>
        href.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || href.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
        || href.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) || href.StartsWith("ftp://", StringComparison.OrdinalIgnoreCase);

    static void AddMatches(StyleSheet sheet, Node n, List<(CssDecl, int)> output, int basePrio, string? pseudo)
    {
        var matched = new List<(int Spec, int Order, List<CssDecl> Decls)>();
        foreach (var rule in sheet.Rules)
        {
            int best = -1;
            foreach (var sel in rule.Selectors)
                if (sel.PseudoElement == pseudo && sel.Specificity > best && sel.Matches(n)) best = sel.Specificity;
            if (best >= 0) matched.Add((best, rule.Order, rule.Decls));
        }
        foreach (var m in matched.OrderBy(m => m.Spec).ThenBy(m => m.Order))
            foreach (var d in m.Decls) output.Add((d, basePrio + Math.Min(m.Spec, 999_999)));
    }

    /// <summary>HTML-атрибуты оформления (align, bgcolor, …) как CSS с низким приоритетом.</summary>
    static IEnumerable<CssDecl> PresentationalHints(Node n)
    {
        CssDecl D(string p, string v) => new(p, v, false);
        if (n.Attr("align") is { } align)
        {
            if (n.Tag is "table" or "img") { if (align.Equals("center", StringComparison.OrdinalIgnoreCase)) yield return D("margin-left", "auto"); }
            else yield return D("text-align", align.ToLowerInvariant());
        }
        if (n.Attr("valign") is { } valign) yield return D("vertical-align", valign.ToLowerInvariant());
        if (n.Attr("bgcolor") is { } bg) yield return D("background-color", bg);
        if (n.Tag == "font")
        {
            if (n.Attr("color") is { } color) yield return D("color", color);
            if (n.Attr("face") is { } face) yield return D("font-family", face);
            if (n.Attr("size") is { } size)
            {
                string sz = size.Trim();
                int v = 3;
                if (sz.StartsWith('+') || sz.StartsWith('-')) { if (int.TryParse(sz, out var rel)) v = 3 + rel; }
                else int.TryParse(sz, out v);
                double[] px = [10, 13, 16, 18, 24, 32, 48];
                yield return D("font-size", px[Math.Clamp(v, 1, 7) - 1].ToString(CultureInfo.InvariantCulture) + "px");
            }
        }
        if (n.Tag is "td" or "th" or "table" or "img" or "col" or "tr")
        {
            if (n.Attr("width") is { } w) yield return D("width", w.EndsWith('%') ? w : w.TrimEnd('p', 'x') + "px");
            if (n.Attr("height") is { } h && !h.EndsWith('%')) yield return D("height", h.TrimEnd('p', 'x') + "px");
        }
        if (n.Tag is "td" or "th" && n.Attr("nowrap") != null) yield return D("white-space", "nowrap");
        if (n.Tag is "ol" or "ul" && n.Attr("type") is { } type)
            yield return D("list-style-type", type switch
            {
                "1" => "decimal", "a" => "lower-alpha", "A" => "upper-alpha", "i" => "lower-roman", "I" => "upper-roman",
                _ => type.ToLowerInvariant()
            });
    }

    string PickFont(string value)
    {
        if (_fontPick.TryGetValue(value, out var cached)) return cached!;
        string? result = null;
        string? generic = null;
        foreach (var raw in value.Split(','))
        {
            string f = raw.Trim().Trim('"', '\'').Trim();
            if (f.Length == 0) continue;
            string? g = f.ToLowerInvariant() switch
            {
                "serif" => "Times New Roman",
                "sans-serif" or "system-ui" or "-apple-system" or "blinkmacsystemfont" or "ui-sans-serif" => "Arial",
                "monospace" or "ui-monospace" => "Courier New",
                "cursive" or "fantasy" => "Comic Sans MS",
                _ => null
            };
            if (g != null) { generic ??= g; if (_fonts.HasFamily(g)) { result = g; break; } continue; }
            if (_fonts.HasFamily(f)) { result = f; break; }
        }
        result ??= generic ?? value.Split(',')[0].Trim().Trim('"', '\'');
        _fontPick[value] = result;
        return result;
    }

    void Apply(CssStyle s, CssStyle parent, CssDecl d)
    {
        string v = d.Value.Trim();
        string lv = v.ToLowerInvariant();
        double fs = s.FontSize;
        if (lv == "inherit")
        {
            // Упрощённо: наследуемые свойства уже унаследованы, для остальных — игнорируем
            return;
        }
        switch (d.Property)
        {
            case "color": if (CssValues.Color(v) is Rgb c) s.Color = c; break;
            case "background-color": s.Background = CssValues.Color(v); break;
            case "background":
                s.Background = null;
                foreach (var t in CssValues.Tokens(v)) if (CssValues.Color(t) is Rgb bc) s.Background = bc;
                break;
            case "font-family": s.FontFamily = PickFont(v); break;
            case "font-size": s.FontSize = FontSize(lv, parent.FontSize); break;
            case "font-weight":
                s.Bold = lv switch { "bold" or "bolder" => true, "normal" or "lighter" => false, _ => int.TryParse(lv, out var w) ? w >= 600 : s.Bold };
                s.Weight = int.TryParse(lv, out var nw) && nw is >= 1 and <= 1000 && nw != 400 && nw != 700 ? nw : 0;
                break;
            case "font-style": s.Italic = lv is "italic" or "oblique"; break;
            case "font-kerning": s.Kerning = lv != "none"; break;
            case "hyphens" or "-webkit-hyphens" or "-ms-hyphens": s.Hyphens = lv switch { "auto" => 2, "none" => 0, _ => 1 }; break;
            case "hyphenate-limit-chars":
            {
                var t = lv.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                int P(int i, int d) => i < t.Length && int.TryParse(t[i], out var x) ? x : d;
                s.HyphenLimitChars = (P(0, 5), P(1, t.Length > 1 ? 2 : 2), P(2, t.Length > 2 ? 2 : P(1, 2)));
                break;
            }
            case "font": ApplyFontShorthand(s, parent, v); break;
            case "text-decoration" or "text-decoration-line":
                if (lv.Contains("none")) { s.Underline = false; s.Strike = false; }
                if (lv.Contains("underline")) s.Underline = true;
                if (lv.Contains("line-through")) s.Strike = true;
                break;
            case "text-align":
                s.TextAlign = lv switch { "center" or "-webkit-center" => "center", "right" or "end" or "-webkit-right" => "right", "justify" => "both", _ => "left" };
                break;
            case "text-indent": s.TextIndent = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "line-height":
                if (lv == "normal") { s.LineHeightFactor = null; s.LineHeightPt = null; }
                else if (double.TryParse(lv, NumberStyles.Float, CultureInfo.InvariantCulture, out var factor)) { s.LineHeightFactor = factor; s.LineHeightPt = null; }
                else if (lv.EndsWith('%') && double.TryParse(lv[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pct)) { s.LineHeightFactor = pct / 100; s.LineHeightPt = null; }
                else if (CssValues.Length(lv, fs) is double lh) { s.LineHeightPt = lh; s.LineHeightFactor = null; }
                break;
            case "white-space": s.WhiteSpace = lv; break;
            case "text-transform": s.Upper = lv == "uppercase"; s.Lower = lv == "lowercase"; break;
            case "vertical-align":
                s.VertAlign = lv switch { "super" => 1, "sub" => -1, _ => 0 };
                if (lv is "top" or "middle" or "bottom" or "baseline" or "text-top" or "text-bottom") s.VerticalAlign = lv;
                break;
            case "list-style-type": s.ListStyleType = lv; break;
            case "list-style":
                foreach (var t in CssValues.Tokens(lv))
                    if (t is "none" or "disc" or "circle" or "square" or "decimal" or "decimal-leading-zero" or "lower-alpha" or "upper-alpha"
                        or "lower-latin" or "upper-latin" or "lower-roman" or "upper-roman" or "lower-greek" or "cyrillic-lower" or "cyrillic-upper")
                        s.ListStyleType = t;
                break;
            case "display":
                s.MasonryDisplay = lv is "grid-lanes" or "inline-grid-lanes" or "masonry" or "inline-masonry";
                s.Display = lv switch
                {
                    "none" => "none",
                    "inline" or "inline-block" or "contents" => "inline",
                    "list-item" => "list-item",
                    "flex" or "inline-flex" => "flex",
                    "grid" or "inline-grid" => "grid",
                    "grid-lanes" or "inline-grid-lanes" or "masonry" or "inline-masonry" => "grid",
                    "table" or "inline-table" => "table",
                    "table-row" or "table-cell" or "table-row-group" or "table-header-group" or "table-footer-group" or "table-caption" => lv,
                    _ => "block"
                };
                break;
            case "box-sizing": s.BoxSizing = lv; break;
            case "transform" or "-webkit-transform" or "-ms-transform" or "-moz-transform" or "-o-transform":
                s.Transform = lv is "none" ? null : lv; break;
            case "transform-origin" or "-webkit-transform-origin" or "-ms-transform-origin": s.TransformOrigin = lv; break;
            case "translate": s.Translate = lv is "none" ? null : lv; break;
            case "perspective" or "-webkit-perspective": s.Perspective = lv is "none" ? null : lv; break;
            case "perspective-origin" or "-webkit-perspective-origin": s.PerspectiveOrigin = lv; break;
            case "backface-visibility" or "-webkit-backface-visibility": s.BackfaceHidden = lv == "hidden"; break;
            case "transform-style" or "-webkit-transform-style": s.Preserve3D = lv == "preserve-3d"; break;
            case "rotate": s.Rotate = lv is "none" ? null : lv; break;
            case "scale": s.Scale = lv is "none" ? null : lv; break;
            case "word-spacing": s.WordSpacing = lv == "normal" ? 0 : CssValues.Length(lv, fs) ?? 0; break;
            case "border-radius" or "-webkit-border-radius" or "-moz-border-radius":
            {
                // "гориз. 1–4 значения [/ верт. 1–4 значения]", порядок углов TL TR BR BL
                static string[] Expand(List<string> t) => t.Count switch
                {
                    0 => ["0", "0", "0", "0"],
                    1 => [t[0], t[0], t[0], t[0]],
                    2 => [t[0], t[1], t[0], t[1]],
                    3 => [t[0], t[1], t[2], t[1]],
                    _ => [t[0], t[1], t[2], t[3]],
                };
                int slash = lv.IndexOf('/');
                var hr = Expand(CssValues.Tokens(slash < 0 ? lv : lv[..slash]));
                var vr = slash < 0 ? hr : Expand(CssValues.Tokens(lv[(slash + 1)..]));
                s.Radius = [.. hr, .. vr];
                break;
            }
            case "border-top-left-radius" or "border-top-right-radius" or "border-bottom-right-radius" or "border-bottom-left-radius":
            {
                // одно значение — круглый угол, два — эллиптический (гориз. верт.)
                int corner = d.Property switch
                {
                    "border-top-left-radius" => 0, "border-top-right-radius" => 1, "border-bottom-right-radius" => 2, _ => 3
                };
                var t = CssValues.Tokens(lv);
                if (t.Count == 0) break;
                s.Radius[corner] = t[0];
                s.Radius[corner + 4] = t.Count > 1 ? t[1] : t[0];
                break;
            }
            case "initial-letter" or "-webkit-initial-letter": s.InitialLetter = lv is "normal" ? null : lv; break;
            case "letter-spacing": s.LetterSpacing = lv == "normal" ? 0 : CssValues.Length(lv, fs) ?? 0; break;
            case "font-variant" or "font-variant-caps":
                if (lv.Contains("small-caps")) s.SmallCaps = true;
                else if (lv == "normal") s.SmallCaps = false;
                break;
            case "content": s.Content = v; break;
            case "counter-reset": s.CounterReset = v; break;
            case "counter-increment": s.CounterIncrement = v; break;
            case "counter-set": s.CounterSet = v; break;
            case "quotes": s.Quotes = v; break;
            case "float": s.Float = lv switch { "left" or "inline-start" => "left", "right" or "inline-end" => "right", _ => null }; break;
            case "clear": s.Clear = lv switch { "left" or "right" or "both" => lv, "inline-start" => "left", "inline-end" => "right", _ => null }; break;
            case "position": s.Position = lv switch { "sticky" => "relative", "relative" or "absolute" or "fixed" => lv, _ => "static" }; break;
            case "left": s.Left = lv == "auto" ? null : lv; break;
            case "top": s.Top = lv == "auto" ? null : lv; break;
            case "right": s.Right = lv == "auto" ? null : lv; break;
            case "bottom": s.Bottom = lv == "auto" ? null : lv; break;
            case "inset":
            {
                var (t, r, bt, l) = CssValues.Box(lv);
                s.Top = t == "auto" ? null : t; s.Right = r == "auto" ? null : r; s.Bottom = bt == "auto" ? null : bt; s.Left = l == "auto" ? null : l;
                break;
            }
            case "z-index": s.ZIndex = int.TryParse(lv, out var z) ? z : 0; break;
            case "flex-direction": s.FlexDirection = lv; break;
            case "flex-wrap": s.FlexWrap = lv is "wrap" or "wrap-reverse"; break;
            case "flex-flow":
                foreach (var t in CssValues.Tokens(lv))
                {
                    if (t.StartsWith("row") || t.StartsWith("column")) s.FlexDirection = t;
                    else s.FlexWrap = t is "wrap" or "wrap-reverse";
                }
                break;
            case "justify-content": s.JustifyContent = lv.Replace("safe ", "").Replace("unsafe ", ""); break;
            case "align-items": s.AlignItems = lv; break;
            case "align-self": s.AlignSelf = lv == "auto" ? null : lv; break;
            case "flex-grow": s.FlexGrow = ParseNum(lv, 0); break;
            case "flex-shrink": s.FlexShrink = ParseNum(lv, 1); break;
            case "flex-basis": s.FlexBasis = lv; break;
            case "flex":
            {
                var t = CssValues.Tokens(lv);
                if (lv == "none") { s.FlexGrow = 0; s.FlexShrink = 0; s.FlexBasis = "auto"; }
                else if (lv == "auto") { s.FlexGrow = 1; s.FlexShrink = 1; s.FlexBasis = "auto"; }
                else if (lv == "initial") { s.FlexGrow = 0; s.FlexShrink = 1; s.FlexBasis = "auto"; }
                else
                {
                    var nums = t.Where(x => double.TryParse(x, NumberStyles.Float, CultureInfo.InvariantCulture, out _)).ToList();
                    var basis = t.FirstOrDefault(x => !double.TryParse(x, NumberStyles.Float, CultureInfo.InvariantCulture, out _));
                    s.FlexGrow = nums.Count > 0 ? ParseNum(nums[0], 0) : 1;
                    s.FlexShrink = nums.Count > 1 ? ParseNum(nums[1], 1) : 1;
                    s.FlexBasis = basis ?? "0";
                }
                break;
            }
            case "order": s.Order = int.TryParse(lv, out var ord) ? ord : 0; break;
            case "gap" or "grid-gap":
            {
                var t = CssValues.Tokens(lv);
                s.RowGap = t[0]; s.ColumnGap = t.Count > 1 ? t[1] : t[0];
                break;
            }
            case "row-gap" or "grid-row-gap": s.RowGap = lv; break;
            case "column-gap" or "grid-column-gap": s.ColumnGap = lv; break;
            case "grid-template-columns": s.GridTemplateColumns = lv; s.MasonryCols = lv == "masonry"; break;
            case "align-tracks": s.AlignTracks = lv; break;
            case "justify-tracks": s.JustifyTracks = lv; break;
            case "grid-template-rows": s.GridTemplateRows = lv; s.MasonryRows = lv == "masonry"; break;
            case "grid-auto-flow":
            {
                var t = CssValues.Tokens(lv);
                s.GridFlowColumn = t.Contains("column");
                s.GridFlowDense = t.Contains("dense");
                break;
            }
            case "grid-auto-columns": s.GridAutoColumns = lv; break;
            case "grid-auto-rows": s.GridAutoRows = lv; break;
            case "grid-column": (s.GridColumnStart, s.GridColumnEnd) = SplitSlash(lv); break;
            case "grid-row": (s.GridRowStart, s.GridRowEnd) = SplitSlash(lv); break;
            case "grid-column-start": s.GridColumnStart = lv; break;
            case "grid-column-end": s.GridColumnEnd = lv; break;
            case "grid-row-start": s.GridRowStart = lv; break;
            case "grid-row-end": s.GridRowEnd = lv; break;
            case "grid-template-areas": s.GridTemplateAreas = lv == "none" ? null : lv; break;
            case "align-content": s.AlignContent = lv.Replace("safe ", "").Replace("unsafe ", "").Trim(); break;
            case "place-content":
            {
                var pc = lv.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (pc.Length > 0) { s.AlignContent = pc[0]; s.JustifyContent = pc.Length > 1 ? pc[1] : pc[0]; }
                break;
            }
            case "grid-area":
            {
                var parts = lv.Split('/').Select(x => x.Trim()).ToArray();
                // grid-area: имя — все четыре линии по имени области
                if (parts.Length == 1 && parts[0].Length > 0 && !char.IsDigit(parts[0][0]) && parts[0][0] != '-' && !parts[0].StartsWith("span") && parts[0] != "auto")
                { s.GridRowStart = s.GridColumnStart = s.GridRowEnd = s.GridColumnEnd = parts[0]; break; }
                if (parts.Length >= 1) s.GridRowStart = parts[0];
                if (parts.Length >= 2) s.GridColumnStart = parts[1];
                if (parts.Length >= 3) s.GridRowEnd = parts[2];
                if (parts.Length >= 4) s.GridColumnEnd = parts[3];
                break;
            }
            case "visibility": s.Hidden = lv is "hidden" or "collapse"; break;
            case "margin":
            {
                var (t, r, b, l) = CssValues.Box(lv);
                s.MarginTop = CssValues.Length(t, fs, _contentWidth) ?? 0;
                s.MarginRight = CssValues.Length(r, fs, _contentWidth) ?? 0;
                s.MarginBottom = CssValues.Length(b, fs, _contentWidth) ?? 0;
                s.MarginLeft = CssValues.Length(l, fs, _contentWidth) ?? 0;
                s.MarginLeftAuto = l == "auto"; s.MarginRightAuto = r == "auto";
                break;
            }
            case "margin-top": s.MarginTop = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "margin-bottom": s.MarginBottom = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "margin-left" or "margin-inline-start": s.MarginLeft = CssValues.Length(lv, fs, _contentWidth) ?? 0; s.MarginLeftAuto = lv == "auto"; break;
            case "margin-right" or "margin-inline-end": s.MarginRight = CssValues.Length(lv, fs, _contentWidth) ?? 0; s.MarginRightAuto = lv == "auto"; break;
            case "padding":
            {
                var (t, r, b, l) = CssValues.Box(lv);
                s.PaddingTop = CssValues.Length(t, fs, _contentWidth) ?? 0;
                s.PaddingRight = CssValues.Length(r, fs, _contentWidth) ?? 0;
                s.PaddingBottom = CssValues.Length(b, fs, _contentWidth) ?? 0;
                s.PaddingLeft = CssValues.Length(l, fs, _contentWidth) ?? 0;
                break;
            }
            case "padding-top": s.PaddingTop = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "padding-bottom": s.PaddingBottom = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "padding-left" or "padding-inline-start": s.PaddingLeft = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "padding-right" or "padding-inline-end": s.PaddingRight = CssValues.Length(lv, fs, _contentWidth) ?? 0; break;
            case "border":
                s.BorderTop = s.BorderRight = s.BorderBottom = s.BorderLeft = CssValues.Border(v, fs, null);
                break;
            case "border-top": s.BorderTop = CssValues.Border(v, fs, null); break;
            case "border-bottom": s.BorderBottom = CssValues.Border(v, fs, null); break;
            case "border-left": s.BorderLeft = CssValues.Border(v, fs, null); break;
            case "border-right": s.BorderRight = CssValues.Border(v, fs, null); break;
            case "border-width" or "border-style" or "border-color":
            {
                var (t, r, b, l) = CssValues.Box(v);
                s.BorderTop = CssValues.Border(t, fs, s.BorderTop);
                s.BorderRight = CssValues.Border(r, fs, s.BorderRight);
                s.BorderBottom = CssValues.Border(b, fs, s.BorderBottom);
                s.BorderLeft = CssValues.Border(l, fs, s.BorderLeft);
                break;
            }
            case "border-top-width" or "border-top-style" or "border-top-color": s.BorderTop = CssValues.Border(v, fs, s.BorderTop); break;
            case "border-bottom-width" or "border-bottom-style" or "border-bottom-color": s.BorderBottom = CssValues.Border(v, fs, s.BorderBottom); break;
            case "border-left-width" or "border-left-style" or "border-left-color": s.BorderLeft = CssValues.Border(v, fs, s.BorderLeft); break;
            case "border-right-width" or "border-right-style" or "border-right-color": s.BorderRight = CssValues.Border(v, fs, s.BorderRight); break;
            case "width": s.Width = lv; break;
            case "height": s.Height = lv; break;
            case "page-break-before" or "break-before": s.PageBreakBefore = lv is "always" or "page" or "left" or "right"; break;
            case "page-break-after" or "break-after": s.PageBreakAfter = lv is "always" or "page" or "left" or "right"; break;
            case "page-break-inside" or "break-inside": s.AvoidBreakInside = lv.StartsWith("avoid"); break;
        }
    }

    static double ParseNum(string v, double def) =>
        double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : def;

    static (string? A, string? B) SplitSlash(string v)
    {
        int i = v.IndexOf('/');
        return i < 0 ? (v.Trim(), null) : (v[..i].Trim(), v[(i + 1)..].Trim());
    }

    static double FontSize(string v, double parent) => v switch
    {
        "xx-small" => 7.5, "x-small" => 7.5, "small" => 9.75, "medium" => 12, "large" => 13.5,
        "x-large" => 18, "xx-large" => 24, "xxx-large" => 36,
        "smaller" => parent / 1.2, "larger" => parent * 1.2,
        _ => CssValues.Length(v, parent, parent) is double d && d > 0 ? d : parent
    };

    void ApplyFontShorthand(CssStyle s, CssStyle parent, string v)
    {
        var tokens = CssValues.Tokens(v);
        s.Weight = 0;
        int i = 0;
        for (; i < tokens.Count; i++)
        {
            string t = tokens[i].ToLowerInvariant();
            if (t is "italic" or "oblique") s.Italic = true;
            else if (t is "bold" or "bolder") { s.Bold = true; s.Weight = 0; }
            else if (int.TryParse(t, out var w) && w >= 100 && w <= 900) { s.Bold = w >= 600; s.Weight = w is 400 or 700 ? 0 : w; }
            else if (t == "small-caps") s.SmallCaps = true;
            else if (t is "normal") { }
            else break;
        }
        if (i >= tokens.Count) return;
        string size = tokens[i];
        string? lh = null;
        int slash = size.IndexOf('/');
        if (slash >= 0) { lh = size[(slash + 1)..]; size = size[..slash]; }
        s.FontSize = FontSize(size.ToLowerInvariant(), parent.FontSize);
        if (lh != null) Apply(s, parent, new CssDecl("line-height", lh, false));
        string family = string.Join(' ', tokens.Skip(i + 1));
        if (family.Length > 0) s.FontFamily = PickFont(family);
    }

    // =================================================================
    //  Поток блоков
    // =================================================================

    sealed class Flow(List<Block> target)
    {
        public readonly List<Block> Target = target;
        public Paragraph? Cur;
        public BlockCtx? CurCtx;
        public bool LastSpace = true;
        public double PendingMargin;
        public bool PendingBreak;
        public string? PendingLabel;
        public string? PendingClear;
        public TextStyle? LabelStyle;
        public bool FirstNewlineInPre;
    }

    sealed class BlockCtx
    {
        public double IndL, IndR;
        public required CssStyle Owner;
        public double Width;              // доступная ширина
        public Rgb? Shading;
        public ListState? List;
        public AnchorBlock? Anchor;
    }

    sealed class ListState
    {
        public int Counter;
        public bool Ordered;
        public int Depth;
    }

    TextStyle ToTextStyle(CssStyle s) => new()
    {
        FontAscii = s.FontFamily, FontHAnsi = s.FontFamily, FontEastAsia = s.FontFamily,
        Bold = s.Bold, Weight = s.Weight, Italic = s.Italic, Underline = s.Underline, Strike = s.Strike, Caps = s.Upper, SmallCaps = s.SmallCaps, LetterSpacing = s.LetterSpacing, WordSpacing = s.WordSpacing,
        Size = Math.Max(1, s.FontSize), Color = s.Color, Background = s.TextBackground, VertAlign = s.VertAlign, Link = s.Link,
        Box = s.InlineBox,
        Kerning = s.Kerning,
        Hyphens = s.Hyphens,
        Lang = s.Lang,
    };

    Paragraph EnsureParagraph(Flow f, BlockCtx b)
    {
        if (f.Cur != null) return f.Cur;
        var o = b.Owner;
        var p = new Paragraph
        {
            Align = o.TextAlign,
            IndLeft = b.IndL,
            IndRight = b.IndR,
            IndFirst = o.TextIndent,
            SpaceBefore = f.PendingMargin,
            PageBreakBefore = f.PendingBreak,
            Shading = b.Shading,
            Clear = f.PendingClear,
            MarkStyle = ToTextStyle(o),
            WidowControl = true,
            // Переносы как в Chromium: слово от 5 букв, не меньше 2 букв с каждой стороны (hyphenate-limit-chars)
            HyphenMinWord = o.HyphenLimitChars.Word, HyphenMinLeft = o.HyphenLimitChars.Before, HyphenMinRight = o.HyphenLimitChars.After,
        };
        if (o.ExactLineHeight is double ex) { p.LineRule = "exact"; p.LineValue = ex; }
        else if (o.LineHeightFactor is double lf) { p.LineRule = "auto"; p.LineValue = 240 * lf / 1.15; }
        else if (o.LineHeightPt is double lp) { p.LineRule = "css"; p.LineValue = lp; }
        f.PendingMargin = 0;
        f.PendingBreak = false;
        f.PendingClear = null;
        if (f.PendingLabel != null)
        {
            var ls = f.LabelStyle!;
            p.Inlines.Add(new TextInline(f.PendingLabel, ls));
            p.Inlines.Add(new TabInline(ls));
            double hang = Math.Min(Math.Max(18, ls.Size * 1.6), Math.Max(10, b.IndL));
            p.IndFirst = -hang;
            p.Tabs = [];
            f.PendingLabel = null;
        }
        f.Cur = p;
        f.CurCtx = b;
        return p;
    }

    static void Flush(Flow f)
    {
        f.LastSpace = true;
        var p = f.Cur;
        if (p == null) return;
        f.Cur = null;
        // Обрезаем конечные пробелы и завершающий <br>
        while (p.Inlines.Count > 0)
        {
            var last = p.Inlines[^1];
            if (last is TextInline t)
            {
                t.Text = t.Text.TrimEnd(' ');
                if (t.Text.Length == 0) { p.Inlines.RemoveAt(p.Inlines.Count - 1); continue; }
            }
            break;
        }
        if (p.Inlines.Count > 0 && p.Inlines[^1] is BreakInline { Kind: BreakKind.Line } && p.Inlines.Count > 1)
            p.Inlines.RemoveAt(p.Inlines.Count - 1);
        bool meaningful = p.Inlines.Any(i => i is not TextInline || ((TextInline)i).Text.Trim(' ').Length > 0);
        if (!meaningful)
        {
            f.PendingMargin = Math.Max(f.PendingMargin, p.SpaceBefore);
            f.PendingBreak |= p.PageBreakBefore;
            return;
        }
        f.Target.Add(p);
    }

    void AddInline(Flow f, BlockCtx b, Inline inl)
    {
        var p = EnsureParagraph(f, b);
        if (inl is TextInline t && p.Inlines.Count > 0 && p.Inlines[^1] is TextInline prev && prev.Style == t.Style)
            prev.Text += t.Text;
        else p.Inlines.Add(inl);
    }

    void AddText(string text, CssStyle s, Flow f, BlockCtx b)
    {
        if (s.Hidden) return;
        if (s.Lower) text = text.ToLowerInvariant();
        var ts = ToTextStyle(s);
        string ws = s.WhiteSpace;
        if (ws is "pre" or "pre-wrap" or "break-spaces" or "pre-line")
        {
            bool collapse = ws == "pre-line";
            var sb = new StringBuilder();
            void FlushSb() { if (sb.Length > 0) { AddInline(f, b, new TextInline(sb.ToString(), ts)); sb.Clear(); } }
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r') continue;
                if (c == '\n')
                {
                    if (f.FirstNewlineInPre) { f.FirstNewlineInPre = false; continue; }
                    FlushSb();
                    AddInline(f, b, new BreakInline(BreakKind.Line, ts));
                    f.LastSpace = true;
                    continue;
                }
                f.FirstNewlineInPre = false;
                if (c == '\t' && !collapse) { FlushSb(); AddInline(f, b, new TabInline(ts)); continue; }
                if (collapse && char.IsWhiteSpace(c))
                {
                    if (!f.LastSpace) sb.Append(' ');
                    f.LastSpace = true;
                    continue;
                }
                sb.Append(c == ' ' && !collapse ? ' ' : c);
                f.LastSpace = false;
            }
            FlushSb();
            return;
        }

        var outSb = new StringBuilder();
        bool nowrap = ws == "nowrap";
        foreach (char c in text)
        {
            if (c is ' ' or '\t' or '\n' or '\r' or '\f')
            {
                if (!f.LastSpace) outSb.Append(nowrap ? ' ' : ' ');
                f.LastSpace = true;
            }
            else
            {
                outSb.Append(c);
                f.LastSpace = false;
            }
        }
        if (outSb.Length == 0) return;
        string str = outSb.ToString();
        if (f.Cur == null && str.Trim(' ').Length == 0) return;
        if (f.Cur == null) str = str.TrimStart(' ');
        AddInline(f, b, new TextInline(str, ts));
    }

    void WalkChildren(Node n, Flow f, BlockCtx b)
    {
        foreach (var c in n.Children) Walk(c, f, b);
    }

    void Walk(Node n, Flow f, BlockCtx b)
    {
        if (n.IsText)
        {
            AddText(n.Text, n.Parent?.Style ?? b.Owner, f, b);
            return;
        }
        if (!n.IsElement) { WalkChildren(n, f, b); return; }
        var s = n.Style!;
        if (s.Display == "none") return;
        bool root = n.Tag is "html" or "body";

        if (!root && s.Position is "absolute" or "fixed") { AddPositioned(n, s, f, b); return; }
        if (!root && s.Float is "left" or "right") { AddFloat(n, s, f, b); return; }

        // transform у блочного элемента: всё, что он создаст, рисуется с преобразованием
        if (!root && s.Display != "inline" && n.Tag != "img" && TransformOf(s) is { } tr && _transforming.Add(n))
        {
            Flush(f);
            int idx = f.Target.Count;
            double avail = Math.Max(10, _contentWidth - b.IndL - b.IndR - s.MarginLeft - s.MarginRight);
            double w = Math.Min(avail, BorderBoxWidth(s, avail) ?? avail);
            double bx = b.IndL + s.MarginLeft + (s.MarginLeftAuto && s.MarginRightAuto ? (avail - w) / 2 : s.MarginLeftAuto ? avail - w : 0);
            Walk(n, f, b);
            Flush(f);
            var made = f.Target.Skip(idx).ToList();
            f.Target.RemoveRange(idx, made.Count);
            if (made.Count > 0) f.Target.Add(new TransformBlock { Blocks = made, BoxX = bx, BoxWidth = w, Resolve = tr });
            _transforming.Remove(n);
            return;
        }

        switch (n.Tag)
        {
            case "br":
                AddInline(f, b, new BreakInline(BreakKind.Line, ToTextStyle(s)));
                f.LastSpace = true;
                return;
            case "wbr":
                AddInline(f, b, new TextInline("\u200B", ToTextStyle(s)));
                return;
            case "img":
                AddImage(n, s, f, b);
                return;
            case "hr":
                AddHr(s, f, b);
                return;
        }
        // flex/grid с фоном, рамкой, отступами или высотой — рамка-ячейка: фон и рамка рисуются одним
        // прямоугольником (в т.ч. под промежутками строк и при распределении align-content)
        if (s.Display is "flex" or "grid" && n.Tag is not ("html" or "body") && HasBoxDecoration(s)) { AddBoxTable(n, s, f, b); return; }
        switch (s.Display)
        {
            case "table": AddTable(n, s, f, b); return;
            case "flex": AddFlex(n, s, f, b); return;
            case "grid": AddGrid(n, s, f, b); return;
            case "inline": WalkChildren(n, f, b); return;
        }

        // Блок с боковыми рамками или высотой — как рамка-ячейка (умеет разрываться между страницами)
        if (NeedsBox(n, s)) { AddBoxTable(n, s, f, b); return; }

        // Блочный элемент
        var inner = BeginBox(n, s, f, b, out int startIdx);

        if (n.Tag is "ul" or "ol" or "menu" or "dir")
        {
            int start = 1;
            if (n.Tag == "ol" && int.TryParse(n.Attr("start"), out var st)) start = st;
            inner.List = new ListState { Counter = start - 1, Ordered = n.Tag == "ol", Depth = (b.List?.Depth ?? -1) + 1 };
        }
        if (s.Display == "list-item")
        {
            var list = b.List ?? new ListState();
            list.Counter++;
            if (int.TryParse(n.Attr("value"), out var val)) list.Counter = val;
            string label = ListLabel(s.ListStyleType, list.Counter);
            // ::marker — цвет/шрифт/содержимое маркера
            var ms = s;
            var markerDecls = CollectDecls(n, "marker");
            if (markerDecls.Count > 0)
            {
                ms = s.InheritFrom();
                ApplyAll(ms, s, markerDecls);
                if (ms.Content is { } mc && mc.Trim() is not ("normal" or ""))
                    label = mc.Trim() == "none" ? "" : string.Concat(ParseContent(mc).Where(x => x.Kind == "string").Select(x => x.Value));
            }
            if (label.Length > 0)
            {
                f.PendingLabel = label;
                var ls = ToTextStyle(ms) with { Underline = false, Strike = false, Link = null, Background = null };
                f.LabelStyle = ls;
            }
        }
        if (n.Tag is "pre" or "listing" or "xmp") f.FirstNewlineInPre = true;

        WalkChildren(n, f, b: inner);
        Flush(f);
        f.PendingLabel = null;   // пустой элемент списка
        f.FirstNewlineInPre = false;
        EndBox(n, s, f, startIdx);
    }

    /// <summary>Начало блочного контейнера: отступы, разрывы, clear, якорь для absolute-потомков.</summary>
    BlockCtx BeginBox(Node n, CssStyle s, Flow f, BlockCtx b, out int startIdx)
    {
        Flush(f);
        f.PendingMargin = Math.Max(f.PendingMargin, s.MarginTop);
        if (s.PageBreakBefore) f.PendingBreak = true;
        if (s.Clear is "left" or "right" or "both") f.PendingClear = s.Clear;

        var inner = new BlockCtx
        {
            IndL = b.IndL + s.MarginLeft + s.PaddingLeft + (s.BorderLeft?.Width ?? 0),
            IndR = b.IndR + s.MarginRight + s.PaddingRight + (s.BorderRight?.Width ?? 0),
            Owner = s,
            Shading = s.Background ?? b.Shading,
            List = b.List,
            Anchor = b.Anchor,
        };
        inner.Width = Math.Max(20, _contentWidth - inner.IndL - inner.IndR);
        double? fixedW = CssValues.Length(s.Width, s.FontSize, Math.Max(20, _contentWidth - b.IndL - b.IndR - s.MarginLeft - s.MarginRight));
        if (fixedW != null && s.BoxSizing == "border-box") fixedW -= s.PaddingLeft + s.PaddingRight + (s.BorderLeft?.Width ?? 0) + (s.BorderRight?.Width ?? 0);
        if (fixedW is > 0 && fixedW < inner.Width)
        {
            double spare = inner.Width - fixedW.Value;
            if (s.MarginLeftAuto && s.MarginRightAuto) { inner.IndL += spare / 2; inner.IndR += spare / 2; }
            else if (s.MarginLeftAuto) inner.IndL += spare;
            else inner.IndR += spare;
            inner.Width = fixedW.Value;
        }
        if (s.PaddingTop > 0) f.PendingMargin += s.PaddingTop;
        startIdx = f.Target.Count;
        if (s.Position == "relative")
        {
            var anchor = new AnchorBlock
            {
                X = inner.IndL - s.PaddingLeft, Width = inner.Width + s.PaddingLeft + s.PaddingRight, TopOffset = f.PendingMargin - s.PaddingTop,
                Height = PaddingBoxHeight(s),
            };
            f.Target.Add(anchor);
            inner.Anchor = anchor;
        }
        return inner;
    }

    /// <summary>Конец блочного контейнера: рамки, сдвиг position: relative, нижний отступ.</summary>
    void EndBox(Node n, CssStyle s, Flow f, int startIdx)
    {
        var created = f.Target.Skip(startIdx).ToList();
        ApplyFirstLine(n, s, created);
        var paras = created.OfType<Paragraph>().ToList();
        if (paras.Count > 0)
        {
            if (s.BorderTop?.ToBorderDef() is { } bt) { bt.Space = s.PaddingTop; paras[0].BorderTop = bt; }
            if (s.BorderBottom?.ToBorderDef() is { } bb) { bb.Space = s.PaddingBottom; paras[^1].BorderBottom = bb; }
            if (n.Tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6") foreach (var p in paras) p.KeepNext = true;
            if (s.AvoidBreakInside) foreach (var p in paras) p.KeepLines = true;
        }
        if (s.Position == "relative")
        {
            double dx = CssValues.Length(s.Left, s.FontSize, _contentWidth) ?? -(CssValues.Length(s.Right, s.FontSize, _contentWidth) ?? 0);
            double dy = CssValues.Length(s.Top, s.FontSize) ?? -(CssValues.Length(s.Bottom, s.FontSize) ?? 0);
            foreach (var blk in created)
            {
                switch (blk)
                {
                    case Paragraph p: p.OffsetX += dx; p.OffsetY += dy; break;
                    case Table t: t.Indent += dx; break;
                    case FloatBlock fb: fb.MarginLeft += fb.Side == "left" ? dx : 0; fb.MarginTop += dy; break;
                    case AnchorBlock ab: ab.X += dx; ab.TopOffset += dy; break;
                }
            }
        }
        // Конец relative-предка: по метке известна его высота (для bottom у absolute-потомков)
        if (s.Position == "relative" && created.OfType<AnchorBlock>().FirstOrDefault() is { } anc)
            f.Target.Add(new AnchorEndBlock { Anchor = anc, PadBottom = s.PaddingBottom + f.PendingMargin });
        if (s.PaddingBottom > 0) f.PendingMargin += s.PaddingBottom;
        f.PendingMargin = Math.Max(f.PendingMargin, s.MarginBottom);
        if (s.PageBreakAfter) f.PendingBreak = true;
    }

    static string ListLabel(string type, int n) => type switch
    {
        "none" => "",
        "disc" => "•",
        "circle" => "◦",
        "square" => "▪",
        "decimal" => n + ".",
        "decimal-leading-zero" => (n < 10 ? "0" : "") + n + ".",
        "lower-alpha" or "lower-latin" => Letters(n, "abcdefghijklmnopqrstuvwxyz") + ".",
        "upper-alpha" or "upper-latin" => Letters(n, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") + ".",
        "lower-roman" => Roman(n).ToLowerInvariant() + ".",
        "upper-roman" => Roman(n) + ".",
        "lower-greek" => Letters(n, "αβγδεζηθικλμνξοπρστυφχψω") + ".",
        "cyrillic-lower" => Letters(n, "абвгдежзиклмнопрстуфхцчшщэюя") + ")",
        "cyrillic-upper" => Letters(n, "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ") + ")",
        _ => "•"
    };

    static string Letters(int n, string abc)
    {
        if (n <= 0) return n.ToString(CultureInfo.InvariantCulture);
        var sb = new StringBuilder();
        while (n > 0) { n--; sb.Insert(0, abc[n % abc.Length]); n /= abc.Length; }
        return sb.ToString();
    }

    static string Roman(int n)
    {
        if (n <= 0 || n >= 4000) return n.ToString(CultureInfo.InvariantCulture);
        int[] v = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
        string[] r = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
        var sb = new StringBuilder();
        for (int i = 0; i < v.Length; i++) while (n >= v[i]) { sb.Append(r[i]); n -= v[i]; }
        return sb.ToString();
    }

    void AddHr(CssStyle s, Flow f, BlockCtx b)
    {
        Flush(f);
        f.PendingMargin = Math.Max(f.PendingMargin, s.MarginTop);
        var color = s.BorderTop?.Color ?? s.Background ?? new Rgb(0.6, 0.6, 0.6);
        double width = s.BorderTop?.Width ?? (CssValues.Length(s.Height, s.FontSize) ?? 0.75);
        var tiny = ToTextStyle(s) with { Size = 1, Underline = false, Link = null };
        var p = new Paragraph
        {
            IndLeft = b.IndL + s.MarginLeft, IndRight = b.IndR + s.MarginRight,
            SpaceBefore = f.PendingMargin, PageBreakBefore = f.PendingBreak, MarkStyle = tiny, LineRule = "exact", LineValue = 1,
            BorderBottom = new BorderDef { Width = Math.Max(0.5, width), Color = color },
        };
        f.PendingMargin = 0; f.PendingBreak = false;
        f.Target.Add(p);
        f.PendingMargin = Math.Max(0, s.MarginBottom);
    }

    // =================================================================
    //  Изображения
    // =================================================================

    string _svgCss = "";
    readonly Dictionary<string, byte[]> _inlineSvg = new(StringComparer.Ordinal);

    byte[]? LoadResource(string? src)
    {
        // Встроенный SVG: с CSS документа в начале (тот же массив для одного и того же src)
        if (src != null && src.StartsWith("data:image/svg+xml;inline;base64,", StringComparison.Ordinal))
        {
            if (_inlineSvg.TryGetValue(src, out var ready)) return ready;
            try
            {
                string raw = Encoding.UTF8.GetString(System.Convert.FromBase64String(src[(src.IndexOf(',') + 1)..]));
                if (_svgCss.Trim().Length > 0)
                {
                    int s0 = raw.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
                    int gt = s0 >= 0 ? raw.IndexOf('>', s0) : -1;
                    if (gt > 0 && raw[gt - 1] != '/')
                        raw = raw[..(gt + 1)] + "<style><![CDATA[" + _svgCss.Replace("]]>", "") + "]]></style>" + raw[(gt + 1)..];
                }
                return _inlineSvg[src] = Encoding.UTF8.GetBytes(raw);
            }
            catch { return null; }
        }
        if (string.IsNullOrWhiteSpace(src)) return null;
        src = src.Trim();
        try
        {
            // Веб-архив: сначала ищем ресурс среди частей MHTML (по cid: или адресу)
            if (_mht != null && !src.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && _mht.Find(src, _baseUrl) is { } part)
                return part.Data;
            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                int comma = src.IndexOf(',');
                if (comma < 0) return null;
                string meta = src[5..comma];
                string data = src[(comma + 1)..];
                return meta.EndsWith(";base64", StringComparison.OrdinalIgnoreCase)
                    ? System.Convert.FromBase64String(Uri.UnescapeDataString(data).Replace(" ", "").Replace("\n", "").Replace("\r", ""))
                    : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(data));
            }
            if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || src.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
            {
                if (!_opt.AllowRemote) return null;
                return Http.GetByteArrayAsync(src).GetAwaiter().GetResult();
            }
            if (src.StartsWith("//")) return _opt.AllowRemote ? Http.GetByteArrayAsync("https:" + src).GetAwaiter().GetResult() : null;
            string path = src.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(src).LocalPath : Uri.UnescapeDataString(src.Split('?', '#')[0]);
            if (!Path.IsPathRooted(path)) path = Path.Combine(_baseDir, path.Replace('/', Path.DirectorySeparatorChar));
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }
        catch { return null; }
    }

    void AddImage(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        var bytes = LoadResource(n.Attr("src"));
        var info = bytes != null ? ImageDecoder.Decode(bytes) : null;
        if (info == null)
        {
            string? alt = n.Attr("alt");
            if (!string.IsNullOrWhiteSpace(alt)) AddText("[" + alt + "]", s, f, b);
            return;
        }
        double iw = info.Width * 0.75, ih = info.Height * 0.75;
        double? w = CssValues.Length(s.Width, s.FontSize, b.Width);
        double? h = CssValues.Length(s.Height, s.FontSize, b.Width);
        if (s.Height?.EndsWith('%') == true) h = null;
        if (w != null && h == null) h = w * ih / iw;
        else if (h != null && w == null) w = h * iw / ih;
        w ??= iw; h ??= ih;
        string key = "html-image-" + (++_imageCounter);
        _doc.Media[key] = bytes!;
        bool block = s.Display == "block" || s.MarginLeftAuto;
        if (block) { Flush(f); f.PendingMargin = Math.Max(f.PendingMargin, s.MarginTop); }
        var p = EnsureParagraph(f, b);
        if (block && s.MarginLeftAuto && s.MarginRightAuto) p.Align = "center";
        AddInline(f, b, new ImageInline(key, w.Value, h.Value, ToTextStyle(s)) { Transform = TransformOf(s), Radius = RadiusOf(s), VAlign = s.VerticalAlign });
        f.LastSpace = false;
        if (block) { Flush(f); f.PendingMargin = Math.Max(f.PendingMargin, s.MarginBottom); }
    }

    // =================================================================
    //  Таблицы
    // =================================================================

    sealed record HtmlCell(Node Node, CssStyle Style, int ColSpan, int RowSpan, CssStyle RowStyle, CssStyle? GroupStyle);

    void AddTable(Node tableNode, CssStyle ts, Flow f, BlockCtx b)
    {
        Flush(f);
        f.PendingMargin = Math.Max(f.PendingMargin, ts.MarginTop);

        // Строки: thead → tbody/tr → tfoot
        f.PendingClear = null;
        var head = new List<(Node Tr, CssStyle? Group)>();
        var body = new List<(Node Tr, CssStyle? Group)>();
        var foot = new List<(Node Tr, CssStyle? Group)>();
        var colWidths = new List<string?>();
        Node? caption = null;
        // Строки и ячейки определяются по display (работает и для CSS-таблиц на div)
        static bool IsRow(Node x) => x.Style?.Display == "table-row";
        static bool IsCell(Node x) => x.Style?.Display == "table-cell";
        Node? anon = null;   // анонимная строка для ячеек без родителя-строки
        foreach (var c in tableNode.Children.Where(c => c.IsElement && c.Style?.Display != "none"))
        {
            if (IsCell(c))
            {
                if (anon == null) { anon = new Node { Tag = "tr", Style = new CssStyle { Display = "table-row" } }; body.Add((anon, null)); }
                anon.Children.Add(c);
                continue;
            }
            anon = null;
            switch (c.Style!.Display)
            {
                case "table-caption": caption = c; break;
                case "table-header-group": head.AddRange(c.Children.Where(IsRow).Select(r => (r, (CssStyle?)c.Style))); break;
                case "table-footer-group": foot.AddRange(c.Children.Where(IsRow).Select(r => (r, (CssStyle?)c.Style))); break;
                case "table-row-group": body.AddRange(c.Children.Where(IsRow).Select(r => (r, (CssStyle?)c.Style))); break;
                case "table-row": body.Add((c, null)); break;
            }
            switch (c.Tag)
            {
                case "colgroup":
                    foreach (var col in c.Children.Where(x => x.Tag == "col"))
                        for (int k = 0; k < Math.Max(1, int.TryParse(col.Attr("span"), out var sp) ? sp : 1); k++) colWidths.Add(col.Style?.Width);
                    break;
                case "col":
                    for (int k = 0; k < Math.Max(1, int.TryParse(c.Attr("span"), out var sp2) ? sp2 : 1); k++) colWidths.Add(c.Style?.Width);
                    break;
            }
        }
        bool headerFromFirstRow = head.Count == 0 && body.Count > 1 && body[0].Tr.Children.Where(x => x.IsElement).All(x => x.Tag == "th");
        var allRows = head.Concat(body).Concat(foot).Where(r => r.Tr.Style?.Display != "none").ToList();
        int headerCount = head.Count > 0 ? head.Count : headerFromFirstRow ? 1 : 0;

        if (caption != null)
        {
            var capCtx = new BlockCtx { IndL = b.IndL, IndR = b.IndR, Owner = caption.Style!, Width = b.Width, List = b.List };
            WalkChildren(caption, f, capCtx);
            Flush(f);
        }
        if (allRows.Count == 0) return;

        // Размещение ячеек в сетке с учётом colspan/rowspan
        var grid = new List<Dictionary<int, (HtmlCell Cell, bool Origin)>>();
        for (int r = 0; r < allRows.Count; r++) grid.Add([]);
        int cols = 0;
        for (int r = 0; r < allRows.Count; r++)
        {
            var (tr, group) = allRows[r];
            int c = 0;
            foreach (var td in tr.Children.Where(IsCell))
            {
                while (grid[r].ContainsKey(c)) c++;
                int cs = Math.Clamp(int.TryParse(td.Attr("colspan"), out var csv) ? csv : 1, 1, 1000);
                int rs = Math.Max(1, int.TryParse(td.Attr("rowspan"), out var rsv) ? (rsv == 0 ? allRows.Count - r : rsv) : 1);
                rs = Math.Min(rs, allRows.Count - r);
                var cell = new HtmlCell(td, td.Style!, cs, rs, tr.Style!, group);
                for (int rr = r; rr < r + rs; rr++)
                    for (int cc = c; cc < c + cs; cc++)
                        grid[rr][cc] = (cell, rr == r && cc == c);
                c += cs;
                cols = Math.Max(cols, c);
            }
        }
        if (cols == 0) return;

        // Рамки и отступы
        int borderAttr = 0;
        if (tableNode.Attr("border") is { } ba) borderAttr = int.TryParse(ba, out var bv) ? bv : 1;
        double? cellPadAttr = tableNode.Attr("cellpadding") is { } cp && double.TryParse(cp, NumberStyles.Float, CultureInfo.InvariantCulture, out var cpv) ? cpv * 0.75 : null;
        var gridColor = new Rgb(0.5, 0.5, 0.5);

        var table = new Table
        {
            Indent = b.IndL + ts.MarginLeft,
            PageBreakBefore = f.PendingBreak,
            SpaceBefore = f.PendingMargin,
            Clear = ts.Clear,
        };
        f.PendingBreak = false; f.PendingMargin = 0;
        if (borderAttr > 0)
        {
            var outer = new BorderDef { Width = borderAttr * 0.75, Color = gridColor };
            var innerB = new BorderDef { Width = 0.75, Color = gridColor };
            table.Borders = new CellBorders { Top = outer, Bottom = outer, Left = outer, Right = outer, InsideH = innerB, InsideV = innerB };
        }
        else table.Borders = new CellBorders();
        table.Borders.Top = ts.BorderTop?.ToBorderDef() ?? table.Borders.Top;
        table.Borders.Bottom = ts.BorderBottom?.ToBorderDef() ?? table.Borders.Bottom;
        table.Borders.Left = ts.BorderLeft?.ToBorderDef() ?? table.Borders.Left;
        table.Borders.Right = ts.BorderRight?.ToBorderDef() ?? table.Borders.Right;
        table.CellMargins = new Margins { Top = 0.75, Bottom = 0.75, Left = 1.5, Right = 1.5 };
        if (cellPadAttr != null) table.CellMargins = new Margins { Top = cellPadAttr, Bottom = cellPadAttr, Left = cellPadAttr, Right = cellPadAttr };

        // Ширины колонок
        double avail = Math.Max(30, _contentWidth - b.IndL - b.IndR - ts.MarginLeft - ts.MarginRight);
        double? tableWidth = CssValues.Length(ts.Width, ts.FontSize, avail);
        if (tableWidth != null) tableWidth = Math.Min(tableWidth.Value, avail);
        var minW = new double[cols];
        var prefW = new double[cols];
        var fixedW = new double?[cols];
        for (int c = 0; c < cols && c < colWidths.Count; c++)
            fixedW[c] = CssValues.Length(colWidths[c], ts.FontSize, tableWidth ?? avail);
        var spanned = new List<(int C, int Span, double Min, double Pref)>();
        for (int r = 0; r < allRows.Count; r++)
        {
            foreach (var (c, (cell, origin)) in grid[r])
            {
                if (!origin) continue;
                double padH = HPad(cell.Style, cellPadAttr) + 1;
                var (mn, pf) = Measure(cell.Node, avail);
                mn += padH; pf += padH;
                double? fw = CssValues.Length(cell.Style.Width, cell.Style.FontSize, tableWidth ?? avail);
                if (fw != null) { fw += cell.Style.Width!.EndsWith('%') ? 0 : padH; pf = Math.Max(mn, fw.Value); }
                if (cell.ColSpan == 1)
                {
                    minW[c] = Math.Max(minW[c], mn);
                    prefW[c] = Math.Max(prefW[c], pf);
                    if (fw != null) fixedW[c] = Math.Max(fixedW[c] ?? 0, fw.Value);
                }
                else spanned.Add((c, cell.ColSpan, mn, pf));
            }
        }
        foreach (var (c, span, mn, pf) in spanned)
        {
            double curMin = 0, curPref = 0;
            for (int k = c; k < c + span && k < cols; k++) { curMin += minW[k]; curPref += prefW[k]; }
            for (int k = c; k < c + span && k < cols; k++)
            {
                if (mn > curMin) minW[k] += (mn - curMin) / span;
                if (pf > curPref) prefW[k] += (pf - curPref) / span;
            }
        }
        for (int c = 0; c < cols; c++)
        {
            if (fixedW[c] is double fw) { prefW[c] = Math.Max(minW[c], fw); }
            prefW[c] = Math.Max(prefW[c], Math.Max(minW[c], 6));
        }
        double sumMin = minW.Sum(), sumPref = prefW.Sum();
        double target = tableWidth ?? Math.Min(avail, sumPref);
        var widths = new double[cols];
        if (sumPref <= target)
        {
            // Лишнее место раздаём колонкам без фиксированной ширины
            double extra = target - sumPref;
            var flex = Enumerable.Range(0, cols).Where(c => fixedW[c] == null).ToList();
            if (flex.Count == 0) flex = Enumerable.Range(0, cols).ToList();
            double flexPref = flex.Sum(c => prefW[c]);
            for (int c = 0; c < cols; c++)
                widths[c] = prefW[c] + (flex.Contains(c) ? extra * (flexPref > 0 ? prefW[c] / flexPref : 1.0 / flex.Count) : 0);
        }
        else if (sumMin < target)
        {
            double room = target - sumMin, want = sumPref - sumMin;
            for (int c = 0; c < cols; c++) widths[c] = minW[c] + room * (want > 0 ? (prefW[c] - minW[c]) / want : 1.0 / cols);
        }
        else
        {
            for (int c = 0; c < cols; c++) widths[c] = minW[c] * target / sumMin;
        }
        table.Grid = widths.ToList();
        double total = widths.Sum();
        if (ts.MarginLeftAuto && ts.MarginRightAuto) table.Indent = b.IndL + Math.Max(0, (avail - total) / 2);
        else if (ts.MarginLeftAuto) table.Indent = b.IndL + Math.Max(0, avail - total);

        // Строки модели
        for (int r = 0; r < allRows.Count; r++)
        {
            var (tr, group) = allRows[r];
            var row = new TableRow { IsHeader = r < headerCount };
            double? rh = CssValues.Length(tr.Style?.Height, tr.Style?.FontSize ?? 12);
            if (rh != null) row.Height = rh.Value;
            int c = 0;
            while (c < cols)
            {
                if (!grid[r].TryGetValue(c, out var entry))
                {
                    row.Cells.Add(new TableCell { GridSpan = 1, Borders = new CellBorders() });
                    c++;
                    continue;
                }
                var (cell, origin) = entry;
                int span = Math.Min(cell.ColSpan, cols - c);
                bool topRow = grid[r].TryGetValue(c, out var e2) && (r == 0 || !grid[r - 1].TryGetValue(c, out var up) || up.Cell != cell);
                var tc = new TableCell
                {
                    GridSpan = span,
                    VMerge = cell.RowSpan > 1 ? (topRow ? VMerge.Restart : VMerge.Continue) : VMerge.None,
                    Shading = cell.Style.Background ?? cell.RowStyle.Background ?? cell.GroupStyle?.Background,
                    VAlign = (cell.Style.VerticalAlign ?? cell.RowStyle.VerticalAlign ?? "middle") switch { "top" or "baseline" => "top", "bottom" => "bottom", _ => "center" },
                    Margins = CellMargins(cell.Style, cellPadAttr),
                };
                var cs = cell.Style;
                if (cs.BorderTop != null || cs.BorderBottom != null || cs.BorderLeft != null || cs.BorderRight != null)
                {
                    tc.Borders = new CellBorders
                    {
                        Top = cs.BorderTop != null ? cs.BorderTop.ToBorderDef() ?? new BorderDef { Val = "nil" } : null,
                        Bottom = cs.BorderBottom != null ? cs.BorderBottom.ToBorderDef() ?? new BorderDef { Val = "nil" } : null,
                        Left = cs.BorderLeft != null ? cs.BorderLeft.ToBorderDef() ?? new BorderDef { Val = "nil" } : null,
                        Right = cs.BorderRight != null ? cs.BorderRight.ToBorderDef() ?? new BorderDef { Val = "nil" } : null,
                    };
                }
                if (cs.Height != null && CssValues.Length(cs.Height, cs.FontSize) is double ch && origin) row.Height = Math.Max(row.Height, ch);
                if (origin)
                {
                    var cellFlow = new Flow(tc.Blocks);
                    double cw = widths.Skip(c).Take(span).Sum() - (tc.Margins.Left ?? 0) - (tc.Margins.Right ?? 0);
                    var cellCtx = new BlockCtx { IndL = 0, IndR = 0, Owner = cs, Width = Math.Max(10, cw) };
                    double savedWidth = _contentWidth;
                    _contentWidth = cellCtx.Width;
                    WalkChildren(cell.Node, cellFlow, cellCtx);
                    Flush(cellFlow);
                    // Ячейка — отдельный контекст форматирования: нижнее поле последнего блока остаётся внутри неё
                    if (cellFlow.PendingMargin > 0.01)
                    {
                        if (tc.Blocks.Count > 0 && tc.Blocks[^1] is Paragraph lastP) lastP.SpaceAfter = Math.Max(lastP.SpaceAfter, cellFlow.PendingMargin);
                        else if (tc.Blocks.Count > 0)
                            tc.Blocks.Add(new Paragraph { SpaceBefore = Math.Max(0, cellFlow.PendingMargin - 1), MarkStyle = ToTextStyle(cs) with { Size = 1 }, LineRule = "exact", LineValue = 1 });
                    }
                    ApplyFirstLine(cell.Node, cs, tc.Blocks);
                    _contentWidth = savedWidth;
                }
                row.Cells.Add(tc);
                c += span;
            }
            table.Rows.Add(row);
        }
        f.Target.Add(table);
        f.PendingMargin = Math.Max(f.PendingMargin, ts.MarginBottom);
        if (ts.PageBreakAfter) f.PendingBreak = true;
    }

    static double HPad(CssStyle s, double? cellPad) =>
        (cellPad ?? Math.Max(1.5, s.PaddingLeft)) + (cellPad ?? Math.Max(1.5, s.PaddingRight));

    static Margins CellMargins(CssStyle s, double? cellPad) => new()
    {
        Top = cellPad ?? Math.Max(0.75, s.PaddingTop),
        Bottom = cellPad ?? Math.Max(0.75, s.PaddingBottom),
        Left = cellPad ?? Math.Max(1.5, s.PaddingLeft),
        Right = cellPad ?? Math.Max(1.5, s.PaddingRight),
    };

    /// <summary>Минимальная (самое длинное слово) и предпочтительная (без переносов) ширина содержимого.</summary>
    (double Min, double Pref) Measure(Node n, double avail)
    {
        double min = 0, pref = 0, line = 0;
        void Walk(Node x)
        {
            if (x.IsText)
            {
                var st = x.Parent?.Style;
                if (st == null || st.Hidden) return;
                bool nowrap = st.WhiteSpace is "nowrap" or "pre";
                var face = _fonts.Get(st.FontFamily, st.Bold, st.Italic, st.Weight);
                double size = st.VertAlign != 0 ? st.FontSize * 0.65 : st.FontSize;
                double spaceW = Width(face, " ", size);
                string text = st.Upper ? x.Text.ToUpperInvariant() : x.Text;
                if (nowrap)
                {
                    foreach (var seg in text.Split('\n'))
                    {
                        double w = Width(face, seg.Replace('\t', ' '), size);
                        line += w;
                        min = Math.Max(min, line);
                        if (st.WhiteSpace == "pre") { pref = Math.Max(pref, line); line = 0; }
                    }
                    return;
                }
                // Пробелы — только между словами (и по краям, если они есть в тексте)
                var words = text.Split([' ', '\n', '\t', '\r'], StringSplitOptions.RemoveEmptyEntries);
                if (words.Length == 0) { if (text.Length > 0 && line > 0) line += spaceW + st.WordSpacing; return; }
                if (char.IsWhiteSpace(text[0]) && line > 0) line += spaceW + st.WordSpacing;
                for (int wi = 0; wi < words.Length; wi++)
                {
                    double w = Width(face, words[wi], size) + st.LetterSpacing * words[wi].Length;
                    min = Math.Max(min, w);
                    line += w + (wi < words.Length - 1 ? spaceW + st.WordSpacing : 0);
                }
                if (char.IsWhiteSpace(text[^1])) line += spaceW + st.WordSpacing;
                return;
            }
            if (!x.IsElement) { foreach (var c in x.Children) Walk(c); return; }
            var s = x.Style!;
            if (s.Display == "none") return;
            if (x.Tag == "br") { pref = Math.Max(pref, line); line = 0; return; }
            if (x.Tag == "img")
            {
                double w = CssValues.Length(s.Width, s.FontSize, avail) ?? 0;
                if (w <= 0)
                {
                    var bytes = LoadResource(x.Attr("src"));
                    var info = bytes != null ? ImageDecoder.Decode(bytes) : null;
                    if (info != null)
                    {
                        w = info.Width * 0.75;
                        if (CssValues.Length(s.Height, s.FontSize) is double h && info.Height > 0) w = h * info.Width / info.Height;
                    }
                }
                min = Math.Max(min, w);
                line += w;
                return;
            }
            bool block = s.Display != "inline";
            if (block) { pref = Math.Max(pref, line); line = 0; }
            else if (s.InlineBox is { } ib && ib != x.Parent?.Style?.InlineBox)
            {
                line += ib.StartWidth + ib.EndWidth;
                min = Math.Max(min, ib.StartWidth + ib.EndWidth);
            }
            double extra = block ? s.MarginLeft + s.MarginRight + s.PaddingLeft + s.PaddingRight : 0;
            if (block && x.Tag is "ul" or "ol") extra = Math.Max(extra, 30);
            double savedMin = min, savedPref = pref;
            min = 0; pref = 0;
            foreach (var c in x.Children) Walk(c);
            if (block) { pref = Math.Max(pref, line); line = 0; }
            min = Math.Max(savedMin, min + extra);
            pref = Math.Max(savedPref, pref + extra);
        }
        foreach (var c in n.Children) Walk(c);
        pref = Math.Max(pref, line);
        return (min, Math.Min(Math.Max(pref, min), Math.Max(avail, min)));
    }

    static double Width(FontFace face, string s, double size)
    {
        double w = 0;
        foreach (var r in s.EnumerateRunes())
            w += face.Font.TryGetGlyph(r.Value, out var g) ? face.Font.Advance(g) * size : size * 0.55;
        return w;
    }
}

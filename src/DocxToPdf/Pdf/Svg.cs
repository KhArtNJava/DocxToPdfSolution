using System.Globalization;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;

namespace DocxToPdf.Pdf;

/// <summary>
/// Векторный вывод SVG 1.1/2 в PDF: фигуры и контуры (вся грамматика path, дуги), группы, use/symbol, вложенные svg,
/// преобразования, viewBox и preserveAspectRatio, заливка и обводка (пунктир, концы, стыки), прозрачность,
/// градиенты (линейный, радиальный, objectBoundingBox, gradientTransform, наследование href), узоры (pattern),
/// обрезка (clipPath), маркеры, текст (text/tspan/textPath, text-anchor, базовая линия, межбуквенный интервал),
/// встроенные картинки, стили (атрибуты, &lt;style&gt; с селекторами, style="").
/// </summary>
public static class Svg
{
    static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    // ================================================================
    //  Определение формата и размер
    // ================================================================

    public static bool IsSvg(byte[] d)
    {
        if (d.Length < 4) return false;
        if (d[0] == 0x1F && d[1] == 0x8B) return Decompress(d) is { } u && IsSvgText(u);
        return IsSvgText(d);
    }

    static bool IsSvgText(byte[] d)
    {
        int n = Math.Min(d.Length, 4096);
        string head = Encoding.UTF8.GetString(d, 0, n);
        int i = head.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
        if (i < 0) return false;
        // Перед <svg> — только пролог XML, DOCTYPE, комментарии и пробелы
        string pre = Regex.Replace(head[..i], @"<\?.*?\?>|<!--.*?-->|<!DOCTYPE[^>\[]*(\[.*?\])?\s*>", "", RegexOptions.Singleline);
        return pre.Trim('﻿', ' ', '\t', '\r', '\n').Length == 0;
    }

    static byte[]? Decompress(byte[] d)
    {
        try
        {
            using var gz = new GZipStream(new MemoryStream(d), CompressionMode.Decompress);
            var ms = new MemoryStream();
            gz.CopyTo(ms);
            return ms.ToArray();
        }
        catch { return null; }
    }

    /// <summary>Собственный размер рисунка, pt (width/height; иначе viewBox; иначе 300×150 px).</summary>
    public static (double W, double H)? Size(byte[] d)
    {
        var doc = Load(d);
        if (doc == null) return null;
        var r = doc.Root;
        var vb = ViewBox(r);
        double? w = AbsLength(Attr(r, "width")), h = AbsLength(Attr(r, "height"));
        if (w == null && h == null && vb != null) { w = vb.Value.W; h = vb.Value.H; }
        else if (w == null && vb != null && h != null) w = h * vb.Value.W / vb.Value.H;
        else if (h == null && vb != null && w != null) h = w * vb.Value.H / vb.Value.W;
        return ((w ?? 300) * 0.75, (h ?? 150) * 0.75);
    }

    /// <summary>Длина без процентов в px (единица пользователя SVG).</summary>
    static double? AbsLength(string? v)
    {
        if (string.IsNullOrWhiteSpace(v) || v.Trim().EndsWith('%')) return null;
        var l = Len(v, 0, 16);
        return l > 0 ? l : null;
    }

    // ================================================================
    //  Документ: разбор, идентификаторы, стили
    // ================================================================

    sealed class Rule
    {
        public required List<(string Tag, string? Id, List<string> Classes)> Chain;   // по предкам: ... > последний
        public required int Spec, Order;
        public required List<(string Name, string Value)> Decls;
    }

    sealed class SvgDocument
    {
        public required XElement Root;
        public readonly Dictionary<string, XElement> Ids = new(StringComparer.Ordinal);
        public readonly List<Rule> Rules = [];
        public readonly ConditionalWeakTable<XElement, Dictionary<string, string>> Specified = new();
    }

    static readonly ConditionalWeakTable<byte[], SvgDocument?> Cache = new();

    static SvgDocument? Load(byte[] data)
    {
        if (Cache.TryGetValue(data, out var cached)) return cached;
        var doc = Parse(data);
        Cache.AddOrUpdate(data, doc);
        return doc;
    }

    static SvgDocument? Parse(byte[] data)
    {
        if (data.Length > 1 && data[0] == 0x1F && data[1] == 0x8B) data = Decompress(data) ?? data;
        XDocument? x = null;
        try { x = XDocument.Load(new MemoryStream(data), LoadOptions.PreserveWhitespace); }
        catch
        {
            // Разметка из HTML: именованные сущности, необъявленные префиксы
            try
            {
                string t = Encoding.UTF8.GetString(data).TrimStart('﻿');
                t = Regex.Replace(t, @"&([a-zA-Z][a-zA-Z0-9]*);", m => m.Groups[1].Value is "amp" or "lt" or "gt" or "quot" or "apos" ? m.Value
                    : Html.Entities.Decode(m.Value) is { } dv && dv != m.Value ? "&#" + char.ConvertToUtf32(dv, 0) + ";" : "");
                t = Regex.Replace(t, @"&(?![a-zA-Z#][a-zA-Z0-9]*;)", "&amp;");
                int s = t.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
                if (s >= 0)
                {
                    int e = t.IndexOf('>', s);
                    string open = t[s..e];
                    var add = new StringBuilder();
                    if (!open.Contains("xmlns:xlink") && t.Contains("xlink:")) add.Append(" xmlns:xlink=\"http://www.w3.org/1999/xlink\"");
                    foreach (Match m in Regex.Matches(t, @"\s([a-zA-Z]+):[a-zA-Z-]+="))
                    {
                        string px = m.Groups[1].Value;
                        if (px is "xlink" or "xml" or "xmlns") continue;
                        if (!open.Contains("xmlns:" + px) && !add.ToString().Contains("xmlns:" + px)) add.Append($" xmlns:{px}=\"urn:x-{px}\"");
                    }
                    t = t[..(s + 4)] + add + t[(s + 4)..];
                }
                x = XDocument.Parse(t, LoadOptions.PreserveWhitespace);
            }
            catch { return null; }
        }
        var root = x.Root;
        if (root == null || root.Name.LocalName != "svg") return null;
        var doc = new SvgDocument { Root = root };
        foreach (var e in root.DescendantsAndSelf())
            if ((string?)e.Attribute("id") is { Length: > 0 } id) doc.Ids.TryAdd(id, e);
        int order = 0;
        foreach (var st in root.Descendants().Where(e => e.Name.LocalName == "style"))
            ParseCss(st.Value, doc.Rules, ref order);
        return doc;
    }

    static string? Attr(XElement e, string name)
    {
        var a = e.Attribute(name);
        if (a != null) return a.Value;
        foreach (var at in e.Attributes()) if (at.Name.LocalName == name && !at.IsNamespaceDeclaration) return at.Value;
        return null;
    }

    static string? Href(XElement e) => Attr(e, "href");

    static XElement? Ref(SvgDocument doc, string? iri)
    {
        if (string.IsNullOrWhiteSpace(iri)) return null;
        iri = iri.Trim();
        if (iri.StartsWith("url(", StringComparison.OrdinalIgnoreCase)) iri = iri[4..].TrimEnd(')').Trim(' ', '"', '\'');
        int h = iri.IndexOf('#');
        if (h < 0) return null;
        return doc.Ids.GetValueOrDefault(iri[(h + 1)..]);
    }

    // ---------------- CSS ----------------

    static void ParseCss(string css, List<Rule> rules, ref int order)
    {
        css = Regex.Replace(css, @"/\*.*?\*/", "", RegexOptions.Singleline).Replace("<![CDATA[", "").Replace("]]>", "");
        int i = 0;
        while (i < css.Length)
        {
            int ob = css.IndexOf('{', i);
            if (ob < 0) break;
            string sel = css[i..ob].Trim();
            // @-правила: блок пропускаем (с учётом вложенности), @media — разбираем содержимое
            int depth = 1, j = ob + 1;
            while (j < css.Length && depth > 0) { if (css[j] == '{') depth++; else if (css[j] == '}') depth--; j++; }
            string body = css[(ob + 1)..Math.Max(ob + 1, j - 1)];
            i = j;
            if (sel.StartsWith('@'))
            {
                if (sel.StartsWith("@media", StringComparison.OrdinalIgnoreCase) && !sel.Contains("screen") || sel.StartsWith("@supports")) ParseCss(body, rules, ref order);
                continue;
            }
            var decls = new List<(string, string)>();
            foreach (var d in body.Split(';'))
            {
                int c = d.IndexOf(':');
                if (c <= 0) continue;
                decls.Add((d[..c].Trim().ToLowerInvariant(), d[(c + 1)..].Replace("!important", "").Trim()));
            }
            foreach (var one in sel.Split(','))
            {
                var parts = Regex.Split(one.Trim().Replace(">", " ").Replace("+", " ").Replace("~", " "), @"\s+").Where(p => p.Length > 0).ToList();
                if (parts.Count == 0) continue;
                var chain = new List<(string, string?, List<string>)>();
                int spec = 0;
                bool ok = true;
                foreach (var p in parts)
                {
                    var m = Regex.Match(p, @"^([a-zA-Z*][\w-]*)?((?:[#.][\w-]+|\[[^\]]*\]|:[\w-]+(?:\([^)]*\))?)*)$");
                    if (!m.Success) { ok = false; break; }
                    string tag = m.Groups[1].Value is { Length: > 0 } t && t != "*" ? t : "*";
                    string? id = null; var cls = new List<string>();
                    foreach (Match q in Regex.Matches(m.Groups[2].Value, @"([#.])([\w-]+)|(:[\w-]+)"))
                    {
                        if (q.Groups[1].Value == "#") { id = q.Groups[2].Value; spec += 10000; }
                        else if (q.Groups[1].Value == ".") { cls.Add(q.Groups[2].Value); spec += 100; }
                        else if (q.Groups[3].Success && q.Groups[3].Value is not (":root" or ":first-child")) ok = false;   // динамические псевдоклассы не действуют
                    }
                    if (tag != "*") spec += 1;
                    chain.Add((tag, id, cls));
                }
                if (ok) rules.Add(new Rule { Chain = chain, Spec = spec, Order = order++, Decls = decls });
            }
        }
    }

    static bool MatchCompound(XElement e, (string Tag, string? Id, List<string> Classes) c)
    {
        if (c.Tag != "*" && !e.Name.LocalName.Equals(c.Tag, StringComparison.OrdinalIgnoreCase)) return false;
        if (c.Id != null && (string?)e.Attribute("id") != c.Id) return false;
        if (c.Classes.Count > 0)
        {
            var have = (Attr(e, "class") ?? "").Split([' ', '\t', '\n', '\r'], StringSplitOptions.RemoveEmptyEntries);
            foreach (var cl in c.Classes) if (!have.Contains(cl)) return false;
        }
        return true;
    }

    static bool Matches(XElement e, Rule r)
    {
        if (!MatchCompound(e, r.Chain[^1])) return false;
        var anc = e.Parent;
        for (int k = r.Chain.Count - 2; k >= 0; k--)
        {
            while (anc != null && !MatchCompound(anc, r.Chain[k])) anc = anc.Parent;
            if (anc == null) return false;
            anc = anc.Parent;
        }
        return true;
    }

    static readonly HashSet<string> Props =
    [
        "fill", "fill-opacity", "fill-rule", "stroke", "stroke-width", "stroke-opacity", "stroke-linecap", "stroke-linejoin",
        "stroke-miterlimit", "stroke-dasharray", "stroke-dashoffset", "opacity", "display", "visibility", "font-family", "font-size",
        "font-weight", "font-style", "text-anchor", "dominant-baseline", "alignment-baseline", "baseline-shift", "color", "stop-color",
        "stop-opacity", "clip-path", "clip-rule", "marker-start", "marker-mid", "marker-end", "marker", "letter-spacing", "word-spacing",
        "text-decoration", "overflow", "font", "transform", "vector-effect", "paint-order", "font-variant", "white-space", "mask",
    ];

    static readonly HashSet<string> Inherited =
    [
        "fill", "fill-opacity", "fill-rule", "stroke", "stroke-width", "stroke-opacity", "stroke-linecap", "stroke-linejoin",
        "stroke-miterlimit", "stroke-dasharray", "stroke-dashoffset", "visibility", "font-family", "font-size", "font-weight",
        "font-style", "text-anchor", "color", "clip-rule", "marker-start", "marker-mid", "marker-end", "letter-spacing",
        "word-spacing", "dominant-baseline", "font-variant", "white-space", "paint-order",
    ];

    /// <summary>Заданные свойства элемента: атрибуты оформления → правила &lt;style&gt; → style="".</summary>
    static Dictionary<string, string> SpecifiedOf(SvgDocument doc, XElement e)
    {
        if (doc.Specified.TryGetValue(e, out var s)) return s;
        s = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var a in e.Attributes())
        {
            string n = a.Name.LocalName;
            if (a.IsNamespaceDeclaration || n == "transform" || !Props.Contains(n)) continue;
            s[n] = a.Value.Trim();
        }
        foreach (var r in doc.Rules.Where(r => Matches(e, r)).OrderBy(r => r.Spec).ThenBy(r => r.Order))
            foreach (var (n, v) in r.Decls) Set(s, n, v);
        if (Attr(e, "style") is { } st)
            foreach (var d in st.Split(';'))
            {
                int c = d.IndexOf(':');
                if (c > 0) Set(s, d[..c].Trim().ToLowerInvariant(), d[(c + 1)..].Replace("!important", "").Trim());
            }
        doc.Specified.AddOrUpdate(e, s);
        return s;

        static void Set(Dictionary<string, string> s, string n, string v)
        {
            if (n == "marker") { s["marker-start"] = s["marker-mid"] = s["marker-end"] = v; return; }
            if (n == "font")
            {
                // font: [style] [weight] size[/line-height] family
                var m = Regex.Match(v, @"^\s*((?:italic|oblique|normal|bold|bolder|lighter|\d{3}|small-caps)\s+)*([\d.]+(?:px|pt|em|%)?)(?:/\S+)?\s+(.+)$");
                if (m.Success)
                {
                    foreach (Capture cp in m.Groups[1].Captures)
                    {
                        string t = cp.Value.Trim();
                        if (t is "italic" or "oblique") s["font-style"] = t;
                        else if (t is "bold" or "bolder" || t.Length == 3 && char.IsDigit(t[0])) s["font-weight"] = t;
                    }
                    s["font-size"] = m.Groups[2].Value; s["font-family"] = m.Groups[3].Value;
                }
                return;
            }
            if (Props.Contains(n)) s[n] = v;
        }
    }

    /// <summary>Вычисленный стиль: наследуемые свойства родителя + заданные элементом.</summary>
    sealed class Style
    {
        public readonly Dictionary<string, string> P = new(StringComparer.Ordinal);
        public string Get(string n, string def = "") => P.TryGetValue(n, out var v) ? v : def;
    }

    static Style Compute(SvgDocument doc, XElement e, Style parent)
    {
        var s = new Style();
        foreach (var (k, v) in parent.P) if (Inherited.Contains(k)) s.P[k] = v;
        foreach (var (k, v) in SpecifiedOf(doc, e))
        {
            if (v == "inherit") { if (parent.P.TryGetValue(k, out var pv)) s.P[k] = pv; continue; }
            if (k == "font-size" && parent.P.TryGetValue("font-size", out var pfs))
            {
                double baseSize = Len(pfs, 0, 16);
                if (v.EndsWith('%') || v.EndsWith("em"))
                {
                    s.P[k] = (Len(v, baseSize, baseSize)).ToString(Inv);
                    continue;
                }
                if (v is "larger") { s.P[k] = (baseSize * 1.2).ToString(Inv); continue; }
                if (v is "smaller") { s.P[k] = (baseSize / 1.2).ToString(Inv); continue; }
            }
            s.P[k] = v;
        }
        return s;
    }

    // ================================================================
    //  Значения
    // ================================================================

    static readonly Dictionary<string, double> FontSizes = new() { ["xx-small"] = 9, ["x-small"] = 10, ["small"] = 13, ["medium"] = 16, ["large"] = 18, ["x-large"] = 24, ["xx-large"] = 32 };

    /// <summary>Длина в единицах пользователя (px); % — от reference, em — от fontSize.</summary>
    static double Len(string? v, double reference, double fontSize)
    {
        if (string.IsNullOrWhiteSpace(v)) return 0;
        v = v.Trim();
        if (FontSizes.TryGetValue(v, out var fsv)) return fsv;
        var m = Regex.Match(v, @"^([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)\s*([a-zA-Z%]*)");
        if (!m.Success) return 0;
        double n = double.Parse(m.Groups[1].Value, Inv);
        return m.Groups[2].Value.ToLowerInvariant() switch
        {
            "" or "px" => n,
            "%" => n * reference / 100,
            "em" => n * fontSize,
            "ex" => n * fontSize / 2,
            "rem" => n * 16,
            "pt" => n * 4 / 3,
            "pc" => n * 16,
            "mm" => n * 96 / 25.4,
            "cm" => n * 96 / 2.54,
            "in" => n * 96,
            "q" => n * 96 / 101.6,
            _ => n,
        };
    }

    static double Num(string? v, double def = 0) =>
        v != null && double.TryParse(v.Trim().TrimEnd('%'), NumberStyles.Float, Inv, out var d) ? (v.Trim().EndsWith('%') ? d / 100 : d) : def;

    static List<double> Numbers(string? v)
    {
        var res = new List<double>();
        if (v == null) return res;
        foreach (Match m in Regex.Matches(v, @"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")) res.Add(double.Parse(m.Value, Inv));
        return res;
    }

    static (double X, double Y, double W, double H)? ViewBox(XElement e)
    {
        var n = Numbers(Attr(e, "viewBox"));
        return n.Count == 4 && n[2] > 0 && n[3] > 0 ? (n[0], n[1], n[2], n[3]) : null;
    }

    static readonly Dictionary<string, string> Named = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "ffffff", ["red"] = "ff0000", ["green"] = "008000", ["blue"] = "0000ff", ["yellow"] = "ffff00",
        ["cyan"] = "00ffff", ["aqua"] = "00ffff", ["magenta"] = "ff00ff", ["fuchsia"] = "ff00ff", ["gray"] = "808080", ["grey"] = "808080",
        ["silver"] = "c0c0c0", ["maroon"] = "800000", ["olive"] = "808000", ["lime"] = "00ff00", ["teal"] = "008080", ["navy"] = "000080",
        ["purple"] = "800080", ["orange"] = "ffa500", ["pink"] = "ffc0cb", ["brown"] = "a52a2a", ["gold"] = "ffd700", ["coral"] = "ff7f50",
        ["crimson"] = "dc143c", ["darkblue"] = "00008b", ["darkgreen"] = "006400", ["darkred"] = "8b0000", ["darkgray"] = "a9a9a9",
        ["darkgrey"] = "a9a9a9", ["lightgray"] = "d3d3d3", ["lightgrey"] = "d3d3d3", ["lightblue"] = "add8e6", ["lightgreen"] = "90ee90",
        ["steelblue"] = "4682b4", ["skyblue"] = "87ceeb", ["tomato"] = "ff6347", ["salmon"] = "fa8072", ["indigo"] = "4b0082",
        ["violet"] = "ee82ee", ["orchid"] = "da70d6", ["khaki"] = "f0e68c", ["beige"] = "f5f5dc", ["ivory"] = "fffff0", ["tan"] = "d2b48c",
        ["chocolate"] = "d2691e", ["firebrick"] = "b22222", ["forestgreen"] = "228b22", ["seagreen"] = "2e8b57", ["royalblue"] = "4169e1",
        ["dodgerblue"] = "1e90ff", ["deepskyblue"] = "00bfff", ["slategray"] = "708090", ["slategrey"] = "708090", ["dimgray"] = "696969",
        ["dimgrey"] = "696969", ["gainsboro"] = "dcdcdc", ["whitesmoke"] = "f5f5f5", ["lavender"] = "e6e6fa", ["plum"] = "dda0dd",
        ["turquoise"] = "40e0d0", ["darkorange"] = "ff8c00", ["orangered"] = "ff4500", ["goldenrod"] = "daa520", ["limegreen"] = "32cd32",
        ["yellowgreen"] = "9acd32", ["olivedrab"] = "6b8e23", ["darkviolet"] = "9400d3", ["darkcyan"] = "008b8b", ["cornflowerblue"] = "6495ed",
        ["midnightblue"] = "191970", ["mediumblue"] = "0000cd", ["hotpink"] = "ff69b4", ["deeppink"] = "ff1493", ["sienna"] = "a0522d",
        ["peru"] = "cd853f", ["wheat"] = "f5deb3", ["linen"] = "faf0e6", ["mintcream"] = "f5fffa", ["aliceblue"] = "f0f8ff",
        ["lightyellow"] = "ffffe0", ["lightpink"] = "ffb6c1", ["lightcoral"] = "f08080", ["darkslategray"] = "2f4f4f", ["darkslategrey"] = "2f4f4f",
        ["lightsteelblue"] = "b0c4de", ["powderblue"] = "b0e0e6", ["cadetblue"] = "5f9ea0", ["mediumseagreen"] = "3cb371", ["darkkhaki"] = "bdb76b",
        ["rebeccapurple"] = "663399", ["mediumpurple"] = "9370db", ["slateblue"] = "6a5acd", ["darkslateblue"] = "483d8b", ["lightseagreen"] = "20b2aa",
        ["mediumvioletred"] = "c71585", ["palevioletred"] = "db7093", ["indianred"] = "cd5c5c", ["rosybrown"] = "bc8f8f", ["sandybrown"] = "f4a460",
        ["burlywood"] = "deb887", ["navajowhite"] = "ffdead", ["moccasin"] = "ffe4b5", ["papayawhip"] = "ffefd5", ["lemonchiffon"] = "fffacd",
        ["honeydew"] = "f0fff0", ["azure"] = "f0ffff", ["snow"] = "fffafa", ["seashell"] = "fff5ee", ["oldlace"] = "fdf5e6", ["cornsilk"] = "fff8dc",
        ["springgreen"] = "00ff7f", ["chartreuse"] = "7fff00", ["lawngreen"] = "7cfc00", ["greenyellow"] = "adff2f", ["aquamarine"] = "7fffd4",
        ["mediumaquamarine"] = "66cdaa", ["darkturquoise"] = "00ced1", ["mediumturquoise"] = "48d1cc", ["paleturquoise"] = "afeeee",
        ["lightcyan"] = "e0ffff", ["darkmagenta"] = "8b008b", ["darkorchid"] = "9932cc", ["mediumorchid"] = "ba55d3", ["thistle"] = "d8bfd8",
        ["blueviolet"] = "8a2be2", ["darkgoldenrod"] = "b8860b", ["palegoldenrod"] = "eee8aa", ["lightgoldenrodyellow"] = "fafad2",
        ["saddlebrown"] = "8b4513", ["darkseagreen"] = "8fbc8f", ["palegreen"] = "98fb98", ["darkolivegreen"] = "556b2f", ["lightsalmon"] = "ffa07a",
        ["darksalmon"] = "e9967a", ["mistyrose"] = "ffe4e1", ["lavenderblush"] = "fff0f5", ["antiquewhite"] = "faebd7", ["bisque"] = "ffe4c4",
        ["blanchedalmond"] = "ffebcd", ["peachpuff"] = "ffdab9", ["floralwhite"] = "fffaf0", ["ghostwhite"] = "f8f8ff", ["lightslategray"] = "778899",
        ["lightslategrey"] = "778899", ["lightskyblue"] = "87cefa", ["mediumslateblue"] = "7b68ee", ["mediumspringgreen"] = "00fa9a",
        ["darkslategrey2"] = "2f4f4f",
    };

    /// <summary>Цвет и его непрозрачность; null — нет цвета (none/transparent).</summary>
    static (Rgb Color, double Alpha)? ParseColor(string? v, string? currentColor = null, int depth = 0)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        v = v.Trim();
        string lower = v.ToLowerInvariant();
        if (lower is "none" or "transparent") return null;
        if (lower == "currentcolor") return depth < 2 ? ParseColor(currentColor ?? "black", null, depth + 1) : (Rgb.Black, 1);
        if (v.StartsWith('#'))
        {
            string h = v[1..];
            double a = 1;
            if (h.Length is 3 or 4) h = string.Concat(h.Select(c => new string(c, 2)));
            if (h.Length == 8) { a = Convert.ToInt32(h[6..], 16) / 255.0; h = h[..6]; }
            return Rgb.Parse(h) is Rgb c ? (c, a) : null;
        }
        if (lower.StartsWith("rgb") || lower.StartsWith("hsl"))
        {
            int o = lower.IndexOf('('), cl = lower.LastIndexOf(')');
            if (o < 0 || cl < o) return null;
            var parts = lower[(o + 1)..cl].Replace("/", " ").Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3) return null;
            double A(int i) => parts.Length > i ? (parts[i].EndsWith('%') ? Num(parts[i]) : Num(parts[i], 1)) : 1;
            if (lower.StartsWith("rgb"))
            {
                double C(string s) => s.EndsWith('%') ? Num(s) : Num(s) / 255;
                return (new Rgb(Math.Clamp(C(parts[0]), 0, 1), Math.Clamp(C(parts[1]), 0, 1), Math.Clamp(C(parts[2]), 0, 1)), A(3));
            }
            double hh = Num(parts[0].Replace("deg", "")) / 360, ss = Num(parts[1]), ll = Num(parts[2]);
            if (!parts[1].EndsWith('%')) ss /= 100;
            if (!parts[2].EndsWith('%')) ll /= 100;
            double q = ll < 0.5 ? ll * (1 + ss) : ll + ss - ll * ss, p = 2 * ll - q;
            double Hue(double t) { t = (t % 1 + 1) % 1; return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p; }
            return (new Rgb(Hue(hh + 1.0 / 3), Hue(hh), Hue(hh - 1.0 / 3)), A(3));
        }
        return Named.TryGetValue(lower, out var hex) ? (Rgb.Parse(hex)!.Value, 1) : null;
    }

    // ================================================================
    //  Матрицы
    // ================================================================

    readonly record struct M(double A, double B, double C, double D, double E, double F)
    {
        public static readonly M I = new(1, 0, 0, 1, 0, 0);
        /// <summary>this · o: сначала o, затем this.</summary>
        public M Mul(M o) => new(A * o.A + C * o.B, B * o.A + D * o.B, A * o.C + C * o.D, B * o.C + D * o.D, A * o.E + C * o.F + E, B * o.E + D * o.F + F);
        public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
        public bool IsIdentity => A == 1 && B == 0 && C == 0 && D == 1 && E == 0 && F == 0;
        public double Scale => Math.Sqrt(Math.Abs(A * D - B * C));
    }

    static M ParseTransform(string? t)
    {
        var m = M.I;
        if (string.IsNullOrWhiteSpace(t)) return m;
        foreach (Match f in Regex.Matches(t, @"(matrix|translate|scale|rotate|skewX|skewY)\s*\(([^)]*)\)"))
        {
            var a = Numbers(f.Groups[2].Value);
            // CSS-единицы (deg, px) в transform свойства стиля
            M op = f.Groups[1].Value switch
            {
                "matrix" when a.Count >= 6 => new M(a[0], a[1], a[2], a[3], a[4], a[5]),
                "translate" when a.Count >= 1 => new M(1, 0, 0, 1, a[0], a.Count > 1 ? a[1] : 0),
                "scale" when a.Count >= 1 => new M(a[0], 0, 0, a.Count > 1 ? a[1] : a[0], 0, 0),
                "rotate" when a.Count >= 1 => Rot(a[0], a.Count >= 3 ? a[1] : 0, a.Count >= 3 ? a[2] : 0),
                "skewX" when a.Count >= 1 => new M(1, 0, Math.Tan(a[0] * Math.PI / 180), 1, 0, 0),
                "skewY" when a.Count >= 1 => new M(1, Math.Tan(a[0] * Math.PI / 180), 0, 1, 0, 0),
                _ => M.I,
            };
            m = m.Mul(op);
        }
        return m;

        static M Rot(double deg, double cx, double cy)
        {
            double r = deg * Math.PI / 180, c = Math.Cos(r), s = Math.Sin(r);
            return new M(1, 0, 0, 1, cx, cy).Mul(new M(c, s, -s, c, 0, 0)).Mul(new M(1, 0, 0, 1, -cx, -cy));
        }
    }

    /// <summary>viewBox → область (x, y, w, h) по preserveAspectRatio.</summary>
    static M ViewMatrix((double X, double Y, double W, double H) vb, double x, double y, double w, double h, string? par)
    {
        par = (par ?? "xMidYMid meet").Trim();
        if (par.StartsWith("defer")) par = par[5..].Trim();
        var t = par.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        string align = t.Length > 0 ? t[0] : "xMidYMid";
        bool slice = t.Length > 1 && t[1] == "slice";
        double sx = w / vb.W, sy = h / vb.H;
        if (align != "none")
        {
            double s = slice ? Math.Max(sx, sy) : Math.Min(sx, sy);
            sx = sy = s;
        }
        double tx = x - vb.X * sx, ty = y - vb.Y * sy;
        if (align != "none")
        {
            double ew = vb.W * sx, eh = vb.H * sy;
            if (align.Contains("xMid")) tx += (w - ew) / 2; else if (align.Contains("xMax")) tx += w - ew;
            if (align.Contains("YMid")) ty += (h - eh) / 2; else if (align.Contains("YMax")) ty += h - eh;
        }
        return new M(sx, 0, 0, sy, tx, ty);
    }

    // ================================================================
    //  Геометрия
    // ================================================================

    /// <summary>Контур: команды M/L/C/Z в координатах пользователя.</summary>
    sealed class Geo
    {
        public readonly List<(char Cmd, double[] P)> Cmds = [];
        double _cx, _cy, _sx, _sy;
        public void Move(double x, double y) { Cmds.Add(('M', [x, y])); _cx = _sx = x; _cy = _sy = y; }
        public void Line(double x, double y) { if (Cmds.Count == 0) Move(_cx, _cy); Cmds.Add(('L', [x, y])); _cx = x; _cy = y; }
        public void Curve(double x1, double y1, double x2, double y2, double x, double y) { if (Cmds.Count == 0) Move(_cx, _cy); Cmds.Add(('C', [x1, y1, x2, y2, x, y])); _cx = x; _cy = y; }
        public void Close() { if (Cmds.Count > 0) Cmds.Add(('Z', [])); _cx = _sx; _cy = _sy; }
        public (double X, double Y) Cur => (_cx, _cy);

        public (double X, double Y, double W, double H) Bounds()
        {
            double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
            foreach (var (c, p) in Cmds)
                for (int i = 0; i + 1 < p.Length; i += 2) { x0 = Math.Min(x0, p[i]); y0 = Math.Min(y0, p[i + 1]); x1 = Math.Max(x1, p[i]); y1 = Math.Max(y1, p[i + 1]); }
            return x0 > x1 ? (0, 0, 0, 0) : (x0, y0, x1 - x0, y1 - y0);
        }

        public void Emit(PdfCanvas.VectorPath vp, M m)
        {
            foreach (var (c, p) in Cmds)
            {
                switch (c)
                {
                    case 'M': { var (x, y) = m.Apply(p[0], p[1]); vp.M(x, y); break; }
                    case 'L': { var (x, y) = m.Apply(p[0], p[1]); vp.L(x, y); break; }
                    case 'C':
                    {
                        var (a, b) = m.Apply(p[0], p[1]); var (cc, d) = m.Apply(p[2], p[3]); var (e, f) = m.Apply(p[4], p[5]);
                        vp.C(a, b, cc, d, e, f);
                        break;
                    }
                    case 'Z': vp.Z(); break;
                }
            }
        }
    }

    static void Ellipse(Geo g, double cx, double cy, double rx, double ry)
    {
        const double k = 0.5522847498;
        g.Move(cx + rx, cy);
        g.Curve(cx + rx, cy + ry * k, cx + rx * k, cy + ry, cx, cy + ry);
        g.Curve(cx - rx * k, cy + ry, cx - rx, cy + ry * k, cx - rx, cy);
        g.Curve(cx - rx, cy - ry * k, cx - rx * k, cy - ry, cx, cy - ry);
        g.Curve(cx + rx * k, cy - ry, cx + rx, cy - ry * k, cx + rx, cy);
        g.Close();
    }

    /// <summary>Разбор данных контура (все команды, относительные, сокращённые кривые, дуги).</summary>
    static void ParsePath(string d, Geo g)
    {
        var toks = Regex.Matches(d, @"[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?").Select(m => m.Value).ToList();
        int i = 0;
        char cmd = ' ';
        double lcx = 0, lcy = 0;          // последняя контрольная точка (S/T)
        char prev = ' ';
        bool IsNum() => i < toks.Count && !char.IsLetter(toks[i][0]);
        double N() => i < toks.Count && !char.IsLetter(toks[i][0]) ? double.Parse(toks[i++], Inv) : 0;
        // Флаги дуги могут быть слитны («a10 10 0 01 5 5»)
        double Flag()
        {
            if (i >= toks.Count || char.IsLetter(toks[i][0])) return 0;
            string t = toks[i];
            if (t.Length > 1 && t[0] is '0' or '1' && !t.Contains('.'))
            {
                toks[i] = t[1..];
                return t[0] - '0';
            }
            i++;
            return double.Parse(t, Inv);
        }
        while (i < toks.Count)
        {
            if (char.IsLetter(toks[i][0])) cmd = toks[i++][0];
            else if (cmd == ' ') { i++; continue; }
            var (cx, cy) = g.Cur;
            bool rel = char.IsLower(cmd);
            double ox = rel ? cx : 0, oy = rel ? cy : 0;
            switch (char.ToUpperInvariant(cmd))
            {
                case 'M':
                    g.Move(ox + N(), oy + N());
                    cmd = rel ? 'l' : 'L';
                    break;
                case 'L': g.Line(ox + N(), oy + N()); break;
                case 'H': g.Line(ox + N(), cy); break;
                case 'V': g.Line(cx, oy + N()); break;
                case 'C':
                {
                    double x1 = ox + N(), y1 = oy + N(), x2 = ox + N(), y2 = oy + N(), x = ox + N(), y = oy + N();
                    g.Curve(x1, y1, x2, y2, x, y); lcx = x2; lcy = y2;
                    break;
                }
                case 'S':
                {
                    double x1 = char.ToUpperInvariant(prev) is 'C' or 'S' ? 2 * cx - lcx : cx, y1 = char.ToUpperInvariant(prev) is 'C' or 'S' ? 2 * cy - lcy : cy;
                    double x2 = ox + N(), y2 = oy + N(), x = ox + N(), y = oy + N();
                    g.Curve(x1, y1, x2, y2, x, y); lcx = x2; lcy = y2;
                    break;
                }
                case 'Q':
                {
                    double qx = ox + N(), qy = oy + N(), x = ox + N(), y = oy + N();
                    g.Curve(cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy), x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y);
                    lcx = qx; lcy = qy;
                    break;
                }
                case 'T':
                {
                    double qx = char.ToUpperInvariant(prev) is 'Q' or 'T' ? 2 * cx - lcx : cx, qy = char.ToUpperInvariant(prev) is 'Q' or 'T' ? 2 * cy - lcy : cy;
                    double x = ox + N(), y = oy + N();
                    g.Curve(cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy), x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y);
                    lcx = qx; lcy = qy;
                    break;
                }
                case 'A':
                {
                    double rx = N(), ry = N(), rot = N(), large = Flag(), sweep = Flag(), x = ox + N(), y = oy + N();
                    Arc(g, cx, cy, rx, ry, rot, large != 0, sweep != 0, x, y);
                    break;
                }
                case 'Z': g.Close(); break;
                default: i++; break;
            }
            prev = cmd;
            if (char.ToUpperInvariant(cmd) == 'Z' && IsNum()) cmd = 'L';
        }
    }

    /// <summary>Эллиптическая дуга (параметры SVG) → кубические кривые.</summary>
    static void Arc(Geo g, double x1, double y1, double rx, double ry, double phiDeg, bool large, bool sweep, double x2, double y2)
    {
        if (rx == 0 || ry == 0 || x1 == x2 && y1 == y2) { g.Line(x2, y2); return; }
        rx = Math.Abs(rx); ry = Math.Abs(ry);
        double phi = phiDeg * Math.PI / 180, cp = Math.Cos(phi), sp = Math.Sin(phi);
        double dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
        double x1p = cp * dx + sp * dy, y1p = -sp * dx + cp * dy;
        double lam = x1p * x1p / (rx * rx) + y1p * y1p / (ry * ry);
        if (lam > 1) { double s = Math.Sqrt(lam); rx *= s; ry *= s; }
        double num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
        double den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
        double co = Math.Sqrt(Math.Max(0, num / den)) * (large == sweep ? -1 : 1);
        double cxp = co * rx * y1p / ry, cyp = -co * ry * x1p / rx;
        double cx = cp * cxp - sp * cyp + (x1 + x2) / 2, cy = sp * cxp + cp * cyp + (y1 + y2) / 2;
        double Ang(double ux, double uy, double vx, double vy)
        {
            double a = Math.Atan2(ux * vy - uy * vx, ux * vx + uy * vy);
            return a;
        }
        double t1 = Ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
        double dt = Ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
        if (!sweep && dt > 0) dt -= 2 * Math.PI;
        else if (sweep && dt < 0) dt += 2 * Math.PI;
        int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(dt) / (Math.PI / 2) - 1e-9));
        double d = dt / segs, k = 4.0 / 3 * Math.Tan(d / 4);
        (double, double) P(double t) => (cx + rx * Math.Cos(t) * cp - ry * Math.Sin(t) * sp, cy + rx * Math.Cos(t) * sp + ry * Math.Sin(t) * cp);
        (double, double) Dv(double t) => (-rx * Math.Sin(t) * cp - ry * Math.Cos(t) * sp, -rx * Math.Sin(t) * sp + ry * Math.Cos(t) * cp);
        double ta = t1;
        for (int s = 0; s < segs; s++)
        {
            double tb = ta + d;
            var (ax, ay) = P(ta); var (bx, by) = P(tb); var (dax, day) = Dv(ta); var (dbx, dby) = Dv(tb);
            g.Curve(ax + k * dax, ay + k * day, bx - k * dbx, by - k * dby, s == segs - 1 ? x2 : bx, s == segs - 1 ? y2 : by);
            ta = tb;
        }
    }

    // ================================================================
    //  Отрисовка
    // ================================================================

    sealed class Ctx
    {
        public required PdfCanvas C;
        public required SvgDocument Doc;
        public double VW, VH;            // область просмотра (для %)
        public double Alpha = 1;
        public int Depth;
        public FontManager? Fonts;
    }

    public static void Render(PdfCanvas c, byte[] d, double x, double y, double w, double h)
    {
        if (w <= 0 || h <= 0) return;
        var doc = Load(d);
        if (doc == null) return;
        var root = doc.Root;
        var vb = ViewBox(root);
        double vw = vb?.W ?? AbsLength(Attr(root, "width")) ?? w / 0.75, vh = vb?.H ?? AbsLength(Attr(root, "height")) ?? h / 0.75;
        // Без viewBox: собственные размеры рисунка растягиваются на рамку
        var m = vb != null ? ViewMatrix(vb.Value, x, y, w, h, Attr(root, "preserveAspectRatio")) : new M(w / vw, 0, 0, h / vh, x, y);
        var ctx = new Ctx { C = c, Doc = doc, VW = vw, VH = vh, Fonts = c.FontProvider };
        var clip = c.NewPath(); clip.Rect(x, y, w, h);
        c.BeginClip(clip, false);
        c.BeginTransform(0, 0, m.A, m.B, m.C, m.D, m.E, m.F);
        try
        {
            var st = Compute(doc, root, new Style());
            RenderChildren(root, st, ctx);
        }
        catch { /* повреждённый рисунок: выводим то, что успели */ }
        c.EndTransform();
        c.EndClip();
    }

    static void RenderChildren(XElement e, Style st, Ctx ctx)
    {
        foreach (var ch in e.Elements())
        {
            if (ch.Name.LocalName == "switch")
            {
                // Первый подходящий потомок
                var first = ch.Elements().FirstOrDefault(x => Attr(x, "requiredExtensions") == null);
                if (first != null) RenderElement(first, Compute(ctx.Doc, ch, st), ctx);
                continue;
            }
            RenderElement(ch, st, ctx);
        }
    }

    static readonly HashSet<string> NonRendering =
        ["defs", "symbol", "clipPath", "mask", "marker", "pattern", "linearGradient", "radialGradient", "style", "title", "desc",
         "metadata", "filter", "script", "foreignObject", "font", "font-face", "cursor", "view", "animate", "animateTransform", "set"];

    static void RenderElement(XElement e, Style parent, Ctx ctx, M? extra = null)
    {
        string tag = e.Name.LocalName;
        if (NonRendering.Contains(tag) || ctx.Depth > 40) return;
        var st = Compute(ctx.Doc, e, parent);
        if (st.Get("display") == "none") return;
        var c = ctx.C;

        // Преобразование: атрибут (или свойство стиля) и смещение use
        var m = ParseTransform(Attr(e, "transform") ?? (SpecifiedOf(ctx.Doc, e).TryGetValue("transform", out var ct) ? ct : null));
        if (extra is { } ex) m = m.Mul(ex);
        bool tr = !m.IsIdentity;
        if (tr) c.BeginTransform(0, 0, m.A, m.B, m.C, m.D, m.E, m.F);
        double op = Math.Clamp(Num(st.Get("opacity", "1"), 1), 0, 1);
        double savedAlpha = ctx.Alpha;
        ctx.Alpha *= op;
        ctx.Depth++;
        // Обрезка clip-path
        bool clipped = false;
        if (st.Get("clip-path") is { Length: > 0 } cp && cp != "none" && Ref(ctx.Doc, cp) is { } clipEl && clipEl.Name.LocalName == "clipPath")
        {
            var bbox = BBox(e, st, ctx);
            var vp = c.NewPath();
            var cm = Attr(clipEl, "clipPathUnits") == "objectBoundingBox" ? new M(bbox.W, 0, 0, bbox.H, bbox.X, bbox.Y) : M.I;
            cm = cm.Mul(ParseTransform(Attr(clipEl, "transform")));
            bool evenOdd = false;
            foreach (var ce in clipEl.Elements())
            {
                var cst = Compute(ctx.Doc, ce, Compute(ctx.Doc, clipEl, st));
                if (cst.Get("display") == "none") continue;
                var g = ShapeGeo(ce, cst, ctx);
                if (g == null) continue;
                g.Emit(vp, cm.Mul(ParseTransform(Attr(ce, "transform"))));
                evenOdd |= cst.Get("clip-rule") == "evenodd";
            }
            c.BeginClip(vp, evenOdd);
            clipped = true;
        }
        try
        {
            switch (tag)
            {
                case "g" or "a": RenderChildren(e, st, ctx); break;
                case "svg": NestedSvg(e, st, ctx, null, null); break;
                case "use": Use(e, st, ctx); break;
                case "image": Image(e, st, ctx); break;
                case "text": Text(e, st, ctx); break;
                default:
                    if (st.Get("visibility") is "hidden" or "collapse") break;
                    var geo = ShapeGeo(e, st, ctx);
                    if (geo != null)
                    {
                        Paint(geo, st, ctx, e);
                        Markers(e, geo, st, ctx);
                    }
                    break;
            }
        }
        finally
        {
            ctx.Depth--;
            ctx.Alpha = savedAlpha;
            if (clipped) c.EndClip();
            if (tr) c.EndTransform();
        }
    }

    static double LenX(XElement e, string a, Ctx ctx, Style st) => Len(Attr(e, a), ctx.VW, FontSize(st));
    static double LenY(XElement e, string a, Ctx ctx, Style st) => Len(Attr(e, a), ctx.VH, FontSize(st));
    static double LenD(string? v, Ctx ctx, Style st) => Len(v, Math.Sqrt((ctx.VW * ctx.VW + ctx.VH * ctx.VH) / 2), FontSize(st));
    static double FontSize(Style st) => Len(st.Get("font-size", "16"), 16, 16) is var f && f > 0 ? f : 16;

    /// <summary>Контур фигуры (rect, circle, ellipse, line, polyline, polygon, path); null — не фигура.</summary>
    static Geo? ShapeGeo(XElement e, Style st, Ctx ctx)
    {
        var g = new Geo();
        switch (e.Name.LocalName)
        {
            case "rect":
            {
                double x = LenX(e, "x", ctx, st), y = LenY(e, "y", ctx, st), w = LenX(e, "width", ctx, st), h = LenY(e, "height", ctx, st);
                if (w <= 0 || h <= 0) return null;
                string? rxa = Attr(e, "rx"), rya = Attr(e, "ry");
                double rx = rxa != null && rxa != "auto" ? LenX(e, "rx", ctx, st) : -1, ry = rya != null && rya != "auto" ? LenY(e, "ry", ctx, st) : -1;
                if (rx < 0) rx = ry; if (ry < 0) ry = rx;
                rx = Math.Clamp(rx, 0, w / 2); ry = Math.Clamp(ry, 0, h / 2);
                if (rx <= 0 || ry <= 0) { g.Move(x, y); g.Line(x + w, y); g.Line(x + w, y + h); g.Line(x, y + h); g.Close(); return g; }
                const double k = 0.5522847498;
                g.Move(x + rx, y);
                g.Line(x + w - rx, y); g.Curve(x + w - rx + rx * k, y, x + w, y + ry - ry * k, x + w, y + ry);
                g.Line(x + w, y + h - ry); g.Curve(x + w, y + h - ry + ry * k, x + w - rx + rx * k, y + h, x + w - rx, y + h);
                g.Line(x + rx, y + h); g.Curve(x + rx - rx * k, y + h, x, y + h - ry + ry * k, x, y + h - ry);
                g.Line(x, y + ry); g.Curve(x, y + ry - ry * k, x + rx - rx * k, y, x + rx, y);
                g.Close();
                return g;
            }
            case "circle":
            {
                double r = LenD(Attr(e, "r"), ctx, st);
                if (r <= 0) return null;
                Ellipse(g, LenX(e, "cx", ctx, st), LenY(e, "cy", ctx, st), r, r);
                return g;
            }
            case "ellipse":
            {
                double rx = LenX(e, "rx", ctx, st), ry = LenY(e, "ry", ctx, st);
                if (Attr(e, "rx") is null or "auto") rx = ry; if (Attr(e, "ry") is null or "auto") ry = rx;
                if (rx <= 0 || ry <= 0) return null;
                Ellipse(g, LenX(e, "cx", ctx, st), LenY(e, "cy", ctx, st), rx, ry);
                return g;
            }
            case "line":
                g.Move(LenX(e, "x1", ctx, st), LenY(e, "y1", ctx, st));
                g.Line(LenX(e, "x2", ctx, st), LenY(e, "y2", ctx, st));
                return g;
            case "polyline" or "polygon":
            {
                var n = Numbers(Attr(e, "points"));
                if (n.Count < 4) return null;
                g.Move(n[0], n[1]);
                for (int i = 2; i + 1 < n.Count; i += 2) g.Line(n[i], n[i + 1]);
                if (e.Name.LocalName == "polygon") g.Close();
                return g;
            }
            case "path":
                ParsePath(Attr(e, "d") ?? "", g);
                return g.Cmds.Count > 0 ? g : null;
        }
        return null;
    }

    /// <summary>Рамка элемента в его системе координат (для objectBoundingBox).</summary>
    static (double X, double Y, double W, double H) BBox(XElement e, Style st, Ctx ctx)
    {
        if (ShapeGeo(e, st, ctx) is { } g) return g.Bounds();
        if (e.Name.LocalName is "image" or "svg" or "use")
            return (LenX(e, "x", ctx, st), LenY(e, "y", ctx, st), LenX(e, "width", ctx, st), LenY(e, "height", ctx, st));
        // Группа: объединение рамок фигур-потомков (без учёта их преобразований)
        double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
        foreach (var d in e.Descendants())
        {
            var gg = ShapeGeo(d, st, ctx);
            if (gg == null) continue;
            var b = gg.Bounds();
            x0 = Math.Min(x0, b.X); y0 = Math.Min(y0, b.Y); x1 = Math.Max(x1, b.X + b.W); y1 = Math.Max(y1, b.Y + b.H);
        }
        return x0 > x1 ? (0, 0, 0, 0) : (x0, y0, x1 - x0, y1 - y0);
    }

    // ---------------- Заливка и обводка ----------------

    static void Paint(Geo geo, Style st, Ctx ctx, XElement el)
    {
        var c = ctx.C;
        var vp = c.NewPath();
        geo.Emit(vp, M.I);
        if (vp.IsEmpty) return;
        bool evenOdd = st.Get("fill-rule") == "evenodd";
        string fillV = st.Get("fill", "black");
        string strokeV = st.Get("stroke", "none");
        double fo = Math.Clamp(Num(st.Get("fill-opacity", "1"), 1), 0, 1) * ctx.Alpha;
        double so = Math.Clamp(Num(st.Get("stroke-opacity", "1"), 1), 0, 1) * ctx.Alpha;
        bool strokeFirst = st.Get("paint-order").TrimStart().StartsWith("stroke");
        // Линия без площади не заливается
        bool open = el.Name.LocalName is "line";
        if (strokeFirst) { Stroke(); if (!open) Fill(); }
        else { if (!open) Fill(); Stroke(); }

        void Fill()
        {
            var paint = ResolvePaint(fillV, st, ctx);
            if (paint == null) return;
            if (paint.Value.Gradient is { } grad) { GradientFill(vp, evenOdd, grad, geo.Bounds(), fo, ctx); return; }
            if (paint.Value.Pattern is { } pat) { PatternFill(vp, evenOdd, pat, geo.Bounds(), st, ctx); return; }
            double a = fo * paint.Value.Alpha;
            if (a <= 0.001) return;
            var sb = new StringBuilder("q ");
            if (a < 0.999) sb.Append('/').Append(c.GState(a, 1)).Append(" gs ");
            sb.Append(Col(paint.Value.Color)).Append(" rg ").Append(vp.Ops).Append(evenOdd ? "f*" : "f").Append(" Q\n");
            c.RawOps(sb.ToString());
        }

        void Stroke()
        {
            var paint = ResolvePaint(strokeV, st, ctx);
            if (paint == null) return;
            double lw = LenD(st.Get("stroke-width", "1"), ctx, st);
            if (lw <= 0) return;
            double a = so * paint.Value.Alpha;
            if (a <= 0.001) return;
            var sb = new StringBuilder("q ");
            if (a < 0.999) sb.Append('/').Append(c.GState(1, a)).Append(" gs ");
            sb.Append(Col(paint.Value.Color)).Append(" RG ").Append(PdfCanvas.Num(lw)).Append(" w ");
            sb.Append(st.Get("stroke-linecap") switch { "round" => "1", "square" => "2", _ => "0" }).Append(" J ");
            sb.Append(st.Get("stroke-linejoin") switch { "round" => "1", "bevel" => "2", _ => "0" }).Append(" j ");
            sb.Append(PdfCanvas.Num(Math.Max(1, Num(st.Get("stroke-miterlimit", "4"), 4)))).Append(" M ");
            var dash = Numbers(st.Get("stroke-dasharray", "none") == "none" ? null : st.Get("stroke-dasharray").Replace("%", ""));
            if (dash.Count > 0 && dash.Any(v => v > 0) && dash.All(v => v >= 0))
            {
                if (dash.Count % 2 == 1) dash.AddRange(dash.ToList());
                sb.Append('[').Append(string.Join(' ', dash.Select(PdfCanvas.Num))).Append("] ").Append(PdfCanvas.Num(LenD(st.Get("stroke-dashoffset", "0"), ctx, st))).Append(" d ");
            }
            // Неравномерное масштабирование: vector-effect non-scaling-stroke не поддерживается — толщина масштабируется
            sb.Append(vp.Ops).Append("S Q\n");
            c.RawOps(sb.ToString());
        }
    }

    static string Col(Rgb c) => $"{PdfCanvas.Num(c.R)} {PdfCanvas.Num(c.G)} {PdfCanvas.Num(c.B)}";

    readonly record struct PaintSpec(Rgb Color, double Alpha, XElement? Gradient, XElement? Pattern);

    /// <summary>Краска: цвет, градиент или узор (с запасным цветом «url(#id) цвет»).</summary>
    static PaintSpec? ResolvePaint(string v, Style st, Ctx ctx)
    {
        v = v.Trim();
        if (v is "" or "none") return null;
        if (v.StartsWith("url(", StringComparison.OrdinalIgnoreCase))
        {
            int close = v.IndexOf(')');
            string fallback = close >= 0 ? v[(close + 1)..].Trim() : "";
            var target = Ref(ctx.Doc, v[..(close >= 0 ? close + 1 : v.Length)]);
            if (target != null && target.Name.LocalName is "linearGradient" or "radialGradient")
            {
                var stops = Stops(target, st, ctx);
                if (stops.Count == 0) return null;
                if (stops.Count == 1) return new PaintSpec(stops[0].Color, stops[0].Alpha, null, null);
                return new PaintSpec(stops[0].Color, 1, target, null);
            }
            if (target != null && target.Name.LocalName == "pattern") return new PaintSpec(Rgb.Black, 1, null, target);
            if (fallback.Length > 0) return ResolvePaint(fallback, st, ctx);
            return null;
        }
        if (v is "context-fill" or "context-stroke") return null;
        var c = ParseColor(v, st.Get("color", "black"));
        return c is { } cc ? new PaintSpec(cc.Color, cc.Alpha, null, null) : null;
    }

    /// <summary>Опорные точки градиента (с наследованием по href).</summary>
    static List<(double Pos, Rgb Color, double Alpha)> Stops(XElement grad, Style st, Ctx ctx)
    {
        var g = grad;
        for (int guard = 0; guard < 8 && g != null; guard++)
        {
            var stops = g.Elements().Where(x => x.Name.LocalName == "stop").ToList();
            if (stops.Count > 0)
            {
                var res = new List<(double, Rgb, double)>();
                double last = 0;
                foreach (var s in stops)
                {
                    var ss = Compute(ctx.Doc, s, st);
                    double off = Math.Clamp(Num(Attr(s, "offset"), 0), 0, 1);
                    off = Math.Max(off, last); last = off;
                    var col = ParseColor(ss.Get("stop-color", "black"), ss.Get("color", "black")) ?? (Rgb.Black, 0);
                    res.Add((off, col.Color, col.Alpha * Math.Clamp(Num(ss.Get("stop-opacity", "1"), 1), 0, 1)));
                }
                return res;
            }
            g = Ref(ctx.Doc, Href(g));
        }
        return [];
    }

    /// <summary>Атрибут градиента с наследованием по href.</summary>
    static string? GradAttr(XElement grad, string name, Ctx ctx)
    {
        var g = grad;
        for (int guard = 0; guard < 8 && g != null; guard++)
        {
            if (Attr(g, name) is { } v) return v;
            g = Ref(ctx.Doc, Href(g));
        }
        return null;
    }

    static void GradientFill(PdfCanvas.VectorPath vp, bool evenOdd, XElement grad, (double X, double Y, double W, double H) bbox, double alpha, Ctx ctx)
    {
        var c = ctx.C;
        var stops = Stops(grad, new Style(), ctx);
        if (stops.Count == 0) return;
        bool obb = GradAttr(grad, "gradientUnits", ctx) != "userSpaceOnUse";
        if (obb && (bbox.W <= 0 || bbox.H <= 0)) return;
        var gm = obb ? new M(bbox.W, 0, 0, bbox.H, bbox.X, bbox.Y) : M.I;
        gm = gm.Mul(ParseTransform(GradAttr(grad, "gradientTransform", ctx)));
        double rw = obb ? 1 : ctx.VW, rh = obb ? 1 : ctx.VH;
        double L(string n, string def, double reference) => GradAttr(grad, n, ctx) is { } v ? (obb && !v.EndsWith('%') ? Num(v) : Len(v, reference, 16)) : (def.EndsWith('%') ? Num(def) * reference : Num(def));
        bool radial = grad.Name.LocalName == "radialGradient";
        double[] coords;
        if (radial)
        {
            double rd = obb ? 1 : Math.Sqrt((ctx.VW * ctx.VW + ctx.VH * ctx.VH) / 2);
            double cx = L("cx", "50%", rw), cy = L("cy", "50%", rh), r = L("r", "50%", rd);
            double fx = GradAttr(grad, "fx", ctx) != null ? L("fx", "50%", rw) : cx, fy = GradAttr(grad, "fy", ctx) != null ? L("fy", "50%", rh) : cy;
            coords = [fx, fy, cx, cy, r];
        }
        else coords = [L("x1", "0%", rw), L("y1", "0%", rh), L("x2", "100%", rw), L("y2", "0%", rh)];
        // Прозрачность точек — средняя (мягкая маска не строится)
        double a = alpha * stops.Average(s => s.Alpha);
        if (a <= 0.001) return;
        string sh = c.GradientShading(stops.Select(s => (s.Pos, s.Color)).ToList(), radial, coords);
        c.RawOps("q " + (a < 0.999 ? "/" + c.GState(a, a) + " gs " : "") + vp.Ops + (evenOdd ? "W* n\n" : "W n\n"));
        c.BeginTransform(0, 0, gm.A, gm.B, gm.C, gm.D, gm.E, gm.F);
        c.RawOps($"/{sh} sh\n");
        c.EndTransform();
        c.RawOps("Q\n");
    }

    /// <summary>Узор: плитка рисуется повторно внутри фигуры (с обрезкой по ней).</summary>
    static void PatternFill(PdfCanvas.VectorPath vp, bool evenOdd, XElement pat, (double X, double Y, double W, double H) bbox, Style st, Ctx ctx)
    {
        var c = ctx.C;
        var p = pat;
        // Содержимое и атрибуты могут наследоваться по href
        XElement? content = p;
        for (int guard = 0; guard < 8 && content != null && !content.Elements().Any(); guard++) content = Ref(ctx.Doc, Href(content));
        if (content == null) return;
        bool obbUnits = GradAttr(pat, "patternUnits", ctx) != "userSpaceOnUse";
        double Lp(string n) => GradAttr(pat, n, ctx) is { } v ? (obbUnits && !v.EndsWith('%') ? Num(v) * (n is "x" or "width" ? bbox.W : bbox.H) + (n == "x" ? bbox.X : n == "y" ? bbox.Y : 0) : Len(v, n is "x" or "width" ? ctx.VW : ctx.VH, 16)) : (n == "x" ? (obbUnits ? bbox.X : 0) : n == "y" ? (obbUnits ? bbox.Y : 0) : 0);
        double px = Lp("x"), py = Lp("y"), pw = Lp("width"), ph = Lp("height");
        if (pw <= 0.01 || ph <= 0.01) return;
        var pm = ParseTransform(GradAttr(pat, "patternTransform", ctx));
        var vbA = GradAttr(pat, "viewBox", ctx);
        bool contentObb = GradAttr(pat, "patternContentUnits", ctx) == "objectBoundingBox";
        c.RawOps("q " + vp.Ops + (evenOdd ? "W* n\n" : "W n\n"));
        c.BeginTransform(0, 0, pm.A, pm.B, pm.C, pm.D, pm.E, pm.F);
        // Плитки, покрывающие рамку фигуры (с запасом на преобразование)
        double span = Math.Max(bbox.W, bbox.H) * 2;
        int i0 = (int)Math.Floor((bbox.X - span / 2 - px) / pw), i1 = (int)Math.Ceiling((bbox.X + bbox.W + span / 2 - px) / pw);
        int j0 = (int)Math.Floor((bbox.Y - span / 2 - py) / ph), j1 = (int)Math.Ceiling((bbox.Y + bbox.H + span / 2 - py) / ph);
        if ((long)(i1 - i0) * (j1 - j0) <= 2500)
        {
            var pst = Compute(ctx.Doc, content, new Style());
            for (int i = i0; i < i1; i++)
                for (int j = j0; j < j1; j++)
                {
                    double tx = px + i * pw, ty = py + j * ph;
                    var tile = c.NewPath(); tile.Rect(tx, ty, pw, ph);
                    c.BeginClip(tile, false);
                    var tm = new M(1, 0, 0, 1, tx, ty);
                    if (vbA != null && Numbers(vbA) is { Count: 4 } vbn && vbn[2] > 0 && vbn[3] > 0)
                        tm = tm.Mul(ViewMatrix((vbn[0], vbn[1], vbn[2], vbn[3]), 0, 0, pw, ph, GradAttr(pat, "preserveAspectRatio", ctx)));
                    else if (contentObb) tm = tm.Mul(new M(bbox.W, 0, 0, bbox.H, 0, 0));
                    c.BeginTransform(0, 0, tm.A, tm.B, tm.C, tm.D, tm.E, tm.F);
                    RenderChildren(content, pst, ctx);
                    c.EndTransform();
                    c.EndClip();
                }
        }
        c.EndTransform();
        c.RawOps("Q\n");
    }

    // ---------------- Маркеры ----------------

    static void Markers(XElement e, Geo geo, Style st, Ctx ctx)
    {
        if (e.Name.LocalName is not ("path" or "line" or "polyline" or "polygon")) return;
        string ms = st.Get("marker-start", "none"), mm = st.Get("marker-mid", "none"), me = st.Get("marker-end", "none");
        if (ms == "none" && mm == "none" && me == "none") return;
        // Вершины и направления касательных
        var verts = new List<(double X, double Y, double In, double Out)>();
        double cx = 0, cy = 0, sx = 0, sy = 0, lastAngle = 0;
        int subStart = -1;
        foreach (var (cmd, p) in geo.Cmds)
        {
            switch (cmd)
            {
                case 'M':
                    cx = sx = p[0]; cy = sy = p[1];
                    verts.Add((cx, cy, double.NaN, double.NaN)); subStart = verts.Count - 1;
                    break;
                case 'L':
                {
                    double a = Math.Atan2(p[1] - cy, p[0] - cx);
                    if (verts.Count > 0) verts[^1] = verts[^1] with { Out = a };
                    cx = p[0]; cy = p[1];
                    verts.Add((cx, cy, a, double.NaN)); lastAngle = a;
                    break;
                }
                case 'C':
                {
                    double a0 = Math.Atan2(p[1] - cy, p[0] - cx), a1 = Math.Atan2(p[5] - p[3], p[4] - p[2]);
                    if (p[0] == cx && p[1] == cy) a0 = Math.Atan2(p[3] - cy, p[2] - cx);
                    if (p[4] == p[2] && p[5] == p[3]) a1 = Math.Atan2(p[5] - p[1], p[4] - p[0]);
                    if (verts.Count > 0) verts[^1] = verts[^1] with { Out = a0 };
                    cx = p[4]; cy = p[5];
                    verts.Add((cx, cy, a1, double.NaN)); lastAngle = a1;
                    break;
                }
                case 'Z':
                {
                    double a = Math.Atan2(sy - cy, sx - cx);
                    if (verts.Count > 0) verts[^1] = verts[^1] with { Out = a };
                    cx = sx; cy = sy;
                    if (subStart >= 0) verts[subStart] = verts[subStart] with { In = a };
                    break;
                }
            }
        }
        if (verts.Count == 0) return;
        double sw = LenD(st.Get("stroke-width", "1"), ctx, st);
        for (int i = 0; i < verts.Count; i++)
        {
            string which = i == 0 ? ms : i == verts.Count - 1 ? me : mm;
            if (which == "none" || Ref(ctx.Doc, which) is not { } marker || marker.Name.LocalName != "marker") continue;
            var v = verts[i];
            double inA = double.IsNaN(v.In) ? v.Out : v.In, outA = double.IsNaN(v.Out) ? v.In : v.Out;
            if (double.IsNaN(inA)) inA = outA = 0;
            double bis = Math.Atan2(Math.Sin(inA) + Math.Sin(outA), Math.Cos(inA) + Math.Cos(outA));
            string orient = Attr(marker, "orient") ?? "0";
            double ang = orient switch
            {
                "auto" => i == 0 ? outA : i == verts.Count - 1 ? inA : bis,
                "auto-start-reverse" => i == 0 ? outA + Math.PI : i == verts.Count - 1 ? inA : bis,
                _ => Num(orient.Replace("deg", "")) * Math.PI / 180,
            };
            double mw = Attr(marker, "markerWidth") is { } mwa ? Len(mwa, ctx.VW, 16) : 3, mh = Attr(marker, "markerHeight") is { } mha ? Len(mha, ctx.VH, 16) : 3;
            double scale = Attr(marker, "markerUnits") == "userSpaceOnUse" ? 1 : sw;
            var mt = new M(1, 0, 0, 1, v.X, v.Y).Mul(new M(Math.Cos(ang), Math.Sin(ang), -Math.Sin(ang), Math.Cos(ang), 0, 0)).Mul(new M(scale, 0, 0, scale, 0, 0));
            var vb = ViewBox(marker);
            double refX = Num(Attr(marker, "refX")), refY = Num(Attr(marker, "refY"));
            if (vb != null)
            {
                var vm = ViewMatrix(vb.Value, 0, 0, mw, mh, Attr(marker, "preserveAspectRatio"));
                var (rx, ry) = vm.Apply(refX, refY);
                mt = mt.Mul(new M(1, 0, 0, 1, -rx, -ry)).Mul(vm);
            }
            else mt = mt.Mul(new M(1, 0, 0, 1, -refX, -refY));
            ctx.C.BeginTransform(0, 0, mt.A, mt.B, mt.C, mt.D, mt.E, mt.F);
            // Стиль маркера наследуется от самого маркера, не от фигуры
            RenderChildren(marker, Compute(ctx.Doc, marker, new Style()), ctx);
            ctx.C.EndTransform();
        }
    }

    // ---------------- use, вложенный svg, image ----------------

    static void Use(XElement e, Style st, Ctx ctx)
    {
        var target = Ref(ctx.Doc, Href(e));
        if (target == null || target.AncestorsAndSelf().Contains(e)) return;
        double x = LenX(e, "x", ctx, st), y = LenY(e, "y", ctx, st);
        var t = new M(1, 0, 0, 1, x, y);
        if (target.Name.LocalName == "symbol")
        {
            double w = Attr(e, "width") != null ? LenX(e, "width", ctx, st) : Attr(target, "width") != null ? LenX(target, "width", ctx, st) : ctx.VW;
            double h = Attr(e, "height") != null ? LenY(e, "height", ctx, st) : Attr(target, "height") != null ? LenY(target, "height", ctx, st) : ctx.VH;
            ctx.C.BeginTransform(0, 0, t.A, t.B, t.C, t.D, t.E, t.F);
            var sst = Compute(ctx.Doc, target, st);
            var vb = ViewBox(target);
            var vm = vb != null ? ViewMatrix(vb.Value, 0, 0, w, h, Attr(target, "preserveAspectRatio")) : M.I;
            bool clip = Attr(target, "overflow") is not ("visible" or "auto") && sst.Get("overflow") is not ("visible" or "auto");
            if (clip) { var cp = ctx.C.NewPath(); cp.Rect(0, 0, w, h); ctx.C.BeginClip(cp, false); }
            ctx.C.BeginTransform(0, 0, vm.A, vm.B, vm.C, vm.D, vm.E, vm.F);
            double vw = ctx.VW, vh = ctx.VH;
            if (vb != null) { ctx.VW = vb.Value.W; ctx.VH = vb.Value.H; }
            RenderChildren(target, sst, ctx);
            ctx.VW = vw; ctx.VH = vh;
            ctx.C.EndTransform();
            if (clip) ctx.C.EndClip();
            ctx.C.EndTransform();
            return;
        }
        if (target.Name.LocalName == "svg")
        {
            ctx.C.BeginTransform(0, 0, t.A, t.B, t.C, t.D, t.E, t.F);
            NestedSvg(target, st, ctx, Attr(e, "width") != null ? LenX(e, "width", ctx, st) : null, Attr(e, "height") != null ? LenY(e, "height", ctx, st) : null);
            ctx.C.EndTransform();
            return;
        }
        RenderElement(target, st, ctx, t);
    }

    static void NestedSvg(XElement e, Style st, Ctx ctx, double? wOverride, double? hOverride)
    {
        double x = LenX(e, "x", ctx, st), y = LenY(e, "y", ctx, st);
        double w = wOverride ?? (Attr(e, "width") is { } wa && wa != "auto" ? LenX(e, "width", ctx, st) : ctx.VW);
        double h = hOverride ?? (Attr(e, "height") is { } ha && ha != "auto" ? LenY(e, "height", ctx, st) : ctx.VH);
        if (w <= 0 || h <= 0) return;
        var vb = ViewBox(e);
        var vm = vb != null ? ViewMatrix(vb.Value, x, y, w, h, Attr(e, "preserveAspectRatio")) : new M(1, 0, 0, 1, x, y);
        bool clip = st.Get("overflow") is not ("visible" or "auto");
        if (clip) { var cp = ctx.C.NewPath(); cp.Rect(x, y, w, h); ctx.C.BeginClip(cp, false); }
        ctx.C.BeginTransform(0, 0, vm.A, vm.B, vm.C, vm.D, vm.E, vm.F);
        double vw = ctx.VW, vh = ctx.VH;
        ctx.VW = vb?.W ?? w; ctx.VH = vb?.H ?? h;
        RenderChildren(e, st, ctx);
        ctx.VW = vw; ctx.VH = vh;
        ctx.C.EndTransform();
        if (clip) ctx.C.EndClip();
    }

    static void Image(XElement e, Style st, Ctx ctx)
    {
        if (st.Get("visibility") is "hidden" or "collapse") return;
        string? href = Href(e);
        if (string.IsNullOrWhiteSpace(href) || !href.TrimStart().StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return;
        byte[]? bytes;
        try
        {
            string h = href.Trim();
            int comma = h.IndexOf(',');
            string meta = h[5..comma];
            string data = h[(comma + 1)..];
            bytes = meta.Contains(";base64", StringComparison.OrdinalIgnoreCase)
                ? Convert.FromBase64String(Regex.Replace(Uri.UnescapeDataString(data), @"\s", ""))
                : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(data));
        }
        catch { return; }
        double x = LenX(e, "x", ctx, st), y = LenY(e, "y", ctx, st);
        // Собственный размер картинки
        double iw, ih;
        bool svg = IsSvg(bytes);
        if (svg && Size(bytes) is { } ss) { iw = ss.W / 0.75; ih = ss.H / 0.75; }
        else if (DocxToPdf.Images.ImageDecoder.Decode(bytes) is { } info) { iw = info.Width; ih = info.Height; }
        else return;
        double w = Attr(e, "width") is { } wa && wa != "auto" ? LenX(e, "width", ctx, st) : iw;
        double h2 = Attr(e, "height") is { } ha && ha != "auto" ? LenY(e, "height", ctx, st) : ih;
        if (Attr(e, "width") is null or "auto" && Attr(e, "height") is { } && ih > 0) w = h2 * iw / ih;
        if (Attr(e, "height") is null or "auto" && Attr(e, "width") is { } && iw > 0) h2 = w * ih / iw;
        if (w <= 0 || h2 <= 0) return;
        var vm = ViewMatrix((0, 0, iw, ih), x, y, w, h2, Attr(e, "preserveAspectRatio"));
        var clip = ctx.C.NewPath(); clip.Rect(x, y, w, h2);
        ctx.C.BeginClip(clip, false);
        var (dx, dy) = vm.Apply(0, 0);
        double dw = iw * vm.A, dh = ih * vm.D;
        string key = "svg-img-" + RuntimeHelpers.GetHashCode(bytes).ToString(Inv) + "-" + bytes.Length.ToString(Inv);
        if (svg) Render(ctx.C, bytes, dx, dy, dw, dh);
        else ctx.C.Image(key, bytes, dx, dy, dw, dh, ctx.Alpha);
        ctx.C.EndClip();
    }

    // ---------------- Текст ----------------

    sealed class TChar
    {
        public required string Ch;
        public required Style St;
        public double? X, Y;
        public double Dx, Dy;
        public double PosX, PosY, Adv;
        public FontFace? Face;
        public ushort Gid;
        public double Size;
    }

    static void Text(XElement e, Style st, Ctx ctx)
    {
        var fonts = ctx.Fonts;
        if (fonts == null) return;
        var chars = new List<TChar>();
        bool preserve = (string?)e.Attribute(XNamespace.Xml + "space") == "preserve" || st.Get("white-space") is "pre" or "pre-wrap";
        CollectText(e, st, ctx, chars, preserve, true);
        if (!preserve)
        {
            // Крайние пробелы текста отбрасываются
            while (chars.Count > 0 && chars[0].Ch == " " && chars[0].X == null) chars.RemoveAt(0);
            while (chars.Count > 0 && chars[^1].Ch == " ") chars.RemoveAt(chars.Count - 1);
        }
        if (chars.Count == 0) return;
        // Глифы и ширины
        foreach (var t in chars)
        {
            var s = t.St;
            string family = FirstFamily(s.Get("font-family", "serif"));
            string fw = s.Get("font-weight", "normal");
            bool bold = fw is "bold" or "bolder" || int.TryParse(fw, out var wn) && wn >= 600;
            bool italic = s.Get("font-style") is "italic" or "oblique";
            int weight = int.TryParse(fw, out var wv) && wv is not (400 or 700) ? wv : 0;
            t.Size = FontSize(s);
            var face = fonts.Get(family, bold, italic, weight);
            int cp = char.ConvertToUtf32(t.Ch, 0);
            if (!face.Font.TryGetGlyph(cp, out var gid) || gid == 0)
            {
                int mapped = FontManager.MapSymbolChar(family, cp);
                if (fonts.FallbackFor(mapped, family, bold, italic, weight) is { } fb && fb.Font.TryGetGlyph(mapped, out var g2)) { face = fb; gid = g2; }
            }
            t.Face = face; t.Gid = gid;
            double ls = s.Get("letter-spacing") is { Length: > 0 } lsv && lsv != "normal" ? Len(lsv, 0, t.Size) : 0;
            double ws = t.Ch == " " && s.Get("word-spacing") is { Length: > 0 } wsv && wsv != "normal" ? Len(wsv, 0, t.Size) : 0;
            t.Adv = face.Font.Advance(gid) * t.Size + ls + ws;
        }
        // Позиции и фрагменты (новый фрагмент — с абсолютной координатой x)
        double cx = 0, cy = 0;
        var chunks = new List<(int Start, int End)>();
        int chunkStart = 0;
        for (int i = 0; i < chars.Count; i++)
        {
            var t = chars[i];
            if (t.X != null) { if (i > chunkStart) chunks.Add((chunkStart, i)); chunkStart = i; cx = t.X.Value; }
            if (t.Y != null) cy = t.Y.Value;
            cx += t.Dx; cy += t.Dy;
            t.PosX = cx; t.PosY = cy;
            cx += t.Adv;
        }
        chunks.Add((chunkStart, chars.Count));
        foreach (var (a, b) in chunks)
        {
            string anchor = chars[a].St.Get("text-anchor", "start");
            if (anchor is not ("middle" or "end")) continue;
            double w = chars[b - 1].PosX + chars[b - 1].Adv - chars[a].PosX;
            double shift = anchor == "middle" ? -w / 2 : -w;
            for (int k = a; k < b; k++) chars[k].PosX += shift;
        }
        // Вывод: подряд идущие символы одного оформления на одной линии — одним вызовом
        var c = ctx.C;
        int j = 0;
        while (j < chars.Count)
        {
            var f = chars[j];
            int k = j + 1;
            while (k < chars.Count && chars[k].Face == f.Face && chars[k].Size == f.Size && chars[k].St == f.St
                   && Math.Abs(chars[k].PosY - f.PosY) < 1e-6 && Math.Abs(chars[k].PosX - (chars[k - 1].PosX + chars[k - 1].Adv)) < 1e-6) k++;
            DrawRun(chars, j, k, ctx);
            j = k;
        }
    }

    static string FirstFamily(string list)
    {
        foreach (var raw in list.Split(','))
        {
            string f = raw.Trim().Trim('"', '\'');
            if (f.Length == 0) continue;
            return f.ToLowerInvariant() switch
            {
                "serif" => "Times New Roman",
                "sans-serif" or "system-ui" or "-apple-system" => "Arial",
                "monospace" => "Courier New",
                "cursive" or "fantasy" => "Arial",
                _ => f,
            };
        }
        return "Times New Roman";
    }

    static void CollectText(XElement e, Style st, Ctx ctx, List<TChar> output, bool preserve, bool isText)
    {
        var xs = Attr(e, "x") is { } xa ? Numbers(xa).Select(v => v).ToList() : null;
        var ys = Attr(e, "y") is { } ya ? Numbers(ya).ToList() : null;
        var dxs = Attr(e, "dx") is { } dxa ? Numbers(dxa) : null;
        var dys = Attr(e, "dy") is { } dya ? Numbers(dya) : null;
        // Единицы em/% в одиночных значениях
        if (xs is { Count: 1 }) xs[0] = Len(Attr(e, "x"), ctx.VW, FontSize(st));
        if (ys is { Count: 1 }) ys[0] = Len(Attr(e, "y"), ctx.VH, FontSize(st));
        if (dxs is { Count: 1 }) dxs[0] = Len(Attr(e, "dx"), ctx.VW, FontSize(st));
        if (dys is { Count: 1 }) dys[0] = Len(Attr(e, "dy"), ctx.VH, FontSize(st));
        int idx = 0;
        // Сдвиг базовой линии
        double fs = FontSize(st);
        double baseShift = !isText ? 0 : st.Get("dominant-baseline", st.Get("alignment-baseline")) switch
        {
            "middle" => fs * 0.33, "central" => fs * 0.36, "hanging" or "text-before-edge" or "before-edge" or "text-top" => fs * 0.78,
            "mathematical" => fs * 0.5, "text-after-edge" or "after-edge" or "ideographic" or "text-bottom" => -fs * 0.22, _ => 0,
        };
        double shift = isText ? 0 : ShiftOf(st);
        bool first = true;
        double pendingDy = 0;   // возврат на базовую линию после tspan со сдвигом
        foreach (var node in e.Nodes())
        {
            if (node is XText t)
            {
                string s = t.Value;
                if (!preserve)
                {
                    s = s.Replace("\r", "").Replace("\n", "").Replace('\t', ' ');
                    s = Regex.Replace(s, " {2,}", " ");
                    if (output.Count > 0 && output[^1].Ch == " " && s.StartsWith(' ')) s = s[1..];
                }
                else s = s.Replace("\r", "").Replace('\n', ' ').Replace('\t', ' ');
                foreach (var rune in s.EnumerateRunes())
                {
                    var ch = new TChar { Ch = rune.ToString(), St = st };
                    if (xs != null && idx < xs.Count) ch.X = xs[idx];
                    if (ys != null && idx < ys.Count) ch.Y = ys[idx] + baseShift + shift;
                    else if (first && (baseShift != 0 || shift != 0)) ch.Dy += baseShift + shift;
                    if (dxs != null && idx < dxs.Count) ch.Dx = dxs[idx];
                    if (dys != null && idx < dys.Count) ch.Dy += dys[idx];
                    ch.Dy += pendingDy; pendingDy = 0;
                    output.Add(ch);
                    idx++;
                    first = false;
                }
            }
            else if (node is XElement ce && ce.Name.LocalName is "tspan" or "textPath" or "a" or "altGlyph")
            {
                var cst = Compute(ctx.Doc, ce, st);
                if (cst.Get("display") == "none") continue;
                int before = output.Count;
                CollectText(ce, cst, ctx, output, preserve, false);
                double cs = ShiftOf(cst);
                if (cs != 0 && output.Count > before && Attr(ce, "y") == null) pendingDy -= cs;
            }
        }
    }

    /// <summary>baseline-shift (вниз — положительный).</summary>
    static double ShiftOf(Style st)
    {
        double fs = FontSize(st);
        return st.P.TryGetValue("baseline-shift", out var v) ? v switch { "super" => -fs * 0.4, "sub" => fs * 0.2, "" or "baseline" => 0, _ => -Len(v, fs, fs) } : 0;
    }

    static void DrawRun(List<TChar> chars, int a, int b, Ctx ctx)
    {
        var f = chars[a];
        var s = f.St;
        if (s.Get("visibility") is "hidden" or "collapse" || f.Face == null) return;
        var c = ctx.C;
        var fill = ResolvePaint(s.Get("fill", "black"), s, ctx);
        var stroke = ResolvePaint(s.Get("stroke", "none"), s, ctx);
        double fo = Math.Clamp(Num(s.Get("fill-opacity", "1"), 1), 0, 1) * ctx.Alpha;
        var gids = new List<ushort>(); var texts = new List<string>(); var extra = new List<double>();
        bool any = false;
        for (int i = a; i < b; i++)
        {
            gids.Add(chars[i].Gid); texts.Add(chars[i].Ch);
            double e = chars[i].Adv - chars[i].Face!.Font.Advance(chars[i].Gid) * chars[i].Size;
            extra.Add(e); any |= Math.Abs(e) > 1e-6;
        }
        Rgb col; double alpha;
        bool outline = false;
        if (fill is { } fp) { col = fp.Color; alpha = fo * fp.Alpha; }
        else if (stroke is { } sp) { col = sp.Color; alpha = Math.Clamp(Num(s.Get("stroke-opacity", "1"), 1), 0, 1) * ctx.Alpha * sp.Alpha; outline = true; }
        else return;
        if (alpha <= 0.001) return;
        if (alpha < 0.999) c.RawOps("q /" + c.GState(alpha, alpha) + " gs\n");
        c.Glyphs(f.Face, f.Size, f.PosX, f.PosY, gids, texts, col, any ? extra : null, 1, outline);
        // Подчёркивание и зачёркивание
        string deco = s.Get("text-decoration");
        if (deco.Contains("underline") || deco.Contains("line-through") || deco.Contains("overline"))
        {
            double w = chars[b - 1].PosX + chars[b - 1].Adv - f.PosX;
            double th = Math.Max(0.5, f.Size / 18);
            if (deco.Contains("underline")) c.Line(f.PosX, f.PosY + f.Size * 0.12, f.PosX + w, f.PosY + f.Size * 0.12, th, col);
            if (deco.Contains("line-through")) c.Line(f.PosX, f.PosY - f.Size * 0.28, f.PosX + w, f.PosY - f.Size * 0.28, th, col);
            if (deco.Contains("overline")) c.Line(f.PosX, f.PosY - f.Size * 0.8, f.PosX + w, f.PosY - f.Size * 0.8, th, col);
        }
        if (alpha < 0.999) c.RawOps("Q\n");
    }
}

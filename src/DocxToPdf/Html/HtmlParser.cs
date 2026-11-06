using System.Globalization;
using System.Text;

namespace DocxToPdf.Html;

/// <summary>Узел DOM-дерева.</summary>
public sealed class Node
{
    public string Tag = "";               // имя тега в нижнем регистре; "#text" — текст; "#root" — корень
    public string Text = "";
    public Dictionary<string, string> Attrs = new(StringComparer.OrdinalIgnoreCase);
    public List<Node> Children = [];
    public Node? Parent;
    public CssStyle? Style;

    public bool IsText => Tag == "#text";
    public bool IsElement => Tag.Length > 0 && Tag[0] != '#';
    /// <summary>Сгенерированный узел ::before / ::after.</summary>
    public bool IsPseudo => Tag.StartsWith("::", StringComparison.Ordinal);
    public string? Attr(string name) => Attrs.TryGetValue(name, out var v) ? v : null;
    public IEnumerable<string> Classes => (Attr("class") ?? "").Split(' ', '\t', '\n', '\r', '\f').Where(c => c.Length > 0);

    public int ElementIndex()
    {
        if (Parent == null) return 0;
        int i = 0;
        foreach (var c in Parent.Children)
        {
            if (c == this) return i;
            if (c.IsElement && !c.IsPseudo) i++;
        }
        return i;
    }

    public IEnumerable<Node> Descendants()
    {
        foreach (var c in Children)
        {
            yield return c;
            foreach (var d in c.Descendants()) yield return d;
        }
    }

    public override string ToString() => IsText ? $"\"{Text}\"" : $"<{Tag}>";
}

/// <summary>
/// Терпимый к ошибкам HTML-парсер: токенизатор + построение дерева
/// с неявным закрытием тегов (p, li, td, tr, …) как в браузерах.
/// </summary>
public static class HtmlParser
{
    static readonly HashSet<string> Void = ["area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr", "basefont", "frame", "keygen"];
    static readonly HashSet<string> RawText = ["script", "style", "textarea", "title", "xmp", "noscript", "iframe", "noembed", "noframes"];

    /// <summary>Теги, которые неявно закрывают открытый &lt;p&gt;.</summary>
    static readonly HashSet<string> ClosesP =
    [
        "address", "article", "aside", "blockquote", "center", "details", "dialog", "dir", "div", "dl", "fieldset", "figcaption", "figure",
        "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hgroup", "hr", "main", "menu", "nav", "ol", "p", "pre",
        "section", "summary", "table", "ul", "li", "dd", "dt"
    ];

    public static Node Parse(string html)
    {
        var root = new Node { Tag = "#root" };
        var stack = new List<Node> { root };
        Node Current() => stack[^1];

        void Append(Node n)
        {
            n.Parent = Current();
            Current().Children.Add(n);
        }

        void AddText(string text)
        {
            if (text.Length == 0) return;
            var cur = Current();
            if (cur.Children.Count > 0 && cur.Children[^1].IsText) cur.Children[^1].Text += text;
            else Append(new Node { Tag = "#text", Text = text });
        }

        int FindOpen(string tag, params string[] stopAt)
        {
            for (int i = stack.Count - 1; i > 0; i--)
            {
                if (stack[i].Tag == tag) return i;
                if (stopAt.Contains(stack[i].Tag)) return -1;
            }
            return -1;
        }

        void PopTo(int index) { if (index > 0) stack.RemoveRange(index, stack.Count - index); }

        void CloseImplied(string tag)
        {
            if (ClosesP.Contains(tag)) PopTo(FindOpen("p", "table", "td", "th", "li", "div", "button"));
            switch (tag)
            {
                case "li": PopTo(FindOpen("li", "ul", "ol", "menu", "table")); break;
                case "dt" or "dd":
                    int i = Math.Max(FindOpen("dt", "dl", "table"), FindOpen("dd", "dl", "table"));
                    PopTo(i); break;
                case "tr": PopTo(FindOpen("tr", "table")); break;
                case "td" or "th":
                    PopTo(Math.Max(FindOpen("td", "table", "tr"), FindOpen("th", "table", "tr"))); break;
                case "thead" or "tbody" or "tfoot":
                    PopTo(Math.Max(FindOpen("thead", "table"), Math.Max(FindOpen("tbody", "table"), FindOpen("tfoot", "table")))); break;
                case "option": PopTo(FindOpen("option", "select")); break;
                case "caption" or "colgroup":
                    PopTo(FindOpen("tr", "table")); break;
            }
            if (tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6" && Current().Tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                stack.RemoveAt(stack.Count - 1);
        }

        int pos = 0, n = html.Length;
        var text = new StringBuilder();
        while (pos < n)
        {
            char c = html[pos];
            if (c != '<' || pos + 1 >= n)
            {
                int next = html.IndexOf('<', pos + 1);
                if (next < 0) next = n;
                text.Append(html, pos, next - pos);
                pos = next;
                continue;
            }
            char c1 = html[pos + 1];
            if (html.AsSpan(pos).StartsWith("<!--"))
            {
                int end = html.IndexOf("-->", pos + 4, StringComparison.Ordinal);
                pos = end < 0 ? n : end + 3;
                continue;
            }
            if (c1 is '!' or '?')
            {
                if (html.AsSpan(pos).StartsWith("<![CDATA[", StringComparison.Ordinal))
                {
                    int end = html.IndexOf("]]>", pos, StringComparison.Ordinal);
                    if (end < 0) end = n;
                    text.Append(html, pos + 9, end - pos - 9);
                    pos = Math.Min(n, end + 3);
                    continue;
                }
                int gt = html.IndexOf('>', pos);
                pos = gt < 0 ? n : gt + 1;
                continue;
            }
            if (c1 == '/')
            {
                int p2 = pos + 2;
                int nameStart = p2;
                while (p2 < n && (char.IsLetterOrDigit(html[p2]) || html[p2] is '-' or ':' or '_')) p2++;
                string name = html[nameStart..p2].ToLowerInvariant();
                int gt = html.IndexOf('>', p2);
                pos = gt < 0 ? n : gt + 1;
                if (name.Length == 0) continue;
                AddText(Entities.Decode(text.ToString())); text.Clear();
                if (name == "br") { Append(new Node { Tag = "br" }); continue; }
                if (name == "p" && FindOpen("p") < 0) { Append(new Node { Tag = "p" }); continue; }
                int idx = FindOpen(name);
                if (idx > 0) PopTo(idx);
                continue;
            }
            if (!char.IsLetter(c1)) { text.Append(c); pos++; continue; }

            // Открывающий тег
            AddText(Entities.Decode(text.ToString())); text.Clear();
            int tagStart = pos;
            int p = pos + 1;
            int ns = p;
            while (p < n && !char.IsWhiteSpace(html[p]) && html[p] is not ('>' or '/')) p++;
            string tag = html[ns..p].ToLowerInvariant();
            var node = new Node { Tag = tag };
            bool selfClose = false;
            while (p < n)
            {
                while (p < n && char.IsWhiteSpace(html[p])) p++;
                if (p >= n) break;
                if (html[p] == '>') { p++; break; }
                if (html[p] == '/') { selfClose = true; p++; continue; }
                int an = p;
                while (p < n && !char.IsWhiteSpace(html[p]) && html[p] is not ('>' or '=' or '/')) p++;
                if (p == an) { p++; continue; }
                string attr = html[an..p].ToLowerInvariant();
                while (p < n && char.IsWhiteSpace(html[p])) p++;
                string value = "";
                if (p < n && html[p] == '=')
                {
                    p++;
                    while (p < n && char.IsWhiteSpace(html[p])) p++;
                    if (p < n && html[p] is '"' or '\'')
                    {
                        char q = html[p++];
                        int end = html.IndexOf(q, p);
                        if (end < 0) end = n;
                        value = html[p..end];
                        p = Math.Min(n, end + 1);
                    }
                    else
                    {
                        int vs = p;
                        while (p < n && !char.IsWhiteSpace(html[p]) && html[p] != '>') p++;
                        value = html[vs..p];
                    }
                }
                node.Attrs.TryAdd(attr, Entities.Decode(value));
            }
            pos = p;

            if (tag is "html" or "head" or "body")
            {
                // Атрибуты html/body сливаем в единственный узел
                var existing = stack.FirstOrDefault(s => s.Tag == tag) ?? root.Descendants().FirstOrDefault(d => d.Tag == tag);
                if (existing != null)
                {
                    foreach (var (k, v) in node.Attrs) existing.Attrs.TryAdd(k, v);
                    continue;
                }
            }

            if (tag == "svg")
            {
                // Встроенный SVG: разметка целиком (до парного </svg>) становится картинкой data:image/svg+xml
                int end = p;
                if (!selfClose)
                {
                    int depth = 1, q = p;
                    while (depth > 0)
                    {
                        int o = html.IndexOf("<svg", q, StringComparison.OrdinalIgnoreCase);
                        int c2 = html.IndexOf("</svg", q, StringComparison.OrdinalIgnoreCase);
                        if (c2 < 0) { end = n; break; }
                        if (o >= 0 && o < c2 && o + 4 < n && (char.IsWhiteSpace(html[o + 4]) || html[o + 4] is '>' or '/')) { depth++; q = o + 4; continue; }
                        depth--;
                        int gt2 = html.IndexOf('>', c2);
                        q = end = gt2 < 0 ? n : gt2 + 1;
                    }
                }
                string raw = html[tagStart..end];
                if (selfClose) raw = raw.TrimEnd().TrimEnd('>').TrimEnd('/') + "/>";
                node.Tag = "img";
                node.Attrs["data-tag"] = "svg";
                node.Attrs["src"] = "data:image/svg+xml;inline;base64," + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(raw));
                // Без размеров, но с viewBox: во всю ширину контейнера (как в браузерах)
                if (!node.Attrs.ContainsKey("width") && !node.Attrs.ContainsKey("height") && node.Attrs.ContainsKey("viewbox")) node.Attrs["width"] = "100%";
                CloseImplied("img");
                Append(node);
                pos = end;
                continue;
            }

            CloseImplied(tag);
            Append(node);

            if (RawText.Contains(tag))
            {
                int end = IndexOfCloseTag(html, tag, pos);
                string raw = html[pos..end];
                if (tag is "textarea" or "title") raw = Entities.Decode(raw);
                if (raw.Length > 0) node.Children.Add(new Node { Tag = "#text", Text = raw, Parent = node });
                int gt = html.IndexOf('>', end);
                pos = end >= n ? n : (gt < 0 ? n : gt + 1);
                continue;
            }
            if (!Void.Contains(tag) && !(selfClose && tag is "svg" or "math" or "span" or "div" or "a" or "p"))
                stack.Add(node);
        }
        AddText(Entities.Decode(text.ToString()));
        return root;
    }

    static int IndexOfCloseTag(string html, string tag, int from)
    {
        int i = from;
        while (true)
        {
            i = html.IndexOf("</", i, StringComparison.Ordinal);
            if (i < 0) return html.Length;
            if (string.Compare(html, i + 2, tag, 0, tag.Length, StringComparison.OrdinalIgnoreCase) == 0) return i;
            i += 2;
        }
    }
}

/// <summary>Декодирование HTML-сущностей.</summary>
public static class Entities
{
    static readonly Dictionary<string, string> Named = new()
    {
        ["amp"] = "&", ["lt"] = "<", ["gt"] = ">", ["quot"] = "\"", ["apos"] = "'", ["nbsp"] = " ",
        ["copy"] = "©", ["reg"] = "®", ["trade"] = "™", ["hellip"] = "…", ["mdash"] = "—", ["ndash"] = "–",
        ["laquo"] = "«", ["raquo"] = "»", ["lsaquo"] = "‹", ["rsaquo"] = "›", ["bull"] = "•", ["middot"] = "·",
        ["deg"] = "°", ["plusmn"] = "±", ["times"] = "×", ["divide"] = "÷", ["euro"] = "€", ["pound"] = "£",
        ["yen"] = "¥", ["cent"] = "¢", ["sect"] = "§", ["para"] = "¶", ["shy"] = "­", ["lsquo"] = "‘",
        ["rsquo"] = "’", ["ldquo"] = "“", ["rdquo"] = "”", ["sbquo"] = "‚", ["bdquo"] = "„", ["larr"] = "←",
        ["rarr"] = "→", ["uarr"] = "↑", ["darr"] = "↓", ["harr"] = "↔", ["lArr"] = "⇐", ["rArr"] = "⇒", ["hArr"] = "⇔",
        ["le"] = "≤", ["ge"] = "≥", ["ne"] = "≠", ["asymp"] = "≈", ["equiv"] = "≡", ["infin"] = "∞", ["minus"] = "−",
        ["sum"] = "∑", ["prod"] = "∏", ["radic"] = "√", ["int"] = "∫", ["part"] = "∂", ["nabla"] = "∇", ["isin"] = "∈",
        ["notin"] = "∉", ["cap"] = "∩", ["cup"] = "∪", ["sub"] = "⊂", ["sup"] = "⊃", ["and"] = "∧", ["or"] = "∨",
        ["forall"] = "∀", ["exist"] = "∃", ["empty"] = "∅", ["prime"] = "′", ["Prime"] = "″", ["permil"] = "‰",
        ["frac12"] = "½", ["frac14"] = "¼", ["frac34"] = "¾", ["sup1"] = "¹", ["sup2"] = "²", ["sup3"] = "³",
        ["micro"] = "µ", ["iexcl"] = "¡", ["iquest"] = "¿", ["ensp"] = " ", ["emsp"] = " ", ["thinsp"] = " ",
        ["zwnj"] = "‌", ["zwj"] = "‍", ["lrm"] = "‎", ["rlm"] = "‏", ["dagger"] = "†", ["Dagger"] = "‡",
        ["spades"] = "♠", ["clubs"] = "♣", ["hearts"] = "♥", ["diams"] = "♦", ["check"] = "✓", ["numero"] = "№",
        ["alpha"] = "α", ["beta"] = "β", ["gamma"] = "γ", ["delta"] = "δ", ["epsilon"] = "ε", ["lambda"] = "λ",
        ["mu"] = "μ", ["pi"] = "π", ["sigma"] = "σ", ["omega"] = "ω", ["theta"] = "θ", ["phi"] = "φ",
        ["Alpha"] = "Α", ["Beta"] = "Β", ["Gamma"] = "Γ", ["Delta"] = "Δ", ["Pi"] = "Π", ["Sigma"] = "Σ", ["Omega"] = "Ω",
        ["acute"] = "´", ["cedil"] = "¸", ["uml"] = "¨", ["macr"] = "¯", ["ordf"] = "ª", ["ordm"] = "º", ["not"] = "¬",
        ["curren"] = "¤", ["brvbar"] = "¦", ["agrave"] = "à", ["aacute"] = "á", ["eacute"] = "é", ["egrave"] = "è",
        ["ecirc"] = "ê", ["euml"] = "ë", ["iacute"] = "í", ["oacute"] = "ó", ["uacute"] = "ú", ["ntilde"] = "ñ",
        ["ccedil"] = "ç", ["auml"] = "ä", ["ouml"] = "ö", ["uuml"] = "ü", ["Auml"] = "Ä", ["Ouml"] = "Ö", ["Uuml"] = "Ü",
        ["szlig"] = "ß", ["oslash"] = "ø", ["aring"] = "å", ["aelig"] = "æ", ["Eacute"] = "É",
    };

    public static string Decode(string s)
    {
        if (s.IndexOf('&') < 0) return s;
        var sb = new StringBuilder(s.Length);
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (c != '&') { sb.Append(c); continue; }
            if (i + 1 < s.Length && s[i + 1] == '#')
            {
                int j = i + 2;
                bool hex = j < s.Length && s[j] is 'x' or 'X';
                if (hex) j++;
                int start = j;
                while (j < s.Length && (hex ? Uri.IsHexDigit(s[j]) : char.IsDigit(s[j]))) j++;
                if (j > start && int.TryParse(s.AsSpan(start, j - start), hex ? NumberStyles.HexNumber : NumberStyles.Integer, CultureInfo.InvariantCulture, out int cp))
                {
                    if (cp is >= 0x80 and <= 0x9F) cp = Encoding.GetEncoding(1252).GetString([(byte)cp])[0];
                    if (cp is > 0 and <= 0x10FFFF and not (>= 0xD800 and <= 0xDFFF)) sb.Append(char.ConvertFromUtf32(cp));
                    i = j < s.Length && s[j] == ';' ? j : j - 1;
                    continue;
                }
            }
            else
            {
                int j = i + 1;
                while (j < s.Length && char.IsLetterOrDigit(s[j]) && j - i < 12) j++;
                string name = s[(i + 1)..j];
                if (Named.TryGetValue(name, out var v))
                {
                    sb.Append(v);
                    i = j < s.Length && s[j] == ';' ? j : j - 1;
                    continue;
                }
            }
            sb.Append(c);
        }
        return sb.ToString();
    }
}

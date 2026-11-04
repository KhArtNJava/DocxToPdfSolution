using System.Globalization;
using System.Text;
using DocxToPdf.Docx;

namespace DocxToPdf.Html;

/// <summary>Одно объявление CSS.</summary>
public sealed record CssDecl(string Property, string Value, bool Important);

/// <summary>Простой селектор (тег, #id, .класс, [атрибут], :first-child) с комбинатором слева.</summary>
public sealed class CompoundSelector
{
    public string? Tag;
    public string? Id;
    public List<string> Classes = [];
    public List<(string Name, string Op, string? Value)> AttrConds = [];
    /// <summary>Псевдоклассы (:first-child, :nth-child(…), :not(…) …) как условия.</summary>
    public List<Func<Node, bool>> Conditions = [];
    public string? PseudoElement;           // before | after
    public bool Unsupported;
    public int ExtraSpecificity;            // вклад псевдоклассов
    public char Combinator = ' '; // связь с предыдущим (левым) селектором: ' ' — потомок, '>' — ребёнок, '+' — соседний
}

public sealed class Selector
{
    public List<CompoundSelector> Parts = [];
    public int Specificity;
    public string? PseudoElement => Parts.Count > 0 ? Parts[^1].PseudoElement : null;

    public bool Matches(Node n) => MatchFrom(n, Parts.Count - 1);

    bool MatchFrom(Node? n, int idx)
    {
        if (n == null || !MatchCompound(n, Parts[idx])) return false;
        if (idx == 0) return true;
        char comb = Parts[idx].Combinator;
        switch (comb)
        {
            case '>': return MatchFrom(ParentElement(n), idx - 1);
            case '+': return MatchFrom(PrevElement(n), idx - 1);
            case '~':
                for (var s = PrevElement(n); s != null; s = PrevElement(s))
                    if (MatchFrom(s, idx - 1)) return true;
                return false;
            default:
                for (var a = ParentElement(n); a != null; a = ParentElement(a))
                    if (MatchFrom(a, idx - 1)) return true;
                return false;
        }
    }

    static Node? ParentElement(Node n) => n.Parent is { IsElement: true } p ? p : null;

    /// <summary>Соседи-элементы (без текста и сгенерированных ::before/::after).</summary>
    public static List<Node> Siblings(Node n) =>
        n.Parent == null ? [n] : n.Parent.Children.Where(c => c.IsElement && !c.IsPseudo).ToList();

    public static Node? PrevElement(Node n)
    {
        var sib = Siblings(n);
        int i = sib.IndexOf(n);
        return i > 0 ? sib[i - 1] : null;
    }

    static bool MatchCompound(Node n, CompoundSelector c)
    {
        if (c.Unsupported || !n.IsElement || n.IsPseudo) return false;
        if (c.Tag != null && c.Tag != "*" && c.Tag != n.Tag) return false;
        if (c.Id != null && n.Attr("id") != c.Id) return false;
        if (c.Classes.Count > 0)
        {
            var cls = n.Classes.ToHashSet();
            if (!c.Classes.All(cls.Contains)) return false;
        }
        foreach (var (name, op, value) in c.AttrConds)
        {
            var v = n.Attr(name);
            if (v == null) return false;
            if (value == null) continue;
            bool ok = op switch
            {
                "~=" => v.Split(' ', StringSplitOptions.RemoveEmptyEntries).Contains(value),
                "|=" => v == value || v.StartsWith(value + "-"),
                "^=" => value.Length > 0 && v.StartsWith(value),
                "$=" => value.Length > 0 && v.EndsWith(value),
                "*=" => value.Length > 0 && v.Contains(value),
                _ => v == value
            };
            if (!ok) return false;
        }
        foreach (var cond in c.Conditions)
            if (!cond(n)) return false;
        return true;
    }
}

public sealed class CssRule
{
    public List<Selector> Selectors = [];
    public List<CssDecl> Decls = [];
    public int Order;
}

/// <summary>Таблица стилей: правила + свойства @page.</summary>
public sealed class StyleSheet
{
    public List<CssRule> Rules = [];
    public List<CssDecl> PageDecls = [];
    int _order;

    public void Parse(string css)
    {
        css = StripComments(css);
        int i = 0;
        ParseBlock(css, ref i, topLevel: true);
    }

    static string StripComments(string css)
    {
        var sb = new StringBuilder(css.Length);
        for (int i = 0; i < css.Length; i++)
        {
            if (css[i] == '/' && i + 1 < css.Length && css[i + 1] == '*')
            {
                int end = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = end < 0 ? css.Length : end + 1;
                continue;
            }
            sb.Append(css[i]);
        }
        return sb.ToString();
    }

    void ParseBlock(string css, ref int i, bool topLevel)
    {
        while (i < css.Length)
        {
            while (i < css.Length && (char.IsWhiteSpace(css[i]) || css[i] == ';')) i++;
            if (i >= css.Length) return;
            if (css[i] == '}') { i++; if (!topLevel) return; continue; }
            if (css[i] == '@')
            {
                int brace = css.IndexOf('{', i), semi = css.IndexOf(';', i);
                if (semi >= 0 && (brace < 0 || semi < brace)) { i = semi + 1; continue; } // @import, @charset
                if (brace < 0) return;
                string at = css[i..brace].Trim().ToLowerInvariant();
                i = brace + 1;
                if (at.StartsWith("@page"))
                {
                    int end = FindClose(css, i);
                    PageDecls.AddRange(ParseDeclarations(css[i..end]));
                    i = end + 1;
                }
                else if (at.StartsWith("@media") && (at.Contains("print") || at.Contains("all")) && !at.Contains("not print"))
                    ParseBlock(css, ref i, topLevel: false);   // правила для печати применяем
                else i = FindClose(css, i) + 1;
                continue;
            }
            int open = css.IndexOf('{', i);
            if (open < 0) return;
            string selText = css[i..open];
            int close = FindClose(css, open + 1);
            var decls = ParseDeclarations(css[(open + 1)..close]);
            i = close + 1;
            var rule = new CssRule { Decls = decls, Order = _order++ };
            foreach (var s in SplitTopLevel(selText, ','))
            {
                var sel = ParseSelector(s.Trim());
                if (sel != null) rule.Selectors.Add(sel);
            }
            if (rule.Selectors.Count > 0 && decls.Count > 0) Rules.Add(rule);
        }
    }

    static int FindClose(string css, int from)
    {
        int depth = 1;
        for (int i = from; i < css.Length; i++)
        {
            if (css[i] == '{') depth++;
            else if (css[i] == '}' && --depth == 0) return i;
        }
        return css.Length;
    }

    public static List<CssDecl> ParseDeclarations(string text)
    {
        var list = new List<CssDecl>();
        foreach (var part in SplitTopLevel(text, ';'))
        {
            int colon = part.IndexOf(':');
            if (colon <= 0) continue;
            string prop = part[..colon].Trim().ToLowerInvariant();
            string value = part[(colon + 1)..].Trim();
            bool important = false;
            int imp = value.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
            if (imp >= 0) { important = true; value = value[..imp].Trim(); }
            if (prop.Length > 0 && value.Length > 0) list.Add(new CssDecl(prop, value, important));
        }
        return list;
    }

    static IEnumerable<string> SplitTopLevel(string s, char sep)
    {
        int depth = 0, start = 0;
        char quote = '\0';
        for (int i = 0; i < s.Length; i++)
        {
            char c = s[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c == '(') depth++;
            else if (c == ')') depth--;
            else if (c == sep && depth == 0) { yield return s[start..i]; start = i + 1; }
        }
        yield return s[start..];
    }

    public static Selector? ParseSelector(string s)
    {
        s = s.Trim();
        if (s.Length == 0) return null;
        var sel = new Selector();
        int i = 0;
        char comb = ' ';
        while (i < s.Length)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i >= s.Length) break;
            if (s[i] is '>' or '+' or '~') { comb = s[i]; i++; continue; }
            var c = new CompoundSelector { Combinator = comb };
            comb = ' ';
            while (i < s.Length && !char.IsWhiteSpace(s[i]) && s[i] is not ('>' or '+' or '~'))
            {
                char ch = s[i];
                if (ch == '#') { i++; c.Id = ReadIdent(s, ref i); }
                else if (ch == '.') { i++; c.Classes.Add(ReadIdent(s, ref i)); }
                else if (ch == '[')
                {
                    int end = s.IndexOf(']', i);
                    if (end < 0) end = s.Length - 1;
                    string body = s[(i + 1)..end];
                    int eq = body.IndexOf('=');
                    if (eq < 0) c.AttrConds.Add((body.Trim().ToLowerInvariant(), "", null));
                    else
                    {
                        string op = eq > 0 && body[eq - 1] is '~' or '|' or '^' or '$' or '*' ? body.Substring(eq - 1, 2) : "=";
                        string name = body[..(op == "=" ? eq : eq - 1)].Trim().ToLowerInvariant();
                        string val = body[(eq + 1)..].Trim();
                        if (val.EndsWith(" i") || val.EndsWith(" s")) val = val[..^2].Trim();
                        c.AttrConds.Add((name, op, val.Trim('"', '\'')));
                    }
                    i = end + 1;
                }
                else if (ch == ':')
                {
                    i++;
                    bool element = false;
                    if (i < s.Length && s[i] == ':') { element = true; i++; }
                    string name = ReadIdent(s, ref i).ToLowerInvariant();
                    string? arg = null;
                    if (i < s.Length && s[i] == '(')
                    {
                        int depth = 0, st = i + 1;
                        for (; i < s.Length; i++)
                        {
                            if (s[i] == '(') depth++;
                            else if (s[i] == ')' && --depth == 0) break;
                        }
                        arg = s[st..Math.Min(i, s.Length)];
                        i++;
                    }
                    if (element || name is "before" or "after" or "first-line" or "first-letter")
                    {
                        if (name is "before" or "after" or "marker" or "first-letter" or "first-line") c.PseudoElement = name;
                        else c.Unsupported = true;   // ::first-letter, ::marker, ::selection …
                    }
                    else AddPseudoClass(c, name, arg);
                }
                else if (ch == '*') { c.Tag = "*"; i++; }
                else c.Tag = ReadIdent(s, ref i).ToLowerInvariant();
            }
            sel.Parts.Add(c);
        }
        if (sel.Parts.Count == 0) return null;
        int a2 = 0, b2 = 0, d2 = 0, extra = 0;
        foreach (var p in sel.Parts)
        {
            if (p.Id != null) a2++;
            b2 += p.Classes.Count + p.AttrConds.Count;
            if (p.Tag is not (null or "*")) d2++;
            if (p.PseudoElement != null) d2++;
            extra += p.ExtraSpecificity;
        }
        sel.Specificity = a2 * 10000 + b2 * 100 + d2 + extra;
        return sel;
    }

    static List<Selector> ParseSelectorList(string? s) =>
        s == null ? [] : SplitTopLevel(s, ',').Select(ParseSelector).Where(x => x != null).Cast<Selector>().ToList();

    /// <summary>Разбор "an+b" (odd, even, 3, 2n+1, -n+3 …).</summary>
    static (int A, int B)? ParseNth(string? arg)
    {
        if (arg == null) return null;
        string t = arg.Replace(" ", "").ToLowerInvariant();
        int of = t.IndexOf("of", StringComparison.Ordinal);
        if (of > 0) t = t[..of];
        if (t == "odd") return (2, 1);
        if (t == "even") return (2, 0);
        int n = t.IndexOf('n');
        if (n < 0) return int.TryParse(t, out var b0) ? (0, b0) : null;
        string ap = t[..n];
        int a = ap switch { "" or "+" => 1, "-" => -1, _ => int.TryParse(ap, out var av) ? av : 1 };
        string bp = t[(n + 1)..];
        int b = bp.Length == 0 ? 0 : int.TryParse(bp, out var bv) ? bv : 0;
        return (a, b);
    }

    static bool NthMatch((int A, int B) f, int k) =>
        f.A == 0 ? k == f.B : (k - f.B) % f.A == 0 && (k - f.B) / f.A >= 0;

    static void AddPseudoClass(CompoundSelector c, string name, string? arg)
    {
        void Add(Func<Node, bool> f, int spec = 100) { c.Conditions.Add(f); c.ExtraSpecificity += spec; }
        static List<Node> OfType(Node n) => Selector.Siblings(n).Where(x => x.Tag == n.Tag).ToList();

        switch (name)
        {
            case "first-child": Add(n => Selector.Siblings(n).FirstOrDefault() == n); break;
            case "last-child": Add(n => Selector.Siblings(n).LastOrDefault() == n); break;
            case "only-child": Add(n => Selector.Siblings(n).Count == 1); break;
            case "first-of-type": Add(n => OfType(n).FirstOrDefault() == n); break;
            case "last-of-type": Add(n => OfType(n).LastOrDefault() == n); break;
            case "only-of-type": Add(n => OfType(n).Count == 1); break;
            case "nth-child" or "nth-last-child" or "nth-of-type" or "nth-last-of-type":
            {
                var f = ParseNth(arg);
                if (f == null) { c.Unsupported = true; break; }
                bool last = name.Contains("last"), ofType = name.EndsWith("of-type");
                Add(n =>
                {
                    var list = ofType ? OfType(n) : Selector.Siblings(n);
                    int idx = list.IndexOf(n);
                    if (idx < 0) return false;
                    int k = last ? list.Count - idx : idx + 1;
                    return NthMatch(f.Value, k);
                });
                break;
            }
            case "root": Add(n => n.Tag == "html" || n.Parent?.Tag == "#root" && n.Parent.Children.Count(x => x.IsElement) == 1); break;
            case "empty": Add(n => n.Children.All(x => x.IsPseudo || (x.IsText && x.Text.Length == 0))); break;
            case "link" or "any-link": Add(n => n.Tag is "a" or "area" && n.Attr("href") != null); break;
            case "checked": Add(n => n.Attr("checked") != null || n.Attr("selected") != null); break;
            case "disabled": Add(n => n.Attr("disabled") != null); break;
            case "enabled": Add(n => n.Tag is "input" or "button" or "select" or "textarea" && n.Attr("disabled") == null); break;
            case "required": Add(n => n.Attr("required") != null); break;
            case "optional": Add(n => n.Tag is "input" or "select" or "textarea" && n.Attr("required") == null); break;
            case "read-only": Add(n => n.Attr("readonly") != null); break;
            case "lang":
            {
                string lang = (arg ?? "").Trim().Trim('"', '\'').ToLowerInvariant();
                Add(n =>
                {
                    for (var x = n; x != null; x = x.Parent)
                        if (x.Attr("lang") is { } l) return l.ToLowerInvariant() == lang || l.ToLowerInvariant().StartsWith(lang + "-");
                    return false;
                });
                break;
            }
            case "not":
            {
                var list = ParseSelectorList(arg);
                if (list.Count == 0) { c.Unsupported = true; break; }
                Add(n => !list.Any(sel => sel.Matches(n)), list.Max(x => x.Specificity));
                break;
            }
            case "is" or "matches" or "-webkit-any" or "-moz-any" or "where":
            {
                var list = ParseSelectorList(arg);
                if (list.Count == 0) { c.Unsupported = true; break; }
                Add(n => list.Any(sel => sel.Matches(n)), name == "where" ? 0 : list.Max(x => x.Specificity));
                break;
            }
            case "has":
            {
                // Упрощённо: относительные селекторы ("> a", "+ b") трактуются как поиск среди потомков/соседей
                string t = (arg ?? "").Trim();
                char rel = t.Length > 0 && t[0] is '>' or '+' or '~' ? t[0] : ' ';
                var list = ParseSelectorList(rel == ' ' ? t : t[1..]);
                if (list.Count == 0) { c.Unsupported = true; break; }
                Add(n =>
                {
                    IEnumerable<Node> cand = rel switch
                    {
                        '>' => n.Children.Where(x => x.IsElement),
                        '+' => Selector.Siblings(n).SkipWhile(x => x != n).Skip(1).Take(1),
                        '~' => Selector.Siblings(n).SkipWhile(x => x != n).Skip(1),
                        _ => n.Descendants().Where(x => x.IsElement)
                    };
                    return cand.Any(x => list.Any(sel => sel.Matches(x)));
                }, list.Max(x => x.Specificity));
                break;
            }
            case "scope": Add(n => n.Tag == "html"); break;
            default: c.Unsupported = true; break;   // :hover, :focus, :visited … в печати не выполняются
        }
    }

    static string ReadIdent(string s, ref int i)
    {
        int st = i;
        while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] is '-' or '_' || s[i] > 127)) i++;
        if (i == st && i < s.Length) i++;
        return s[st..i];
    }
}

/// <summary>Граница блока/ячейки.</summary>
public sealed record CssBorder(double Width, string Style, Rgb Color)
{
    public BorderDef? ToBorderDef() =>
        Width <= 0 || Style is "none" or "hidden" ? null : new BorderDef { Val = Style switch { "dashed" => "dashed", "dotted" => "dotted", "double" => "double", _ => "single" }, Width = Width, Color = Color };
}

/// <summary>Вычисленный стиль элемента.</summary>
public sealed class CssStyle
{
    // Наследуемые
    public string FontFamily = "Times New Roman";
    public double FontSize = 12;
    public bool Bold, Italic, Underline, Strike, Upper, Lower, SmallCaps;
    public Rgb Color = Rgb.Black;
    public string TextAlign = "left";
    public double? LineHeightFactor;
    public double? LineHeightPt;
    public string WhiteSpace = "normal";
    public double TextIndent;
    public string ListStyleType = "disc";
    public string? Link;
    public int VertAlign;
    public Rgb? TextBackground;
    public bool Hidden;     // visibility: hidden
    public double LetterSpacing, WordSpacing;
    public Dictionary<string, string>? Vars;   // пользовательские свойства (--x), наследуются
    public InlineBox? InlineBox;               // ближайший строчный блок с рамкой/отступами
    public string? Quotes;

    // Ненаследуемые
    public string Display = "inline";
    public double MarginTop, MarginBottom, MarginLeft, MarginRight;
    public bool MarginLeftAuto, MarginRightAuto;
    public double PaddingTop, PaddingBottom, PaddingLeft, PaddingRight;
    public Rgb? Background;
    public CssBorder? BorderTop, BorderBottom, BorderLeft, BorderRight;
    public string? Width, Height;      // исходные значения (могут быть в %)
    public bool PageBreakBefore, PageBreakAfter, AvoidBreakInside;
    public string? VerticalAlign;
    public string BoxSizing = "content-box";

    // Позиционирование и float
    public string? Float, Clear;
    public string Position = "static";
    public string? Left, Top, Right, Bottom;
    public int ZIndex;

    // Flexbox
    public string FlexDirection = "row";
    public bool FlexWrap;
    public string JustifyContent = "flex-start";
    public string AlignItems = "stretch";
    public string? AlignSelf;
    public double FlexGrow, FlexShrink = 1;
    public string? FlexBasis;
    public int Order;
    public string? RowGap, ColumnGap;

    // Grid
    public string? GridTemplateColumns, GridTemplateRows;
    /// <summary>grid-template-areas: строки шаблона с именами областей.</summary>
    public string? GridTemplateAreas;
    /// <summary>align-content: распределение строк сетки / строк flex-контейнера по высоте контейнера.</summary>
    public string AlignContent = "normal";
    public string? GridColumnStart, GridColumnEnd, GridRowStart, GridRowEnd;

    // Скругление углов (TL, TR, BR, BL — исходные значения, % — от ширины)
    // rx TL, TR, BR, BL, затем ry TL, TR, BR, BL
    public string?[] Radius = new string?[8];
    public bool HasRadius => Radius.Any(r => r != null && r != "0" && r != "0px");
    // initial-letter: размер в строках и «погружение»
    public string? InitialLetter;
    public double? ExactLineHeight;

    // Преобразования
    public string? Transform, TransformOrigin, Translate, Rotate, Scale;
    public string? Perspective, PerspectiveOrigin;              // задаются у родителя, действуют на детей
    public string? ParentPerspective, ParentPerspectiveOrigin;  // перспектива, унаследованная от родителя
    public bool BackfaceHidden;
    public bool Preserve3D;                                     // transform-style: preserve-3d
    public bool ParentPreserve3D;                               // родитель — 3D-сцена

    // Генерируемое содержимое и счётчики
    public string? Content;
    public string? CounterReset, CounterIncrement, CounterSet;

    public CssStyle Clone() => (CssStyle)MemberwiseClone();

    public CssStyle InheritFrom()
    {
        return new CssStyle
        {
            FontFamily = FontFamily, FontSize = FontSize, Bold = Bold, Italic = Italic, Underline = Underline, Strike = Strike,
            Upper = Upper, Lower = Lower, SmallCaps = SmallCaps, LetterSpacing = LetterSpacing, WordSpacing = WordSpacing, Color = Color, TextAlign = TextAlign, LineHeightFactor = LineHeightFactor, LineHeightPt = LineHeightPt,
            WhiteSpace = WhiteSpace, TextIndent = TextIndent, ListStyleType = ListStyleType, Link = Link, VertAlign = VertAlign,
            TextBackground = TextBackground, Hidden = Hidden, Vars = Vars, Quotes = Quotes, InlineBox = InlineBox,
        };
    }
}

/// <summary>Разбор значений CSS: длины, цвета.</summary>
public static class CssValues
{
    public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

    /// <summary>Длина в пунктах. em — относительно fontSize, % — относительно percentBase.</summary>
    public static double? Length(string? v, double fontSize, double percentBase = 0)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        v = v.Trim().ToLowerInvariant();
        if (v is "0" or "auto" or "none" or "normal" or "initial" or "inherit") return v == "0" ? 0 : null;
        if (v.Contains('(') && (v.StartsWith("calc(") || v.StartsWith("min(") || v.StartsWith("max(") || v.StartsWith("clamp(")
            || v.StartsWith("-webkit-calc(")))
            return Calc(v.StartsWith("-webkit-") ? v[8..] : v, fontSize, percentBase);
        if (v is "thin") return 0.75;
        if (v is "medium") return 2.25;
        if (v is "thick") return 3.75;
        int k = 0;
        while (k < v.Length && (char.IsDigit(v[k]) || v[k] is '.' or '-' or '+')) k++;
        if (k == 0 || !double.TryParse(v[..k], NumberStyles.Float, Inv, out double num)) return null;
        string unit = v[k..].Trim();
        return unit switch
        {
            "px" or "" => num * 0.75,
            "pt" => num,
            "pc" => num * 12,
            "in" => num * 72,
            "cm" => num * 72 / 2.54,
            "mm" => num * 72 / 25.4,
            "q" => num * 72 / 101.6,
            "em" => num * fontSize,
            "rem" => num * 12,
            "ex" or "ch" => num * fontSize * 0.5,
            "%" => num / 100 * percentBase,
            "vw" or "vh" => num / 100 * 500,
            _ => null
        };
    }

    // ---------------------------------------------------------------
    //  calc() / min() / max() / clamp()
    // ---------------------------------------------------------------

    /// <summary>Вычисляет выражение CSS с единицами измерения; результат в пунктах.</summary>
    public static double? Calc(string expr, double fontSize, double percentBase)
    {
        var p = new CalcParser(expr, fontSize, percentBase);
        try
        {
            var r = p.ParseExpr();
            p.SkipWs();
            return p.AtEnd ? r.Value : null;
        }
        catch { return null; }
    }

    readonly record struct CalcValue(double Value, bool HasUnit);

    sealed class CalcParser(string s, double fontSize, double percentBase)
    {
        int _i;
        public bool AtEnd => _i >= s.Length;
        public void SkipWs() { while (_i < s.Length && char.IsWhiteSpace(s[_i])) _i++; }

        public CalcValue ParseExpr()
        {
            var left = ParseTerm();
            while (true)
            {
                SkipWs();
                if (_i >= s.Length || s[_i] is not ('+' or '-')) return left;
                char op = s[_i++];
                var right = ParseTerm();
                left = new CalcValue(op == '+' ? left.Value + right.Value : left.Value - right.Value, left.HasUnit || right.HasUnit);
            }
        }

        CalcValue ParseTerm()
        {
            var left = ParseFactor();
            while (true)
            {
                SkipWs();
                if (_i >= s.Length || s[_i] is not ('*' or '/')) return left;
                char op = s[_i++];
                var right = ParseFactor();
                if (op == '*') left = new CalcValue(left.Value * right.Value, left.HasUnit || right.HasUnit);
                else left = new CalcValue(right.Value == 0 ? 0 : left.Value / right.Value, left.HasUnit);
            }
        }

        List<CalcValue> Args()
        {
            var list = new List<CalcValue>();
            while (true)
            {
                list.Add(ParseExpr());
                SkipWs();
                if (_i < s.Length && s[_i] == ',') { _i++; continue; }
                if (_i < s.Length && s[_i] == ')') { _i++; return list; }
                throw new FormatException();
            }
        }

        CalcValue ParseFactor()
        {
            SkipWs();
            if (_i >= s.Length) throw new FormatException();
            if (s[_i] == '(') { _i++; var v = ParseExpr(); SkipWs(); _i++; return v; }
            if (s[_i] == '-' && _i + 1 < s.Length && (s[_i + 1] == '(' || char.IsLetter(s[_i + 1])))
            {
                _i++;
                var v = ParseFactor();
                return v with { Value = -v.Value };
            }
            if (char.IsLetter(s[_i]))
            {
                int st = _i;
                while (_i < s.Length && (char.IsLetter(s[_i]) || s[_i] == '-')) _i++;
                string fn = s[st.._i];
                if (_i < s.Length && s[_i] == '(')
                {
                    _i++;
                    if (fn == "calc") { var v = ParseExpr(); SkipWs(); _i++; return v; }
                    var args = Args();
                    return fn switch
                    {
                        "min" => args.MinBy(a => a.Value),
                        "max" => args.MaxBy(a => a.Value),
                        "clamp" when args.Count == 3 => new CalcValue(Math.Max(args[0].Value, Math.Min(args[1].Value, args[2].Value)), true),
                        _ => throw new FormatException()
                    };
                }
                throw new FormatException();
            }
            int ns = _i;
            if (s[_i] is '+' or '-') _i++;
            while (_i < s.Length && (char.IsDigit(s[_i]) || s[_i] == '.')) _i++;
            if (_i < s.Length && s[_i] is 'e' or 'E' && _i + 1 < s.Length && (char.IsDigit(s[_i + 1]) || s[_i + 1] is '-' or '+'))
            {
                _i += 2;
                while (_i < s.Length && char.IsDigit(s[_i])) _i++;
            }
            double num = double.Parse(s[ns.._i], NumberStyles.Float, Inv);
            int us = _i;
            while (_i < s.Length && (char.IsLetter(s[_i]) || s[_i] == '%')) _i++;
            string unit = s[us.._i];
            if (unit.Length == 0) return new CalcValue(num, false);
            double? v2 = Length(num.ToString("R", Inv) + unit, fontSize, percentBase);
            if (v2 == null) throw new FormatException();
            return new CalcValue(v2.Value, true);
        }
    }

    // ---------------------------------------------------------------
    //  Переменные var(--x, fallback)
    // ---------------------------------------------------------------

    /// <summary>Подставляет значения пользовательских свойств; null — значение недействительно.</summary>
    public static string? SubstituteVars(string value, IReadOnlyDictionary<string, string>? vars, int depth = 0)
    {
        if (depth > 20) return null;
        int idx = value.IndexOf("var(", StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return value;
        var sb = new StringBuilder();
        int pos = 0;
        while (idx >= 0)
        {
            sb.Append(value, pos, idx - pos);
            int st = idx + 4, d = 1, i = st;
            for (; i < value.Length && d > 0; i++)
            {
                if (value[i] == '(') d++;
                else if (value[i] == ')') d--;
            }
            if (d != 0) return null;
            string inner = value[st..(i - 1)];
            int comma = -1, dd = 0;
            for (int k = 0; k < inner.Length; k++)
            {
                if (inner[k] == '(') dd++;
                else if (inner[k] == ')') dd--;
                else if (inner[k] == ',' && dd == 0) { comma = k; break; }
            }
            string name = (comma < 0 ? inner : inner[..comma]).Trim().ToLowerInvariant();
            string? fallback = comma < 0 ? null : inner[(comma + 1)..].Trim();
            string? replacement = null;
            if (vars != null && vars.TryGetValue(name, out var v)) replacement = SubstituteVars(v, vars, depth + 1);
            if (replacement == null && fallback != null) replacement = SubstituteVars(fallback, vars, depth + 1);
            if (replacement == null) return null;
            sb.Append(replacement);
            pos = i;
            idx = value.IndexOf("var(", pos, StringComparison.OrdinalIgnoreCase);
        }
        sb.Append(value, pos, value.Length - pos);
        return sb.ToString();
    }

    static readonly Dictionary<string, string> NamedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        ["black"] = "000000", ["white"] = "ffffff", ["red"] = "ff0000", ["green"] = "008000", ["blue"] = "0000ff",
        ["yellow"] = "ffff00", ["cyan"] = "00ffff", ["aqua"] = "00ffff", ["magenta"] = "ff00ff", ["fuchsia"] = "ff00ff",
        ["gray"] = "808080", ["grey"] = "808080", ["silver"] = "c0c0c0", ["maroon"] = "800000", ["olive"] = "808000",
        ["lime"] = "00ff00", ["teal"] = "008080", ["navy"] = "000080", ["purple"] = "800080", ["orange"] = "ffa500",
        ["pink"] = "ffc0cb", ["brown"] = "a52a2a", ["gold"] = "ffd700", ["darkgray"] = "a9a9a9", ["darkgrey"] = "a9a9a9",
        ["lightgray"] = "d3d3d3", ["lightgrey"] = "d3d3d3", ["gainsboro"] = "dcdcdc", ["whitesmoke"] = "f5f5f5",
        ["darkred"] = "8b0000", ["darkgreen"] = "006400", ["darkblue"] = "00008b", ["lightblue"] = "add8e6",
        ["lightgreen"] = "90ee90", ["lightyellow"] = "ffffe0", ["beige"] = "f5f5dc", ["ivory"] = "fffff0",
        ["indigo"] = "4b0082", ["violet"] = "ee82ee", ["coral"] = "ff7f50", ["tomato"] = "ff6347", ["salmon"] = "fa8072",
        ["khaki"] = "f0e68c", ["crimson"] = "dc143c", ["steelblue"] = "4682b4", ["royalblue"] = "4169e1",
        ["skyblue"] = "87ceeb", ["dodgerblue"] = "1e90ff", ["slategray"] = "708090", ["darkslategray"] = "2f4f4f",
        ["dimgray"] = "696969", ["lavender"] = "e6e6fa", ["linen"] = "faf0e6", ["honeydew"] = "f0fff0",
        ["aliceblue"] = "f0f8ff", ["mintcream"] = "f5fffa", ["seashell"] = "fff5ee", ["wheat"] = "f5deb3",
        ["tan"] = "d2b48c", ["chocolate"] = "d2691e", ["firebrick"] = "b22222", ["orangered"] = "ff4500",
        ["darkorange"] = "ff8c00", ["seagreen"] = "2e8b57", ["forestgreen"] = "228b22", ["limegreen"] = "32cd32",
        ["midnightblue"] = "191970", ["cornflowerblue"] = "6495ed", ["cadetblue"] = "5f9ea0", ["turquoise"] = "40e0d0",
        ["orchid"] = "da70d6", ["plum"] = "dda0dd", ["peru"] = "cd853f", ["sienna"] = "a0522d", ["snow"] = "fffafa",
    };

    /// <summary>Цвет; null — прозрачный/не распознан.</summary>
    public static Rgb? Color(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        v = v.Trim();
        if (v.StartsWith('#'))
        {
            string h = v[1..];
            if (h.Length is 3 or 4) h = string.Concat(h.Take(3).Select(c => new string(c, 2)));
            if (h.Length == 8) h = h[..6];
            return Rgb.Parse(h);
        }
        string lower = v.ToLowerInvariant();
        if (lower.StartsWith("rgb"))
        {
            int a = lower.IndexOf('('), b = lower.LastIndexOf(')');
            if (a < 0 || b < a) return null;
            var parts = lower[(a + 1)..b].Replace("/", " ").Split([',', ' '], StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3) return null;
            if (parts.Length >= 4 && double.TryParse(parts[3].TrimEnd('%'), NumberStyles.Float, Inv, out var alpha)
                && (parts[3].EndsWith('%') ? alpha / 100 : alpha) < 0.1) return null;
            double C(string s) => s.EndsWith('%')
                ? double.Parse(s[..^1], Inv) / 100
                : double.Parse(s, Inv) / 255;
            try { return new Rgb(Math.Clamp(C(parts[0]), 0, 1), Math.Clamp(C(parts[1]), 0, 1), Math.Clamp(C(parts[2]), 0, 1)); }
            catch { return null; }
        }
        if (lower is "transparent" or "none" or "inherit" or "initial") return null;
        return NamedColors.TryGetValue(lower, out var hex) ? Rgb.Parse(hex) : null;
    }

    /// <summary>Разбивает значение на токены по пробелам (с учётом скобок).</summary>
    public static List<string> Tokens(string v)
    {
        var list = new List<string>();
        var sb = new StringBuilder();
        int depth = 0;
        foreach (char c in v)
        {
            if (c == '(') depth++;
            if (c == ')') depth--;
            if (char.IsWhiteSpace(c) && depth == 0) { if (sb.Length > 0) { list.Add(sb.ToString()); sb.Clear(); } }
            else sb.Append(c);
        }
        if (sb.Length > 0) list.Add(sb.ToString());
        return list;
    }

    /// <summary>Раскладывает сокращённую запись (1–4 значения) на top/right/bottom/left.</summary>
    public static (string T, string R, string B, string L) Box(string v)
    {
        var t = Tokens(v);
        return t.Count switch
        {
            1 => (t[0], t[0], t[0], t[0]),
            2 => (t[0], t[1], t[0], t[1]),
            3 => (t[0], t[1], t[2], t[1]),
            >= 4 => (t[0], t[1], t[2], t[3]),
            _ => ("0", "0", "0", "0")
        };
    }

    static readonly HashSet<string> BorderStyles = ["none", "hidden", "dotted", "dashed", "solid", "double", "groove", "ridge", "inset", "outset"];

    public static CssBorder? Border(string v, double fontSize, CssBorder? current)
    {
        double width = current?.Width ?? 2.25;
        string style = current?.Style ?? "none";
        Rgb color = current?.Color ?? Rgb.Black;
        foreach (var tok in Tokens(v))
        {
            string t = tok.ToLowerInvariant();
            if (BorderStyles.Contains(t)) style = t;
            else if (Length(t, fontSize) is double w) width = w;
            else if (Color(t) is Rgb c) color = c;
        }
        return new CssBorder(width, style, color);
    }
}

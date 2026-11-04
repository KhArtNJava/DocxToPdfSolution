using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Pdf;

namespace DocxToPdf.Layout;

/// <summary>Свёрстанная формула: операции рисования относительно (x, базовая линия).</summary>
public sealed class MathFrag : Frag
{
    public double Width, Ascent, Descent;
    internal readonly List<Action<PdfCanvas, double, double>> Ops = [];

    /// <summary>(x, y) — левый верхний угол; базовая линия — y + Ascent.</summary>
    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var op in Ops) op(c, x, y + Ascent);
    }
}

/// <summary>
/// Двумерная вёрстка формул OMML: дроби, индексы (в т. ч. предварительные), радикалы, n-арные операторы с пределами,
/// растягиваемые скобки, матрицы, функции, акценты, черты, групповые знаки, пределы, системы уравнений, рамки.
/// </summary>
public sealed class MathLayout(LayoutContext ctx, TextStyle style)
{
    const string MathFont = "Cambria Math";
    readonly LayoutContext _ctx = ctx;
    readonly TextStyle _style = style;
    Rgb Color => _style.Color;

    /// <summary>Блок формулы: ширина, высота над базовой линией и глубина под ней; операции рисования от (x, базовая линия).</summary>
    sealed class Box
    {
        public double W, A, D;
        public readonly List<Action<PdfCanvas, double, double>> Ops = [];
        /// <summary>Тип для межэлементных интервалов: 0 — обычный, 1 — бинарный оператор, 2 — отношение, 3 — пунктуация, 4 — открывающая скобка.</summary>
        public int Kind;
        public void Put(Box b, double dx, double up)
        {
            foreach (var op in b.Ops) Ops.Add((c, x, y) => op(c, x + dx, y - up));
        }
    }

    public MathFrag Build(MathNode node, bool display)
    {
        var b = Layout(node, _style.Size, display, 0);
        var f = new MathFrag { Width = b.W, Ascent = b.A, Descent = b.D, Height = b.A + b.D };
        f.Ops.AddRange(b.Ops);
        return f;
    }

    // ---------------- Текст ----------------

    static bool IsBinary(int cp) => cp is '+' or '-' or 0x2212 or 0xB1 or 0x2213 or 0xD7 or 0xF7 or 0x22C5 or 0x2218 or 0x2219 or 0x2227 or 0x2228 or 0x2229 or 0x222A or 0x2295 or 0x2297 or '*';
    static bool IsRelation(int cp) => cp is '=' or '<' or '>' or 0x2260 or 0x2264 or 0x2265 or 0x2248 or 0x2261 or 0x223C or 0x2245 or 0x221D
        or 0x2192 or 0x2190 or 0x2194 or 0x21D2 or 0x21D0 or 0x21D4 or 0x2208 or 0x2209 or 0x2282 or 0x2283 or 0x2286 or 0x2287 or 0x226A or 0x226B or 0x2223 or 0x2225 or ':';

    Box Glyphs(string text, double size, bool italic, bool bold)
    {
        var b = new Box();
        var fonts = _ctx.Fonts;
        var main = fonts.Get(MathFont, bold, italic);
        var runs = new List<(FontFace F, List<ushort> G, List<string> T)>();
        foreach (var r in text.EnumerateRunes())
        {
            int cp = r.Value == '-' ? 0x2212 : r.Value;   // минус вместо дефиса
            var face = main;
            if (!face.Font.TryGetGlyph(cp, out var gid) || gid == 0)
            {
                var fb = fonts.FallbackFor(cp, MathFont, bold, italic);
                if (fb != null && fb.Font.TryGetGlyph(cp, out var g2)) { face = fb; gid = g2; }
            }
            if (runs.Count == 0 || runs[^1].F != face) runs.Add((face, [], []));
            runs[^1].G.Add(gid); runs[^1].T.Add(char.ConvertFromUtf32(r.Value));
            b.W += face.Font.Advance(gid) * size;
            var fnt = face.Font;
            b.A = Math.Max(b.A, fnt.Ascent / (double)fnt.UnitsPerEm * size * 0.82);
            b.D = Math.Max(b.D, fnt.Descent / (double)fnt.UnitsPerEm * size * 0.75);
        }
        var col = Color;
        double x0 = 0;
        foreach (var (f, g, t) in runs)
        {
            double xs = x0;
            b.Ops.Add((c, x, y) => c.Glyphs(f, size, x + xs, y, g, t, col));
            foreach (var gid in g) x0 += f.Font.Advance(gid) * size;
        }
        // Курсив чуть выступает вправо
        if (italic) b.W += size * 0.03;
        return b;
    }

    /// <summary>Текст формулы: буквы курсивом (кроме «прямого» стиля), цифры и операторы — прямо, с интервалами вокруг операторов.</summary>
    Box TextBox(MText t, double size, int level) => Hrow(TextAtoms(t, size), size, level);

    List<Box> TextAtoms(MText t, double size)
    {
        var row = new List<Box>();
        var sb = new System.Text.StringBuilder();
        bool? curItalic = null;
        void Flush()
        {
            if (sb.Length == 0) return;
            row.Add(Glyphs(sb.ToString(), size, curItalic == true, t.Bold));
            sb.Clear();
        }
        foreach (var r in t.Text.EnumerateRunes())
        {
            int cp = r.Value;
            if (IsBinary(cp) || IsRelation(cp) || cp is ',' or ';')
            {
                Flush();
                var ob = Glyphs(r.ToString(), size, false, t.Bold);
                ob.Kind = IsRelation(cp) ? 2 : IsBinary(cp) ? 1 : 3;
                row.Add(ob);
                curItalic = null;
                continue;
            }
            bool italic = !t.Plain && System.Text.Rune.IsLetter(r) && !(cp is >= 0x0391 and <= 0x03A9);   // прописные греческие — прямо
            if (curItalic != null && curItalic != italic) Flush();
            curItalic = italic;
            sb.Append(r.ToString());
        }
        Flush();
        return row;
    }

    /// <summary>Горизонтальная строка с интервалами TeX-подобных классов (бинарные операторы, отношения, пунктуация).</summary>
    Box Hrow(List<Box> items, double size, int level = 0)
    {
        var b = new Box();
        double x = 0;
        double k = level > 0 ? 0.35 : 1;   // в индексах интервалы вокруг операторов меньше
        for (int i = 0; i < items.Count; i++)
        {
            var it = items[i];
            bool unary = it.Kind == 1 && (i == 0 || items[i - 1].Kind is 1 or 2 or 3 or 4);
            double before = (it.Kind == 2 ? size * 0.28 : it.Kind == 1 && !unary ? size * 0.22 : 0) * k;
            double after = (it.Kind == 2 ? size * 0.28 : it.Kind == 1 && !unary ? size * 0.22 : it.Kind == 3 ? size * 0.17 : 0) * k;
            if (i == 0) before = 0;
            if (i == items.Count - 1 && it.Kind is 1 or 2) after = 0;
            x += before;
            b.Put(it, x, 0);
            x += it.W + after;
            b.A = Math.Max(b.A, it.A); b.D = Math.Max(b.D, it.D);
        }
        b.W = x;
        if (items.Count == 1) b.Kind = items[0].Kind;
        return b;
    }

    /// <summary>Вертикальные границы контура глифа (в долях em, ось y вверх) по его пути.</summary>
    static (double Min, double Max)? GlyphBounds(FontFace face, ushort gid)
    {
        var path = face.Font.GlyphPath(gid);
        if (path.Length == 0) return null;
        var tok = path.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        double min = double.MaxValue, max = double.MinValue;
        var nums = new List<double>();
        foreach (var t in tok)
        {
            if (t is "m" or "l" or "c")
            {
                for (int i = 1; i < nums.Count; i += 2) { min = Math.Min(min, nums[i]); max = Math.Max(max, nums[i]); }
                nums.Clear();
            }
            else if (t != "h" && double.TryParse(t, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v)) nums.Add(v);
        }
        double em = face.Font.UnitsPerEm;
        return min <= max ? (min / em, max / em) : null;
    }

    /// <summary>Один глиф с его фактическими границами по вертикали (для крупных операторов).</summary>
    (Box Box, double Lo, double Hi) BigGlyph(string ch, double size)
    {
        var box = Glyphs(ch, size, false, false);
        var fonts = _ctx.Fonts;
        var face = fonts.Get(MathFont, false, false);
        int cp = char.ConvertToUtf32(ch, 0);
        if (!face.Font.TryGetGlyph(cp, out var gid) || gid == 0)
        {
            var fb = fonts.FallbackFor(cp, MathFont, false, false);
            if (fb != null && fb.Font.TryGetGlyph(cp, out var g2)) { face = fb; gid = g2; }
        }
        var bb = GlyphBounds(face, gid) ?? (-0.2, 0.75);
        return (box, bb.Min * size, bb.Max * size);
    }

    // ---------------- Вёрстка узлов ----------------

    static double Axis(double size) => size * 0.25;
    static double ScriptSize(double size, int level) => size * (level == 0 ? 0.73 : 0.82);

    Box Layout(MathNode node, double size, bool display, int level)
    {
        switch (node)
        {
            case MText t: return TextBox(t, size, level);
            case MRow r:
            {
                var items = new List<Box>();
                foreach (var it in r.Items)
                {
                    // Текст строки разбивается на атомы, чтобы интервалы операторов учитывали соседей (унарный минус и т. п.)
                    if (it is MText mt) items.AddRange(TextAtoms(mt, size));
                    else items.Add(Layout(it, size, display, level));
                }
                return Hrow(items, size, level);
            }
            case MFrac f: return Frac(f, size, display, level);
            case MScript s: return Script(s, size, display, level);
            case MRad rad: return Radical(rad, size, display, level);
            case MNary nary: return Nary(nary, size, display, level);
            case MDelim d: return Delim(d, size, display, level);
            case MMatrix m: return Matrix(m, size, display, level);
            case MFunc fn:
            {
                var name = Layout(fn.Name is MText nt ? new MText(nt.Text, true, nt.Bold) : fn.Name, size, display, level);
                var arg = Layout(fn.E, size, display, level);
                var b = new Box { W = name.W + size * 0.17 + arg.W, A = Math.Max(name.A, arg.A), D = Math.Max(name.D, arg.D) };
                b.Put(name, 0, 0); b.Put(arg, name.W + size * 0.17, 0);
                return b;
            }
            case MAcc acc: return Accent(acc, size, display, level);
            case MBar bar:
            {
                var e = Layout(bar.E, size, display, level);
                double th = size * 0.05, gap = size * 0.08;
                var b = new Box { W = e.W, A = e.A + (bar.Top ? gap + th : 0), D = e.D + (bar.Top ? 0 : gap + th) };
                b.Put(e, 0, 0);
                double yy = bar.Top ? -(e.A + gap + th / 2) : e.D + gap + th / 2;
                var col = Color;
                b.Ops.Add((c, x, y) => c.Line(x, y + yy, x + e.W, y + yy, th, col));
                return b;
            }
            case MGroupChr g: return GroupChr(g, size, display, level);
            case MLim lim:
            {
                var e = Layout(lim.E, size, display, level);
                var l = Layout(lim.Lim, ScriptSize(size, level), display, level + 1);
                return Stack(e, lim.Upper ? l : null, lim.Upper ? null : l, size);
            }
            case MEqArr eq:
            {
                var rows = eq.Rows.Select(r => Layout(r, size, display, level)).ToList();
                double w = rows.Max(r => r.W), gap = size * 0.25;
                double total = rows.Sum(r => r.A + r.D) + gap * (rows.Count - 1);
                var b = new Box { W = w };
                double top = total / 2 + Axis(size), cy = 0;
                foreach (var r in rows)
                {
                    double baseUp = top - cy - r.A;
                    b.Put(r, (w - r.W) / 2, baseUp);
                    cy += r.A + r.D + gap;
                }
                b.A = top; b.D = total - top;
                return b;
            }
            case MBox bx:
            {
                var e = Layout(bx.E, size, display, level);
                if (!bx.Border) return e;
                double pad = size * 0.12;
                var b = new Box { W = e.W + 2 * pad, A = e.A + pad, D = e.D + pad };
                b.Put(e, pad, 0);
                var col = Color;
                b.Ops.Add((c, x, y) => c.Rect(x, y - e.A - pad, e.W + 2 * pad, e.A + e.D + 2 * pad, size * 0.04, col));
                return b;
            }
        }
        return new Box();
    }

    Box Frac(MFrac f, double size, bool display, int level)
    {
        // Вложенные дроби — чуть меньше
        double inner = level > 0 || !display ? Math.Max(size * 0.85, _style.Size * 0.6) : size;
        if (f.Type is "lin" or "skw")
        {
            var a = Layout(f.Num, size, display, level);
            var slash = Glyphs("/", size, false, false);
            var c2 = Layout(f.Den, size, display, level);
            return Hrow([a, slash, c2], size, level);
        }
        var num = Layout(f.Num, inner, display, level + 1);
        var den = Layout(f.Den, inner, display, level + 1);
        double th = f.Type == "noBar" ? 0 : size * 0.05, gap = size * 0.12, axis = Axis(size), pad = size * 0.1;
        double w = Math.Max(num.W, den.W) + 2 * pad;
        var b = new Box { W = w };
        double numUp = axis + th / 2 + gap + num.D, denDown = gap + th / 2 + den.A - axis;
        b.Put(num, (w - num.W) / 2, numUp);
        b.Put(den, (w - den.W) / 2, -denDown);
        b.A = numUp + num.A; b.D = denDown + den.D;
        if (th > 0)
        {
            var col = Color;
            b.Ops.Add((c, x, y) => c.Line(x + pad * 0.5, y - axis, x + w - pad * 0.5, y - axis, th, col));
        }
        return b;
    }

    Box Script(MScript s, double size, bool display, int level)
    {
        var bas = Layout(s.Base, size, display, level);
        double ss = ScriptSize(size, level);
        var sub = s.Sub != null ? Layout(s.Sub, ss, display, level + 1) : null;
        var sup = s.Sup != null ? Layout(s.Sup, ss, display, level + 1) : null;
        double supUp = Math.Max(size * 0.42, bas.A - (sup?.A ?? 0) * 0.55);
        double subDown = Math.Max(size * 0.18, bas.D + (sub?.A ?? 0) * 0.35 - size * 0.05);
        if (sub != null && sup != null)
        {
            // Зазор между индексами
            double clash = (supUp - sup.D) - (sub.A - subDown);
            if (clash < size * 0.15) { double fix = size * 0.15 - clash; supUp += fix / 2; subDown += fix / 2; }
        }
        double sw = Math.Max(sub?.W ?? 0, sup?.W ?? 0) + size * 0.04;
        var b = new Box { W = bas.W + sw, A = Math.Max(bas.A, sup != null ? supUp + sup.A : 0), D = Math.Max(bas.D, sub != null ? subDown + sub.D : 0) };
        double bx = s.Pre ? sw : 0, sx = s.Pre ? 0 : bas.W + size * 0.02;
        b.Put(bas, bx, 0);
        if (sup != null) b.Put(sup, s.Pre ? sw - sup.W - size * 0.02 : sx, supUp);
        if (sub != null) b.Put(sub, s.Pre ? sw - sub.W - size * 0.02 : sx, -subDown);
        b.Kind = bas.Kind;
        return b;
    }

    Box Radical(MRad r, double size, bool display, int level)
    {
        var e = Layout(r.E, size, display, level);
        double th = size * 0.05, clear = size * 0.12;
        double top = e.A + clear + th, bottom = e.D + size * 0.05;
        double h = top + bottom, signW = Math.Min(size * 0.6, size * 0.35 + h * 0.12);
        Box? deg = r.Deg != null ? Layout(r.Deg, size * 0.5, display, level + 2) : null;
        double degW = deg != null ? Math.Max(0, deg.W - signW * 0.45) : 0;
        var b = new Box { W = degW + signW + e.W + size * 0.1, A = top + th, D = bottom };
        b.Put(e, degW + signW + size * 0.05, 0);
        var col = Color;
        double x0 = degW;
        b.Ops.Add((c, x, y) =>
        {
            // Знак корня: засечка, спуск, подъём и черта над подкоренным выражением
            var p = c.NewPath();
            p.M(x + x0, y - h * 0.38 + bottom);
            p.L(x + x0 + signW * 0.22, y - h * 0.45 + bottom);
            p.L(x + x0 + signW * 0.5, y + bottom);
            p.L(x + x0 + signW, y - top);
            p.L(x + x0 + signW + e.W + size * 0.1, y - top);
            c.PaintPath(p, null, col, th, false, null, 0, 1);
        });
        if (deg != null) b.Put(deg, 0, -bottom + h * 0.55 + deg.D);
        if (deg != null) b.A = Math.Max(b.A, -bottom + h * 0.55 + deg.D + deg.A);
        return b;
    }

    Box Nary(MNary n, double size, bool display, int level)
    {
        // Знак оператора — крупнее, центр — на оси формулы
        bool integral = n.Chr is "∫" or "∬" or "∭" or "∮" or "∯" or "∰";
        double axis = Axis(size);
        // Видимая высота знака: в формуле-блоке ~1,8 em (интеграл ~2,2 em), в строке ~1,25 em
        var (probe, lo0, hi0) = BigGlyph(n.Chr, size);
        double target = (display ? (integral ? 2.2 : 1.8) : (integral ? 1.5 : 1.25)) * size;
        double big = size * target / Math.Max(size * 0.2, hi0 - lo0);
        var (op, lo, hi) = BigGlyph(n.Chr, big);
        // Центр видимой части — на оси формулы
        double opUp = axis - (lo + hi) / 2;
        op.A = opUp + hi; op.D = Math.Max(0, -(opUp + lo));
        double ss = ScriptSize(size, level);
        var sub = n.Sub != null ? Layout(n.Sub, ss, display, level + 1) : null;
        var sup = n.Sup != null ? Layout(n.Sup, ss, display, level + 1) : null;
        Box opBox;
        if (n.UndOvr)
        {
            var shifted = new Box { W = op.W, A = op.A, D = op.D };
            shifted.Put(op, 0, opUp);
            opBox = Stack(shifted, sup, sub, size);
        }
        else
        {
            opBox = new Box { W = op.W, A = op.A, D = op.D };
            opBox.Put(op, 0, opUp);
            double sw = Math.Max(sub?.W ?? 0, sup?.W ?? 0);
            double supUp = op.A - (sup?.A ?? 0) * 0.6, subDown = op.D + (sub?.A ?? 0) * 0.2;
            if (sup != null) { opBox.Put(sup, op.W + size * 0.05 - (integral ? 0 : 0), supUp); opBox.A = Math.Max(opBox.A, supUp + sup.A); }
            if (sub != null) { opBox.Put(sub, op.W - (integral ? size * 0.15 : 0), -subDown); opBox.D = Math.Max(opBox.D, subDown + sub.D); }
            opBox.W = op.W + sw + size * 0.08;
        }
        var e = Layout(n.E, size, display, level);
        var b = new Box { W = opBox.W + size * 0.12 + e.W, A = Math.Max(opBox.A, e.A), D = Math.Max(opBox.D, e.D) };
        b.Put(opBox, 0, 0);
        b.Put(e, opBox.W + size * 0.12, 0);
        return b;
    }

    /// <summary>Элемент с пределами над и под ним (по центру).</summary>
    static Box Stack(Box e, Box? over, Box? under, double size)
    {
        double gap = size * 0.08;
        double w = Math.Max(e.W, Math.Max(over?.W ?? 0, under?.W ?? 0));
        var b = new Box { W = w, A = e.A, D = e.D };
        b.Put(e, (w - e.W) / 2, 0);
        if (over != null) { double up = e.A + gap + over.D; b.Put(over, (w - over.W) / 2, up); b.A = up + over.A; }
        if (under != null) { double down = e.D + gap + under.A; b.Put(under, (w - under.W) / 2, -down); b.D = down + under.D; }
        return b;
    }

    Box Delim(MDelim d, double size, bool display, int level)
    {
        var parts = new List<Box>();
        for (int i = 0; i < d.Items.Count; i++)
        {
            if (i > 0 && d.Sep.Length > 0) parts.Add(new Box { Kind = 99 });   // место разделителя
            parts.Add(Layout(d.Items[i], size, display, level));
        }
        double axis = Axis(size);
        double contentA = parts.Where(p => p.Kind != 99).Select(p => p.A).DefaultIfEmpty(size * 0.7).Max();
        double contentD = parts.Where(p => p.Kind != 99).Select(p => p.D).DefaultIfEmpty(size * 0.2).Max();
        double half = Math.Max(Math.Max(contentA - axis, contentD + axis) + size * 0.08, size * 0.6);
        // Однострочное содержимое — обычные глифы скобок (как в тексте)
        var paren = Glyphs("(", size, false, false);
        bool small = contentA <= paren.A + size * 0.02 && contentD <= paren.D + size * 0.02;
        var b = new Box();
        double x = 0;
        void Fence(string ch, bool left)
        {
            if (ch.Length == 0) return;
            var fb = small ? Glyphs(ch, size, false, false) : FenceBox(ch, half, axis, size, left);
            b.Put(fb, x, 0);
            x += fb.W;
            b.A = Math.Max(b.A, fb.A); b.D = Math.Max(b.D, fb.D);
        }
        Fence(d.Beg, true);
        x += size * 0.04;
        foreach (var p in parts)
        {
            if (p.Kind == 99) { x += size * 0.08; Fence(d.Sep, true); x += size * 0.08; continue; }
            b.Put(p, x, 0);
            x += p.W;
            b.A = Math.Max(b.A, p.A); b.D = Math.Max(b.D, p.D);
        }
        x += size * 0.04;
        Fence(d.End, false);
        b.W = x;
        return b;
    }

    /// <summary>Растягиваемая скобка высотой 2·half вокруг оси: контуром (круглые, квадратные, фигурные, модули, угловые) или глифом.</summary>
    Box FenceBox(string ch, double half, double axis, double size, bool left)
    {
        double top = axis + half, bottom = half - axis;   // над и под базовой линией
        double h = top + bottom, th = size * 0.055;
        double w = ch switch { "(" or ")" => Math.Min(size * 0.45, size * 0.22 + h * 0.08), "{" or "}" => size * 0.4, "|" or "‖" => size * 0.25, _ => size * 0.32 };
        var b = new Box { W = w, A = top, D = bottom };
        var col = Color;
        switch (ch)
        {
            case "(" or ")":
            {
                bool l = ch == "(";
                b.Ops.Add((c, x, y) =>
                {
                    var p = c.NewPath();
                    double xo = l ? x + w * 0.85 : x + w * 0.15, xi = l ? x + w * 0.15 : x + w * 0.85;
                    p.M(xo, y - top); p.C(xi, y - top + h * 0.2, xi, y + bottom - h * 0.2, xo, y + bottom);
                    c.PaintPath(p, null, col, th, false, null, 1, 1);
                });
                return b;
            }
            case "[" or "]" or "⌈" or "⌉" or "⌊" or "⌋":
            {
                bool l = ch is "[" or "⌈" or "⌊";
                bool topBar = ch is "[" or "]" or "⌈" or "⌉", botBar = ch is "[" or "]" or "⌊" or "⌋";
                b.Ops.Add((c, x, y) =>
                {
                    var p = c.NewPath();
                    double xv = l ? x + w * 0.3 : x + w * 0.7, xe = l ? x + w * 0.85 : x + w * 0.15;
                    if (topBar) { p.M(xe, y - top); p.L(xv, y - top); } else p.M(xv, y - top);
                    p.L(xv, y + bottom);
                    if (botBar) p.L(xe, y + bottom);
                    c.PaintPath(p, null, col, th, false, null, 0, 0);
                });
                return b;
            }
            case "{" or "}":
            {
                bool l = ch == "{";
                b.Ops.Add((c, x, y) =>
                {
                    var p = c.NewPath();
                    double s = l ? 1 : -1, xm = x + w * 0.5, xa = xm + s * w * 0.35, xb = xm - s * w * 0.35, ym = y - axis;
                    p.M(xa, y - top);
                    p.C(xm, y - top, xm, y - top + h * 0.05, xm, y - top + h * 0.15);
                    p.L(xm, ym - h * 0.1);
                    p.C(xm, ym - h * 0.02, xb, ym, xb, ym);
                    p.C(xb, ym, xm, ym + h * 0.02, xm, ym + h * 0.1);
                    p.L(xm, y + bottom - h * 0.15);
                    p.C(xm, y + bottom - h * 0.05, xm, y + bottom, xa, y + bottom);
                    c.PaintPath(p, null, col, th, false, null, 1, 1);
                });
                return b;
            }
            case "|" or "‖" or "∣" or "∥":
            {
                bool dbl = ch is "‖" or "∥";
                b.Ops.Add((c, x, y) =>
                {
                    var p = c.NewPath();
                    double xm = x + w / 2;
                    if (dbl) { p.M(xm - size * 0.05, y - top); p.L(xm - size * 0.05, y + bottom); p.M(xm + size * 0.05, y - top); p.L(xm + size * 0.05, y + bottom); }
                    else { p.M(xm, y - top); p.L(xm, y + bottom); }
                    c.PaintPath(p, null, col, th, false, null, 0, 0);
                });
                return b;
            }
            case "⟨" or "⟩" or "〈" or "〉":
            {
                bool l = ch is "⟨" or "〈";
                b.Ops.Add((c, x, y) =>
                {
                    var p = c.NewPath();
                    double xo = l ? x + w * 0.85 : x + w * 0.15, xi = l ? x + w * 0.15 : x + w * 0.85;
                    p.M(xo, y - top); p.L(xi, y - axis); p.L(xo, y + bottom);
                    c.PaintPath(p, null, col, th, false, null, 1, 1);
                });
                return b;
            }
        }
        // Прочие символы — глифом, растянутым по высоте
        var g = Glyphs(ch, size, false, false);
        double k = Math.Max(1, h / Math.Max(1, g.A + g.D));
        var gb = new Box { W = g.W, A = top, D = bottom };
        gb.Ops.Add((c, x, y) =>
        {
            // Центр глифа — на оси формулы, растяжение по вертикали вокруг оси
            double yb = y - axis + (g.A - g.D) / 2;
            c.BeginTransform(x, y - axis, 1, 0, 0, k, 0, 0);
            foreach (var op in g.Ops) op(c, x, yb);
            c.EndTransform();
        });
        return gb;
    }

    Box Matrix(MMatrix m, double size, bool display, int level)
    {
        if (m.Rows.Count == 0) return new Box();
        var cells = m.Rows.Select(r => r.Select(e => Layout(e, size, display, level)).ToList()).ToList();
        int cols = cells.Max(r => r.Count);
        var colW = new double[cols];
        foreach (var r in cells) for (int i = 0; i < r.Count; i++) colW[i] = Math.Max(colW[i], r[i].W);
        var rowA = cells.Select(r => r.Select(c => c.A).DefaultIfEmpty(size * 0.7).Max()).ToList();
        var rowD = cells.Select(r => r.Select(c => c.D).DefaultIfEmpty(size * 0.2).Max()).ToList();
        double colGap = size * 0.9, rowGap = size * 0.3;
        double total = rowA.Sum() + rowD.Sum() + rowGap * (cells.Count - 1);
        double w = colW.Sum() + colGap * (cols - 1);
        var b = new Box { W = w };
        double axis = Axis(size), top = total / 2 + axis, cy = 0;
        for (int ri = 0; ri < cells.Count; ri++)
        {
            double baseUp = top - cy - rowA[ri], x = 0;
            for (int ci = 0; ci < cells[ri].Count; ci++)
            {
                b.Put(cells[ri][ci], x + (colW[ci] - cells[ri][ci].W) / 2, baseUp);
                x += colW[ci] + colGap;
            }
            cy += rowA[ri] + rowD[ri] + rowGap;
        }
        b.A = top; b.D = total - top;
        return b;
    }

    Box Accent(MAcc acc, double size, bool display, int level)
    {
        var e = Layout(acc.E, size, display, level);
        int cp = acc.Chr.Length > 0 ? char.ConvertToUtf32(acc.Chr, 0) : 0x302;
        var col = Color;
        double gap = size * 0.05;
        var b = new Box { W = e.W, D = e.D };
        b.Put(e, 0, 0);
        if (cp is 0x20D7 or 0x20D6 or 0x20E1)
        {
            // Стрелка-вектор над выражением
            double yy = e.A + gap + size * 0.08, ah = size * 0.08;
            b.Ops.Add((c, x, y) =>
            {
                c.Line(x, y - yy, x + e.W, y - yy, size * 0.045, col);
                if (cp != 0x20D6) { c.Line(x + e.W - ah * 1.6, y - yy - ah, x + e.W, y - yy, size * 0.045, col); c.Line(x + e.W - ah * 1.6, y - yy + ah, x + e.W, y - yy, size * 0.045, col); }
                if (cp != 0x20D7) { c.Line(x + ah * 1.6, y - yy - ah, x, y - yy, size * 0.045, col); c.Line(x + ah * 1.6, y - yy + ah, x, y - yy, size * 0.045, col); }
            });
            b.A = yy + ah;
            return b;
        }
        if (cp is 0x305 or 0x304 or 0xAF or 0x203E)
        {
            double yy = e.A + gap + size * 0.03;
            b.Ops.Add((c, x, y) => c.Line(x, y - yy, x + e.W, y - yy, size * 0.045, col));
            b.A = yy + size * 0.03;
            return b;
        }
        // Комбинируемый знак → отдельный (не комбинируемый) символ, по центру над выражением
        string spacing = cp switch
        {
            0x302 => "ˆ", 0x303 => "˜", 0x307 => "˙", 0x308 => "¨", 0x301 => "´", 0x300 => "`", 0x306 => "˘", 0x30C => "ˇ", 0x30A => "˚",
            _ => char.ConvertFromUtf32(cp)
        };
        var ab = Glyphs(spacing, size, false, false);
        double up = Math.Max(0, e.A - size * 0.62) + gap;
        b.Put(ab, (e.W - ab.W) / 2 + (e.W > 0 ? size * 0.05 : 0), up);
        b.A = Math.Max(e.A, up + ab.A);
        return b;
    }

    Box GroupChr(MGroupChr g, double size, bool display, int level)
    {
        var e = Layout(g.E, size, display, level);
        var col = Color;
        double gap = size * 0.08, bh = size * 0.22;
        var b = new Box { W = e.W, A = e.A, D = e.D };
        b.Put(e, 0, 0);
        bool brace = g.Chr is "⏟" or "⏞" or "︸" or "︷";
        bool top = g.Top || g.Chr is "⏞" or "︷";
        double yEdge = top ? -(e.A + gap) : e.D + gap;
        double dir = top ? -1 : 1;
        if (brace)
        {
            // Горизонтальная фигурная скобка
            b.Ops.Add((c, x, y) =>
            {
                var p = c.NewPath();
                double y0 = y + yEdge, y1 = y0 + dir * bh, xm = x + e.W / 2, ym = y0 + dir * bh / 2;
                p.M(x, y0);
                p.C(x, ym, x + e.W * 0.05, ym, x + e.W * 0.15, ym);
                p.L(xm - e.W * 0.08, ym);
                p.C(xm - e.W * 0.02, ym, xm, y1, xm, y1);
                p.C(xm, y1, xm + e.W * 0.02, ym, xm + e.W * 0.08, ym);
                p.L(x + e.W * 0.85, ym);
                p.C(x + e.W * 0.95, ym, x + e.W, ym, x + e.W, y0);
                c.PaintPath(p, null, col, size * 0.045, false, null, 1, 1);
            });
        }
        else
        {
            var cb = Glyphs(g.Chr, size, false, false);
            double up = top ? e.A + gap + cb.D : -(e.D + gap + cb.A);
            b.Put(cb, (e.W - cb.W) / 2, up);
        }
        if (top) b.A = e.A + gap + bh; else b.D = e.D + gap + bh;
        return b;
    }
}

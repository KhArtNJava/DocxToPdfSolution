using System.Globalization;
using System.Text.RegularExpressions;
using DocxToPdf.Images;

namespace DocxToPdf.Pdf;

/// <summary>
/// Растеризатор операторов содержимого PDF — того подмножества, которое выводит рисование SVG и метафайлов:
/// контуры (m l c v y h re), заливка по правилам nonzero/evenodd со сглаживанием, обводка (толщина, концы, стыки,
/// пунктир), обрезка (W/W*), цвета RGB/серый, прозрачность (gs: ca/CA), мягкие маски (SMask), градиенты (sh),
/// картинки и формы (Do), вложенные q/Q и cm. Нужен для фильтров SVG, которые работают с пикселями.
/// Пиксели — RGBA с предумноженной альфой, значения 0…1.
/// </summary>
sealed class Raster
{
    public readonly int W, H;
    public readonly float[] Px;
    readonly PdfCanvas _c;
    /// <summary>Встретилось то, что нарисовать нельзя (JPEG и т.п.) — результат неточен.</summary>
    public bool Unsupported;
    int _depth;

    public Raster(int w, int h, PdfCanvas c)
    {
        W = w; H = h; _c = c;
        Px = new float[w * h * 4];
    }

    // ================================================================
    //  Матрица
    // ================================================================

    public readonly record struct Mat(double A, double B, double C, double D, double E, double F)
    {
        public static readonly Mat I = new(1, 0, 0, 1, 0, 0);
        /// <summary>Сначала this, затем o (как в PDF: cm умножает слева на CTM).</summary>
        public Mat Then(Mat o) => new(A * o.A + B * o.C, A * o.B + B * o.D, C * o.A + D * o.C, C * o.B + D * o.D, E * o.A + F * o.C + o.E, E * o.B + F * o.D + o.F);
        public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
        public double Det => A * D - B * C;
        public Mat Inverse()
        {
            double d = Det;
            if (Math.Abs(d) < 1e-12) return I;
            return new Mat(D / d, -B / d, -C / d, A / d, (C * F - D * E) / d, (B * E - A * F) / d);
        }
        public double Scale => Math.Sqrt(Math.Abs(Det));
    }

    sealed class GS
    {
        public Mat Ctm = Mat.I;
        public double[] Fill = [0, 0, 0], Stroke = [0, 0, 0];
        public double Ca = 1, CA = 1;
        public double Lw = 1, Miter = 10, DashPhase;
        public int Cap, Join;
        public double[]? Dash;
        public float[]? Clip;      // покрытие обрезки (null — без обрезки)
        public float[]? SMask;     // мягкая маска (null — нет)
        public GS Clone() => (GS)MemberwiseClone();
    }

    // ================================================================
    //  Интерпретатор
    // ================================================================

    /// <summary>Рисует операторы с базовой матрицей (система PDF → пиксели).</summary>
    public void Run(string ops, Mat baseCtm) => Run(ops, new GS { Ctm = baseCtm });

    void Run(string ops, GS start)
    {
        if (++_depth > 12) { _depth--; return; }
        var gs = start;
        var stack = new Stack<GS>();
        var operands = new List<object>();
        var path = new List<List<(double X, double Y)>>();   // в координатах пользователя (CTM на момент рисования)
        List<(double X, double Y)>? cur = null;
        double cx = 0, cy = 0;
        int clipPending = 0;   // 1 — nonzero, 2 — evenodd
        bool inText = false;

        double N(int i) => operands.Count > i && operands[i] is double d ? d : 0;

        void Curve(double x1, double y1, double x2, double y2, double x3, double y3)
        {
            cur ??= NewSub(cx, cy);
            var (dx0, dy0) = gs.Ctm.Apply(cx, cy); var (dx3, dy3) = gs.Ctm.Apply(x3, y3);
            var (dx1, dy1) = gs.Ctm.Apply(x1, y1); var (dx2, dy2) = gs.Ctm.Apply(x2, y2);
            double len = Math.Sqrt((dx1 - dx0) * (dx1 - dx0) + (dy1 - dy0) * (dy1 - dy0)) + Math.Sqrt((dx2 - dx1) * (dx2 - dx1) + (dy2 - dy1) * (dy2 - dy1))
                + Math.Sqrt((dx3 - dx2) * (dx3 - dx2) + (dy3 - dy2) * (dy3 - dy2));
            int n = Math.Clamp((int)(len / 2), 4, 200);
            for (int k = 1; k <= n; k++)
            {
                double t = (double)k / n, u = 1 - t;
                cur.Add((u * u * u * cx + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * x3,
                         u * u * u * cy + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * y3));
            }
            cx = x3; cy = y3;
        }

        List<(double, double)> NewSub(double x, double y)
        {
            var s = new List<(double, double)> { (x, y) };
            path.Add(s);
            return s;
        }

        void EndPath(char paint, bool evenOdd)
        {
            if (paint is 'f' or 'B') FillPath(path, evenOdd, gs);
            if (paint is 'S' or 'B') StrokePath(path, gs);
            if (clipPending != 0)
            {
                var cov = Coverage(Device(path, gs.Ctm, closeAll: true), clipPending == 2);
                var nc = new float[W * H];
                for (int i = 0; i < nc.Length; i++) nc[i] = cov[i] * (gs.Clip?[i] ?? 1);
                gs.Clip = nc;
                clipPending = 0;
            }
            path = []; cur = null;
        }

        foreach (var tok in Tokens(ops))
        {
            if (tok.Kind == 0) { operands.Add(tok.Num); continue; }
            if (tok.Kind == 1) { operands.Add(tok.Text); continue; }
            if (tok.Kind == 2) { operands.Add(tok.Arr!); continue; }
            if (tok.Kind == 3) { operands.Add(""); continue; }   // строка (текст)
            string op = tok.Text;
            if (inText) { if (op == "ET") inText = false; operands.Clear(); continue; }
            switch (op)
            {
                case "q": stack.Push(gs); gs = gs.Clone(); break;
                case "Q": if (stack.Count > 0) gs = stack.Pop(); break;
                case "cm": gs.Ctm = new Mat(N(0), N(1), N(2), N(3), N(4), N(5)).Then(gs.Ctm); break;
                case "m": cx = N(0); cy = N(1); cur = NewSub(cx, cy); break;
                case "l": cur ??= NewSub(cx, cy); cx = N(0); cy = N(1); cur.Add((cx, cy)); break;
                case "c": Curve(N(0), N(1), N(2), N(3), N(4), N(5)); break;
                case "v": Curve(cx, cy, N(0), N(1), N(2), N(3)); break;
                case "y": Curve(N(0), N(1), N(2), N(3), N(2), N(3)); break;
                case "h":
                    if (cur != null && cur.Count > 0)
                    {
                        cur.Add(cur[0]);
                        (cx, cy) = cur[0];
                        cur = null;
                    }
                    break;
                case "re":
                {
                    double x = N(0), y = N(1), w = N(2), h = N(3);
                    var s = NewSub(x, y);
                    s.Add((x + w, y)); s.Add((x + w, y + h)); s.Add((x, y + h)); s.Add((x, y));
                    cur = null; cx = x; cy = y;
                    break;
                }
                case "W": clipPending = 1; break;
                case "W*": clipPending = 2; break;
                case "n": EndPath(' ', false); break;
                case "f" or "F": EndPath('f', false); break;
                case "f*": EndPath('f', true); break;
                case "S": EndPath('S', false); break;
                case "s": CloseAll(path); EndPath('S', false); break;
                case "B": EndPath('B', false); break;
                case "B*": EndPath('B', true); break;
                case "b": CloseAll(path); EndPath('B', false); break;
                case "b*": CloseAll(path); EndPath('B', true); break;
                case "rg": gs.Fill = [N(0), N(1), N(2)]; break;
                case "RG": gs.Stroke = [N(0), N(1), N(2)]; break;
                case "g": gs.Fill = [N(0), N(0), N(0)]; break;
                case "G": gs.Stroke = [N(0), N(0), N(0)]; break;
                case "k": gs.Fill = Cmyk(N(0), N(1), N(2), N(3)); break;
                case "K": gs.Stroke = Cmyk(N(0), N(1), N(2), N(3)); break;
                case "w": gs.Lw = N(0); break;
                case "J": gs.Cap = (int)N(0); break;
                case "j": gs.Join = (int)N(0); break;
                case "M": gs.Miter = N(0); break;
                case "d":
                    gs.Dash = operands.Count > 0 && operands[0] is List<double> da && da.Count > 0 && da.Any(v => v > 0) ? [.. da] : null;
                    gs.DashPhase = N(1);
                    break;
                case "gs" when operands.Count > 0 && operands[0] is string gname:
                    if (_c.AlphaByName(gname) is { } al) { gs.Ca = al.Fill; gs.CA = al.Stroke; }
                    else if (_c.SoftMaskByName(gname) is { } sm) gs.SMask = SoftMask(sm.Form, sm.Alpha, gs.Ctm);
                    break;
                case "sh" when operands.Count > 0 && operands[0] is string sname:
                    if (_c.ShadingByName(sname) is { } dict) PaintShading(dict, gs);
                    break;
                case "Do" when operands.Count > 0 && operands[0] is string xname:
                    if (_c.FormByName(xname) is { } form)
                    {
                        var g2 = gs.Clone();
                        Run(form, g2);
                    }
                    else if (_c.ImageByName(xname) is { } img) PaintImage(img, gs);
                    break;
                case "BT": inText = true; Unsupported = true; break;
            }
            operands.Clear();
        }
        _depth--;
    }

    static double[] Cmyk(double c, double m, double y, double k) => [(1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k)];

    static void CloseAll(List<List<(double X, double Y)>> path)
    {
        foreach (var s in path) if (s.Count > 1 && s[0] != s[^1]) s.Add(s[0]);
    }

    // ---------------- Лексер ----------------

    readonly record struct Tok(int Kind, double Num, string Text, List<double>? Arr);   // 0 число, 1 имя, 2 массив, 3 строка, 4 оператор

    static IEnumerable<Tok> Tokens(string s)
    {
        int i = 0, n = s.Length;
        while (i < n)
        {
            char ch = s[i];
            if (char.IsWhiteSpace(ch)) { i++; continue; }
            if (ch == '%') { while (i < n && s[i] != '\n') i++; continue; }
            if (ch == '/')
            {
                int st = ++i;
                while (i < n && !char.IsWhiteSpace(s[i]) && "/[]()<>{}".IndexOf(s[i]) < 0) i++;
                yield return new Tok(1, 0, s[st..i], null);
                continue;
            }
            if (ch == '[')
            {
                int e = s.IndexOf(']', i);
                if (e < 0) e = n - 1;
                var arr = new List<double>();
                bool str = false;
                foreach (var part in s[(i + 1)..e].Split([' ', '\n', '\r', '\t'], StringSplitOptions.RemoveEmptyEntries))
                {
                    if (double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)) arr.Add(v);
                    else str = true;
                }
                i = e + 1;
                yield return str ? new Tok(3, 0, "", null) : new Tok(2, 0, "", arr);
                continue;
            }
            if (ch == '(')
            {
                int depth = 0;
                for (; i < n; i++)
                {
                    if (s[i] == '\\') { i++; continue; }
                    if (s[i] == '(') depth++;
                    else if (s[i] == ')' && --depth == 0) { i++; break; }
                }
                yield return new Tok(3, 0, "", null);
                continue;
            }
            if (ch == '<')
            {
                int e = s.IndexOf('>', i);
                i = e < 0 ? n : e + 1;
                yield return new Tok(3, 0, "", null);
                continue;
            }
            int s0 = i;
            while (i < n && !char.IsWhiteSpace(s[i]) && "/[]()<>%".IndexOf(s[i]) < 0) i++;
            if (i == s0) { i++; continue; }
            string t = s[s0..i];
            if ((char.IsDigit(t[0]) || t[0] is '-' or '+' or '.') && double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var num))
                yield return new Tok(0, num, "", null);
            else yield return new Tok(4, 0, t, null);
        }
    }

    // ================================================================
    //  Покрытие (сглаживание: 4 подстроки на пиксель, точные доли по горизонтали)
    // ================================================================

    static List<List<(double X, double Y)>> Device(List<List<(double X, double Y)>> path, Mat m, bool closeAll)
    {
        var res = new List<List<(double, double)>>();
        foreach (var s in path)
        {
            if (s.Count < 2) continue;
            var d = new List<(double, double)>(s.Count + 1);
            foreach (var (x, y) in s) d.Add(m.Apply(x, y));
            if (closeAll && d[0] != d[^1]) d.Add(d[0]);
            res.Add(d);
        }
        return res;
    }

    const int SS = 4;

    /// <summary>Покрытие пикселей многоугольниками (координаты в пикселях), 0…1.</summary>
    float[] Coverage(List<List<(double X, double Y)>> polys, bool evenOdd)
    {
        var cov = new float[W * H];
        var edges = new List<(double X0, double Y0, double X1, double Y1, int Dir)>();
        double minY = double.MaxValue, maxY = double.MinValue;
        foreach (var p in polys)
            for (int i = 0; i + 1 < p.Count; i++)
            {
                var (x0, y0) = p[i]; var (x1, y1) = p[i + 1];
                if (double.IsNaN(x0 + y0 + x1 + y1) || y0 == y1) continue;
                if (y0 < y1) edges.Add((x0, y0, x1, y1, 1)); else edges.Add((x1, y1, x0, y0, -1));
                minY = Math.Min(minY, Math.Min(y0, y1)); maxY = Math.Max(maxY, Math.Max(y0, y1));
            }
        if (edges.Count == 0) return cov;
        edges.Sort((a, b) => a.Y0.CompareTo(b.Y0));
        int row0 = Math.Max(0, (int)Math.Floor(minY)), row1 = Math.Min(H - 1, (int)Math.Ceiling(maxY));
        var active = new List<int>();
        int next = 0;
        var xs = new List<(double X, int Dir)>();
        var acc = new float[W + 2];
        for (int row = row0; row <= row1; row++)
        {
            Array.Clear(acc);
            bool any = false;
            for (int sub = 0; sub < SS; sub++)
            {
                double sy = row + (sub + 0.5) / SS;
                while (next < edges.Count && edges[next].Y0 <= sy) active.Add(next++);
                xs.Clear();
                for (int k = active.Count - 1; k >= 0; k--)
                {
                    var e = edges[active[k]];
                    if (e.Y1 <= sy) { if (e.Y1 < row) active.RemoveAt(k); continue; }
                    if (e.Y0 > sy) continue;
                    xs.Add((e.X0 + (sy - e.Y0) * (e.X1 - e.X0) / (e.Y1 - e.Y0), e.Dir));
                }
                if (xs.Count < 2) continue;
                xs.Sort((a, b) => a.X.CompareTo(b.X));
                int wind = 0;
                for (int k = 0; k + 1 < xs.Count; k++)
                {
                    wind += xs[k].Dir;
                    bool inside = evenOdd ? (wind & 1) != 0 : wind != 0;
                    if (!inside) continue;
                    double a = Math.Max(0, xs[k].X), b = Math.Min(W, xs[k + 1].X);
                    if (b <= a) continue;
                    any = true;
                    int ia = (int)a, ib = (int)b;
                    const float w = 1f / SS;
                    if (ia == ib) { acc[ia] += (float)(b - a) * w; continue; }
                    acc[ia] += (float)(ia + 1 - a) * w;
                    for (int x = ia + 1; x < ib; x++) acc[x] += w;
                    if (ib < W) acc[ib] += (float)(b - ib) * w;
                }
            }
            if (!any) continue;
            int o = row * W;
            for (int x = 0; x < W; x++) if (acc[x] > 0) cov[o + x] = Math.Min(1f, acc[x]);
        }
        return cov;
    }

    // ================================================================
    //  Заливка, обводка
    // ================================================================

    void FillPath(List<List<(double X, double Y)>> path, bool evenOdd, GS gs)
    {
        var cov = Coverage(Device(path, gs.Ctm, closeAll: true), evenOdd);
        Composite(cov, gs.Fill, gs.Ca, gs);
    }

    void StrokePath(List<List<(double X, double Y)>> path, GS gs)
    {
        double hw = Math.Max(gs.Lw, 0) / 2;
        // Слишком тонкая линия — не тоньше пикселя (как у растеризаторов PDF)
        double minHw = 0.5 / Math.Max(1e-6, gs.Ctm.Scale);
        if (hw < minHw) hw = minHw;
        var pieces = new List<List<(double X, double Y)>>();
        double devR = hw * gs.Ctm.Scale;
        int circleN = Math.Clamp((int)(devR * 1.5), 8, 64);
        foreach (var sub0 in path)
        {
            var sub = Dedup(sub0);
            bool closed = sub.Count > 2 && Near(sub[0], sub[^1]);
            var runs = gs.Dash != null ? Dashed(sub, gs.Dash, gs.DashPhase, closed) : [(sub, closed)];
            foreach (var (pts, isClosed) in runs) StrokeRun(pts, isClosed, hw, gs, pieces, circleN);
        }
        // Все куски — с положительной ориентацией (в пикселях), объединение по nonzero
        var dev = Device(pieces, gs.Ctm, closeAll: true);
        foreach (var p in dev) if (Area(p) < 0) p.Reverse();
        var cov = Coverage(dev, false);
        Composite(cov, gs.Stroke, gs.CA, gs);
    }

    static bool Near((double X, double Y) a, (double X, double Y) b) => Math.Abs(a.X - b.X) < 1e-9 && Math.Abs(a.Y - b.Y) < 1e-9;

    static List<(double X, double Y)> Dedup(List<(double X, double Y)> s)
    {
        var r = new List<(double, double)>(s.Count);
        foreach (var p in s) if (r.Count == 0 || !Near(r[^1], p)) r.Add(p);
        if (r.Count == 1 && s.Count > 1) r.Add(r[0]);   // точка: для круглых концов
        return r;
    }

    static double Area(List<(double X, double Y)> p)
    {
        double a = 0;
        for (int i = 0; i + 1 < p.Count; i++) a += p[i].X * p[i + 1].Y - p[i + 1].X * p[i].Y;
        return a / 2;
    }

    static List<(List<(double X, double Y)>, bool)> Dashed(List<(double X, double Y)> pts, double[] dash, double phase, bool closed)
    {
        var res = new List<(List<(double, double)>, bool)>();
        double total = dash.Sum();
        if (total <= 0) return [(pts, closed)];
        int di = 0; double left = dash[0];
        phase %= total; if (phase < 0) phase += total;
        while (phase > 0) { if (phase >= left) { phase -= left; di = (di + 1) % dash.Length; left = dash[di]; } else { left -= phase; phase = 0; } }
        bool on = di % 2 == 0;
        List<(double, double)>? cur = on ? [pts[0]] : null;
        for (int i = 0; i + 1 < pts.Count; i++)
        {
            var (x0, y0) = pts[i]; var (x1, y1) = pts[i + 1];
            double len = Math.Sqrt((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)), pos = 0;
            while (len - pos > left)
            {
                pos += left;
                double t = pos / len;
                var p = (x0 + (x1 - x0) * t, y0 + (y1 - y0) * t);
                if (on) { cur!.Add(p); res.Add((cur, false)); cur = null; }
                else cur = [p];
                on = !on;
                di = (di + 1) % dash.Length; left = dash[di];
            }
            left -= len - pos;
            cur?.Add((x1, y1));
        }
        if (cur is { Count: > 1 }) res.Add((cur, false));
        return res;
    }

    void StrokeRun(List<(double X, double Y)> p, bool closed, double hw, GS gs, List<List<(double X, double Y)>> pieces, int circleN)
    {
        if (p.Count == 0) return;
        if (p.Count == 1 || p.Count == 2 && Near(p[0], p[1]))
        {
            // Нулевая длина: точка для круглого и квадратного конца
            if (gs.Cap == 1) pieces.Add(Circle(p[0], hw, circleN));
            else if (gs.Cap == 2) pieces.Add([(p[0].X - hw, p[0].Y - hw), (p[0].X + hw, p[0].Y - hw), (p[0].X + hw, p[0].Y + hw), (p[0].X - hw, p[0].Y + hw)]);
            return;
        }
        int segs = p.Count - 1;
        for (int i = 0; i < segs; i++)
        {
            var (x0, y0) = p[i]; var (x1, y1) = p[i + 1];
            double dx = x1 - x0, dy = y1 - y0, len = Math.Sqrt(dx * dx + dy * dy);
            if (len < 1e-12) continue;
            double ux = dx / len, uy = dy / len, nx = -uy * hw, ny = ux * hw;
            double ex0 = 0, ex1 = 0;
            if (!closed && gs.Cap == 2) { if (i == 0) ex0 = hw; if (i == segs - 1) ex1 = hw; }
            double ax = x0 - ux * ex0, ay = y0 - uy * ex0, bx = x1 + ux * ex1, by = y1 + uy * ex1;
            pieces.Add([(ax + nx, ay + ny), (bx + nx, by + ny), (bx - nx, by - ny), (ax - nx, ay - ny)]);
        }
        // Стыки
        if (closed)
            for (int k = 0; k < segs; k++) Join(p[(k - 1 + segs) % segs], p[k], p[(k + 1) % segs], hw, gs, pieces, circleN);
        else
            for (int k = 1; k < segs; k++) Join(p[k - 1], p[k], p[k + 1], hw, gs, pieces, circleN);
        if (!closed && gs.Cap == 1)
        {
            pieces.Add(Circle(p[0], hw, circleN));
            pieces.Add(Circle(p[^1], hw, circleN));
        }
    }

    static void Join((double X, double Y) a, (double X, double Y) v, (double X, double Y) b, double hw, GS gs, List<List<(double X, double Y)>> pieces, int circleN)
    {
        double d1x = v.X - a.X, d1y = v.Y - a.Y, d2x = b.X - v.X, d2y = b.Y - v.Y;
        double l1 = Math.Sqrt(d1x * d1x + d1y * d1y), l2 = Math.Sqrt(d2x * d2x + d2y * d2y);
        if (l1 < 1e-12 || l2 < 1e-12) return;
        d1x /= l1; d1y /= l1; d2x /= l2; d2y /= l2;
        double cross = d1x * d2y - d1y * d2x;
        if (Math.Abs(cross) < 1e-9 && d1x * d2x + d1y * d2y > 0) return;   // прямая
        if (gs.Join == 1) { pieces.Add(Circle(v, hw, circleN)); return; }
        // Внешняя сторона поворота
        double s = cross > 0 ? -1 : 1;
        var p1 = (v.X - d1y * hw * s, v.Y + d1x * hw * s);
        var p2 = (v.X - d2y * hw * s, v.Y + d2x * hw * s);
        if (gs.Join == 0)
        {
            double cosT = Math.Clamp(-(d1x * d2x + d1y * d2y), -1, 1);
            double theta = Math.Acos(cosT);   // угол между сегментами
            double ratio = 1 / Math.Max(1e-9, Math.Sin(theta / 2));
            if (ratio <= gs.Miter)
            {
                double mx = -d1y * s + -d2y * s, my = d1x * s + d2x * s;
                double ml = Math.Sqrt(mx * mx + my * my);
                if (ml > 1e-9)
                {
                    double dist = hw * ratio;
                    var mp = (v.X + mx / ml * dist, v.Y + my / ml * dist);
                    pieces.Add([v, p1, mp, p2]);
                    return;
                }
            }
        }
        pieces.Add([v, p1, p2]);
    }

    static List<(double X, double Y)> Circle((double X, double Y) c, double r, int n)
    {
        var res = new List<(double, double)>(n);
        for (int i = 0; i < n; i++) { double a = 2 * Math.PI * i / n; res.Add((c.X + r * Math.Cos(a), c.Y + r * Math.Sin(a))); }
        return res;
    }

    /// <summary>Наложение цвета с покрытием (source-over).</summary>
    void Composite(float[] cov, double[] col, double alpha, GS gs)
    {
        float r = (float)col[0], g = (float)col[1], b = (float)col[2], al = (float)Math.Clamp(alpha, 0, 1);
        for (int i = 0; i < cov.Length; i++)
        {
            float a = cov[i];
            if (a <= 0) continue;
            a *= al;
            if (gs.Clip != null) a *= gs.Clip[i];
            if (gs.SMask != null) a *= gs.SMask[i];
            if (a <= 0) continue;
            int o = i * 4;
            float k = 1 - a;
            Px[o] = r * a + Px[o] * k; Px[o + 1] = g * a + Px[o + 1] * k; Px[o + 2] = b * a + Px[o + 2] * k; Px[o + 3] = a + Px[o + 3] * k;
        }
    }

    // ================================================================
    //  Мягкая маска, градиент, картинка
    // ================================================================

    float[] SoftMask(string form, bool alpha, Mat ctm)
    {
        var content = _c.FormByName(form);
        var m = new float[W * H];
        if (content == null) return m;
        var r = new Raster(W, H, _c) { _depth = _depth };
        r.Run(content, new GS { Ctm = ctm });
        Unsupported |= r.Unsupported;
        for (int i = 0; i < m.Length; i++)
        {
            int o = i * 4;
            // Яркость содержимого на чёрном фоне (предумноженные значения)
            m[i] = alpha ? r.Px[o + 3] : Math.Clamp(0.2125f * r.Px[o] + 0.7154f * r.Px[o + 1] + 0.0721f * r.Px[o + 2], 0, 1);
        }
        return m;
    }

    static double[] Arr(string dict, string key)
    {
        var m = Regex.Match(dict, "/" + key + @"\s*\[([^\]]*)\]");
        if (!m.Success) return [];
        return m.Groups[1].Value.Split([' ', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .Select(v => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0).ToArray();
    }

    void PaintShading(string dict, GS gs)
    {
        bool radial = dict.Contains("/ShadingType 3");
        var co = Arr(dict, "Coords");
        if (co.Length < (radial ? 6 : 4)) return;
        // Функция: пары C0/C1 по порядку и границы
        var c0s = Regex.Matches(dict, @"/C0\s*\[([^\]]*)\]").Select(m => Nums(m.Groups[1].Value)).ToList();
        var c1s = Regex.Matches(dict, @"/C1\s*\[([^\]]*)\]").Select(m => Nums(m.Groups[1].Value)).ToList();
        if (c0s.Count == 0 || c1s.Count != c0s.Count) return;
        var bounds = Arr(dict, "Bounds");
        var inv = gs.Ctm.Inverse();
        for (int py = 0; py < H; py++)
            for (int px = 0; px < W; px++)
            {
                int i = py * W + px;
                float a = (float)gs.Ca;
                if (gs.Clip != null) a *= gs.Clip[i];
                if (gs.SMask != null) a *= gs.SMask[i];
                if (a <= 0) continue;
                var (x, y) = inv.Apply(px + 0.5, py + 0.5);
                double t;
                if (!radial)
                {
                    double dx = co[2] - co[0], dy = co[3] - co[1], dd = dx * dx + dy * dy;
                    t = dd < 1e-12 ? 0 : ((x - co[0]) * dx + (y - co[1]) * dy) / dd;
                }
                else
                {
                    // Окружности: центр f + t·(c − f), радиус r0 + t·(r1 − r0)
                    double fx = co[0], fy = co[1], r0 = co[2], cdx = co[3] - fx, cdy = co[4] - fy, dr = co[5] - r0;
                    double qx = x - fx, qy = y - fy;
                    double A = cdx * cdx + cdy * cdy - dr * dr, B = qx * cdx + qy * cdy + r0 * dr, C = qx * qx + qy * qy - r0 * r0;
                    if (Math.Abs(A) < 1e-12) t = Math.Abs(B) < 1e-12 ? 0 : C / (2 * B);
                    else
                    {
                        double disc = B * B - A * C;
                        if (disc < 0) continue;
                        double sq = Math.Sqrt(disc), t1 = (B + sq) / A, t2 = (B - sq) / A;
                        t = Math.Max(t1, t2);
                        if (r0 + t * dr < 0) t = Math.Min(t1, t2);
                    }
                }
                t = Math.Clamp(t, 0, 1);
                int seg = 0;
                double lo = 0, hi = 1;
                if (c0s.Count > 1)
                {
                    while (seg < bounds.Length && t >= bounds[seg]) seg++;
                    seg = Math.Min(seg, c0s.Count - 1);
                    lo = seg == 0 ? 0 : bounds[seg - 1]; hi = seg < bounds.Length ? bounds[seg] : 1;
                }
                double u = hi - lo > 1e-12 ? (t - lo) / (hi - lo) : 0;
                var a0 = c0s[seg]; var a1 = c1s[seg];
                int o = i * 4;
                float k = 1 - a;
                for (int ch = 0; ch < 3; ch++)
                {
                    double v = (ch < a0.Length ? a0[ch] : 0) * (1 - u) + (ch < a1.Length ? a1[ch] : 0) * u;
                    Px[o + ch] = (float)v * a + Px[o + ch] * k;
                }
                Px[o + 3] = a + Px[o + 3] * k;
            }
    }

    static double[] Nums(string s) => s.Split([' ', '\n'], StringSplitOptions.RemoveEmptyEntries)
        .Select(v => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0).ToArray();

    void PaintImage(PdfImageData img, GS gs)
    {
        if (img.Filter != "/FlateDecode" || img.BitsPerComponent != 8 || img.Width <= 0 || img.Height <= 0) { Unsupported = true; return; }
        int comps = img.ColorSpace switch { "/DeviceGray" => 1, "/DeviceCMYK" => 4, "/DeviceRGB" => 3, _ => 0 };
        if (comps == 0 || img.Data.Length < img.Width * img.Height * comps) { Unsupported = true; return; }
        var inv = gs.Ctm.Inverse();
        // Рамка картинки в пикселях
        double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
        foreach (var (u, v) in new[] { (0.0, 0.0), (1.0, 0.0), (0.0, 1.0), (1.0, 1.0) })
        {
            var (dx, dy) = gs.Ctm.Apply(u, v);
            x0 = Math.Min(x0, dx); y0 = Math.Min(y0, dy); x1 = Math.Max(x1, dx); y1 = Math.Max(y1, dy);
        }
        int px0 = Math.Max(0, (int)Math.Floor(x0)), px1 = Math.Min(W - 1, (int)Math.Ceiling(x1));
        int py0 = Math.Max(0, (int)Math.Floor(y0)), py1 = Math.Min(H - 1, (int)Math.Ceiling(y1));
        int iw = img.Width, ih = img.Height;
        for (int py = py0; py <= py1; py++)
            for (int px = px0; px <= px1; px++)
            {
                var (u, v) = inv.Apply(px + 0.5, py + 0.5);
                if (u < 0 || u > 1 || v < 0 || v > 1) continue;
                // Билинейная выборка
                double fx = u * iw - 0.5, fy = (1 - v) * ih - 0.5;
                int ix = (int)Math.Floor(fx), iy = (int)Math.Floor(fy);
                double tx = fx - ix, ty = fy - iy;
                double r = 0, g = 0, b = 0, al = 0;
                for (int k = 0; k < 4; k++)
                {
                    int sx = Math.Clamp(ix + (k & 1), 0, iw - 1), sy = Math.Clamp(iy + (k >> 1), 0, ih - 1);
                    double wgt = ((k & 1) == 1 ? tx : 1 - tx) * ((k >> 1) == 1 ? ty : 1 - ty);
                    int si = sy * iw + sx;
                    double sa = img.Alpha != null && si < img.Alpha.Length ? img.Alpha[si] / 255.0 : 1;
                    double cr, cg, cb;
                    if (comps == 1) cr = cg = cb = img.Data[si] / 255.0;
                    else if (comps == 3) { cr = img.Data[si * 3] / 255.0; cg = img.Data[si * 3 + 1] / 255.0; cb = img.Data[si * 3 + 2] / 255.0; }
                    else
                    {
                        var c4 = Cmyk(img.Data[si * 4] / 255.0, img.Data[si * 4 + 1] / 255.0, img.Data[si * 4 + 2] / 255.0, img.Data[si * 4 + 3] / 255.0);
                        cr = c4[0]; cg = c4[1]; cb = c4[2];
                    }
                    r += cr * sa * wgt; g += cg * sa * wgt; b += cb * sa * wgt; al += sa * wgt;
                }
                int i = py * W + px;
                float m = (float)gs.Ca;
                if (gs.Clip != null) m *= gs.Clip[i];
                if (gs.SMask != null) m *= gs.SMask[i];
                float a = (float)al * m;
                if (a <= 0) continue;
                int o = i * 4;
                float kk = 1 - a;
                Px[o] = (float)r * m + Px[o] * kk; Px[o + 1] = (float)g * m + Px[o + 1] * kk; Px[o + 2] = (float)b * m + Px[o + 2] * kk; Px[o + 3] = a + Px[o + 3] * kk;
            }
    }
}

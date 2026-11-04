using System.Globalization;
using DocxToPdf.Docx;
using DocxToPdf.Pdf;

namespace DocxToPdf.Layout;

/// <summary>Свёрстанная диаграмма: примитивы в координатах рамки (ось y вниз).</summary>
public sealed class ChartFrag : Frag
{
    /// <summary>Операции рисования в порядке добавления (фон → сетка → данные → подписи).</summary>
    readonly List<Action<PdfCanvas, double, double>> _ops = [];

    public void Rect(double x, double y, double w, double h, Rgb col) => _ops.Add((c, ox, oy) => c.FillRect(ox + x, oy + y, w, h, col));
    public void Line(List<(double X, double Y)> pts, double width, Rgb col) =>
        _ops.Add((c, ox, oy) => c.Polyline(pts.Select(p => (ox + p.X, oy + p.Y)).ToList(), width, col));
    public void Polygon(List<(double X, double Y)> pts, Rgb col) =>
        _ops.Add((c, ox, oy) => c.FillPolygon(pts.Select(p => (ox + p.X, oy + p.Y)).ToList(), col));
    public void Sector(double cx, double cy, double r, double a0, double a1, Rgb col, double inner) =>
        _ops.Add((c, ox, oy) => c.FillSector(ox + cx, oy + cy, r, a0, a1, col, inner));
    public void Text(Frag t, double x, double y) => _ops.Add((c, ox, oy) => t.Draw(c, ox + x, oy + y));

    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var op in _ops) op(c, x, y);
    }
}

/// <summary>Вёрстка диаграмм: гистограммы (столбцы/полосы, с группировкой и накоплением), графики, области, круговые.</summary>
public sealed class ChartLayout(BlockLayout layout)
{
    static readonly Rgb Axis = new(0.85, 0.85, 0.85), TextCol = new(0.35, 0.35, 0.35);

    (LineFrag Line, double W, double H) Text(string s, double size, bool bold = false, Rgb? color = null)
    {
        var st = new TextStyle { FontAscii = "Calibri", FontHAnsi = "Calibri", FontEastAsia = "Calibri", Size = size, Bold = bold, Color = color ?? TextCol };
        var p = new Paragraph { MarkStyle = st, SnapToGrid = false };
        p.Inlines.Add(new TextInline(s, st));
        var lf = layout.LayoutParagraph(p, 2000)[0];
        double w = lf.Pieces.Count > 0 ? lf.Pieces.Max(q => q.X + q.Width) : 0;
        return (lf, w, lf.Height);
    }

    void PutText(ChartFrag f, string s, double size, double x, double y, string align = "left", bool bold = false, Rgb? color = null)
    {
        var (lf, w, h) = Text(s, size, bold, color);
        double dx = align switch { "center" => -w / 2, "right" => -w, _ => 0 };
        f.Text(lf, x + dx, y - h / 2);
    }

    static string Num(double v) => Math.Abs(v - Math.Round(v)) < 1e-9 ? Math.Round(v).ToString(CultureInfo.InvariantCulture) : v.ToString("0.##", CultureInfo.InvariantCulture);

    public ChartFrag Build(ChartSpec ch, double width, double height)
    {
        var f = new ChartFrag { Height = height };
        f.Rect(0, 0, width, height, Rgb.White);
        // Рамка области диаграммы
        var border = new Rgb(0.85, 0.85, 0.85);
        f.Line([(0, 0), (width, 0), (width, height), (0, height), (0, 0)], 0.75, border);
        double top = 8, bottom = height - 8, left = 8, right = width - 8;
        if (ch.Title != null)
        {
            PutText(f, ch.Title, 14, width / 2, top + 10, "center", false, new Rgb(0.35, 0.35, 0.35));
            top += 26;
        }
        bool pie = ch.Type is "pie" or "doughnut";
        // Легенда: элементы — ряды (или категории у круговой)
        var entries = pie ? ch.Categories.Select((c, i) => (c, ch.Series[0].PointColors.GetValueOrDefault(i, DocxReader.AccentColor(i)))).ToList()
                          : ch.Series.Select(s => (s.Name, s.Color ?? Rgb.Black)).ToList();
        if (ch.Legend && entries.Count > 0)
        {
            const double sz = 9;
            if (ch.LegendPos is "r" or "l")
            {
                double lw = entries.Max(e => Text(e.Item1, sz).W) + 14, lh = entries.Count * 14;
                double lx = ch.LegendPos == "r" ? right - lw : left, ly = (top + bottom) / 2 - lh / 2;
                for (int i = 0; i < entries.Count; i++)
                {
                    double ey = ly + i * 14 + 7;
                    f.Rect(lx, ey - 3.5, 7, 7, entries[i].Item2);
                    PutText(f, entries[i].Item1, sz, lx + 11, ey);
                }
                if (ch.LegendPos == "r") right -= lw + 8; else left += lw + 8;
            }
            else
            {
                double total = entries.Sum(e => Text(e.Item1, sz).W + 22);
                double lx = width / 2 - total / 2, ly = ch.LegendPos == "t" ? top + 6 : bottom - 6;
                foreach (var (name, col) in entries)
                {
                    f.Rect(lx, ly - 3.5, 7, 7, col);
                    PutText(f, name, sz, lx + 11, ly);
                    lx += Text(name, sz).W + 22;
                }
                if (ch.LegendPos == "t") top += 16; else bottom -= 18;
            }
        }
        if (pie) { BuildPie(f, ch, left, top, right, bottom); return f; }
        BuildAxes(f, ch, left, top, right, bottom);
        return f;
    }

    void BuildPie(ChartFrag f, ChartSpec ch, double l, double t, double r, double b)
    {
        var se = ch.Series[0];
        double sum = se.Values.Where(v => v > 0).Sum();
        if (sum <= 0) return;
        double rad = Math.Max(5, Math.Min(r - l, b - t) / 2 - 4), cx = (l + r) / 2, cy = (t + b) / 2, a = 0;
        for (int i = 0; i < se.Values.Count; i++)
        {
            if (se.Values[i] <= 0) continue;
            double sweep = se.Values[i] / sum * 360;
            f.Sector(cx, cy, rad, a, a + sweep, se.PointColors.GetValueOrDefault(i, DocxReader.AccentColor(i)), ch.Type == "doughnut" ? rad * 0.5 : 0);
            if (ch.DataLabels)
            {
                double mid = (a + sweep / 2 - 90) * Math.PI / 180, lr = ch.Type == "doughnut" ? rad * 0.75 : rad * 0.65;
                PutText(f, Num(se.Values[i]), 9, cx + lr * Math.Cos(mid), cy + lr * Math.Sin(mid), "center", false, Rgb.White);
            }
            a += sweep;
        }
    }

    void BuildAxes(ChartFrag f, ChartSpec ch, double l, double t, double r, double b)
    {
        int n = ch.Categories.Count;
        if (n == 0) return;
        bool stacked = ch.Grouping is "stacked" or "percentStacked", percent = ch.Grouping == "percentStacked";
        bool horizontal = ch.Type == "bar";
        // Диапазон значений
        double min = 0, max = 0;
        for (int i = 0; i < n; i++)
        {
            if (stacked)
            {
                double pos = 0, neg = 0;
                foreach (var s in ch.Series) { double v = i < s.Values.Count ? s.Values[i] : 0; if (v >= 0) pos += v; else neg += v; }
                if (percent) { pos = pos > 0 ? 100 : 0; neg = neg < 0 ? -100 : 0; }
                max = Math.Max(max, pos); min = Math.Min(min, neg);
            }
            else foreach (var s in ch.Series) if (i < s.Values.Count) { max = Math.Max(max, s.Values[i]); min = Math.Min(min, s.Values[i]); }
        }
        if (max - min < 1e-9) max = min + 1;
        // «Красивый» шаг шкалы: 1, 2, 5 × 10^n, около 5–8 делений
        double raw = (max - min) / 6, mag = Math.Pow(10, Math.Floor(Math.Log10(raw))), step = mag;
        foreach (var m in new[] { 1, 2, 5, 10 }) { step = m * mag; if (step >= raw) break; }
        double vmin = Math.Floor(min / step) * step, vmax = Math.Ceiling(max / step - 1e-9) * step;
        if (vmax <= vmin) vmax = vmin + step;
        const double fs = 9;
        // Подписи значений и категорий занимают место у осей
        var valLabels = new List<double>();
        for (double v = vmin; v <= vmax + step / 2; v += step) valLabels.Add(Math.Round(v / step) * step);
        double valW = valLabels.Max(v => Text(Num(v) + (percent ? "%" : ""), fs).W) + 6;
        double catW = ch.Categories.Max(c => Text(c, fs).W) + 6;
        double pl = l + (horizontal ? catW : valW), pr = r - 4, pt = t + 6, pb = b - 16;
        if (pr - pl < 10 || pb - pt < 10) return;
        double VPos(double v) => horizontal ? pl + (v - vmin) / (vmax - vmin) * (pr - pl) : pb - (v - vmin) / (vmax - vmin) * (pb - pt);
        // Сетка и подписи шкалы
        foreach (var v in valLabels)
        {
            double p = VPos(v);
            string lbl = Num(v) + (percent ? "%" : "");
            if (horizontal)
            {
                if (ch.Gridlines) f.Line([(p, pt), (p, pb)], 0.75, Axis);
                PutText(f, lbl, fs, p, pb + 8, "center");
            }
            else
            {
                if (ch.Gridlines) f.Line([(pl, p), (pr, p)], 0.75, Axis);
                PutText(f, lbl, fs, pl - 4, p, "right");
            }
        }
        // Ось категорий
        double zero = VPos(Math.Clamp(0, vmin, vmax));
        double band = (horizontal ? pb - pt : pr - pl) / n;
        for (int i = 0; i < n; i++)
        {
            // У горизонтальных полос первая категория внизу (как в Word)
            double c = horizontal ? pb - band * (i + 0.5) : pl + band * (i + 0.5);
            if (horizontal) PutText(f, ch.Categories[i], fs, pl - 4, c, "right");
            else PutText(f, ch.Categories[i], fs, c, pb + 8, "center");
        }
        if (horizontal) f.Line([(zero, pt), (zero, pb)], 0.75, new Rgb(0.75, 0.75, 0.75));
        else f.Line([(pl, zero), (pr, zero)], 0.75, new Rgb(0.75, 0.75, 0.75));

        if (ch.Type is "col" or "bar")
        {
            int ns = ch.Series.Count;
            double gap = band * 0.4, groupW = band - gap, barW = stacked ? groupW : groupW / ns;
            var posAcc = new double[n]; var negAcc = new double[n];
            for (int si = 0; si < ns; si++)
            {
                var s = ch.Series[si];
                for (int i = 0; i < n && i < s.Values.Count; i++)
                {
                    double v = s.Values[i];
                    if (percent)
                    {
                        double tot = ch.Series.Sum(q => i < q.Values.Count ? Math.Abs(q.Values[i]) : 0);
                        v = tot > 0 ? v / tot * 100 : 0;
                    }
                    double from = 0;
                    if (stacked) { from = v >= 0 ? posAcc[i] : negAcc[i]; if (v >= 0) posAcc[i] += v; else negAcc[i] += v; }
                    double a = VPos(from), e = VPos(from + v);
                    double off = horizontal ? pb - band * (i + 1) + gap / 2 + (stacked ? 0 : barW * (ns - 1 - si))
                                            : pl + band * i + gap / 2 + (stacked ? 0 : barW * si);
                    var col = s.PointColors.GetValueOrDefault(i, s.Color ?? Rgb.Black);
                    if (horizontal) f.Rect(Math.Min(a, e), off, Math.Abs(e - a), barW, col);
                    else f.Rect(off, Math.Min(a, e), barW, Math.Abs(e - a), col);
                    if (ch.DataLabels)
                    {
                        if (horizontal) PutText(f, Num(s.Values[i]), fs, e + 3, off + barW / 2);
                        else PutText(f, Num(s.Values[i]), fs, off + barW / 2, e - 7, "center");
                    }
                }
            }
        }
        else
        {
            // График / область / точечная: точки в центрах категорий
            var acc = new double[n];
            foreach (var s in ch.Series)
            {
                var pts = new List<(double X, double Y)>();
                for (int i = 0; i < n && i < s.Values.Count; i++)
                {
                    double v = s.Values[i];
                    if (stacked) { acc[i] += v; v = acc[i]; }
                    pts.Add((pl + band * (i + 0.5), VPos(v)));
                }
                if (pts.Count == 0) continue;
                var col = s.Color ?? Rgb.Black;
                if (ch.Type == "area")
                {
                    var poly = new List<(double X, double Y)>(pts) { (pts[^1].X, zero), (pts[0].X, zero) };
                    f.Polygon(poly, col);
                }
                else
                {
                    f.Line(pts, ch.Type == "scatter" ? 0 : 2.25, col);
                    if (ch.Type == "scatter" || pts.Count == 1)
                        foreach (var (px, py) in pts) f.Rect(px - 2.5, py - 2.5, 5, 5, col);
                }
                if (ch.DataLabels)
                    for (int i = 0; i < pts.Count; i++) PutText(f, Num(s.Values[i]), fs, pts[i].X, pts[i].Y - 8, "center");
            }
        }
    }
}

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
    public void SectorE(double cx, double cy, double r, double a0, double a1, Rgb col, double yScale, bool outline) =>
        _ops.Add((c, ox, oy) => c.FillSector(ox + cx, oy + cy, r, a0, a1, col, 0, yScale, outline));
    public void Marker(string kind, double x, double y, double size, Rgb col) => _ops.Add((c, ox, oy) =>
    {
        var spec = kind switch
        {
            "circle" or "dot" => new ShapeSpec { Geom = "ellipse", Fill = col },
            "diamond" => new ShapeSpec { Geom = "diamond", Fill = col },
            "triangle" => new ShapeSpec { Geom = "triangle", Fill = col },
            "x" or "star" or "plus" or "dash" => null,
            _ => new ShapeSpec { Geom = "rect", Fill = col },
        };
        double sz = kind == "dot" ? size * 0.5 : size;
        if (spec != null) { c.Shape(spec, ox + x, oy + y, sz, sz, 0, false, false); return; }
        double h = sz / 2, px = ox + x, py = oy + y;
        if (kind is "x" or "star") { c.Line(px - h, py - h, px + h, py + h, 1, col); c.Line(px - h, py + h, px + h, py - h, 1, col); }
        if (kind is "plus" or "star") { c.Line(px - h, py, px + h, py, 1, col); c.Line(px, py - h, px, py + h, 1, col); }
        if (kind == "dash") c.Line(px - h, py, px + h, py, 1.5, col);
    });

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

    static Rgb Shade(Rgb c, double k) => new(c.R * k, c.G * k, c.B * k);

    void BuildPie(ChartFrag f, ChartSpec ch, double l, double t, double r, double b)
    {
        var se = ch.Series[0];
        double sum = se.Values.Where(v => v > 0).Sum();
        if (sum <= 0) return;
        bool d3 = ch.ThreeD && ch.Type == "pie";
        double yScale = d3 ? 0.55 : 1;
        double rad = Math.Max(5, Math.Min((r - l) / 2, (b - t) / 2 / (d3 ? 0.75 : 1)) - 4);
        double depth = d3 ? rad * 0.18 : 0;
        double cx = (l + r) / 2, cy = (t + b) / 2 - depth / 2;
        var colors = Enumerable.Range(0, se.Values.Count).Select(i => se.PointColors.GetValueOrDefault(i, DocxReader.AccentColor(i))).ToList();
        if (d3)
        {
            // Боковая стенка: слои затемнённых секторов снизу вверх
            int layers = Math.Max(4, (int)depth);
            for (int k = layers; k >= 1; k--)
            {
                double a = 0;
                for (int i = 0; i < se.Values.Count; i++)
                {
                    if (se.Values[i] <= 0) continue;
                    double sweep = se.Values[i] / sum * 360;
                    f.SectorE(cx, cy + depth * k / layers, rad, a, a + sweep, Shade(colors[i], 0.7), yScale, false);
                    a += sweep;
                }
            }
        }
        double acc = 0;
        for (int i = 0; i < se.Values.Count; i++)
        {
            if (se.Values[i] <= 0) continue;
            double sweep = se.Values[i] / sum * 360;
            if (d3) f.SectorE(cx, cy, rad, acc, acc + sweep, colors[i], yScale, true);
            else f.Sector(cx, cy, rad, acc, acc + sweep, colors[i], ch.Type == "doughnut" ? rad * 0.5 : 0);
            if (ch.DataLabels)
            {
                double mid = (acc + sweep / 2 - 90) * Math.PI / 180, lr = ch.Type == "doughnut" ? rad * 0.75 : rad * 0.65;
                PutText(f, Num(se.Values[i]), 9, cx + lr * Math.Cos(mid), cy + lr * yScale * Math.Sin(mid), "center", false, Rgb.White);
            }
            acc += sweep;
        }
    }

    /// <summary>Шкала оси значений: «красивый» шаг 1, 2, 5 × 10^n и границы.</summary>
    static (double Min, double Max, double Step) NiceScale(double min, double max)
    {
        if (max - min < 1e-9) max = min + 1;
        double raw = (max - min) / 6, mag = Math.Pow(10, Math.Floor(Math.Log10(raw))), step = mag;
        foreach (var m in new[] { 1, 2, 5, 10 }) { step = m * mag; if (step >= raw) break; }
        double vmin = Math.Floor(min / step) * step, vmax = Math.Ceiling(max / step - 1e-9) * step;
        if (vmax <= vmin) vmax = vmin + step;
        return (vmin, vmax, step);
    }

    /// <summary>Диапазон значений рядов одной оси (с учётом накопления).</summary>
    static (double Min, double Max) Range(List<ChartSeries> series, int n)
    {
        double min = 0, max = 0;
        foreach (var grp in series.GroupBy(s => s.Type + "/" + s.Grouping))
        {
            var list = grp.ToList();
            bool stacked = list[0].Grouping is "stacked" or "percentStacked", percent = list[0].Grouping == "percentStacked";
            for (int i = 0; i < n; i++)
            {
                if (stacked)
                {
                    double pos = 0, neg = 0;
                    foreach (var s in list) { double v = i < s.Values.Count ? s.Values[i] : 0; if (v >= 0) pos += v; else neg += v; }
                    if (percent) { pos = pos > 0 ? 100 : 0; neg = neg < 0 ? -100 : 0; }
                    max = Math.Max(max, pos); min = Math.Min(min, neg);
                }
                else foreach (var s in list) if (i < s.Values.Count) { max = Math.Max(max, s.Values[i]); min = Math.Min(min, s.Values[i]); }
            }
        }
        return (min, max);
    }

    void BuildAxes(ChartFrag f, ChartSpec ch, double l, double t, double r, double b)
    {
        int n = ch.Categories.Count;
        if (n == 0) return;
        bool horizontal = ch.Series.Any(s => s.Type == "bar") && ch.Series.All(s => s.Type == "bar");
        var primary = ch.Series.Where(s => !s.Secondary).ToList();
        var secondary = ch.Series.Where(s => s.Secondary).ToList();
        if (primary.Count == 0) { primary = secondary; secondary = []; }
        bool percent = primary.Any(s => s.Grouping == "percentStacked");
        var (pmin, pmax) = Range(primary, n);
        var (vmin, vmax, step) = NiceScale(pmin, pmax);
        (double Min, double Max, double Step) sec = default;
        if (secondary.Count > 0) { var (smin, smax) = Range(secondary, n); sec = NiceScale(smin, smax); }
        const double fs = 9;
        var valLabels = new List<double>();
        for (double v = vmin; v <= vmax + step / 2; v += step) valLabels.Add(Math.Round(v / step) * step);
        // Вторичная шкала — столько же делений, что и основная (как в Word при общей сетке)
        int divs = Math.Max(1, valLabels.Count - 1);
        var secLabels = secondary.Count > 0 ? Enumerable.Range(0, divs + 1).Select(i => sec.Min + (sec.Max - sec.Min) * i / divs).ToList() : [];
        double valW = valLabels.Max(v => Text(Num(v) + (percent ? "%" : ""), fs).W) + 6;
        double secW = secLabels.Count > 0 ? secLabels.Max(v => Text(Num(v), fs).W) + 6 : 0;
        double catW = ch.Categories.Max(c => Text(c, fs).W) + 6;
        // Объёмный вид: глубина столбцов и смещение задней стенки
        double depth = ch.ThreeD ? 8 : 0;
        double pl = l + (horizontal ? catW : valW), pr = r - 4 - secW - depth, pt = t + 6 + depth, pb = b - 16;
        if (pr - pl < 10 || pb - pt < 10) return;
        double VPos(double v, bool secAxis = false)
        {
            double lo = secAxis ? sec.Min : vmin, hi = secAxis ? sec.Max : vmax;
            double k = hi > lo ? (v - lo) / (hi - lo) : 0;
            return horizontal ? pl + k * (pr - pl) : pb - k * (pb - pt);
        }
        var gridCol = new Rgb(0.85, 0.85, 0.85);
        if (ch.ThreeD)
        {
            // Задняя и боковая стенки
            f.Polygon([(pl, pt), (pl + depth, pt - depth), (pl + depth, pb - depth), (pl, pb)], new Rgb(0.95, 0.95, 0.95));
            f.Polygon([(pl, pb), (pl + depth, pb - depth), (pr + depth, pb - depth), (pr, pb)], new Rgb(0.9, 0.9, 0.9));
        }
        // Сетка и подписи шкал
        for (int i = 0; i < valLabels.Count; i++)
        {
            double p = VPos(valLabels[i]);
            string lbl = Num(valLabels[i]) + (percent ? "%" : "");
            if (horizontal)
            {
                if (ch.Gridlines) f.Line([(p, pt), (p + depth, pt - depth), (p + depth, pb - depth)], 0.75, gridCol);
                PutText(f, lbl, fs, p, pb + 8, "center");
            }
            else
            {
                if (ch.Gridlines) f.Line(ch.ThreeD ? [(pl, p), (pl + depth, p - depth), (pr + depth, p - depth)] : [(pl, p), (pr, p)], 0.75, gridCol);
                PutText(f, lbl, fs, pl - 4, p, "right");
                if (secLabels.Count > i) PutText(f, Num(secLabels[i]), fs, pr + depth + 4, p, "left");
            }
        }
        double band = (horizontal ? pb - pt : pr - pl) / n;
        for (int i = 0; i < n; i++)
        {
            double c = horizontal ? pb - band * (i + 0.5) : pl + band * (i + 0.5);
            if (horizontal) PutText(f, ch.Categories[i], fs, pl - 4, c, "right");
            else PutText(f, ch.Categories[i], fs, c, pb + 8, "center");
        }
        double zero = VPos(Math.Clamp(0, vmin, vmax));
        var axisCol = new Rgb(0.75, 0.75, 0.75);
        if (horizontal) f.Line([(zero, pt), (zero, pb)], 0.75, axisCol);
        else f.Line([(pl, zero), (pr, zero)], 0.75, axisCol);

        // Столбцы/полосы (все ряды этих типов делят место в категории)
        var barSeries = ch.Series.Where(s => s.Type is "col" or "bar").ToList();
        if (barSeries.Count > 0)
        {
            bool stacked = barSeries[0].Grouping is "stacked" or "percentStacked", pct = barSeries[0].Grouping == "percentStacked";
            int slots = stacked ? 1 : barSeries.Count;
            double gap = band * 0.4, groupW = band - gap, barW = groupW / slots;
            var posAcc = new double[n]; var negAcc = new double[n];
            for (int si = 0; si < barSeries.Count; si++)
            {
                var s = barSeries[si];
                for (int i = 0; i < n && i < s.Values.Count; i++)
                {
                    double v = s.Values[i];
                    if (pct)
                    {
                        double tot = barSeries.Sum(q => i < q.Values.Count ? Math.Abs(q.Values[i]) : 0);
                        v = tot > 0 ? v / tot * 100 : 0;
                    }
                    double from = 0;
                    if (stacked) { from = v >= 0 ? posAcc[i] : negAcc[i]; if (v >= 0) posAcc[i] += v; else negAcc[i] += v; }
                    double a = VPos(from, s.Secondary), e = VPos(from + v, s.Secondary);
                    double off = horizontal ? pb - band * (i + 1) + gap / 2 + (stacked ? 0 : barW * (slots - 1 - si))
                                            : pl + band * i + gap / 2 + (stacked ? 0 : barW * si);
                    var col = s.PointColors.GetValueOrDefault(i, s.Color ?? Rgb.Black);
                    double x0 = horizontal ? Math.Min(a, e) : off, y0 = horizontal ? off : Math.Min(a, e);
                    double bw = horizontal ? Math.Abs(e - a) : barW, bh = horizontal ? barW : Math.Abs(e - a);
                    if (ch.ThreeD)
                    {
                        // Объёмный столбец: верхняя и правая грани
                        double dd = Math.Min(depth, barW * 0.6);
                        f.Polygon([(x0, y0), (x0 + dd, y0 - dd), (x0 + bw + dd, y0 - dd), (x0 + bw, y0)], Lighten(col));
                        f.Polygon([(x0 + bw, y0), (x0 + bw + dd, y0 - dd), (x0 + bw + dd, y0 + bh - dd), (x0 + bw, y0 + bh)], Shade(col, 0.7));
                    }
                    f.Rect(x0, y0, bw, bh, col);
                    if (ch.DataLabels)
                    {
                        if (horizontal) PutText(f, Num(s.Values[i]), fs, e + 3, off + barW / 2);
                        else PutText(f, Num(s.Values[i]), fs, off + barW / 2, e - 7, "center");
                    }
                }
            }
        }

        // Графики, области, точечные (каждый ряд — по своей оси); маркеры поверх линий
        var acc = new Dictionary<string, double[]>();
        foreach (var s in ch.Series.Where(s => s.Type is "line" or "area" or "scatter"))
        {
            var pts = new List<(double X, double Y)>();
            string key = s.Type + "/" + s.Grouping + "/" + s.Secondary;
            if (!acc.TryGetValue(key, out var sum)) acc[key] = sum = new double[n];
            bool stacked = s.Grouping is "stacked" or "percentStacked";
            for (int i = 0; i < n && i < s.Values.Count; i++)
            {
                double v = s.Values[i];
                if (stacked) { sum[i] += v; v = sum[i]; }
                double px = horizontal ? VPos(v, s.Secondary) : pl + band * (i + 0.5), py = horizontal ? pb - band * (i + 0.5) : VPos(v, s.Secondary);
                pts.Add((px, py));
            }
            if (pts.Count == 0) continue;
            var col = s.Color ?? Rgb.Black;
            if (s.Type == "area")
            {
                double z = VPos(Math.Clamp(0, s.Secondary ? sec.Min : vmin, s.Secondary ? sec.Max : vmax), s.Secondary);
                var poly = new List<(double X, double Y)>(pts) { (pts[^1].X, z), (pts[0].X, z) };
                f.Polygon(poly, col);
            }
            else
            {
                if (s.Line && pts.Count > 1) f.Line(pts, 2.25, col);
                if (s.Marker != null || pts.Count == 1)
                    foreach (var (px, py) in pts) f.Marker(s.Marker ?? "square", px, py, s.MarkerSize, col);
            }
            if (ch.DataLabels)
                for (int i = 0; i < pts.Count; i++) PutText(f, Num(s.Values[i]), fs, pts[i].X, pts[i].Y - 8, "center");
        }
    }

    static Rgb Lighten(Rgb c) => new(c.R + (1 - c.R) * 0.3, c.G + (1 - c.G) * 0.3, c.B + (1 - c.B) * 0.3);
}

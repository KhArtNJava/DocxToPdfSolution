using System.Globalization;
using System.Text;

namespace DocxToPdf.Docx;

/// <summary>
/// Контуры VML (атрибут path фигуры или v:shapetype): команды m, l, c, x, e, t, r, v, nf, ns, ae, al, at, ar, wa, wr,
/// qx, qy, qb; координаты — в системе coordsize/coordorigin, значения @n — формулы v:formulas, #n — параметры adj.
/// Результат — подконтуры в долях рамки фигуры.
/// </summary>
static class VmlPath
{
    public static List<CustomPath>? Parse(string path, double cw, double ch, double ox, double oy, double[] formulas)
    {
        if (cw == 0) cw = 1000; if (ch == 0) ch = 1000;
        var result = new List<CustomPath>();
        var cur = new CustomPath();
        double x = 0, y = 0, sx = 0, sy = 0;
        bool hasCmd = false;
        (double, double) N(double px, double py) => ((px - ox) / cw, (py - oy) / ch);
        void Flush()
        {
            if (cur.Commands.Count > 0) result.Add(cur);
            cur = new CustomPath { Fill = cur.Fill, Stroke = cur.Stroke };
        }
        void Move(double px, double py) { var (a, b) = N(px, py); cur.Commands.Add(('M', [a, b])); x = px; y = py; sx = px; sy = py; hasCmd = true; }
        void Line(double px, double py) { if (!hasCmd) Move(x, y); var (a, b) = N(px, py); cur.Commands.Add(('L', [a, b])); x = px; y = py; }
        void Curve(double x1, double y1, double x2, double y2, double x3, double y3)
        {
            if (!hasCmd) Move(x, y);
            var (a1, b1) = N(x1, y1); var (a2, b2) = N(x2, y2); var (a3, b3) = N(x3, y3);
            cur.Commands.Add(('C', [a1, b1, a2, b2, a3, b3]));
            x = x3; y = y3;
        }
        void Arc(double cx, double cy, double rx, double ry, double a0, double sweep, bool lineTo)
        {
            double px = cx + rx * Math.Cos(a0), py = cy + ry * Math.Sin(a0);
            if (lineTo && hasCmd) Line(px, py); else Move(px, py);
            int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 2)));
            double step = sweep / segs;
            for (int i = 0; i < segs; i++)
            {
                double t0 = a0 + i * step, t1 = t0 + step, k = 4.0 / 3 * Math.Tan(step / 4);
                double x0 = cx + rx * Math.Cos(t0), y0 = cy + ry * Math.Sin(t0), x3 = cx + rx * Math.Cos(t1), y3 = cy + ry * Math.Sin(t1);
                Curve(x0 - k * rx * Math.Sin(t0), y0 + k * ry * Math.Cos(t0), x3 + k * rx * Math.Sin(t1), y3 - k * ry * Math.Cos(t1), x3, y3);
            }
        }

        // Разбор на команды и числа
        int p = 0;
        var nums = new List<double>();
        string? cmd = null;
        void Exec()
        {
            if (cmd == null) return;
            var a = nums;
            double A(int i) => i < a.Count ? a[i] : 0;
            switch (cmd)
            {
                case "m": for (int i = 0; i + 1 < Math.Max(2, a.Count); i += 2) { if (i == 0) Move(A(0), A(1)); else Line(A(i), A(i + 1)); } break;
                case "t": for (int i = 0; i + 1 < Math.Max(2, a.Count); i += 2) { if (i == 0) Move(x + A(0), y + A(1)); else Line(x + A(i), y + A(i + 1)); } break;
                case "l": for (int i = 0; i + 1 < a.Count; i += 2) Line(A(i), A(i + 1)); break;
                case "r": for (int i = 0; i + 1 < a.Count; i += 2) Line(x + A(i), y + A(i + 1)); break;
                case "c": for (int i = 0; i + 5 < a.Count; i += 6) Curve(A(i), A(i + 1), A(i + 2), A(i + 3), A(i + 4), A(i + 5)); break;
                case "v":
                    for (int i = 0; i + 5 < a.Count; i += 6)
                    {
                        double bx = x, by = y;
                        Curve(bx + A(i), by + A(i + 1), bx + A(i + 2), by + A(i + 3), bx + A(i + 4), by + A(i + 5));
                    }
                    break;
                case "x": if (hasCmd) { cur.Commands.Add(('Z', [])); x = sx; y = sy; } break;
                case "e": Flush(); hasCmd = false; break;
                case "nf": cur.Fill = false; break;
                case "ns": cur.Stroke = false; break;
                case "ae" or "al":
                    for (int i = 0; i + 5 < a.Count; i += 6)
                    {
                        double st = A(i + 4) / 65536 * Math.PI / 180, sw = A(i + 5) / 65536 * Math.PI / 180;
                        // Углы VML отсчитываются против часовой стрелки (ось y вниз — со знаком минус)
                        Arc(A(i), A(i + 1), A(i + 2), A(i + 3), -st, -sw, cmd == "ae");
                    }
                    break;
                case "at" or "ar" or "wa" or "wr":
                    for (int i = 0; i + 7 < a.Count; i += 8)
                    {
                        double l = A(i), t = A(i + 1), r = A(i + 2), b = A(i + 3);
                        double ecx = (l + r) / 2, ecy = (t + b) / 2, rx = Math.Abs(r - l) / 2, ry = Math.Abs(b - t) / 2;
                        if (rx <= 0 || ry <= 0) continue;
                        double a0 = Math.Atan2((A(i + 5) - ecy) / ry, (A(i + 4) - ecx) / rx), a1 = Math.Atan2((A(i + 7) - ecy) / ry, (A(i + 6) - ecx) / rx);
                        double sweep = a1 - a0;
                        bool clockwise = cmd is "wa" or "wr";
                        if (clockwise) { while (sweep <= 0) sweep += 2 * Math.PI; } else { while (sweep >= 0) sweep -= 2 * Math.PI; }
                        Arc(ecx, ecy, rx, ry, a0, sweep, cmd is "at" or "wa");
                    }
                    break;
                case "qx" or "qy":
                {
                    bool horiz = cmd == "qx";
                    const double K = 0.5523;
                    for (int i = 0; i + 1 < a.Count; i += 2)
                    {
                        double x1 = A(i), y1 = A(i + 1);
                        if (horiz) Curve(x + K * (x1 - x), y, x1, y1 + K * (y - y1), x1, y1);
                        else Curve(x, y + K * (y1 - y), x1 + K * (x - x1), y1, x1, y1);
                        horiz = !horiz;
                    }
                    break;
                }
                case "qb":
                {
                    // Квадратичная B-сплайн: контрольные точки, последняя — конечная
                    var pts = new List<(double, double)>();
                    for (int i = 0; i + 1 < a.Count; i += 2) pts.Add((A(i), A(i + 1)));
                    if (pts.Count == 0) break;
                    if (!hasCmd) { Move(pts[0].Item1, pts[0].Item2); pts.RemoveAt(0); }
                    for (int i = 0; i < pts.Count - 1; i++)
                    {
                        var (cx, cy) = pts[i];
                        var end = i + 2 < pts.Count ? ((cx + pts[i + 1].Item1) / 2, (cy + pts[i + 1].Item2) / 2) : pts[i + 1];
                        Curve(x + 2.0 / 3 * (cx - x), y + 2.0 / 3 * (cy - y), end.Item1 + 2.0 / 3 * (cx - end.Item1), end.Item2 + 2.0 / 3 * (cy - end.Item2), end.Item1, end.Item2);
                    }
                    break;
                }
            }
            nums = [];
        }
        bool pendingEmpty = true;   // пустое значение между запятыми = 0
        bool afterComma = false;
        while (p < path.Length)
        {
            char c = path[p];
            if (char.IsLetter(c))
            {
                // Пустое значение после последней запятой («r21600,x») — тоже 0
                if (afterComma) nums.Add(0);
                afterComma = false;
                Exec();
                // Двухбуквенные команды
                string two = p + 1 < path.Length ? path.Substring(p, 2) : "";
                if (two is "nf" or "ns" or "ae" or "al" or "at" or "ar" or "wa" or "wr" or "qx" or "qy" or "qb" or "ha" or "hb" or "hc" or "hd" or "he" or "hf" or "hg" or "hh" or "hi")
                { cmd = two; p += 2; }
                else { cmd = c.ToString(); p++; }
                pendingEmpty = true;
                continue;
            }
            if (c == ',')
            {
                if (pendingEmpty) nums.Add(0);
                pendingEmpty = true;
                afterComma = true;
                p++;
                continue;
            }
            if (!char.IsWhiteSpace(c)) afterComma = false;
            if (char.IsWhiteSpace(c)) { p++; continue; }
            if (c == '@')
            {
                int q = ++p;
                while (p < path.Length && char.IsDigit(path[p])) p++;
                int fi = int.Parse(path.AsSpan(q, p - q), CultureInfo.InvariantCulture);
                nums.Add(fi < formulas.Length ? formulas[fi] : 0);
                pendingEmpty = false;
                continue;
            }
            if (c is '-' or '+' or '.' || char.IsDigit(c))
            {
                int q = p++;
                while (p < path.Length && (char.IsDigit(path[p]) || path[p] == '.')) p++;
                nums.Add(double.TryParse(path.AsSpan(q, p - q), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0);
                pendingEmpty = false;
                continue;
            }
            p++;
        }
        if (afterComma) nums.Add(0);
        Exec();
        Flush();
        return result.Count > 0 ? result : null;
    }

    /// <summary>Вычисление формул v:formulas (eqn) по порядку; #n — параметры adj, @n — ранее вычисленные формулы.</summary>
    public static double[] EvalFormulas(IReadOnlyList<string> eqns, double[] adj, double cw, double ch, double ox, double oy)
    {
        var res = new double[eqns.Count];
        for (int i = 0; i < eqns.Count; i++)
        {
            var parts = eqns[i].Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) continue;
            double V(int k)
            {
                if (k >= parts.Length) return 0;
                string a = parts[k];
                if (a.StartsWith('#')) { int n = int.Parse(a[1..], CultureInfo.InvariantCulture); return n < adj.Length ? adj[n] : 0; }
                if (a.StartsWith('@')) { int n = int.Parse(a[1..], CultureInfo.InvariantCulture); return n < i ? res[n] : 0; }
                return a switch
                {
                    "width" => cw, "height" => ch, "xcenter" => ox + cw / 2, "ycenter" => oy + ch / 2, "xrange" => cw, "yrange" => ch,
                    "xlimo" or "ylimo" => 0, "pixelWidth" => cw, "pixelHeight" => ch, "pixelLineWidth" => 1, "emuWidth" => cw * 12700, "emuHeight" => ch * 12700,
                    "emuWidth2" => cw * 6350, "emuHeight2" => ch * 6350, "lineDrawn" or "hasStroke" or "hasFill" => 1,
                    _ => double.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0
                };
            }
            double a1 = V(1), b = V(2), c = V(3);
            const double Fd = 65536 * 180 / Math.PI;   // радианы → градусы × 65536
            res[i] = parts[0] switch
            {
                "val" => a1,
                "sum" => a1 + b - c,
                "prod" => c != 0 ? a1 * b / c : 0,
                "mid" => (a1 + b) / 2,
                "abs" => Math.Abs(a1),
                "min" => Math.Min(a1, b),
                "max" => Math.Max(a1, b),
                "if" => a1 > 0 ? b : c,
                "mod" => Math.Sqrt(a1 * a1 + b * b + c * c),
                "atan2" => Math.Atan2(b, a1) * Fd,
                "sin" => a1 * Math.Sin(b / Fd),
                "cos" => a1 * Math.Cos(b / Fd),
                "tan" => a1 * Math.Tan(b / Fd),
                "cosatan2" => a1 * Math.Cos(Math.Atan2(c, b)),
                "sinatan2" => a1 * Math.Sin(Math.Atan2(c, b)),
                "sqrt" => Math.Sqrt(Math.Max(0, a1)),
                "sumangle" => a1 + b * 65536 - c * 65536,
                "ellipse" => b != 0 ? c * Math.Sqrt(Math.Max(0, 1 - a1 / b * (a1 / b))) : 0,
                _ => 0
            };
        }
        return res;
    }
}

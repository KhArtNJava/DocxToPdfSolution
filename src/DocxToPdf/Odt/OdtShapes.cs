using System.Globalization;
using System.Xml.Linq;
using DocxToPdf.Docx;

namespace DocxToPdf.Odt;

/// <summary>
/// Геометрия фигур ODF (draw:enhanced-geometry): формулы draw:equation (арифметика, функции, ссылки ?fN,
/// модификаторы $N, ключевые слова width/height/…), контур draw:enhanced-path со всеми командами
/// (M L C Q Z N F S, дуги T U A B W V, четверти эллипса X Y, G — дуга OOXML) → нормализованные пути.
/// </summary>
static class EnhancedGeometry
{
    static readonly XNamespace Dr = OdtReader.Dr, Svg = OdtReader.Svg;
    static readonly XNamespace DrOoo = "http://openoffice.org/2010/draw";

    public static List<CustomPath>? Build(XElement geo)
    {
        string? path = (string?)geo.Attribute(DrOoo + "enhanced-path") ?? (string?)geo.Attribute(Dr + "enhanced-path");
        if (string.IsNullOrWhiteSpace(path)) return null;
        var vb = ((string?)geo.Attribute(Svg + "viewBox") ?? "0 0 21600 21600").Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(t => double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0).ToArray();
        if (vb.Length != 4 || vb[2] <= 0 || vb[3] <= 0) vb = [0, 0, 21600, 21600];
        // Подобласти OOXML-фигур (drawooo:sub-view-size): координаты пути — в своих единицах на каждый подконтур
        var subViews = new List<(double W, double H)>();
        if ((string?)geo.Attribute(DrOoo + "sub-view-size") is { } svs)
        {
            var nums = svs.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(t => double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0).ToArray();
            for (int i = 0; i + 1 < nums.Length; i += 2) subViews.Add((nums[i], nums[i + 1]));
        }
        var mods = ((string?)geo.Attribute(Dr + "modifiers") ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(t => double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0).ToArray();
        var eqs = geo.Elements(Dr + "equation").ToDictionary(e => (string?)e.Attribute(Dr + "name") ?? "", e => (string?)e.Attribute(Dr + "formula") ?? "0");
        var ev = new Evaluator(eqs, mods, vb[0], vb[1], vb[2], vb[3]);
        try { return Interpret(path, ev, vb, subViews); }
        catch { return null; }
    }

    // ---------------- Формулы ----------------

    sealed class Evaluator(Dictionary<string, string> eqs, double[] mods, double x0, double y0, double vbW, double vbH)
    {
        readonly Dictionary<string, double> _cache = [];
        readonly HashSet<string> _busy = [];
        readonly double _right = x0 + vbW, _bottom = y0 + vbH;
        public double LogW = vbW, LogH = vbH;

        public double Eq(string name)
        {
            if (_cache.TryGetValue(name, out var v)) return v;
            if (!eqs.TryGetValue(name, out var f) || !_busy.Add(name)) return 0;
            v = Eval(f);
            _busy.Remove(name);
            return _cache[name] = v;
        }

        public double Value(string token)
        {
            if (token.StartsWith('?')) return Eq(token[1..]);
            if (token.StartsWith('$')) return int.TryParse(token[1..], out var i) && i < mods.Length ? mods[i] : 0;
            if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var d)) return d;
            return Keyword(token) ?? 0;
        }

        double? Keyword(string k) => k switch
        {
            "pi" => Math.PI,
            "left" => x0, "top" => y0, "right" => _right, "bottom" => _bottom,
            "width" or "logwidth" => LogW, "height" or "logheight" => LogH,
            "xstretch" or "ystretch" => 0,
            "hasstroke" or "hasfill" => 1,
            _ => null,
        };

        string _s = "";
        int _p;

        public double Eval(string f)
        {
            var saveS = _s; var saveP = _p;
            _s = f; _p = 0;
            double r = Expr();
            _s = saveS; _p = saveP;
            return double.IsFinite(r) ? r : 0;
        }

        void Ws() { while (_p < _s.Length && char.IsWhiteSpace(_s[_p])) _p++; }

        double Expr()
        {
            double v = Term();
            while (true)
            {
                Ws();
                if (_p < _s.Length && _s[_p] == '+') { _p++; v += Term(); }
                else if (_p < _s.Length && _s[_p] == '-') { _p++; v -= Term(); }
                else return v;
            }
        }

        double Term()
        {
            double v = Unary();
            while (true)
            {
                Ws();
                if (_p < _s.Length && _s[_p] == '*') { _p++; v *= Unary(); }
                else if (_p < _s.Length && _s[_p] == '/') { _p++; double d = Unary(); v = d != 0 ? v / d : 0; }
                else return v;
            }
        }

        double Unary()
        {
            Ws();
            if (_p < _s.Length && _s[_p] == '-') { _p++; return -Unary(); }
            if (_p < _s.Length && _s[_p] == '+') { _p++; return Unary(); }
            return Primary();
        }

        double Primary()
        {
            Ws();
            if (_p >= _s.Length) return 0;
            char c = _s[_p];
            if (c == '(') { _p++; double v = Expr(); Ws(); if (_p < _s.Length && _s[_p] == ')') _p++; return v; }
            if (c == '?' || c == '$')
            {
                int st = _p++;
                while (_p < _s.Length && (char.IsLetterOrDigit(_s[_p]) || _s[_p] == '_')) _p++;
                return Value(_s[st.._p]);
            }
            if (char.IsDigit(c) || c == '.')
            {
                int st = _p;
                while (_p < _s.Length && (char.IsDigit(_s[_p]) || _s[_p] == '.' || (_s[_p] is 'e' or 'E' && _p + 1 < _s.Length && (char.IsDigit(_s[_p + 1]) || _s[_p + 1] == '-')))) { if (_s[_p] is 'e' or 'E') _p++; _p++; }
                return double.TryParse(_s[st.._p], NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;
            }
            if (char.IsLetter(c))
            {
                int st = _p;
                while (_p < _s.Length && char.IsLetterOrDigit(_s[_p])) _p++;
                string name = _s[st.._p];
                Ws();
                if (_p < _s.Length && _s[_p] == '(')
                {
                    _p++;
                    var args = new List<double>();
                    Ws();
                    if (_p < _s.Length && _s[_p] == ')') _p++;
                    else
                    {
                        while (true)
                        {
                            args.Add(Expr());
                            Ws();
                            if (_p < _s.Length && _s[_p] == ',') { _p++; continue; }
                            if (_p < _s.Length && _s[_p] == ')') _p++;
                            break;
                        }
                    }
                    double A(int i) => i < args.Count ? args[i] : 0;
                    return name switch
                    {
                        "abs" => Math.Abs(A(0)), "sqrt" => Math.Sqrt(Math.Max(0, A(0))),
                        "sin" => Math.Sin(A(0)), "cos" => Math.Cos(A(0)), "tan" => Math.Tan(A(0)), "atan" => Math.Atan(A(0)),
                        "atan2" => Math.Atan2(A(0), A(1)),
                        "min" => Math.Min(A(0), A(1)), "max" => Math.Max(A(0), A(1)),
                        "if" => A(0) > 0 ? A(1) : A(2),
                        _ => 0,
                    };
                }
                return Keyword(name) ?? 0;
            }
            _p++;
            return 0;
        }
    }

    // ---------------- Контур ----------------

    static List<CustomPath> Interpret(string path, Evaluator ev, double[] vb, List<(double W, double H)> subViews)
    {
        var toks = Tokenize(path);
        var res = new List<CustomPath>();
        var cur = new CustomPath();
        int sub = 0;
        double bx = vb[0], by = vb[1], bw = vb[2], bh = vb[3];
        void SetView()
        {
            if (sub < subViews.Count && subViews[sub].W > 0 && subViews[sub].H > 0) { bx = 0; by = 0; bw = subViews[sub].W; bh = subViews[sub].H; ev.LogW = bw; ev.LogH = bh; }
        }
        SetView();
        double cx = 0, cy = 0, sx = 0, sy = 0;
        int i = 0;
        char cmd = 'M';
        bool quadX = true;
        bool IsNum() => i < toks.Count && !(toks[i].Length == 1 && char.IsLetter(toks[i][0]));
        double V() => i < toks.Count ? ev.Value(toks[i++]) : 0;
        void Add(char c, params double[] abs)
        {
            var r = new double[abs.Length];
            for (int k = 0; k < abs.Length; k += 2) { r[k] = (abs[k] - bx) / bw; r[k + 1] = (abs[k + 1] - by) / bh; }
            cur.Commands.Add((c, r));
        }
        void MoveTo(double x, double y) { if (cur.Commands.Count > 0 && cur.Commands[^1].Cmd == 'M') cur.Commands.RemoveAt(cur.Commands.Count - 1); Add('M', x, y); cx = sx = x; cy = sy = y; }
        void LineTo(double x, double y) { if (cur.Commands.Count == 0) MoveTo(cx, cy); Add('L', x, y); cx = x; cy = y; }
        // Дуга эллипса (центр, радиусы, углы в радианах, параметрически; y вниз) кубическими кривыми
        void Arc(double ocx, double ocy, double rx, double ry, double a0, double a1, bool move)
        {
            double px = ocx + rx * Math.Cos(a0), py = ocy + ry * Math.Sin(a0);
            if (move || cur.Commands.Count == 0) MoveTo(px, py); else if (Math.Abs(px - cx) > 1e-9 || Math.Abs(py - cy) > 1e-9) LineTo(px, py);
            double sweep = a1 - a0;
            int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 2) - 1e-9));
            double d = sweep / segs;
            double k = 4.0 / 3 * Math.Tan(d / 4);
            double a = a0;
            for (int s = 0; s < segs; s++)
            {
                double b = a + d;
                double x1 = ocx + rx * (Math.Cos(a) - k * Math.Sin(a)), y1 = ocy + ry * (Math.Sin(a) + k * Math.Cos(a));
                double x2 = ocx + rx * (Math.Cos(b) + k * Math.Sin(b)), y2 = ocy + ry * (Math.Sin(b) - k * Math.Cos(b));
                double x3 = ocx + rx * Math.Cos(b), y3 = ocy + ry * Math.Sin(b);
                Add('C', x1, y1, x2, y2, x3, y3);
                a = b;
            }
            cx = ocx + rx * Math.Cos(a1); cy = ocy + ry * Math.Sin(a1);
        }
        while (i < toks.Count)
        {
            if (!IsNum()) { cmd = toks[i++][0]; quadX = cmd == 'X'; }
            switch (cmd)
            {
                case 'M': MoveTo(V(), V()); cmd = 'L'; break;
                case 'L': LineTo(V(), V()); break;
                case 'C': { double x1 = V(), y1 = V(), x2 = V(), y2 = V(), x = V(), y = V(); if (cur.Commands.Count == 0) MoveTo(cx, cy); Add('C', x1, y1, x2, y2, x, y); cx = x; cy = y; break; }
                case 'Q':
                {
                    double qx = V(), qy = V(), x = V(), y = V();
                    if (cur.Commands.Count == 0) MoveTo(cx, cy);
                    Add('C', cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy), x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y);
                    cx = x; cy = y;
                    break;
                }
                case 'Z': if (cur.Commands.Count > 0) cur.Commands.Add(('Z', [])); cx = sx; cy = sy; break;
                case 'N':
                    if (cur.Commands.Count > 0) res.Add(cur);
                    cur = new CustomPath { Fill = cur.Fill, Stroke = cur.Stroke };
                    cur.Fill = true; cur.Stroke = true;
                    sub++; SetView();
                    break;
                case 'F': cur.Fill = false; break;
                case 'S': cur.Stroke = false; break;
                case 'T' or 'U':
                {
                    // Эллиптическая дуга по углам (градусы, против часовой стрелки на экране)
                    double ocx = V(), ocy = V(), rx = V(), ry = V(), t0 = V(), t1 = V();
                    if (t1 <= t0) t1 += 360;
                    Arc(ocx, ocy, rx, ry, -t0 * Math.PI / 180, -t1 * Math.PI / 180, cmd == 'U');
                    break;
                }
                case 'A' or 'B' or 'W' or 'V':
                {
                    double x1 = V(), y1 = V(), x2 = V(), y2 = V(), x3 = V(), y3 = V(), x4 = V(), y4 = V();
                    double ocx = (x1 + x2) / 2, ocy = (y1 + y2) / 2, rx = Math.Abs(x2 - x1) / 2, ry = Math.Abs(y2 - y1) / 2;
                    if (rx <= 0 || ry <= 0) { LineTo(x4, y4); break; }
                    double a0 = Math.Atan2((y3 - ocy) / ry, (x3 - ocx) / rx), a1 = Math.Atan2((y4 - ocy) / ry, (x4 - ocx) / rx);
                    bool clockwise = cmd is 'W' or 'V';
                    // На экране (y вниз) «по часовой» — увеличение угла
                    if (clockwise) { while (a1 <= a0) a1 += 2 * Math.PI; }
                    else { while (a1 >= a0) a1 -= 2 * Math.PI; }
                    Arc(ocx, ocy, rx, ry, a0, a1, cmd is 'B' or 'V');
                    break;
                }
                case 'X' or 'Y':
                {
                    double x = V(), y = V();
                    if (cur.Commands.Count == 0) MoveTo(cx, cy);
                    const double kq = 0.5522847498;
                    if (quadX) Add('C', cx + kq * (x - cx), cy, x, y + kq * (cy - y), x, y);
                    else Add('C', cx, cy + kq * (y - cy), x + kq * (cx - x), y, x, y);
                    cx = x; cy = y;
                    quadX = !quadX;
                    break;
                }
                case 'G':
                {
                    // Дуга OOXML arcTo: радиусы и углы (градусы, по часовой), начало — текущая точка
                    double wr = V(), hr = V(), st = V() * Math.PI / 180, sw = V() * Math.PI / 180;
                    double ocx = cx - wr * Math.Cos(st), ocy = cy - hr * Math.Sin(st);
                    Arc(ocx, ocy, wr, hr, st, st + sw, false);
                    break;
                }
                case 'H' or 'I' or 'J' or 'K': break;   // затемнение/осветление заливки
                default: if (IsNum()) i++; break;
            }
        }
        if (cur.Commands.Count > 0) res.Add(cur);
        return res.Where(p => p.Commands.Count > 1).ToList();
    }

    static List<string> Tokenize(string path)
    {
        var res = new List<string>();
        int i = 0;
        while (i < path.Length)
        {
            char c = path[i];
            if (char.IsWhiteSpace(c) || c == ',') { i++; continue; }
            if (char.IsLetter(c) && (i + 1 >= path.Length || !char.IsLetterOrDigit(path[i + 1])))
            {
                res.Add(c.ToString()); i++; continue;
            }
            int st = i;
            if (c is '?' or '$') i++;
            while (i < path.Length && !char.IsWhiteSpace(path[i]) && path[i] != ',') i++;
            res.Add(path[st..i]);
        }
        return res;
    }
}

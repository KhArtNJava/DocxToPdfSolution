using System.Buffers.Binary;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;

namespace DocxToPdf.Pdf;

/// <summary>
/// Векторный вывод метафайлов Windows (EMF и WMF, в т. ч. с заголовком Aldus Placeable):
/// записи GDI переводятся в контуры, текст и растры PDF. EMF+ (GDI+) не разбирается — в «двойных» EMF
/// выводится их GDI-часть.
/// </summary>
public static class Metafile
{
    const uint PlaceableKey = 0x9AC6CDD7;

    // Кодовые страницы Windows (cp1251 и др.) для текста WMF/EMF в однобайтовых кодировках
    static Metafile() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    public static bool IsMetafile(byte[] d) => Kind(d) != 0;

    /// <summary>1 — EMF, 2 — WMF с заголовком Placeable, 3 — WMF без него, 0 — не метафайл.</summary>
    static int Kind(byte[] d)
    {
        if (d.Length >= 88 && BinaryPrimitives.ReadUInt32LittleEndian(d) == 1 && BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(40)) == 0x464D4520) return 1;
        if (d.Length >= 40 && BinaryPrimitives.ReadUInt32LittleEndian(d) == PlaceableKey) return 2;
        if (d.Length >= 18 && BinaryPrimitives.ReadUInt16LittleEndian(d) is 1 or 2 && BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(2)) == 9
            && BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(4)) is 0x0100 or 0x0300) return 3;
        return 0;
    }

    /// <summary>Собственный размер метафайла, pt.</summary>
    public static (double W, double H)? Size(byte[] d)
    {
        switch (Kind(d))
        {
            case 1:
            {
                // rclFrame — в сотых долях миллиметра
                int l = I32(d, 24), t = I32(d, 28), r = I32(d, 32), b = I32(d, 36);
                return (Math.Abs(r - l) / 100.0 / 25.4 * 72, Math.Abs(b - t) / 100.0 / 25.4 * 72);
            }
            case 2:
            {
                int l = I16(d, 6), t = I16(d, 8), r = I16(d, 10), b = I16(d, 12), inch = U16(d, 14);
                if (inch <= 0) inch = 1440;
                return (Math.Abs(r - l) * 72.0 / inch, Math.Abs(b - t) * 72.0 / inch);
            }
            case 3:
                return (Math.Max(1, (double)WmfWindow(d, 0).W) * 0.75, Math.Max(1, (double)WmfWindow(d, 0).H) * 0.75);
        }
        return null;
    }

    /// <summary>Рисует метафайл в прямоугольник (x, y, w, h) страницы (ось y вниз).</summary>
    public static void Render(PdfCanvas c, byte[] d, double x, double y, double w, double h)
    {
        if (w <= 0 || h <= 0) return;
        var player = new Player(c, d, x, y, w, h);
        // Всё содержимое обрезается по рамке картинки
        var clip = c.NewPath();
        clip.Rect(x, y, w, h);
        c.BeginClip(clip, false);
        try
        {
            switch (Kind(d))
            {
                case 1: player.PlayEmf(); break;
                case 2: player.PlayWmf(22); break;
                case 3: player.PlayWmf(0); break;
            }
        }
        catch { /* повреждённая запись — выводим то, что успели */ }
        c.EndClip();
    }

    static int I32(byte[] d, int o) => o + 4 <= d.Length ? BinaryPrimitives.ReadInt32LittleEndian(d.AsSpan(o)) : 0;
    static uint U32(byte[] d, int o) => o + 4 <= d.Length ? BinaryPrimitives.ReadUInt32LittleEndian(d.AsSpan(o)) : 0;
    static int I16(byte[] d, int o) => o + 2 <= d.Length ? BinaryPrimitives.ReadInt16LittleEndian(d.AsSpan(o)) : 0;
    static int U16(byte[] d, int o) => o + 2 <= d.Length ? BinaryPrimitives.ReadUInt16LittleEndian(d.AsSpan(o)) : 0;
    static float F32(byte[] d, int o) => o + 4 <= d.Length ? BinaryPrimitives.ReadSingleLittleEndian(d.AsSpan(o)) : 0;

    /// <summary>Окно WMF без заголовка Placeable: первые SETWINDOWORG/EXT.</summary>
    static (int X, int Y, int W, int H) WmfWindow(byte[] d, int start)
    {
        int pos = start + U16(d, start + 2) * 2, ox = 0, oy = 0, ew = 0, eh = 0;
        while (pos + 6 <= d.Length)
        {
            int size = (int)U32(d, pos) * 2, fn = U16(d, pos + 4);
            if (size < 6 || fn == 0) break;
            if (fn == 0x020B) { oy = I16(d, pos + 6); ox = I16(d, pos + 8); }
            if (fn == 0x020C) { eh = I16(d, pos + 6); ew = I16(d, pos + 8); break; }
            pos += size;
        }
        return (ox, oy, ew, eh);
    }

    // ================================================================
    //  Объекты и состояние контекста GDI
    // ================================================================

    sealed class Pen { public int Style; public double Width; public Rgb Color; public bool Cosmetic; }
    sealed class Brush { public int Style; public Rgb Color; public int Hatch; }
    sealed class Font { public double Height; public int Escapement; public int Weight; public bool Italic, Underline, Strike; public int Charset; public string Face = "Arial"; }

    sealed class State
    {
        public Pen? Pen = new() { Style = 0, Width = 0, Color = Rgb.Black, Cosmetic = true };
        public Brush? Brush = new() { Style = 0, Color = Rgb.White };
        public Font Font = new() { Height = 12, Face = "Arial" };
        public Rgb TextColor = Rgb.Black, BkColor = Rgb.White;
        public int BkMode = 2, TextAlign, PolyFill = 1, MapMode = 1, ArcDir = 1;
        public double WinOrgX, WinOrgY, WinExtX = 1, WinExtY = 1, VpOrgX, VpOrgY, VpExtX = 1, VpExtY = 1;
        public bool WinSet;
        // Мировое преобразование EMF: x' = m11 x + m21 y + dx, y' = m12 x + m22 y + dy
        public double M11 = 1, M12, M21, M22 = 1, Dx, Dy;
        public double CurX, CurY;
        public List<(double X, double Y, double W, double H)>? Clip;   // пересечение прямоугольников (в координатах страницы)
        public PdfCanvas.VectorPath? ClipPath;
        public State Clone() => (State)MemberwiseClone();
    }

    sealed class Player
    {
        readonly PdfCanvas _c;
        readonly byte[] _d;
        readonly double _x, _y, _w, _h;
        State _s = new();
        readonly Stack<State> _saved = new();
        readonly Dictionary<int, object> _objects = [];
        // EMF: соответствие «устройство → миллиметры» и рамка
        double _mmPerPxX = 0.264583, _mmPerPxY = 0.264583, _frameL, _frameT, _frameW = 1, _frameH = 1;
        bool _emf;
        // WMF Placeable: прямоугольник в логических единицах
        double _bbL, _bbT, _bbW = 1, _bbH = 1;
        PdfCanvas.VectorPath? _path;     // строящийся контур (BEGINPATH … ENDPATH)
        bool _inPath;
        PdfCanvas.VectorPath? _lastPath;
        bool _figureOpen;

        public Player(PdfCanvas c, byte[] d, double x, double y, double w, double h) { _c = c; _d = d; _x = x; _y = y; _w = w; _h = h; }

        // ---------------- Преобразование координат ----------------

        (double X, double Y) LogicalToDevice(double lx, double ly)
        {
            var s = _s;
            if (_emf)
            {
                double px = s.M11 * lx + s.M21 * ly + s.Dx, py = s.M12 * lx + s.M22 * ly + s.Dy;
                lx = px; ly = py;
            }
            switch (s.MapMode)
            {
                case 7 or 8:
                {
                    double sx = s.WinExtX != 0 ? s.VpExtX / s.WinExtX : 1, sy = s.WinExtY != 0 ? s.VpExtY / s.WinExtY : 1;
                    if (s.MapMode == 7)
                    {
                        double m = Math.Min(Math.Abs(sx), Math.Abs(sy));
                        sx = Math.Sign(sx) * m; sy = Math.Sign(sy) * m;
                    }
                    return ((lx - s.WinOrgX) * sx + s.VpOrgX, (ly - s.WinOrgY) * sy + s.VpOrgY);
                }
                case 2 or 3 or 4 or 5 or 6:
                {
                    // Метрические режимы: ось y вверх
                    double mm = s.MapMode switch { 2 => 0.1, 3 => 0.01, 4 => 0.254, 5 => 0.0254, _ => 25.4 / 1440 };
                    return ((lx - s.WinOrgX) * mm / _mmPerPxX + s.VpOrgX, -(ly - s.WinOrgY) * mm / _mmPerPxY + s.VpOrgY);
                }
                default:
                    return (lx - s.WinOrgX + s.VpOrgX, ly - s.WinOrgY + s.VpOrgY);
            }
        }

        /// <summary>Логические координаты → координаты страницы PDF (pt, ось y вниз).</summary>
        (double X, double Y) P(double lx, double ly)
        {
            if (_emf)
            {
                var (dx, dy) = LogicalToDevice(lx, ly);
                double mx = dx * _mmPerPxX * 100, my = dy * _mmPerPxY * 100;
                return (_x + (mx - _frameL) / _frameW * _w, _y + (my - _frameT) / _frameH * _h);
            }
            // WMF: текущее окно отображается на рамку картинки
            var s = _s;
            double wx = s.WinSet ? s.WinOrgX : _bbL, wy = s.WinSet ? s.WinOrgY : _bbT;
            double ww = s.WinSet ? s.WinExtX : _bbW, wh = s.WinSet ? s.WinExtY : _bbH;
            if (ww == 0) ww = 1; if (wh == 0) wh = 1;
            return (_x + (lx - wx) / ww * _w, _y + (ly - wy) / wh * _h);
        }

        /// <summary>Масштаб длины (логические единицы → pt) — для толщины пера и размера шрифта.</summary>
        double Scale(bool vertical = false)
        {
            var (x0, y0) = P(0, 0); var (x1, y1) = P(100, 0); var (x2, y2) = P(0, 100);
            double sx = Math.Sqrt((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)) / 100, sy = Math.Sqrt((x2 - x0) * (x2 - x0) + (y2 - y0) * (y2 - y0)) / 100;
            return vertical ? sy : Math.Sqrt(sx * sy);
        }

        double ScaleX()
        {
            var (x0, y0) = P(0, 0); var (x1, y1) = P(100, 0);
            return Math.Sqrt((x1 - x0) * (x1 - x0) + (y1 - y0) * (y1 - y0)) / 100;
        }

        double DeviceScale()
        {
            // Один пиксель устройства, pt (для косметических перьев)
            if (_emf) return _mmPerPxX * 100 / _frameW * _w;
            return 0.75;
        }

        // ---------------- Рисование ----------------

        PdfCanvas.VectorPath Target() => _inPath ? (_path ??= _c.NewPath()) : _c.NewPath();

        void Paint(PdfCanvas.VectorPath p, bool fill, bool stroke, bool close = false)
        {
            if (_inPath) return;   // внутри BEGINPATH только накапливаем
            var s = _s;
            Rgb? fc = fill && s.Brush is { Style: not 1 } br ? br.Color : null;
            bool hatched = fill && s.Brush is { Style: 2 };
            var pen = stroke ? s.Pen : null;
            if (pen is { Style: 5 }) pen = null;
            double lw = 0; double[]? dash = null;
            if (pen != null)
            {
                lw = pen.Cosmetic || pen.Width <= 0 ? Math.Max(0.25, DeviceScale()) : Math.Max(0.1, pen.Width * Scale());
                double u = Math.Max(lw, 0.75);
                dash = (pen.Style & 0xF) switch
                {
                    1 => [u * 6, u * 2], 2 => [u, u * 2], 3 => [u * 6, u * 2, u, u * 2], 4 => [u * 6, u * 2, u, u * 2, u, u * 2], _ => null
                };
            }
            int cap = pen == null ? 0 : (pen.Style & 0xF00) switch { 0x100 => 2, 0x200 => 0, _ => 1 };
            int join = pen == null ? 0 : (pen.Style & 0xF000) switch { 0x1000 => 2, 0x2000 => 0, _ => 1 };
            BeginClipState();
            if (hatched)
            {
                // Штриховка: фон (если непрозрачный режим) и линии внутри контура
                if (s.BkMode == 2) _c.PaintPath(p, s.BkColor, null, 0, s.PolyFill == 1, null, 0, 0);
                _c.BeginClip(p, s.PolyFill == 1);
                Hatch(p, s.Brush!.Hatch, s.Brush.Color);
                _c.EndClip();
                fc = null;
            }
            if (fc != null || pen != null) _c.PaintPath(p, fc, pen?.Color, lw, s.PolyFill == 1, dash, cap, join);
            EndClipState();
        }

        void Hatch(PdfCanvas.VectorPath p, int style, Rgb color)
        {
            var (bx, by, bw, bh) = p.Bounds();
            double step = 4;
            var lines = _c.NewPath();
            if (style is 0 or 4) for (double yy = by; yy <= by + bh; yy += step) { lines.M(bx, yy); lines.L(bx + bw, yy); }
            if (style is 1 or 4) for (double xx = bx; xx <= bx + bw; xx += step) { lines.M(xx, by); lines.L(xx, by + bh); }
            if (style is 2 or 5) for (double k = -bh; k <= bw; k += step) { lines.M(bx + k, by); lines.L(bx + k + bh, by + bh); }
            if (style is 3 or 5) for (double k = 0; k <= bw + bh; k += step) { lines.M(bx + k, by); lines.L(bx + k - bh, by + bh); }
            _c.PaintPath(lines, null, color, 0.5, false, null, 0, 0);
        }

        int _clipDepth;
        void BeginClipState()
        {
            _clipDepth = 0;
            if (_s.Clip is { Count: > 0 } rects)
            {
                foreach (var (cx, cy, cw, ch) in rects)
                {
                    var cp = _c.NewPath(); cp.Rect(cx, cy, Math.Max(0, cw), Math.Max(0, ch));
                    _c.BeginClip(cp, false); _clipDepth++;
                }
            }
            if (_s.ClipPath != null) { _c.BeginClip(_s.ClipPath, _s.PolyFill == 1); _clipDepth++; }
        }

        void EndClipState() { for (int i = 0; i < _clipDepth; i++) _c.EndClip(); _clipDepth = 0; }

        void IntersectClip(double l, double t, double r, double b)
        {
            var (x1, y1) = P(l, t); var (x2, y2) = P(r, b);
            var rect = (Math.Min(x1, x2), Math.Min(y1, y2), Math.Abs(x2 - x1), Math.Abs(y2 - y1));
            _s.Clip = [.. _s.Clip ?? [], rect];
        }

        void Poly(IReadOnlyList<(double X, double Y)> pts, bool close, bool fill)
        {
            if (pts.Count < 2) return;
            var p = Target();
            var (x0, y0) = P(pts[0].X, pts[0].Y);
            p.M(x0, y0);
            for (int i = 1; i < pts.Count; i++) { var (x, y) = P(pts[i].X, pts[i].Y); p.L(x, y); }
            if (close) p.Z();
            Paint(p, fill && close, true);
        }

        void PolyPoly(List<List<(double X, double Y)>> polys, bool close)
        {
            var p = Target();
            foreach (var pts in polys)
            {
                if (pts.Count < 2) continue;
                var (x0, y0) = P(pts[0].X, pts[0].Y);
                p.M(x0, y0);
                for (int i = 1; i < pts.Count; i++) { var (x, y) = P(pts[i].X, pts[i].Y); p.L(x, y); }
                if (close) p.Z();
            }
            Paint(p, close, true);
        }

        void Bezier(IReadOnlyList<(double X, double Y)> pts, bool fromCurrent)
        {
            var p = Target();
            int i = 0;
            if (!fromCurrent)
            {
                if (pts.Count < 1) return;
                var (x0, y0) = P(pts[0].X, pts[0].Y); p.M(x0, y0); i = 1;
                _s.CurX = pts[0].X; _s.CurY = pts[0].Y;
            }
            else if (!_inPath || !_figureOpen) { var (cx, cy) = P(_s.CurX, _s.CurY); p.M(cx, cy); }
            for (; i + 2 < pts.Count; i += 3)
            {
                var a = P(pts[i].X, pts[i].Y); var b = P(pts[i + 1].X, pts[i + 1].Y); var e = P(pts[i + 2].X, pts[i + 2].Y);
                p.C(a.X, a.Y, b.X, b.Y, e.X, e.Y);
                _s.CurX = pts[i + 2].X; _s.CurY = pts[i + 2].Y;
            }
            _figureOpen = true;
            Paint(p, false, true);
        }

        void LineTo(double x, double y)
        {
            var p = Target();
            if (!_inPath || !_figureOpen) { var (cx, cy) = P(_s.CurX, _s.CurY); p.M(cx, cy); }
            var (tx, ty) = P(x, y);
            p.L(tx, ty);
            _figureOpen = true;
            _s.CurX = x; _s.CurY = y;
            Paint(p, false, true);
        }

        void PolyTo(IReadOnlyList<(double X, double Y)> pts)
        {
            var p = Target();
            if (!_inPath || !_figureOpen) { var (cx, cy) = P(_s.CurX, _s.CurY); p.M(cx, cy); }
            foreach (var (x, y) in pts) { var (tx, ty) = P(x, y); p.L(tx, ty); _s.CurX = x; _s.CurY = y; }
            _figureOpen = true;
            Paint(p, false, true);
        }

        void Rect(double l, double t, double r, double b)
        {
            var p = Target();
            var a = P(l, t); var bb = P(r, t); var cc = P(r, b); var dd = P(l, b);
            p.M(a.X, a.Y); p.L(bb.X, bb.Y); p.L(cc.X, cc.Y); p.L(dd.X, dd.Y); p.Z();
            Paint(p, true, true);
        }

        /// <summary>Эллиптическая дуга в логических координатах (углы — в логической системе, ось y вниз).</summary>
        void ArcPath(PdfCanvas.VectorPath p, double cx, double cy, double rx, double ry, double a0, double sweep, bool move)
        {
            int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweep) / (Math.PI / 2)));
            double step = sweep / segs;
            var s0 = P(cx + rx * Math.Cos(a0), cy + ry * Math.Sin(a0));
            if (move) p.M(s0.X, s0.Y); else p.L(s0.X, s0.Y);
            for (int i = 0; i < segs; i++)
            {
                double t0 = a0 + i * step, t1 = t0 + step, k = 4.0 / 3 * Math.Tan(step / 4);
                double x0 = cx + rx * Math.Cos(t0), y0 = cy + ry * Math.Sin(t0), x3 = cx + rx * Math.Cos(t1), y3 = cy + ry * Math.Sin(t1);
                var c1 = P(x0 - k * rx * Math.Sin(t0), y0 + k * ry * Math.Cos(t0));
                var c2 = P(x3 + k * rx * Math.Sin(t1), y3 - k * ry * Math.Cos(t1));
                var e = P(x3, y3);
                p.C(c1.X, c1.Y, c2.X, c2.Y, e.X, e.Y);
            }
        }

        void Ellipse(double l, double t, double r, double b)
        {
            var p = Target();
            ArcPath(p, (l + r) / 2, (t + b) / 2, Math.Abs(r - l) / 2, Math.Abs(b - t) / 2, 0, 2 * Math.PI, true);
            p.Z();
            Paint(p, true, true);
        }

        void RoundRect(double l, double t, double r, double b, double ew, double eh)
        {
            if (l > r) (l, r) = (r, l);
            if (t > b) (t, b) = (b, t);
            double rx = Math.Min(Math.Abs(ew) / 2, (r - l) / 2), ry = Math.Min(Math.Abs(eh) / 2, (b - t) / 2);
            var p = Target();
            ArcPath(p, r - rx, t + ry, rx, ry, -Math.PI / 2, Math.PI / 2, true);
            ArcPath(p, r - rx, b - ry, rx, ry, 0, Math.PI / 2, false);
            ArcPath(p, l + rx, b - ry, rx, ry, Math.PI / 2, Math.PI / 2, false);
            ArcPath(p, l + rx, t + ry, rx, ry, Math.PI, Math.PI / 2, false);
            p.Z();
            Paint(p, true, true);
        }

        /// <summary>ARC/CHORD/PIE: kind 0 — дуга, 1 — хорда, 2 — сектор.</summary>
        void Arc(double l, double t, double r, double b, double xs, double ys, double xe, double ye, int kind, bool lineToStart = false)
        {
            double cx = (l + r) / 2, cy = (t + b) / 2, rx = Math.Abs(r - l) / 2, ry = Math.Abs(b - t) / 2;
            if (rx <= 0 || ry <= 0) return;
            double a0 = Math.Atan2((ys - cy) / ry, (xs - cx) / rx), a1 = Math.Atan2((ye - cy) / ry, (xe - cx) / rx);
            // По умолчанию дуга идёт против часовой стрелки на экране (ось y вниз — уменьшение угла)
            double sweep = a1 - a0;
            if (_s.ArcDir == 1) { while (sweep >= 0) sweep -= 2 * Math.PI; }
            else { while (sweep <= 0) sweep += 2 * Math.PI; }
            var p = Target();
            if (kind == 2) { var c0 = P(cx, cy); p.M(c0.X, c0.Y); ArcPath(p, cx, cy, rx, ry, a0, sweep, false); }
            else if (lineToStart && (_inPath && _figureOpen)) ArcPath(p, cx, cy, rx, ry, a0, sweep, false);
            else ArcPath(p, cx, cy, rx, ry, a0, sweep, true);
            if (kind > 0) p.Z();
            _figureOpen = true;
            Paint(p, kind > 0, true);
        }

        // ---------------- Текст ----------------

        void Text(double lx, double ly, string text, double[]? dx, bool updateCp, (double L, double T, double R, double B)? opaqueRect)
        {
            if (string.IsNullOrEmpty(text) || _c.FontProvider is not { } fm) return;
            var s = _s;
            var f = s.Font;
            double vs = Scale(true), hs = ScaleX();
            double em = Math.Abs(f.Height) * vs;
            // Неравномерное растяжение картинки: текст сжимается/растягивается по ширине (Tz)
            double hscale = vs > 0 && hs > 0 ? Math.Clamp(hs / vs, 0.05, 20) : 1;
            if (f.Height > 0) em *= 0.82;   // положительная высота — высота ячейки (с межстрочным интервалом)
            if (em < 0.5) em = 12 * vs > 0.5 ? 12 * vs : 8;
            bool bold = f.Weight >= 600;
            string family = f.Face.Length > 0 ? f.Face : "Arial";
            var face = fm.Get(family, bold, f.Italic);
            // Глифы с запасными шрифтами для отсутствующих символов
            var runs = new List<(FontFace Face, List<ushort> G, List<string> T, List<double> Adv)>();
            bool symbol = f.Charset == 2;
            double total = 0;
            int ci = 0;
            foreach (var rune in text.EnumerateRunes())
            {
                int cp = rune.Value;
                if (symbol) cp = FontManager.MapSymbolChar(family, cp);
                var useFace = face;
                if (!useFace.Font.TryGetGlyph(cp, out var gid) || gid == 0)
                {
                    var fb = fm.FallbackFor(cp, family, bold, f.Italic);
                    if (fb != null && fb.Font.TryGetGlyph(cp, out var g2)) { useFace = fb; gid = g2; }
                }
                double natural = useFace.Font.Advance(gid) * em * hscale;
                double adv = dx != null && ci < dx.Length ? dx[ci] * hs : natural;
                if (runs.Count == 0 || runs[^1].Face != useFace) runs.Add((useFace, [], [], []));
                runs[^1].G.Add(gid); runs[^1].T.Add(rune.ToString()); runs[^1].Adv.Add(adv - natural);
                total += adv;
                ci += rune.Utf16SequenceLength;
            }
            // Опорная точка и выравнивание (TA_*): по горизонтали — слева/по центру/справа; по вертикали — верх/базовая линия/низ
            var (px, py) = P(lx, ly);
            double asc = face.Font.Ascent / (double)face.Font.UnitsPerEm * em, desc = face.Font.Descent / (double)face.Font.UnitsPerEm * em;
            int ta = s.TextAlign;
            double offX = (ta & 6) switch { 6 => -total / 2, 2 => -total, _ => 0 };
            double offY = (ta & 24) switch { 24 => 0, 8 => -desc, _ => asc };
            double angle = f.Escapement / 10.0;
            // Направление оси y на странице может быть перевёрнуто (метрические режимы)
            bool rotated = Math.Abs(angle) > 0.01;
            // Сначала обрезка (в координатах страницы), затем поворот текста вокруг опорной точки
            BeginClipState();
            if (rotated)
            {
                double a = -angle * Math.PI / 180;
                _c.BeginTransform(px, py, Math.Cos(a), Math.Sin(a), -Math.Sin(a), Math.Cos(a), 0, 0);
            }
            if (opaqueRect is { } or2)
            {
                var a1 = P(or2.L, or2.T); var a2 = P(or2.R, or2.B);
                var bg = _c.NewPath(); bg.Rect(Math.Min(a1.X, a2.X), Math.Min(a1.Y, a2.Y), Math.Abs(a2.X - a1.X), Math.Abs(a2.Y - a1.Y));
                _c.PaintPath(bg, s.BkColor, null, 0, false, null, 0, 0);
            }
            else if (s.BkMode == 2)
            {
                var bg = _c.NewPath(); bg.Rect(px + offX, py + offY - asc, total, asc + desc);
                _c.PaintPath(bg, s.BkColor, null, 0, false, null, 0, 0);
            }
            double gx = px + offX, baseY = py + offY;
            foreach (var (rf, g, t, extra) in runs)
            {
                _c.Glyphs(rf, em, gx, baseY, g, t, s.TextColor, extra, hscale);
                for (int i = 0; i < g.Count; i++) gx += rf.Font.Advance(g[i]) * em * hscale + extra[i];
            }
            if (f.Underline || f.Strike)
            {
                var ln = _c.NewPath();
                if (f.Underline) { ln.M(px + offX, baseY + em * 0.12); ln.L(px + offX + total, baseY + em * 0.12); }
                if (f.Strike) { ln.M(px + offX, baseY - em * 0.28); ln.L(px + offX + total, baseY - em * 0.28); }
                _c.PaintPath(ln, null, s.TextColor, Math.Max(0.3, em * 0.06), false, null, 0, 0);
            }
            if (rotated) _c.EndTransform();
            EndClipState();
            if (updateCp && dx != null) _s.CurX = lx + dx.Sum();
        }

        // ---------------- Растры ----------------

        void Dib(byte[] src, int bmiOff, int bmiLen, int bitsOff, int bitsLen, double dl, double dt, double dw, double dh,
                 int sx, int sy, int sw, int sh)
        {
            if (bmiLen < 12 || bmiOff + bmiLen > src.Length || bitsOff < 0 || bitsOff + bitsLen > src.Length) return;
            // DIB → файл BMP (заголовок BITMAPFILEHEADER + BITMAPINFO + биты)
            var bmp = new byte[14 + bmiLen + bitsLen];
            bmp[0] = (byte)'B'; bmp[1] = (byte)'M';
            BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(2), bmp.Length);
            BinaryPrimitives.WriteInt32LittleEndian(bmp.AsSpan(10), 14 + bmiLen);
            Array.Copy(src, bmiOff, bmp, 14, bmiLen);
            Array.Copy(src, bitsOff, bmp, 14 + bmiLen, bitsLen);
            int bw = I32(src, bmiOff + 4), bh = Math.Abs(I32(src, bmiOff + 8));
            if (bw <= 0 || bh <= 0) return;
            var a = P(dl, dt); var b = P(dl + dw, dt + dh);
            double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y), w = Math.Abs(b.X - a.X), h = Math.Abs(b.Y - a.Y);
            if (w < 0.01 || h < 0.01) return;
            string key = "wmfdib:" + Convert.ToHexString(System.Security.Cryptography.SHA1.HashData(bmp));
            BeginClipState();
            if (sw > 0 && sh > 0 && (sx != 0 || sy != 0 || sw != bw || sh != bh))
            {
                // Часть растра: обрезка по месту назначения, весь растр — в соответствующем масштабе
                bool bottomUp = I32(src, bmiOff + 8) > 0;
                int top = bottomUp ? bh - sy - sh : sy;
                double kx = w / sw, ky = h / sh;
                var cp = _c.NewPath(); cp.Rect(x, y, w, h);
                _c.BeginClip(cp, false);
                _c.Image(key, bmp, x - sx * kx, y - top * ky, bw * kx, bh * ky);
                _c.EndClip();
            }
            else _c.Image(key, bmp, x, y, w, h);
            EndClipState();
        }

        // ---------------- Объекты ----------------

        void Select(object? o)
        {
            switch (o)
            {
                case Pen p: _s.Pen = p; break;
                case Brush b: _s.Brush = b; break;
                case Font f: _s.Font = f; break;
            }
        }

        static object? Stock(uint idx) => (idx & 0x7FFFFFFF) switch
        {
            0 => new Brush { Color = Rgb.White }, 1 => new Brush { Color = new Rgb(0.75, 0.75, 0.75) }, 2 => new Brush { Color = new Rgb(0.5, 0.5, 0.5) },
            3 => new Brush { Color = new Rgb(0.25, 0.25, 0.25) }, 4 => new Brush { Color = Rgb.Black }, 5 => new Brush { Style = 1 },
            6 => new Pen { Color = Rgb.White, Cosmetic = true }, 7 => new Pen { Color = Rgb.Black, Cosmetic = true }, 8 => new Pen { Style = 5 },
            10 or 11 or 16 => new Font { Height = 12, Face = "Courier New" }, 12 or 13 or 14 or 17 => new Font { Height = 12, Face = "Arial" },
            18 => new Brush { Color = Rgb.White }, 19 => new Pen { Color = Rgb.Black, Cosmetic = true },
            _ => null
        };

        static Rgb ColorRef(uint c) => new((c & 0xFF) / 255.0, (c >> 8 & 0xFF) / 255.0, (c >> 16 & 0xFF) / 255.0);

        static string Utf16(byte[] d, int off, int chars)
        {
            if (off < 0 || chars <= 0) return "";
            int n = Math.Min(chars * 2, d.Length - off);
            return n <= 0 ? "" : Encoding.Unicode.GetString(d, off, n);
        }

        static Encoding CharsetEncoding(int charset)
        {
            int cp = charset switch
            {
                204 => 1251, 238 => 1250, 161 => 1253, 162 => 1254, 177 => 1255, 178 => 1256, 186 => 1257, 163 => 1258,
                222 => 874, 128 => 932, 134 => 936, 129 => 949, 136 => 950, _ => 1252
            };
            try { return Encoding.GetEncoding(cp); } catch { return Encoding.Latin1; }
        }

        // ================================================================
        //  EMF
        // ================================================================

        public void PlayEmf()
        {
            var d = _d;
            _emf = true;
            int frL = I32(d, 24), frT = I32(d, 28), frR = I32(d, 32), frB = I32(d, 36);
            int devX = I32(d, 72), devY = I32(d, 76), mmX = I32(d, 80), mmY = I32(d, 84);
            int hdrSize = I32(d, 4);
            if (hdrSize >= 108 && I32(d, 100) > 0 && devX > 0)
            {
                // Размер устройства в микрометрах (точнее, чем в миллиметрах)
                _mmPerPxX = I32(d, 100) / 1000.0 / devX; _mmPerPxY = I32(d, 104) / 1000.0 / Math.Max(1, devY);
            }
            else if (devX > 0 && devY > 0 && mmX > 0 && mmY > 0) { _mmPerPxX = (double)mmX / devX; _mmPerPxY = (double)mmY / devY; }
            _frameL = Math.Min(frL, frR); _frameT = Math.Min(frT, frB);
            _frameW = Math.Max(1, Math.Abs(frR - frL)); _frameH = Math.Max(1, Math.Abs(frB - frT));

            int pos = 0;
            while (pos + 8 <= d.Length)
            {
                int type = I32(d, pos), size = I32(d, pos + 4);
                if (size < 8 || pos + size > d.Length) break;
                if (type == 14) break;   // EOF
                EmfRecord(type, pos + 8, size - 8);
                pos += size;
            }
        }

        List<(double X, double Y)> Points32(int off, int n)
        {
            var l = new List<(double, double)>(n);
            for (int i = 0; i < n; i++) l.Add((I32(_d, off + i * 8), I32(_d, off + i * 8 + 4)));
            return l;
        }

        List<(double X, double Y)> Points16(int off, int n)
        {
            var l = new List<(double, double)>(n);
            for (int i = 0; i < n; i++) l.Add((I16(_d, off + i * 4), I16(_d, off + i * 4 + 2)));
            return l;
        }

        void EmfRecord(int type, int p, int len)
        {
            var d = _d;
            switch (type)
            {
                case 9: _s.WinExtX = I32(d, p); _s.WinExtY = I32(d, p + 4); break;
                case 10: _s.WinOrgX = I32(d, p); _s.WinOrgY = I32(d, p + 4); break;
                case 11: _s.VpExtX = I32(d, p); _s.VpExtY = I32(d, p + 4); break;
                case 12: _s.VpOrgX = I32(d, p); _s.VpOrgY = I32(d, p + 4); break;
                case 17: _s.MapMode = I32(d, p); break;
                case 18: _s.BkMode = I32(d, p); break;
                case 19: _s.PolyFill = I32(d, p); break;
                case 22: _s.TextAlign = I32(d, p); break;
                case 24: _s.TextColor = ColorRef(U32(d, p)); break;
                case 25: _s.BkColor = ColorRef(U32(d, p)); break;
                case 27: _s.CurX = I32(d, p); _s.CurY = I32(d, p + 4); _figureOpen = false; if (_inPath) { var (mx, my) = P(_s.CurX, _s.CurY); (_path ??= _c.NewPath()).M(mx, my); _figureOpen = true; } break;
                case 30: IntersectClip(I32(d, p), I32(d, p + 4), I32(d, p + 8), I32(d, p + 12)); break;
                case 33: _saved.Push(_s.Clone()); break;
                case 34:
                {
                    int n = I32(d, p);
                    int count = n < 0 ? -n : Math.Max(1, _saved.Count - n + 1);
                    for (int i = 0; i < count && _saved.Count > 0; i++) _s = _saved.Pop();
                    break;
                }
                case 35:
                    _s.M11 = F32(d, p); _s.M12 = F32(d, p + 4); _s.M21 = F32(d, p + 8); _s.M22 = F32(d, p + 12); _s.Dx = F32(d, p + 16); _s.Dy = F32(d, p + 20);
                    break;
                case 36:
                {
                    double a = F32(d, p), b = F32(d, p + 4), c = F32(d, p + 8), e = F32(d, p + 12), fx = F32(d, p + 16), fy = F32(d, p + 20);
                    int mode = I32(d, p + 24);
                    var s = _s;
                    if (mode == 1) { s.M11 = 1; s.M12 = 0; s.M21 = 0; s.M22 = 1; s.Dx = 0; s.Dy = 0; }
                    else if (mode == 2)
                    {
                        // Новое = переданное × текущее
                        (s.M11, s.M12, s.M21, s.M22, s.Dx, s.Dy) = (a * s.M11 + b * s.M21, a * s.M12 + b * s.M22, c * s.M11 + e * s.M21, c * s.M12 + e * s.M22,
                            fx * s.M11 + fy * s.M21 + s.Dx, fx * s.M12 + fy * s.M22 + s.Dy);
                    }
                    else if (mode == 3)
                    {
                        (s.M11, s.M12, s.M21, s.M22, s.Dx, s.Dy) = (s.M11 * a + s.M12 * c, s.M11 * b + s.M12 * e, s.M21 * a + s.M22 * c, s.M21 * b + s.M22 * e,
                            s.Dx * a + s.Dy * c + fx, s.Dx * b + s.Dy * e + fy);
                    }
                    else if (mode == 4) { s.M11 = a; s.M12 = b; s.M21 = c; s.M22 = e; s.Dx = fx; s.Dy = fy; }
                    break;
                }
                case 37:
                {
                    uint idx = U32(d, p);
                    Select((idx & 0x80000000) != 0 ? Stock(idx) : _objects.GetValueOrDefault((int)idx));
                    break;
                }
                case 38:   // CREATEPEN
                    _objects[I32(d, p)] = new Pen { Style = I32(d, p + 4), Width = I32(d, p + 8), Color = ColorRef(U32(d, p + 16)), Cosmetic = I32(d, p + 8) <= 1 };
                    break;
                case 95:   // EXTCREATEPEN
                {
                    int style = I32(d, p + 4 + 16);
                    _objects[I32(d, p)] = new Pen
                    {
                        Style = style, Width = I32(d, p + 4 + 20), Color = ColorRef(U32(d, p + 4 + 28)),
                        Cosmetic = (style & 0x10000) == 0 && I32(d, p + 4 + 20) <= 1,
                    };
                    if (I32(d, p + 4 + 24) == 1) ((Pen)_objects[I32(d, p)]).Style = 5;   // BS_NULL
                    break;
                }
                case 39:
                    _objects[I32(d, p)] = new Brush { Style = I32(d, p + 4), Color = ColorRef(U32(d, p + 8)), Hatch = I32(d, p + 12) };
                    break;
                case 93 or 94:   // монохромная / DIB-кисть — приближённо серой сплошной
                    _objects[I32(d, p)] = new Brush { Style = 0, Color = new Rgb(0.6, 0.6, 0.6) };
                    break;
                case 40: _objects.Remove(I32(d, p)); break;
                case 82:   // EXTCREATEFONTINDIRECTW (LOGFONTW)
                {
                    int lf = p + 4;
                    _objects[I32(d, p)] = new Font
                    {
                        Height = I32(d, lf), Escapement = I32(d, lf + 8), Weight = I32(d, lf + 16),
                        Italic = d[lf + 20] != 0, Underline = d[lf + 21] != 0, Strike = d[lf + 22] != 0, Charset = d[lf + 23],
                        Face = Utf16(d, lf + 28, 32).Split('\0')[0],
                    };
                    break;
                }
                case 57: _s.ArcDir = I32(d, p); break;
                case 2: Bezier(Points32(p + 20, I32(d, p + 16)), false); break;
                case 85: Bezier(Points16(p + 20, I32(d, p + 16)), false); break;
                case 5: Bezier(Points32(p + 20, I32(d, p + 16)), true); break;
                case 88: Bezier(Points16(p + 20, I32(d, p + 16)), true); break;
                case 3: Poly(Points32(p + 20, I32(d, p + 16)), true, true); break;
                case 86: Poly(Points16(p + 20, I32(d, p + 16)), true, true); break;
                case 4: Poly(Points32(p + 20, I32(d, p + 16)), false, false); break;
                case 87: Poly(Points16(p + 20, I32(d, p + 16)), false, false); break;
                case 6: PolyTo(Points32(p + 20, I32(d, p + 16))); break;
                case 89: PolyTo(Points16(p + 20, I32(d, p + 16))); break;
                case 7 or 8 or 90 or 91:
                {
                    int nPolys = I32(d, p + 16);
                    bool small = type >= 90;
                    int cntOff = p + 24, ptOff = cntOff + nPolys * 4;
                    var polys = new List<List<(double, double)>>();
                    for (int i = 0; i < nPolys; i++)
                    {
                        int n = I32(d, cntOff + i * 4);
                        polys.Add(small ? Points16(ptOff, n) : Points32(ptOff, n));
                        ptOff += n * (small ? 4 : 8);
                    }
                    PolyPoly(polys, type is 8 or 91);
                    break;
                }
                case 54: LineTo(I32(d, p), I32(d, p + 4)); break;
                case 42: Ellipse(I32(d, p), I32(d, p + 4), I32(d, p + 8), I32(d, p + 12)); break;
                case 43: Rect(I32(d, p), I32(d, p + 4), I32(d, p + 8), I32(d, p + 12)); break;
                case 44: RoundRect(I32(d, p), I32(d, p + 4), I32(d, p + 8), I32(d, p + 12), I32(d, p + 16), I32(d, p + 20)); break;
                case 45 or 46 or 47 or 55:
                    Arc(I32(d, p), I32(d, p + 4), I32(d, p + 8), I32(d, p + 12), I32(d, p + 16), I32(d, p + 20), I32(d, p + 24), I32(d, p + 28),
                        type switch { 46 => 1, 47 => 2, _ => 0 }, type == 55);
                    break;
                case 59: _inPath = true; _path = _c.NewPath(); _figureOpen = false; break;
                case 60: _inPath = false; _lastPath = _path; _path = null; break;
                case 61: if (_inPath) { _path?.Z(); _figureOpen = false; } break;
                case 68: _inPath = false; _path = null; _lastPath = null; break;
                case 62 or 63 or 64:
                    if (_lastPath != null) Paint(_lastPath, type != 64, type != 62);
                    _lastPath = null;
                    break;
                case 67:   // SELECTCLIPPATH
                    if (_lastPath != null) _s.ClipPath = _lastPath;
                    _lastPath = null;
                    break;
                case 75:   // EXTSELECTCLIPRGN: RGN_COPY без данных — сброс обрезки; иначе — прямоугольники региона
                {
                    int cb = I32(d, p), mode = I32(d, p + 4);
                    if (cb == 0 || mode == 5) { _s.Clip = null; _s.ClipPath = null; break; }
                    int rgn = p + 8, nRects = I32(d, rgn + 8);
                    var rp = _c.NewPath();
                    for (int i = 0; i < nRects; i++)
                    {
                        int o = rgn + 32 + i * 16;
                        var (dx1, dy1) = DeviceToPage(I32(d, o), I32(d, o + 4)); var (dx2, dy2) = DeviceToPage(I32(d, o + 8), I32(d, o + 12));
                        rp.Rect(Math.Min(dx1, dx2), Math.Min(dy1, dy2), Math.Abs(dx2 - dx1), Math.Abs(dy2 - dy1));
                    }
                    if (mode == 1 && _s.ClipPath != null) { /* пересечение: оставляем текущий контур и добавляем прямоугольники */ _s.Clip = null; }
                    _s.ClipPath = nRects > 0 ? rp : null;
                    break;
                }
                case 83 or 84:   // EXTTEXTOUTA / W
                {
                    int t = p + 28;   // EMRTEXT после rclBounds, iGraphicsMode, exScale, eyScale
                    int rx = I32(d, t), ry = I32(d, t + 4), nChars = I32(d, t + 8), offString = I32(d, t + 12), options = I32(d, t + 16);
                    int offDx = I32(d, t + 36);
                    int recStart = p - 8;
                    string text = type == 84 ? Utf16(d, recStart + offString, nChars)
                        : CharsetEncoding(_s.Font.Charset).GetString(d, recStart + offString, Math.Min(nChars, d.Length - recStart - offString));
                    double[]? dx = null;
                    if (offDx > 0 && recStart + offDx + nChars * 4 <= d.Length)
                    {
                        dx = new double[nChars];
                        for (int i = 0; i < nChars; i++) dx[i] = I32(d, recStart + offDx + i * 4);
                    }
                    (double, double, double, double)? opaque = (options & 2) != 0 ? (I32(d, t + 20), I32(d, t + 24), I32(d, t + 28), I32(d, t + 32)) : null;
                    if ((_s.TextAlign & 1) != 0) { rx = (int)_s.CurX; ry = (int)_s.CurY; }
                    Text(rx, ry, text, dx, (_s.TextAlign & 1) != 0, opaque);
                    break;
                }
                case 76 or 77:   // BITBLT / STRETCHBLT
                {
                    int xD = I32(d, p + 16), yD = I32(d, p + 20), cxD = I32(d, p + 24), cyD = I32(d, p + 28);
                    uint rop = U32(d, p + 32);
                    int xS = I32(d, p + 36), yS = I32(d, p + 40);
                    int bo = p + 36 + 8 + 24 + 4 + 4;   // после xformSrc и crBkColorSrc, iUsageSrc
                    int offBmi = I32(d, bo), cbBmi = I32(d, bo + 4), offBits = I32(d, bo + 8), cbBits = I32(d, bo + 12);
                    int cxS = type == 77 ? I32(d, bo + 16) : cxD, cyS = type == 77 ? I32(d, bo + 20) : cyD;
                    if (cbBmi == 0)
                    {
                        // Без растра: заливка кистью (PATCOPY) или чёрным/белым
                        var save = _s.Pen; _s.Pen = new Pen { Style = 5 };
                        var saveB = _s.Brush;
                        if (rop == 0x00000042) _s.Brush = new Brush { Color = Rgb.Black };
                        else if (rop == 0x00FF0062) _s.Brush = new Brush { Color = Rgb.White };
                        Rect(xD, yD, xD + cxD, yD + cyD);
                        _s.Pen = save; _s.Brush = saveB;
                    }
                    else Dib(d, p - 8 + offBmi, cbBmi, p - 8 + offBits, cbBits, xD, yD, cxD, cyD, xS, yS, cxS, cyS);
                    break;
                }
                case 81:   // STRETCHDIBITS
                {
                    int xD = I32(d, p + 16), yD = I32(d, p + 20), xS = I32(d, p + 24), yS = I32(d, p + 28), cxS = I32(d, p + 32), cyS = I32(d, p + 36);
                    int offBmi = I32(d, p + 40), cbBmi = I32(d, p + 44), offBits = I32(d, p + 48), cbBits = I32(d, p + 52);
                    int cxD = I32(d, p + 64), cyD = I32(d, p + 68);
                    Dib(d, p - 8 + offBmi, cbBmi, p - 8 + offBits, cbBits, xD, yD, cxD, cyD, xS, yS, cxS, cyS);
                    break;
                }
                case 80:   // SETDIBITSTODEVICE
                {
                    int xD = I32(d, p + 16), yD = I32(d, p + 20), xS = I32(d, p + 24), yS = I32(d, p + 28), cx = I32(d, p + 32), cy = I32(d, p + 36);
                    int offBmi = I32(d, p + 40), cbBmi = I32(d, p + 44), offBits = I32(d, p + 48), cbBits = I32(d, p + 52);
                    Dib(d, p - 8 + offBmi, cbBmi, p - 8 + offBits, cbBits, xD, yD, cx, cy, xS, yS, cx, cy);
                    break;
                }
                case 71:   // FILLRGN: прямоугольники региона кистью
                {
                    int ihBrush = I32(d, p + 20);
                    var saveB = _s.Brush; var saveP = _s.Pen;
                    if (_objects.GetValueOrDefault(ihBrush) is Brush rb) _s.Brush = rb;
                    _s.Pen = new Pen { Style = 5 };
                    int rgn = p + 24, nRects = I32(d, rgn + 8);
                    var rp = _c.NewPath();
                    for (int i = 0; i < nRects; i++)
                    {
                        int o = rgn + 32 + i * 16;
                        var a = P(I32(d, o), I32(d, o + 4)); var b = P(I32(d, o + 8), I32(d, o + 12));
                        rp.Rect(Math.Min(a.X, b.X), Math.Min(a.Y, b.Y), Math.Abs(b.X - a.X), Math.Abs(b.Y - a.Y));
                    }
                    Paint(rp, true, false);
                    _s.Brush = saveB; _s.Pen = saveP;
                    break;
                }
            }
        }

        /// <summary>Координаты устройства (регионы EMF) → страница.</summary>
        (double X, double Y) DeviceToPage(double dx, double dy)
        {
            double mx = dx * _mmPerPxX * 100, my = dy * _mmPerPxY * 100;
            return (_x + (mx - _frameL) / _frameW * _w, _y + (my - _frameT) / _frameH * _h);
        }

        // ================================================================
        //  WMF
        // ================================================================

        public void PlayWmf(int start)
        {
            var d = _d;
            if (start == 22)
            {
                int l = I16(d, 6), t = I16(d, 8), r = I16(d, 10), b = I16(d, 12);
                _bbL = l; _bbT = t; _bbW = r - l; _bbH = b - t;
            }
            else
            {
                var (ox, oy, ew, eh) = WmfWindow(d, 0);
                _bbL = ox; _bbT = oy; _bbW = ew == 0 ? 1 : ew; _bbH = eh == 0 ? 1 : eh;
            }
            int pos = start + U16(d, start + 2) * 2;
            while (pos + 6 <= d.Length)
            {
                int size = (int)U32(d, pos) * 2, fn = U16(d, pos + 4);
                if (size < 6 || fn == 0 || pos + size > d.Length) break;
                WmfRecord(fn, pos + 6, size - 6);
                pos += size;
            }
        }

        void AddObject(object o)
        {
            int i = 0;
            while (_objects.ContainsKey(i)) i++;
            _objects[i] = o;
        }

        void WmfRecord(int fn, int p, int len)
        {
            var d = _d;
            int W(int i) => I16(d, p + i * 2);
            switch (fn)
            {
                case 0x0102: _s.BkMode = W(0); break;
                case 0x0103: _s.MapMode = W(0); break;
                case 0x0106: _s.PolyFill = W(0); break;
                case 0x012E: _s.TextAlign = U16(d, p); break;
                case 0x0201: _s.BkColor = ColorRef(U32(d, p)); break;
                case 0x0209: _s.TextColor = ColorRef(U32(d, p)); break;
                case 0x020B: _s.WinOrgY = W(0); _s.WinOrgX = W(1); _s.WinSet = true; if (_s.WinExtX == 1 && _s.WinExtY == 1) { _s.WinExtX = _bbW; _s.WinExtY = _bbH; } break;
                case 0x020C: _s.WinExtY = W(0); _s.WinExtX = W(1); _s.WinSet = true; break;
                case 0x020F: _s.WinOrgY += W(0); _s.WinOrgX += W(1); break;
                case 0x0410: if (W(2) != 0 && W(0) != 0) { _s.WinExtY = _s.WinExtY * W(1) / W(0); _s.WinExtX = _s.WinExtX * W(3) / W(2); } break;
                case 0x0213: LineTo(W(1), W(0)); break;
                case 0x0214: _s.CurY = W(0); _s.CurX = W(1); _figureOpen = false; break;
                case 0x0416: IntersectClip(W(3), W(2), W(1), W(0)); break;
                case 0x0418: Ellipse(W(3), W(2), W(1), W(0)); break;
                case 0x041B: Rect(W(3), W(2), W(1), W(0)); break;
                case 0x061C: RoundRect(W(5), W(4), W(3), W(2), W(1), W(0)); break;
                case 0x0817 or 0x081A or 0x0830:
                    Arc(W(7), W(6), W(5), W(4), W(3), W(2), W(1), W(0), fn switch { 0x0830 => 1, 0x081A => 2, _ => 0 });
                    break;
                case 0x0324 or 0x0325:
                {
                    int n = W(0);
                    var pts = new List<(double, double)>(n);
                    for (int i = 0; i < n; i++) pts.Add((W(1 + i * 2), W(2 + i * 2)));
                    Poly(pts, fn == 0x0324, fn == 0x0324);
                    break;
                }
                case 0x0538:
                {
                    int np = W(0), k = 1 + np;
                    var polys = new List<List<(double, double)>>();
                    for (int i = 0; i < np; i++)
                    {
                        int n = W(1 + i);
                        var pts = new List<(double, double)>(n);
                        for (int j = 0; j < n; j++) { pts.Add((W(k), W(k + 1))); k += 2; }
                        polys.Add(pts);
                    }
                    PolyPoly(polys, true);
                    break;
                }
                case 0x001E: _saved.Push(_s.Clone()); break;
                case 0x0127:
                {
                    int n = W(0);
                    int count = n < 0 ? -n : 1;
                    for (int i = 0; i < count && _saved.Count > 0; i++) _s = _saved.Pop();
                    break;
                }
                case 0x012D: Select(_objects.GetValueOrDefault(U16(d, p))); break;
                case 0x01F0: _objects.Remove(U16(d, p)); break;
                case 0x02FA:
                {
                    int style = U16(d, p), wdt = W(1);
                    AddObject(new Pen { Style = style, Width = wdt, Color = ColorRef(U32(d, p + 6)), Cosmetic = wdt <= 1 });
                    break;
                }
                case 0x02FC: AddObject(new Brush { Style = U16(d, p), Color = ColorRef(U32(d, p + 2)), Hatch = U16(d, p + 6) }); break;
                case 0x0142 or 0x01F9: AddObject(new Brush { Style = 0, Color = new Rgb(0.6, 0.6, 0.6) }); break;
                case 0x02FB:
                {
                    int charset = d[p + 13];
                    var nameBytes = d.AsSpan(p + 18, Math.Min(32, len - 18)).ToArray();
                    int z = Array.IndexOf(nameBytes, (byte)0);
                    string face = CharsetEncoding(charset is 2 ? 0 : charset).GetString(nameBytes, 0, z < 0 ? nameBytes.Length : z);
                    AddObject(new Font
                    {
                        Height = W(0), Escapement = W(2), Weight = W(4), Italic = d[p + 10] != 0, Underline = d[p + 11] != 0, Strike = d[p + 12] != 0,
                        Charset = charset, Face = face,
                    });
                    break;
                }
                case 0x00F7 or 0x06FF: AddObject(new object()); break;   // палитра/регион — занимают слот
                case 0x0521:
                {
                    int n = W(0), bytes = (n + 1) & ~1;
                    string text = WmfString(p + 2, n);
                    int y = I16(d, p + 2 + bytes), x = I16(d, p + 4 + bytes);
                    Text(x, y, text, null, false, null);
                    break;
                }
                case 0x0A32:
                {
                    int y = W(0), x = W(1), n = W(2), opts = U16(d, p + 6);
                    int o = p + 8;
                    (double, double, double, double)? opaque = null;
                    if ((opts & 6) != 0) { if ((opts & 2) != 0) opaque = (I16(d, o), I16(d, o + 2), I16(d, o + 4), I16(d, o + 6)); o += 8; }
                    string text = WmfString(o, n);
                    int dxOff = o + ((n + 1) & ~1);
                    double[]? dx = null;
                    if (dxOff + n * 2 <= p + len)
                    {
                        dx = new double[text.Length];
                        // Расстояния заданы на байт строки; для однобайтовых кодировок — по символу
                        for (int i = 0; i < text.Length && i < n; i++) dx[i] = I16(d, dxOff + i * 2);
                    }
                    if ((_s.TextAlign & 1) != 0) { x = (int)_s.CurX; y = (int)_s.CurY; }
                    Text(x, y, text, dx, (_s.TextAlign & 1) != 0, opaque);
                    break;
                }
                case 0x0F43:   // STRETCHDIB
                {
                    int sh = W(3), sw = W(4), sy = W(5), sx = W(6), dh = W(7), dw = W(8), dy = W(9), dx = W(10);
                    WmfDib(p + 22, p + len, dx, dy, dw, dh, sx, sy, sw, sh);
                    break;
                }
                case 0x0B41:   // DIBSTRETCHBLT
                {
                    int sh = W(2), sw = W(3), sy = W(4), sx = W(5), dh = W(6), dw = W(7), dy = W(8), dx = W(9);
                    if (len > 20) WmfDib(p + 20, p + len, dx, dy, dw, dh, sx, sy, sw, sh);
                    break;
                }
                case 0x0940:   // DIBBITBLT
                {
                    int sy = W(2), sx = W(3), dh = W(4), dw = W(5), dy = W(6), dx = W(7);
                    if (len > 16) WmfDib(p + 16, p + len, dx, dy, dw, dh, sx, sy, dw, dh);
                    break;
                }
                case 0x061D:   // PATBLT
                {
                    uint rop = U32(d, p);
                    int h = W(2), w = W(3), y = W(4), x = W(5);
                    var save = _s.Pen; _s.Pen = new Pen { Style = 5 };
                    var saveB = _s.Brush;
                    if (rop == 0x00000042) _s.Brush = new Brush { Color = Rgb.Black };
                    else if (rop == 0x00FF0062) _s.Brush = new Brush { Color = Rgb.White };
                    Rect(x, y, x + w, y + h);
                    _s.Pen = save; _s.Brush = saveB;
                    break;
                }
            }
        }

        string WmfString(int off, int n)
        {
            if (n <= 0 || off + n > _d.Length) return "";
            var f = _s.Font;
            if (f.Charset == 2) return string.Concat(_d.AsSpan(off, n).ToArray().Select(b => (char)(0xF000 + b)));
            return CharsetEncoding(f.Charset).GetString(_d, off, n);
        }

        void WmfDib(int off, int end, int dx, int dy, int dw, int dh, int sx, int sy, int sw, int sh)
        {
            var d = _d;
            if (off + 40 > end) return;
            int hdr = I32(d, off);
            if (hdr is not (12 or 40 or 52 or 56 or 108 or 124)) return;
            int bpp = U16(d, off + 14), clrUsed = hdr >= 40 ? I32(d, off + 32) : 0, compression = hdr >= 40 ? I32(d, off + 16) : 0;
            int palette = bpp <= 8 ? (clrUsed > 0 ? clrUsed : 1 << bpp) * (hdr == 12 ? 3 : 4) : 0;
            if (compression == 3 && hdr == 40) palette += 12;
            int bmiLen = hdr + palette;
            Dib(d, off, bmiLen, off + bmiLen, end - off - bmiLen, dx, dy, dw, dh, sx, sy, sw, sh);
        }
    }
}

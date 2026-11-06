using System.Globalization;
using System.Text;

namespace DocxToPdf.Pdf;

/// <summary>
/// Mac PICT (QuickDraw, версии 1 и 2) → SVG: прямоугольники, скруглённые прямоугольники, овалы, дуги, многоугольники,
/// области (по ограничивающему прямоугольнику), линии, текст, растры BitsRect/PackBitsRect/DirectBitsRect
/// (1–32 бит, упаковка PackBits) и JPEG в QuickTime. SVG выводится в PDF векторно (Svg.cs).
/// </summary>
public static class Pict
{
    static Pict() => Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    /// <summary>SVG-документ из данных PICT (с 512-байтным заголовком файла или без него); null — не PICT.</summary>
    public static byte[]? ToSvg(byte[] data)
    {
        try { return new Conv(data).Run(); }
        catch { return null; }
    }

    static bool Looks(byte[] d, int off) =>
        off + 14 <= d.Length && ((d[off + 10] == 0x00 && d[off + 11] == 0x11 && d[off + 12] == 0x02 && d[off + 13] == 0xFF)
                                  || (d[off + 10] == 0x11 && d[off + 11] == 0x01));

    sealed class Conv(byte[] d)
    {
        int _p, _start;
        bool _v2;
        double _vx, _vy, _vw, _vh;           // область координат (viewBox)
        double _outW, _outH;                 // размер, pt (72 dpi)
        readonly StringBuilder _sb = new();
        int _penX, _penY, _penW = 1, _penH = 1, _ovW, _ovH, _orgX, _orgY;
        (int R, int G, int B) _fg = (0, 0, 0), _bg = (255, 255, 255);
        double _penDensity = 1, _fillDensity = 1;
        int _txFont, _txSize = 12, _txFace, _textX, _textY;
        readonly Dictionary<int, string> _fontNames = [];
        (int T, int L, int B, int R) _lastRect, _lastRRect, _lastOval, _lastArc;
        List<(int X, int Y)> _lastPoly = [];

        int U8() => d[_p++];
        int I8() => (sbyte)d[_p++];
        int U16() { int v = d[_p] << 8 | d[_p + 1]; _p += 2; return v; }
        int I16() => (short)U16();
        long U32() { long v = (long)d[_p] << 24 | (long)d[_p + 1] << 16 | (long)d[_p + 2] << 8 | d[_p + 3]; _p += 4; return v; }
        (int T, int L, int B, int R) Rect() { int t = I16(), l = I16(), b = I16(), r = I16(); return (t, l, b, r); }
        (int X, int Y) Pt() { int v = I16(), h = I16(); return (h, v); }

        static string F(double v) => Math.Round(v, 3).ToString(CultureInfo.InvariantCulture);
        string X(double x) => F(x - _orgX);
        string Y(double y) => F(y - _orgY);
        static string Hex((int R, int G, int B) c) => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

        /// <summary>Цвет узора: доля закрашенных точек — смесь цвета переднего плана и фона.</summary>
        string Mix(double density)
        {
            var c = (R: (int)Math.Round(_fg.R * density + _bg.R * (1 - density)),
                     G: (int)Math.Round(_fg.G * density + _bg.G * (1 - density)),
                     B: (int)Math.Round(_fg.B * density + _bg.B * (1 - density)));
            return Hex(c);
        }

        static double Density(byte[] d, int p)
        {
            int bits = 0;
            for (int i = 0; i < 8; i++) bits += System.Numerics.BitOperations.PopCount(d[p + i]);
            return bits / 64.0;
        }

        static (int, int, int) QdColor(long c) => c switch
        {
            30 => (255, 255, 255), 205 => (221, 8, 6), 341 => (0, 128, 17), 409 => (0, 0, 212),
            273 => (2, 171, 234), 137 => (242, 8, 132), 69 => (252, 243, 5), _ => (0, 0, 0),
        };

        public byte[]? Run()
        {
            _start = !Looks(d, 0) && Looks(d, 512) ? 512 : 0;
            if (!Looks(d, _start)) return null;
            _p = _start + 2;
            var frame = Rect();
            _outW = frame.R - frame.L; _outH = frame.B - frame.T;
            if (_outW <= 0 || _outH <= 0) return null;
            (_vx, _vy, _vw, _vh) = (frame.L, frame.T, _outW, _outH);
            if (d[_p] == 0x11) { _v2 = false; _p += 2; }
            else { _v2 = true; _p += 4; }
            for (int guard = 0; _p < d.Length && guard < 1_000_000; guard++)
            {
                int op;
                if (_v2)
                {
                    if (((_p - _start) & 1) != 0) _p++;
                    if (_p + 2 > d.Length) break;
                    op = U16();
                }
                else op = U8();
                if (op == 0xFF || op == 0xFFFF) break;
                Op(op);
            }
            var svg = new StringBuilder();
            svg.Append($"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{F(_outW)}pt\" height=\"{F(_outH)}pt\" ")
               .Append($"viewBox=\"{F(_vx)} {F(_vy)} {F(_vw)} {F(_vh)}\" preserveAspectRatio=\"none\">")
               .Append(_sb).Append("</svg>");
            return Encoding.UTF8.GetBytes(svg.ToString());
        }

        void Op(int op)
        {
            switch (op)
            {
                case 0x00: break;
                case 0x01: { int n = U16(); _p += n - 2; break; }                  // Clip (область)
                case 0x02: _p += 8; break;                                          // BkPat
                case 0x03: _txFont = I16(); break;
                case 0x04: _txFace = U8(); break;
                case 0x05: _p += 2; break;                                          // TxMode
                case 0x06: _p += 4; break;                                          // SpExtra
                case 0x07: { _penH = Math.Max(0, I16()); _penW = Math.Max(0, I16()); break; }
                case 0x08: _p += 2; break;                                          // PnMode
                case 0x09: _penDensity = Density(d, _p); _p += 8; break;
                case 0x0A: _fillDensity = Density(d, _p); _p += 8; break;
                case 0x0B: { _ovH = I16(); _ovW = I16(); break; }
                case 0x0C: { int dh = I16(), dv = I16(); _orgX += dh; _orgY += dv; break; }
                case 0x0D: _txSize = I16(); break;
                case 0x0E: _fg = QdColor(U32()); break;
                case 0x0F: _bg = QdColor(U32()); break;
                case 0x10: _p += 8; break;                                          // TxRatio
                case 0x11: _p += _v2 ? 2 : 1; break;                                // Version
                case 0x12 or 0x13 or 0x14: PixPat(op); break;
                case 0x15 or 0x16: _p += 2; break;
                case >= 0x17 and <= 0x19: break;
                case 0x1A: _fg = Rgb48(); _penDensity = _fillDensity = 1; break;
                case 0x1B: _bg = Rgb48(); break;
                case 0x1C or 0x1E: break;
                case 0x1D or 0x1F: _p += 6; break;
                case 0x20: { var a = Pt(); var b = Pt(); Line(a, b); break; }
                case 0x21: { var b = Pt(); Line((_penX, _penY), b); break; }
                case 0x22: { var a = Pt(); int dh = I8(), dv = I8(); Line(a, (a.X + dh, a.Y + dv)); break; }
                case 0x23: { int dh = I8(), dv = I8(); Line((_penX, _penY), (_penX + dh, _penY + dv)); break; }
                case >= 0x24 and <= 0x27: { int n = U16(); _p += n; break; }
                case 0x28: { var pt = Pt(); _textX = pt.X; _textY = pt.Y; Text(); break; }
                case 0x29: { _textX += U8(); Text(); break; }
                case 0x2A: { _textY += U8(); Text(); break; }
                case 0x2B: { _textX += U8(); _textY += U8(); Text(); break; }
                case 0x2C:
                {
                    int n = U16(), end = _p + n;
                    int id = I16(), len = U8();
                    _fontNames[id] = Encoding.GetEncoding(10000).GetString(d, _p, Math.Min(len, end - _p));
                    _p = end;
                    break;
                }
                case >= 0x2D and <= 0x2F: { int n = U16(); _p += n; break; }
                case >= 0x30 and <= 0x37: _lastRect = Rect(); if (op <= 0x34) RectOp(op & 7, _lastRect, 0, 0); break;
                case >= 0x38 and <= 0x3F: if (op <= 0x3C) RectOp(op & 7, _lastRect, 0, 0); break;
                case >= 0x40 and <= 0x47: _lastRRect = Rect(); if (op <= 0x44) RectOp(op & 7, _lastRRect, _ovW / 2.0, _ovH / 2.0); break;
                case >= 0x48 and <= 0x4F: if (op <= 0x4C) RectOp(op & 7, _lastRRect, _ovW / 2.0, _ovH / 2.0); break;
                case >= 0x50 and <= 0x57: _lastOval = Rect(); if (op <= 0x54) OvalOp(op & 7, _lastOval); break;
                case >= 0x58 and <= 0x5F: if (op <= 0x5C) OvalOp(op & 7, _lastOval); break;
                case >= 0x60 and <= 0x67: { _lastArc = Rect(); int s = I16(), a = I16(); if (op <= 0x64) ArcOp(op & 7, _lastArc, s, a); break; }
                case >= 0x68 and <= 0x6F: { int s = I16(), a = I16(); if (op <= 0x6C) ArcOp(op & 7, _lastArc, s, a); break; }
                case >= 0x70 and <= 0x77:
                {
                    int start = _p, n = U16();
                    _p += 8;
                    var pts = new List<(int, int)>();
                    while (_p + 4 <= start + n) pts.Add(Pt());
                    _p = start + n;
                    _lastPoly = pts;
                    if (op <= 0x74) PolyOp(op & 7, pts);
                    break;
                }
                case >= 0x78 and <= 0x7F: if (op <= 0x7C) PolyOp(op & 7, _lastPoly); break;
                case >= 0x80 and <= 0x87:
                {
                    // Область — по ограничивающему прямоугольнику
                    int start = _p, n = U16();
                    var r = Rect();
                    _p = start + n;
                    _lastRect = r;
                    if (op <= 0x84) RectOp(op & 7, r, 0, 0);
                    break;
                }
                case >= 0x88 and <= 0x8F: break;
                case 0x90 or 0x91 or 0x98 or 0x99: Bits(op, direct: false); break;
                case 0x9A or 0x9B: Bits(op, direct: true); break;
                case >= 0x92 and <= 0x97 or >= 0x9C and <= 0x9F: { int n = U16(); _p += n; break; }
                case 0xA0: _p += 2; break;
                case 0xA1: { _p += 2; int n = U16(); _p += n; break; }
                case >= 0xA2 and <= 0xAF: { int n = U16(); _p += n; break; }
                case >= 0xB0 and <= 0xCF: break;
                case >= 0xD0 and <= 0xFE: { long n = U32(); _p += (int)n; break; }
                case 0x0C00: Header(); break;
                case >= 0x0100 and <= 0x7FFF: _p += (op >> 8) * 2; break;
                case >= 0x8000 and <= 0x80FF: break;
                case 0x8200: QuickTime(); break;
                default: { long n = U32(); _p += (int)n; break; }
            }
        }

        (int, int, int) Rgb48() { int r = U16(), g = U16(), b = U16(); return (r >> 8, g >> 8, b >> 8); }

        /// <summary>Заголовок версии 2: расширенный (−2) задаёт исходный прямоугольник координат.</summary>
        void Header()
        {
            int start = _p;
            int version = I16();
            if (version == -2)
            {
                _p = start + 12;
                var src = Rect();
                if (src.R > src.L && src.B > src.T) (_vx, _vy, _vw, _vh) = (src.L, src.T, src.R - src.L, src.B - src.T);
            }
            _p = start + 24;
        }

        void PixPat(int op)
        {
            int type = U16();
            double dens = Density(d, _p);
            _p += 8;
            if (type == 2) { var c = Rgb48(); if (op == 0x13) _fg = c; }       // смешанный цвет RGB
            else if (type == 1)
            {
                // Пиксельный узор: PixMap, таблица цветов, данные — пропускаются (берётся одноцветный образец)
                int rowBytes = U16() & 0x3FFF;
                var b = Rect();
                _p += 36;
                SkipColorTable();
                ReadRows(rowBytes, b.B - b.T, 0, 1);
            }
            if (op == 0x13) _penDensity = dens; else if (op == 0x14) _fillDensity = dens;
        }

        void SkipColorTable() { _p += 6; int n = U16() + 1; _p += n * 8; }

        string StrokeAttr()
        {
            double w = Math.Max(_penW, _penH);
            return w <= 0 ? "stroke=\"none\"" : $"stroke=\"{Mix(_penDensity)}\" stroke-width=\"{F(w)}\"";
        }

        string FillFor(int verb) => verb switch
        {
            1 => Mix(_penDensity),               // paint — узором пера
            2 => Hex(_bg),                       // erase
            4 => Mix(_fillDensity),              // fill
            _ => "none",
        };

        void RectOp(int verb, (int T, int L, int B, int R) r, double rx, double ry)
        {
            if (verb == 3) return;               // invert
            string round = rx > 0 && ry > 0 ? $" rx=\"{F(rx)}\" ry=\"{F(ry)}\"" : "";
            if (verb == 0)
            {
                // Перо QuickDraw — внутри рамки
                double w = Math.Max(_penW, _penH);
                if (w <= 0) return;
                double h2 = w / 2;
                _sb.Append($"<rect x=\"{X(r.L + h2)}\" y=\"{Y(r.T + h2)}\" width=\"{F(Math.Max(0, r.R - r.L - w))}\" height=\"{F(Math.Max(0, r.B - r.T - w))}\"{round} fill=\"none\" {StrokeAttr()}/>");
            }
            else
                _sb.Append($"<rect x=\"{X(r.L)}\" y=\"{Y(r.T)}\" width=\"{F(r.R - r.L)}\" height=\"{F(r.B - r.T)}\"{round} fill=\"{FillFor(verb)}\"/>");
        }

        void OvalOp(int verb, (int T, int L, int B, int R) r)
        {
            if (verb == 3) return;
            double cx = (r.L + r.R) / 2.0, cy = (r.T + r.B) / 2.0, rx = (r.R - r.L) / 2.0, ry = (r.B - r.T) / 2.0;
            if (verb == 0)
            {
                double w = Math.Max(_penW, _penH);
                if (w <= 0) return;
                _sb.Append($"<ellipse cx=\"{X(cx)}\" cy=\"{Y(cy)}\" rx=\"{F(Math.Max(0, rx - w / 2))}\" ry=\"{F(Math.Max(0, ry - w / 2))}\" fill=\"none\" {StrokeAttr()}/>");
            }
            else _sb.Append($"<ellipse cx=\"{X(cx)}\" cy=\"{Y(cy)}\" rx=\"{F(rx)}\" ry=\"{F(ry)}\" fill=\"{FillFor(verb)}\"/>");
        }

        /// <summary>Дуга: угол 0 — «12 часов», положительные углы — по часовой стрелке.</summary>
        void ArcOp(int verb, (int T, int L, int B, int R) r, int start, int arc)
        {
            if (verb == 3 || arc == 0) return;
            double w = verb == 0 ? Math.Max(_penW, _penH) : 0;
            double cx = (r.L + r.R) / 2.0, cy = (r.T + r.B) / 2.0, rx = (r.R - r.L) / 2.0 - w / 2, ry = (r.B - r.T) / 2.0 - w / 2;
            if (Math.Abs(arc) >= 360) { if (verb == 0) OvalOp(0, r); else OvalOp(verb, r); return; }
            (double, double) P(double deg) { double a = deg * Math.PI / 180; return (cx + rx * Math.Sin(a), cy - ry * Math.Cos(a)); }
            var (x1, y1) = P(start);
            var (x2, y2) = P(start + arc);
            int large = Math.Abs(arc) > 180 ? 1 : 0, sweep = arc > 0 ? 1 : 0;
            string path = $"M{X(x1)},{Y(y1)} A{F(rx)},{F(ry)} 0 {large} {sweep} {X(x2)},{Y(y2)}";
            if (verb == 0) { if (w > 0) _sb.Append($"<path d=\"{path}\" fill=\"none\" {StrokeAttr()}/>"); }
            else _sb.Append($"<path d=\"M{X(cx)},{Y(cy)} L{X(x1)},{Y(y1)} {path[path.IndexOf('A')..]} Z\" fill=\"{FillFor(verb)}\"/>");
        }

        void PolyOp(int verb, List<(int X, int Y)> pts)
        {
            if (verb == 3 || pts.Count < 2) return;
            string list = string.Join(" ", pts.Select(p => $"{X(p.X)},{Y(p.Y)}"));
            if (verb == 0)
            {
                if (Math.Max(_penW, _penH) <= 0) return;
                double o = Math.Max(_penW, _penH) / 2.0;
                string shifted = string.Join(" ", pts.Select(p => $"{X(p.X + o)},{Y(p.Y + o)}"));
                bool closed = pts[0] == pts[^1];
                _sb.Append($"<{(closed ? "polygon" : "polyline")} points=\"{shifted}\" fill=\"none\" {StrokeAttr()} stroke-linejoin=\"miter\"/>");
            }
            else _sb.Append($"<polygon points=\"{list}\" fill=\"{FillFor(verb)}\"/>");
        }

        /// <summary>Линия: перо QuickDraw свисает вправо-вниз от точки — центр штриха смещён на половину пера.</summary>
        void Line((int X, int Y) a, (int X, int Y) b)
        {
            _penX = b.X; _penY = b.Y;
            if (_penW <= 0 || _penH <= 0) return;
            double ox = _penW / 2.0, oy = _penH / 2.0;
            if (a == b)
            {
                _sb.Append($"<rect x=\"{X(a.X)}\" y=\"{Y(a.Y)}\" width=\"{_penW}\" height=\"{_penH}\" fill=\"{Mix(_penDensity)}\"/>");
                return;
            }
            _sb.Append($"<line x1=\"{X(a.X + ox)}\" y1=\"{Y(a.Y + oy)}\" x2=\"{X(b.X + ox)}\" y2=\"{Y(b.Y + oy)}\" {StrokeAttr()} stroke-linecap=\"square\"/>");
        }

        void Text()
        {
            int n = U8();
            string t = Encoding.GetEncoding(10000).GetString(d, _p, Math.Min(n, d.Length - _p));
            _p += n;
            if (t.Trim().Length == 0) return;
            string font = _fontNames.TryGetValue(_txFont, out var fn) ? fn : _txFont switch
            {
                2 or 20 => "Times New Roman", 4 or 22 => "Courier New", 23 => "Symbol", _ => "Arial",
            };
            font = font switch { "Times" or "New York" => "Times New Roman", "Helvetica" or "Geneva" or "Chicago" => "Arial", "Courier" or "Monaco" => "Courier New", _ => font };
            int size = _txSize > 0 ? _txSize : 12;
            var esc = System.Security.SecurityElement.Escape(t);
            _sb.Append($"<text x=\"{X(_textX)}\" y=\"{Y(_textY)}\" font-family=\"'{font}'\" font-size=\"{size}\"")
               .Append((_txFace & 1) != 0 ? " font-weight=\"bold\"" : "")
               .Append((_txFace & 2) != 0 ? " font-style=\"italic\"" : "")
               .Append((_txFace & 4) != 0 ? " text-decoration=\"underline\"" : "")
               .Append($" fill=\"{Hex(_fg)}\" xml:space=\"preserve\">{esc}</text>");
        }

        /// <summary>Строки растра: без упаковки (rowBytes &lt; 8 или packType 1/2) или PackBits со счётчиком байтов строки.</summary>
        List<byte[]> ReadRows(int rowBytes, int height, int packType, int unit, int rawRowLen = 0)
        {
            var rows = new List<byte[]>();
            for (int y = 0; y < height; y++)
            {
                if (rawRowLen > 0 || rowBytes < 8 || packType == 1)
                {
                    int n = rawRowLen > 0 ? rawRowLen : rowBytes;
                    var row = new byte[n];
                    Array.Copy(d, _p, row, 0, Math.Min(n, d.Length - _p));
                    _p += n;
                    rows.Add(row);
                    continue;
                }
                int cnt = rowBytes > 250 ? U16() : U8();
                int end = _p + cnt;
                var outRow = new byte[rowBytes];
                int o = 0;
                while (_p < end && o < rowBytes)
                {
                    int f = (sbyte)d[_p++];
                    if (f >= 0)
                    {
                        int len = (f + 1) * unit;
                        for (int i = 0; i < len && _p < end; i++) { if (o < rowBytes) outRow[o++] = d[_p]; _p++; }
                    }
                    else if (f != -128)
                    {
                        int rep = 1 - f;
                        var u = new byte[unit];
                        for (int i = 0; i < unit; i++) u[i] = d[_p++];
                        for (int k = 0; k < rep; k++) for (int i = 0; i < unit && o < rowBytes; i++) outRow[o++] = u[i];
                    }
                }
                _p = end;
                rows.Add(outRow);
            }
            return rows;
        }

        /// <summary>BitsRect / PackBitsRect / DirectBitsRect (и варианты с областью).</summary>
        void Bits(int op, bool direct)
        {
            if (direct) _p += 4;                                     // baseAddr
            int rowBytesRaw = U16();
            bool pixmap = (rowBytesRaw & 0x8000) != 0 || direct;
            int rowBytes = rowBytesRaw & 0x3FFF;
            var bounds = Rect();
            int packType = 0, pixelSize = 1, cmpCount = 1;
            var palette = new List<(int, int, int)> { (255, 255, 255), (0, 0, 0) };
            if (pixmap)
            {
                _p += 2;                                             // pmVersion
                packType = U16();
                _p += 4 + 8;                                         // packSize, hRes, vRes
                _p += 2;                                             // pixelType
                pixelSize = U16();
                cmpCount = U16();
                _p += 2 + 4 + 4 + 4;                                 // cmpSize, planeBytes, pmTable, pmReserved
                if (!direct)
                {
                    _p += 4;
                    int flags = U16(), n = U16() + 1;
                    palette = Enumerable.Repeat((0, 0, 0), Math.Max(n, 1 << Math.Min(8, pixelSize))).ToList();
                    for (int i = 0; i < n; i++)
                    {
                        int idx = U16();
                        var c = Rgb48();
                        int at = (flags & 0x8000) != 0 ? i : idx;
                        if (at >= 0 && at < palette.Count) palette[at] = c;
                    }
                }
            }
            var src = Rect();
            var dst = Rect();
            _p += 2;                                                 // mode
            if (op is 0x91 or 0x99 or 0x9B) { int n = U16(); _p += n - 2; }
            int w = bounds.R - bounds.L, h = bounds.B - bounds.T;
            if (w <= 0 || h <= 0 || w > 20000 || h > 20000) return;
            List<byte[]> rows;
            if (direct && pixelSize == 32 && packType == 2) rows = ReadRows(rowBytes, h, packType, 1, rawRowLen: w * 3);
            else rows = ReadRows(rowBytes, h, packType, pixelSize == 16 && packType != 1 ? 2 : 1);   // 16 бит — PackBits по словам
            // Пиксели → RGB
            var rgb = new byte[w * h * 3];
            for (int y = 0; y < h; y++)
            {
                var row = rows[y];
                for (int x = 0; x < w; x++)
                {
                    int r, g, b;
                    if (pixelSize == 32 || pixelSize == 24)
                    {
                        if (direct && packType == 2) { r = Get(row, x * 3); g = Get(row, x * 3 + 1); b = Get(row, x * 3 + 2); }
                        else if (rowBytes >= 8 && packType != 1 && direct)
                        {
                            // Покомпонентно: строки R, G, B (при cmpCount 4 — сначала A)
                            int k = cmpCount >= 4 ? 1 : 0;
                            r = Get(row, (k + 0) * w + x); g = Get(row, (k + 1) * w + x); b = Get(row, (k + 2) * w + x);
                        }
                        else { r = Get(row, x * 4 + 1); g = Get(row, x * 4 + 2); b = Get(row, x * 4 + 3); }
                    }
                    else if (pixelSize == 16)
                    {
                        int v = Get(row, x * 2) << 8 | Get(row, x * 2 + 1);
                        r = (v >> 10 & 31) * 255 / 31; g = (v >> 5 & 31) * 255 / 31; b = (v & 31) * 255 / 31;
                    }
                    else
                    {
                        int bitPos = x * pixelSize, idx = Get(row, bitPos >> 3) >> (8 - pixelSize - (bitPos & 7)) & ((1 << pixelSize) - 1);
                        (r, g, b) = idx < palette.Count ? palette[idx] : (0, 0, 0);
                    }
                    int o = (y * w + x) * 3;
                    rgb[o] = (byte)r; rgb[o + 1] = (byte)g; rgb[o + 2] = (byte)b;
                }
            }
            // Часть растра (srcRect) → прямоугольник dstRect
            int sx0 = Math.Clamp(src.L - bounds.L, 0, w), sy0 = Math.Clamp(src.T - bounds.T, 0, h);
            int sx1 = Math.Clamp(src.R - bounds.L, sx0, w), sy1 = Math.Clamp(src.B - bounds.T, sy0, h);
            if (sx1 <= sx0 || sy1 <= sy0) { sx0 = 0; sy0 = 0; sx1 = w; sy1 = h; }
            var bmp = Bmp(rgb, w, sx0, sy0, sx1 - sx0, sy1 - sy0);
            _sb.Append($"<image x=\"{X(dst.L)}\" y=\"{Y(dst.T)}\" width=\"{F(dst.R - dst.L)}\" height=\"{F(dst.B - dst.T)}\" preserveAspectRatio=\"none\" href=\"data:image/bmp;base64,{Convert.ToBase64String(bmp)}\"/>");
        }

        static int Get(byte[] row, int i) => i >= 0 && i < row.Length ? row[i] : 0;

        /// <summary>24-битный BMP (строки сверху вниз) из части RGB-растра.</summary>
        static byte[] Bmp(byte[] rgb, int fullW, int x0, int y0, int w, int h)
        {
            int stride = (w * 3 + 3) / 4 * 4, off = 54;
            var b = new byte[off + stride * h];
            b[0] = (byte)'B'; b[1] = (byte)'M';
            BitConverter.GetBytes(b.Length).CopyTo(b, 2);
            BitConverter.GetBytes(off).CopyTo(b, 10);
            BitConverter.GetBytes(40).CopyTo(b, 14);
            BitConverter.GetBytes(w).CopyTo(b, 18);
            BitConverter.GetBytes(-h).CopyTo(b, 22);
            BitConverter.GetBytes((short)1).CopyTo(b, 26);
            BitConverter.GetBytes((short)24).CopyTo(b, 28);
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    int s = ((y0 + y) * fullW + x0 + x) * 3, t = off + y * stride + x * 3;
                    b[t] = rgb[s + 2]; b[t + 1] = rgb[s + 1]; b[t + 2] = rgb[s];
                }
            return b;
        }

        /// <summary>Сжатый QuickTime (0x8200): JPEG внутри, положение — srcRect, преобразованный матрицей.</summary>
        void QuickTime()
        {
            long size = U32();
            int start = _p, end = (int)Math.Min(d.Length, start + size);
            try
            {
                double Fixed(int at) => (int)((uint)(d[at] << 24 | d[at + 1] << 16 | d[at + 2] << 8 | d[at + 3])) / 65536.0;
                double a = Fixed(start + 2), bb = Fixed(start + 6), c = Fixed(start + 14), dd = Fixed(start + 18), tx = Fixed(start + 26), ty = Fixed(start + 30);
                _p = start + 52;
                var sr = Rect();
                double x0 = a * sr.L + c * sr.T + tx, y0 = bb * sr.L + dd * sr.T + ty;
                double x1 = a * sr.R + c * sr.B + tx, y1 = bb * sr.R + dd * sr.B + ty;
                int soi = -1;
                for (int i = start; i + 2 < end; i++) if (d[i] == 0xFF && d[i + 1] == 0xD8 && d[i + 2] == 0xFF) { soi = i; break; }
                if (soi >= 0)
                {
                    var jpg = d.AsSpan(soi, end - soi).ToArray();
                    _sb.Append($"<image x=\"{X(Math.Min(x0, x1))}\" y=\"{Y(Math.Min(y0, y1))}\" width=\"{F(Math.Abs(x1 - x0))}\" height=\"{F(Math.Abs(y1 - y0))}\" preserveAspectRatio=\"none\" href=\"data:image/jpeg;base64,{Convert.ToBase64String(jpg)}\"/>");
                }
            }
            finally { _p = end; }
        }
    }
}

using System.Globalization;
using System.IO.Compression;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Images;

namespace DocxToPdf.Pdf;

/// <summary>Шрифт, используемый в PDF: копит набор задействованных глифов.</summary>
sealed class PdfFont(TrueTypeFont font, string name)
{
    public readonly TrueTypeFont Font = font;
    public readonly string Name = name;
    public readonly SortedDictionary<ushort, string> Used = new();
}

sealed class PdfImage(PdfImageData data, string name)
{
    public readonly PdfImageData Data = data;
    public readonly string Name = name;
}

sealed class PdfLink(double x1, double y1, double x2, double y2, string uri)
{
    public double X1 = x1, Y1 = y1, X2 = x2, Y2 = y2;
    public string Uri = uri;
}

sealed class PdfPage
{
    public double Width, Height;
    public PdfCanvas Canvas = null!;
}

/// <summary>Генератор PDF 1.7: объекты, потоки, таблица xref.</summary>
public sealed class PdfDocument
{
    readonly List<PdfPage> _pages = [];
    readonly Dictionary<TrueTypeFont, PdfFont> _fonts = new();
    readonly Dictionary<string, PdfImage?> _images = new();
    public string? Title, Author, Subject;
    /// <summary>Цели внутренних ссылок: имя → (индекс страницы, x, y сверху).</summary>
    public readonly Dictionary<string, (int Page, double X, double Y)> Destinations = [];

    public PdfCanvas AddPage(double width, double height)
    {
        var page = new PdfPage { Width = width, Height = height };
        page.Canvas = new PdfCanvas(this, height);
        _pages.Add(page);
        return page.Canvas;
    }

    internal PdfFont GetFont(TrueTypeFont f)
    {
        if (!_fonts.TryGetValue(f, out var pf))
            _fonts[f] = pf = new PdfFont(f, "F" + (_fonts.Count + 1));
        return pf;
    }

    internal PdfImage? GetImage(string key, byte[] bytes, bool gray = false)
    {
        if (gray) key += "#gray";
        if (_images.TryGetValue(key, out var img)) return img;
        var data = ImageDecoder.Decode(bytes);
        if (gray && data != null) data = ToGray(data);
        img = data == null ? null : new PdfImage(data, "Im" + (_images.Count + 1));
        _images[key] = img;
        return img;
    }

    /// <summary>Оттенки серого (a:grayscl): RGB → яркость. JPEG (DCT) не перекодируется и остаётся цветным.</summary>
    static PdfImageData ToGray(PdfImageData d)
    {
        if (d.Filter != "/FlateDecode" || d.ColorSpace != "/DeviceRGB" || d.BitsPerComponent != 8) return d;
        var g = new byte[d.Width * d.Height];
        for (int i = 0; i < g.Length && i * 3 + 2 < d.Data.Length; i++)
            g[i] = (byte)Math.Round(0.299 * d.Data[i * 3] + 0.587 * d.Data[i * 3 + 1] + 0.114 * d.Data[i * 3 + 2]);
        return new PdfImageData { Width = d.Width, Height = d.Height, ColorSpace = "/DeviceGray", Data = g, Alpha = d.Alpha };
    }

    // ================================================================
    //  Запись
    // ================================================================

    internal static string F(double v)
    {
        if (Math.Abs(v) < 0.0005) return "0";
        return Math.Round(v, 3).ToString("0.###", CultureInfo.InvariantCulture);
    }

    sealed class Writer(Stream s)
    {
        readonly Stream _s = s;
        public readonly List<long> Offsets = [0];

        public int Reserve() { Offsets.Add(-1); return Offsets.Count - 1; }

        public void Raw(string text) { var b = Encoding.Latin1.GetBytes(text); _s.Write(b); }
        public void Raw(byte[] b) => _s.Write(b);

        public void Obj(int id, string body)
        {
            Offsets[id] = _s.Position;
            Raw($"{id} 0 obj\n{body}\nendobj\n");
        }

        public void Stream(int id, string dict, byte[] data, bool compress = true)
        {
            Offsets[id] = _s.Position;
            byte[] payload = data;
            string filter = "";
            if (compress)
            {
                payload = Deflate(data);
                filter = " /Filter /FlateDecode";
            }
            Raw($"{id} 0 obj\n<< {dict}{filter} /Length {payload.Length} >>\nstream\n");
            Raw(payload);
            Raw("\nendstream\nendobj\n");
        }

        public long Position => _s.Position;
    }

    static byte[] Deflate(byte[] data)
    {
        using var ms = new MemoryStream();
        using (var z = new ZLibStream(ms, CompressionLevel.Optimal, leaveOpen: true)) z.Write(data);
        return ms.ToArray();
    }

    static string PdfString(string s)
    {
        if (s.All(c => c >= 32 && c < 127))
            return "(" + s.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)") + ")";
        var sb = new StringBuilder("<FEFF");
        foreach (var b in Encoding.BigEndianUnicode.GetBytes(s)) sb.Append(b.ToString("X2"));
        return sb.Append('>').ToString();
    }

    public void Save(string path)
    {
        using var fs = File.Create(path);
        Save(fs);
    }

    public void Save(Stream output)
    {
        var w = new Writer(output);
        w.Raw("%PDF-1.7\n%âãÏÓ\n");

        int catalog = w.Reserve(), pagesId = w.Reserve(), info = w.Reserve();

        var fontIds = _fonts.Values.ToDictionary(f => f, _ => w.Reserve());
        var imageIds = _images.Values.Where(i => i != null).ToDictionary(i => i!, _ => w.Reserve());

        // Номера объектов страниц резервируются заранее — на них ссылаются внутренние ссылки
        var pageIds = _pages.Select(_ => w.Reserve()).ToList();
        for (int pi = 0; pi < _pages.Count; pi++)
        {
            var page = _pages[pi];
            int pageId = pageIds[pi], contentId = w.Reserve();
            var c = page.Canvas;

            var res = new StringBuilder("<< /ProcSet [/PDF /Text /ImageB /ImageC]");
            if (c.Fonts.Count > 0)
            {
                res.Append(" /Font <<");
                foreach (var f in c.Fonts) res.Append($" /{f.Name} {fontIds[f]} 0 R");
                res.Append(" >>");
            }
            var formIds = c.Forms.Select(_ => w.Reserve()).ToList();
            if (c.Images.Count > 0 || c.Forms.Count > 0)
            {
                res.Append(" /XObject <<");
                foreach (var i in c.Images) res.Append($" /{i.Name} {imageIds[i]} 0 R");
                for (int fi = 0; fi < c.Forms.Count; fi++) res.Append($" /{c.Forms[fi].Name} {formIds[fi]} 0 R");
                res.Append(" >>");
            }
            if (c.GStates.Count > 0)
            {
                res.Append(" /ExtGState <<");
                foreach (var (name, a) in c.GStates) res.Append($" /{name} << /ca {F(a)} /CA {F(a)} >>");
                res.Append(" >>");
            }
            res.Append(" >>");
            int resId = w.Reserve();
            w.Obj(resId, res.ToString());
            for (int fi = 0; fi < c.Forms.Count; fi++)
                w.Stream(formIds[fi],
                    $"/Type /XObject /Subtype /Form /BBox [0 0 {F(page.Width)} {F(page.Height)}] /Group << /S /Transparency /I true >> /Resources {resId} 0 R",
                    Encoding.ASCII.GetBytes(c.Forms[fi].Content));

            var annots = new List<int>();
            foreach (var link in c.Links)
            {
                string action;
                if (link.Uri.StartsWith('#'))
                {
                    // Ссылка на закладку внутри документа: переход к странице и месту
                    if (!Destinations.TryGetValue(link.Uri[1..], out var dest) || dest.Page >= _pages.Count) continue;
                    double dy = _pages[dest.Page].Height - dest.Y;
                    action = $"/Dest [{pageIds[dest.Page]} 0 R /XYZ {F(dest.X)} {F(dy)} null]";
                }
                else action = $"/A << /S /URI /URI {PdfString(link.Uri)} >>";
                int id = w.Reserve();
                annots.Add(id);
                w.Obj(id, $"<< /Type /Annot /Subtype /Link /Rect [{F(link.X1)} {F(link.Y1)} {F(link.X2)} {F(link.Y2)}] /Border [0 0 0] {action} >>");
            }
            string annotStr = annots.Count > 0 ? " /Annots [" + string.Join(" ", annots.Select(a => a + " 0 R")) + "]" : "";
            w.Obj(pageId, $"<< /Type /Page /Parent {pagesId} 0 R /MediaBox [0 0 {F(page.Width)} {F(page.Height)}] /Resources {resId} 0 R /Contents {contentId} 0 R{annotStr} >>");
            w.Stream(contentId, "", c.GetContent());
        }

        foreach (var (img, id) in imageIds) WriteImage(w, img, id);
        foreach (var (font, id) in fontIds) WriteFont(w, font, id);

        w.Obj(pagesId, $"<< /Type /Pages /Kids [{string.Join(" ", pageIds.Select(p => p + " 0 R"))}] /Count {pageIds.Count} >>");
        w.Obj(catalog, $"<< /Type /Catalog /Pages {pagesId} 0 R >>");
        var infoSb = new StringBuilder("<< /Producer (docx2pdf .NET)");
        infoSb.Append(" /CreationDate (D:" + DateTime.UtcNow.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) + "Z)");
        if (Title != null) infoSb.Append(" /Title " + PdfString(Title));
        if (Author != null) infoSb.Append(" /Author " + PdfString(Author));
        if (Subject != null) infoSb.Append(" /Subject " + PdfString(Subject));
        w.Obj(info, infoSb.Append(" >>").ToString());

        long xref = w.Position;
        var sb = new StringBuilder();
        sb.Append($"xref\n0 {w.Offsets.Count}\n0000000000 65535 f \n");
        for (int i = 1; i < w.Offsets.Count; i++) sb.Append($"{w.Offsets[i]:D10} 00000 n \n");
        sb.Append($"trailer\n<< /Size {w.Offsets.Count} /Root {catalog} 0 R /Info {info} 0 R >>\nstartxref\n{xref}\n%%EOF\n");
        w.Raw(sb.ToString());
    }

    static void WriteImage(Writer w, PdfImage img, int id)
    {
        var d = img.Data;
        string smask = "";
        if (d.Alpha != null)
        {
            int maskId = w.Reserve();
            w.Stream(maskId, $"/Type /XObject /Subtype /Image /Width {d.Width} /Height {d.Height} /ColorSpace /DeviceGray /BitsPerComponent 8", d.Alpha);
            smask = $" /SMask {maskId} 0 R";
        }
        string decode = d.Decode != null ? " /Decode " + d.Decode : "";
        string dict = $"/Type /XObject /Subtype /Image /Width {d.Width} /Height {d.Height} /ColorSpace {d.ColorSpace} /BitsPerComponent {d.BitsPerComponent}{decode}{smask}";
        if (d.Filter == "/DCTDecode") w.Stream(id, dict + " /Filter /DCTDecode", d.Data, compress: false);
        else w.Stream(id, dict, d.Data);
    }

    static void WriteFont(Writer w, PdfFont pf, int type0Id)
    {
        var f = pf.Font;
        int cidId = w.Reserve(), descId = w.Reserve(), fileId = w.Reserve(), toUniId = w.Reserve();
        string tag = SubsetTag(pf);
        string baseFont = tag + "+" + f.PostScriptName;
        double k = 1000.0 / f.UnitsPerEm;

        w.Obj(type0Id, $"<< /Type /Font /Subtype /Type0 /BaseFont /{baseFont} /Encoding /Identity-H /DescendantFonts [{cidId} 0 R] /ToUnicode {toUniId} 0 R >>");

        // Ширины: группы подряд идущих CID
        var widths = new StringBuilder("[");
        var gids = pf.Used.Keys.ToList();
        for (int i = 0; i < gids.Count;)
        {
            int start = i;
            while (i + 1 < gids.Count && gids[i + 1] == gids[i] + 1) i++;
            widths.Append($" {gids[start]} [");
            for (int j = start; j <= i; j++) widths.Append(' ').Append(F(f.Advance(gids[j]) * 1000));
            widths.Append(" ]");
            i++;
        }
        widths.Append(" ]");

        w.Obj(cidId, $"<< /Type /Font /Subtype /CIDFontType2 /BaseFont /{baseFont} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor {descId} 0 R /DW 1000 /W {widths} /CIDToGIDMap /Identity >>");

        int flags = 4; // symbolic (корректно для Identity-кодировки)
        if (f.IsFixedPitch) flags |= 1;
        if (f.IsSerif) flags |= 2;
        if (f.IsItalic || f.ItalicAngle != 0) flags |= 64;
        w.Obj(descId, $"<< /Type /FontDescriptor /FontName /{baseFont} /Flags {flags} /FontBBox [{F(f.XMin * k)} {F(f.YMin * k)} {F(f.XMax * k)} {F(f.YMax * k)}] /ItalicAngle {F(f.ItalicAngle)} /Ascent {F(f.Ascent * k)} /Descent {F(-f.Descent * k)} /CapHeight {F(f.CapHeight * k)} /StemV {(f.IsBold ? 120 : 80)} /FontFile2 {fileId} 0 R >>");

        var sub = f.Subset(pf.Used.Keys);
        w.Stream(fileId, $"/Length1 {sub.Length}", sub);

        // ToUnicode CMap
        var cm = new StringBuilder();
        cm.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n");
        cm.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
        cm.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");
        var entries = pf.Used.Where(e => e.Value.Length > 0).ToList();
        for (int i = 0; i < entries.Count; i += 100)
        {
            var chunk = entries.Skip(i).Take(100).ToList();
            cm.Append(chunk.Count).Append(" beginbfchar\n");
            foreach (var (gid, text) in chunk)
            {
                cm.Append('<').Append(gid.ToString("X4")).Append("> <");
                foreach (var b in Encoding.BigEndianUnicode.GetBytes(text)) cm.Append(b.ToString("X2"));
                cm.Append(">\n");
            }
            cm.Append("endbfchar\n");
        }
        cm.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        w.Stream(toUniId, "", Encoding.ASCII.GetBytes(cm.ToString()));
    }

    static string SubsetTag(PdfFont pf)
    {
        uint h = 2166136261;
        foreach (var g in pf.Used.Keys) { h ^= g; h *= 16777619; }
        foreach (var c in pf.Name) { h ^= c; h *= 16777619; }
        var sb = new StringBuilder();
        for (int i = 0; i < 6; i++) { sb.Append((char)('A' + h % 26)); h /= 26; }
        return sb.ToString();
    }
}

/// <summary>
/// Холст страницы. Координаты принимаются в системе "сверху-вниз" (y от верхнего края, в пунктах)
/// и переводятся в систему PDF.
/// </summary>
public sealed class PdfCanvas
{
    readonly PdfDocument _doc;
    readonly double _pageHeight;
    StringBuilder _sb = new();
    /// <summary>Формы (Form XObject) страницы: содержимое, нарисованное один раз и используемое многократно.</summary>
    internal readonly List<(string Name, string Content)> Forms = [];

    /// <summary>
    /// Записывает рисование в отдельную форму (прозрачностная группа) и возвращает её имя для DoForm.
    /// Ссылки внутри формы не создаются.
    /// </summary>
    public string CaptureForm(Action draw)
    {
        var saved = _sb;
        _sb = new StringBuilder();
        SuppressLinks++;
        draw();
        SuppressLinks--;
        string content = _sb.ToString();
        _sb = saved;
        string name = "Fm" + (Forms.Count + 1);
        Forms.Add((name, content));
        return name;
    }

    public void DoForm(string name)
    {
        if (Collecting > 0) return;
        _sb.Append('/').Append(name).Append(" Do\n");
    }
    internal readonly List<PdfFont> Fonts = [];
    internal readonly List<PdfImage> Images = [];
    /// <summary>Графические состояния прозрачности страницы: имя → непрозрачность.</summary>
    internal readonly List<(string Name, double Alpha)> GStates = [];

    /// <summary>Имя ExtGState с заданной непрозрачностью (для «/GSn gs»).</summary>
    string GState(double alpha)
    {
        alpha = Math.Round(Math.Clamp(alpha, 0, 1), 3);
        foreach (var (n, a) in GStates) if (a == alpha) return n;
        string name = "GS" + (GStates.Count + 1);
        GStates.Add((name, alpha));
        return name;
    }
    internal readonly List<PdfLink> Links = [];

    internal PdfCanvas(PdfDocument doc, double pageHeight)
    {
        _doc = doc;
        _pageHeight = pageHeight;
    }

    static string F(double v) => PdfDocument.F(v);
    double Y(double top) => _pageHeight - top;

    internal byte[] GetContent() => Encoding.ASCII.GetBytes(_sb.ToString());

    static string Color(Rgb c) => $"{F(c.R)} {F(c.G)} {F(c.B)}";

    public void FillRect(double x, double y, double w, double h, Rgb color)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        if (w <= 0 || h <= 0) return;
        _sb.Append($"{Color(color)} rg {F(x)} {F(Y(y + h))} {F(w)} {F(h)} re f\n");
    }

    public void Line(double x1, double y1, double x2, double y2, double width, Rgb color, bool dashed = false, bool dotted = false)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        _sb.Append("q ");
        if (dashed) _sb.Append($"[{F(width * 3)} {F(width * 2)}] 0 d ");
        else if (dotted) _sb.Append($"[{F(width)} {F(width * 1.5)}] 0 d ");
        _sb.Append($"{F(width)} w {Color(color)} RG {F(x1)} {F(Y(y1))} m {F(x2)} {F(Y(y2))} l S Q\n");
    }

    /// <summary>Волнистая линия (подчёркивание wave) от x длиной len на высоте y.</summary>
    public void WavyLine(double x, double y, double len, double width, Rgb color)
    {
        if (Collecting > 0 || SkipGraphics || len <= 0) return;
        double amp = Math.Max(0.6, width * 1.2), half = Math.Max(1.5, width * 2.5);
        _sb.Append($"q {F(width)} w {Color(color)} RG 1 J {F(x)} {F(Y(y))} m ");
        double cx = x; int sign = -1;
        while (cx < x + len - 0.01)
        {
            double nx = Math.Min(x + len, cx + half), k = (nx - cx) / half;
            _sb.Append($"{F(cx + (nx - cx) / 3)} {F(Y(y + sign * amp * 1.33 * k))} {F(nx - (nx - cx) / 3)} {F(Y(y + sign * amp * 1.33 * k))} {F(nx)} {F(Y(y))} c ");
            cx = nx; sign = -sign;
        }
        _sb.Append("S Q\n");
    }

    /// <summary>Закрашенный многоугольник (точки в координатах страницы, ось y вниз).</summary>
    public void FillPolygon(IReadOnlyList<(double X, double Y)> pts, Rgb color)
    {
        if (Collecting > 0 || SkipGraphics || pts.Count < 3) return;
        _sb.Append($"{Color(color)} rg {F(pts[0].X)} {F(Y(pts[0].Y))} m ");
        for (int i = 1; i < pts.Count; i++) _sb.Append($"{F(pts[i].X)} {F(Y(pts[i].Y))} l ");
        _sb.Append("h f\n");
    }

    /// <summary>Ломаная линия.</summary>
    public void Polyline(IReadOnlyList<(double X, double Y)> pts, double width, Rgb color)
    {
        if (Collecting > 0 || SkipGraphics || pts.Count < 2) return;
        _sb.Append($"q {F(width)} w 1 J 1 j {Color(color)} RG {F(pts[0].X)} {F(Y(pts[0].Y))} m ");
        for (int i = 1; i < pts.Count; i++) _sb.Append($"{F(pts[i].X)} {F(Y(pts[i].Y))} l ");
        _sb.Append("S Q\n");
    }

    /// <summary>Сектор круга (углы в градусах по часовой стрелке от 12 часов), с белой обводкой.</summary>
    public void FillSector(double cx, double cy, double r, double a0, double a1, Rgb color, double innerR = 0)
    {
        if (Collecting > 0 || SkipGraphics || a1 - a0 < 0.01) return;
        (double, double) Pt(double ang, double rr) { double t = (ang - 90) * Math.PI / 180; return (cx + rr * Math.Cos(t), cy + rr * Math.Sin(t)); }
        var sb = new StringBuilder();
        void Arc(double from, double to, double rr, bool move)
        {
            int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(to - from) / 90));
            double step = (to - from) / segs;
            var (x0, y0) = Pt(from, rr);
            if (move) sb.Append($"{F(x0)} {F(Y(y0))} m "); else sb.Append($"{F(x0)} {F(Y(y0))} l ");
            for (int i = 0; i < segs; i++)
            {
                double s0 = from + i * step, s1 = s0 + step, k = 4.0 / 3 * Math.Tan(step * Math.PI / 180 / 4);
                var (ax, ay) = Pt(s0, rr); var (bx, by) = Pt(s1, rr);
                double t0 = (s0 - 90) * Math.PI / 180, t1 = (s1 - 90) * Math.PI / 180;
                double c1x = ax - k * rr * Math.Sin(t0), c1y = ay + k * rr * Math.Cos(t0);
                double c2x = bx + k * rr * Math.Sin(t1), c2y = by - k * rr * Math.Cos(t1);
                sb.Append($"{F(c1x)} {F(Y(c1y))} {F(c2x)} {F(Y(c2y))} {F(bx)} {F(Y(by))} c ");
            }
        }
        if (innerR > 0) { Arc(a0, a1, r, true); Arc(a1, a0, innerR, false); }
        else { sb.Append($"{F(cx)} {F(Y(cy))} m "); Arc(a0, a1, r, false); }
        _sb.Append($"q {Color(color)} rg 1 1 1 RG 0.75 w ").Append(sb).Append("h B Q\n");
    }

    public void Rect(double x, double y, double w, double h, double width, Rgb color)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        _sb.Append($"q {F(width)} w {Color(color)} RG {F(x)} {F(Y(y + h))} {F(w)} {F(h)} re S Q\n");
    }

    /// <summary>Выводит глифы по базовой линии baseline.</summary>
    /// <summary>Больше нуля — текст выводится контурами (без текстового слоя).</summary>
    public int TextAsPaths;
    /// <summary>Больше нуля — выводится только невидимый текстовый слой (для поиска и копирования).</summary>
    public int InvisibleTextOnly;

    bool SkipGraphics => InvisibleTextOnly > 0;

    /// <summary>Номер текущей страницы и их общее число (для полей PAGE/NUMPAGES в тексте).</summary>
    public int PageNumber, PageCount;
    public string? PageNumberText;
    /// <summary>Номер страницы закладки для полей PAGEREF (известен после разбивки на страницы).</summary>
    public Func<string, string?>? PageRefText;

    public void Glyphs(FontFace face, double size, double x, double baseline, IReadOnlyList<ushort> gids, IReadOnlyList<string> texts, Rgb color,
        IReadOnlyList<double>? extra = null, double hscale = 1, bool outline = false)
    {
        if (Collecting > 0) return;
        if (gids.Count == 0) return;
        if (TextAsPaths > 0 && InvisibleTextOnly == 0)
        {
            // Контуры глифов: каждый — в своей системе координат через cm
            var font = face.Font;
            double k = size / font.UnitsPerEm, skewK = face.SynthItalic ? 0.2126 : 0;
            double gx = x;
            _sb.Append($"q {Color(color)} rg ");
            if (face.SynthBold || outline) _sb.Append($"{Color(color)} RG {F(size * (outline ? 0.04 : 0.03) / k)} w ");
            for (int i = 0; i < gids.Count; i++)
            {
                string path = font.GlyphPath(gids[i]);
                if (path.Length > 0)
                    _sb.Append($"q {F(k * hscale)} 0 {F(k * skewK)} {F(k)} {F(gx)} {F(Y(baseline))} cm ").Append(path).Append(outline ? "S Q " : face.SynthBold ? "B Q " : "f Q ");
                gx += font.Advance(gids[i]) * size * hscale + (extra != null && i < extra.Count ? extra[i] : 0);
            }
            _sb.Append("Q\n");
            return;
        }
        var pf = _doc.GetFont(face.Font);
        if (!Fonts.Contains(pf)) Fonts.Add(pf);
        for (int i = 0; i < gids.Count; i++) pf.Used.TryAdd(gids[i], texts[i]);
        if (pf.Used.TryGetValue(gids[0], out var existing) && existing.Length == 0) pf.Used[gids[0]] = texts[0];

        _sb.Append("BT ");
        _sb.Append($"{Color(color)} rg ");
        if (InvisibleTextOnly > 0) _sb.Append("3 Tr ");
        else if (outline) _sb.Append($"1 Tr {F(size * 0.04)} w {Color(color)} RG ");
        else if (face.SynthBold) _sb.Append($"2 Tr {F(size * 0.03)} w {Color(color)} RG ");
        _sb.Append($"/{pf.Name} {F(size)} Tf ");
        if (Math.Abs(hscale - 1) > 1e-6) _sb.Append($"{F(hscale * 100)} Tz ");
        double skew = face.SynthItalic ? 0.2126 : 0;
        _sb.Append($"1 0 {F(skew)} 1 {F(x)} {F(Y(baseline))} Tm ");
        if (extra == null)
        {
            _sb.Append('<');
            foreach (var g in gids) _sb.Append(g.ToString("X4"));
            _sb.Append("> Tj");
        }
        else
        {
            // TJ: дополнительный сдвиг после каждого глифа (для выравнивания по ширине)
            _sb.Append("[<");
            for (int i = 0; i < gids.Count; i++)
            {
                _sb.Append(gids[i].ToString("X4"));
                if (i < gids.Count - 1 && Math.Abs(extra[i]) > 0.0001)
                    _sb.Append("> ").Append(F(-extra[i] * 1000 / size / hscale)).Append(" <");
            }
            _sb.Append(">] TJ");
        }
        _sb.Append(" ET\n");
        if (face.SynthBold || outline || InvisibleTextOnly > 0) _sb.Append("0 Tr\n");
        // Tz — параметр текстового состояния, сохраняется после ET: возвращаем 100%
        if (Math.Abs(hscale - 1) > 1e-6) _sb.Append("100 Tz\n");
    }

    // ---------------------------------------------------------------
    //  Скруглённые прямоугольники (border-radius)
    // ---------------------------------------------------------------

    /// <summary>
    /// Радиусы (pt): rx TL, TR, BR, BL, затем ry TL, TR, BR, BL. Как в CSS: если сумма радиусов
    /// вдоль стороны больше стороны — все радиусы уменьшаются в одной пропорции.
    /// </summary>
    static double[] ClampRadii(double w, double h, double[] r)
    {
        var c = r.Select(v => Math.Max(0, v)).ToArray();
        double k = 1;
        void Side(double len, double sum) { if (sum > len && sum > 0) k = Math.Min(k, len / sum); }
        Side(w, c[0] + c[1]);   // верх: rx TL + rx TR
        Side(w, c[3] + c[2]);   // низ: rx BL + rx BR
        Side(h, c[4] + c[7]);   // лево: ry TL + ry BL
        Side(h, c[5] + c[6]);   // право: ry TR + ry BR
        for (int i = 0; i < 8; i++) c[i] *= k;
        // Угол с одним нулевым радиусом — прямой
        for (int i = 0; i < 4; i++) if (c[i] <= 0 || c[i + 4] <= 0) c[i] = c[i + 4] = 0;
        return c;
    }

    void RoundedPath(double x, double y, double w, double h, double[] radii)
    {
        var r = ClampRadii(w, h, radii);
        const double K = 0.5523;   // аппроксимация четверти эллипса кривой Безье
        double x2 = x + w, y2 = y + h;
        double txl = r[0], trx = r[1], brx = r[2], blx = r[3], tly = r[4], try_ = r[5], bry = r[6], bly = r[7];
        void M(double px, double py) => _sb.Append($"{F(px)} {F(Y(py))} m ");
        void L(double px, double py) => _sb.Append($"{F(px)} {F(Y(py))} l ");
        void C(double ax, double ay, double bx, double by, double px, double py) =>
            _sb.Append($"{F(ax)} {F(Y(ay))} {F(bx)} {F(Y(by))} {F(px)} {F(Y(py))} c ");
        M(x + txl, y);
        L(x2 - trx, y);
        if (trx > 0) C(x2 - trx + trx * K, y, x2, y + try_ - try_ * K, x2, y + try_);
        L(x2, y2 - bry);
        if (brx > 0) C(x2, y2 - bry + bry * K, x2 - brx + brx * K, y2, x2 - brx, y2);
        L(x + blx, y2);
        if (blx > 0) C(x + blx - blx * K, y2, x, y2 - bly + bly * K, x, y2 - bly);
        L(x, y + tly);
        if (txl > 0) C(x, y + tly - tly * K, x + txl - txl * K, y, x + txl, y);
        _sb.Append("h ");
    }

    public static bool HasRadius(double[]? r) => r != null && r.Any(v => v > 0);

    /// <summary>Уменьшает радиусы на величину inset (для внутренних контуров рамки).</summary>
    static double[] Inset(double[] r, double d) => r.Select(v => Math.Max(0, v - d)).ToArray();

    public void FillRoundedRect(double x, double y, double w, double h, double[] radii, Rgb color)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        if (w <= 0 || h <= 0) return;
        _sb.Append($"{Color(color)} rg ");
        RoundedPath(x, y, w, h, radii);
        _sb.Append("f\n");
    }

    /// <summary>Обводка скруглённого прямоугольника; линия идёт по середине толщины (внутри рамки).</summary>
    public void StrokeRoundedRect(double x, double y, double w, double h, double[] radii, BorderDef b)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        if (b.IsNone || w <= 0 || h <= 0) return;
        if (b.Val == "double")
        {
            double d = b.Width / 3;
            var thin = new BorderDef { Val = "single", Width = d, Color = b.Color };
            var outer = ClampRadii(w, h, radii);
            StrokeRoundedRect(x, y, w, h, outer, thin);
            var inner = Inset(outer, 2 * d);
            StrokeRoundedRect(x + 2 * d, y + 2 * d, w - 4 * d, h - 4 * d, inner, thin);
            return;
        }
        double t = b.Width, half = t / 2;
        var rr = Inset(ClampRadii(w, h, radii), half);
        _sb.Append("q ");
        if (b.Val.Contains("dash")) _sb.Append($"[{F(t * 3)} {F(t * 2)}] 0 d ");
        else if (b.Val.Contains("dot")) _sb.Append($"1 J [0 {F(t * 2)}] 0 d ");
        _sb.Append($"{F(t)} w {Color(b.Color)} RG ");
        RoundedPath(x + half, y + half, w - t, h - t, rr);
        _sb.Append("S Q\n");
    }

    /// <summary>Ограничивает рисование скруглённой областью до EndClip().</summary>
    public void BeginRoundedClip(double x, double y, double w, double h, double[] radii)
    {
        if (Collecting > 0) return;
        _sb.Append("q ");
        RoundedPath(x, y, w, h, radii);
        _sb.Append("W n\n");
    }

    public void EndClip() { if (Collecting == 0) _sb.Append("Q\n"); }

    /// <summary>
    /// Начинает преобразованную область: матрица задана в координатах "y вниз" относительно точки (ox, oy) страницы.
    /// </summary>
    public void BeginTransform(double ox, double oy, double a, double b, double c, double d, double e, double f)
    {
        if (Collecting > 0) return;
        double yo = Y(oy);
        // В системе PDF (y вверх) матрица CSS превращается в [a, -b, -c, d, e, -f]
        _sb.Append($"q 1 0 0 1 {F(ox)} {F(yo)} cm {F(a)} {F(-b)} {F(-c)} {F(d)} {F(e)} {F(-f)} cm 1 0 0 1 {F(-ox)} {F(-yo)} cm\n");
    }

    public void EndTransform() { if (Collecting == 0) _sb.Append("Q\n"); }

    /// <summary>Больше нуля — холостой проход (сбор плоскостей 3D-сцены): ничего не выводится.</summary>
    public int Collecting;
    /// <summary>Текущая собираемая 3D-сцена (transform-style: preserve-3d).</summary>
    public object? Scene;
    /// <summary>Больше нуля — преобразованные потомки пропускаются (они рисуются отдельными плоскостями сцены).</summary>
    public int SkipTransformed;

    /// <summary>Обрезка по многоугольнику (координаты страницы, y вниз) до EndClip().</summary>
    public void BeginClipPolygon(IReadOnlyList<(double X, double Y)> pts)
    {
        if (Collecting > 0) return;
        _sb.Append("q ");
        for (int i = 0; i < pts.Count; i++)
            _sb.Append($"{F(pts[i].X)} {F(Y(pts[i].Y))} {(i == 0 ? "m" : "l")} ");
        _sb.Append("h W n\n");
    }

    /// <summary>Пока больше нуля — ссылки не добавляются (повторная отрисовка того же содержимого).</summary>
    public int SuppressLinks;

    /// <summary>
    /// Картинка с обрезкой, поворотом (по часовой стрелке, градусы) и отражением.
    /// (cx, cy) — центр рамки, w×h — размер неповёрнутой рамки (видимой части).
    /// </summary>
    public void ImageGeometry(string key, byte[] bytes, double cx, double cy, double w, double h,
        double cropL, double cropT, double cropR, double cropB, double rotation, bool flipH, bool flipV, double opacity = 1, bool gray = false)
    {
        if (Collecting > 0 || SkipGraphics) return;
        double a = -rotation * Math.PI / 180, cos = Math.Cos(a), sin = Math.Sin(a);
        double sx = flipH ? -1 : 1, sy = flipV ? -1 : 1;
        // Переход в систему рамки: начало в центре, ось Y вверх
        _sb.Append($"q {F(cos * sx)} {F(sin * sx)} {F(-sin * sy)} {F(cos * sy)} {F(cx)} {F(Y(cy))} cm ");
        if (opacity < 0.999) _sb.Append($"/{GState(opacity)} gs ");
        bool crop = cropL != 0 || cropT != 0 || cropR != 0 || cropB != 0;
        if (crop) _sb.Append($"{F(-w / 2)} {F(-h / 2)} {F(w)} {F(h)} re W n ");
        double fw = w / (1 - cropL - cropR), fh = h / (1 - cropT - cropB);
        double left = -w / 2 - cropL * fw, top = h / 2 + cropT * fh;
        var img = _doc.GetImage(key, bytes, gray);
        if (img == null)
        {
            _sb.Append($"0.93 0.93 0.93 rg {F(-w / 2)} {F(-h / 2)} {F(w)} {F(h)} re f ");
        }
        else
        {
            if (!Images.Contains(img)) Images.Add(img);
            _sb.Append($"{F(fw)} 0 0 {F(fh)} {F(left)} {F(top - fh)} cm /{img.Name} Do ");
        }
        _sb.Append("Q\n");
    }

    /// <summary>
    /// Фигура (ShapeSpec) в рамке w×h с центром (cx, cy), повёрнутая по часовой стрелке и отражённая.
    /// </summary>
    public void Shape(DocxToPdf.Docx.ShapeSpec sh, double cx, double cy, double w, double h, double rotation, bool flipH, bool flipV)
    {
        if (Collecting > 0 || SkipGraphics) return;
        double a = -rotation * Math.PI / 180, cos = Math.Cos(a), sin = Math.Sin(a);
        double sx = flipH ? -1 : 1, sy = flipV ? -1 : 1;
        _sb.Append($"q {F(cos * sx)} {F(sin * sx)} {F(-sin * sy)} {F(cos * sy)} {F(cx)} {F(Y(cy))} cm ");
        // Точка рамки (ось y вниз, от левого верхнего угла) → локальная система (центр, ось y вверх)
        string P(double px, double py) => $"{F(px - w / 2)} {F(h / 2 - py)}";
        var path = new StringBuilder();
        void Poly(params (double X, double Y)[] pts)
        {
            path.Append(P(pts[0].X, pts[0].Y)).Append(" m ");
            for (int i = 1; i < pts.Length; i++) path.Append(P(pts[i].X, pts[i].Y)).Append(" l ");
            path.Append("h ");
        }
        bool open = false;
        switch (sh.Geom)
        {
            case "line" or "straightConnector1" or "connector":
                path.Append(P(0, 0)).Append(" m ").Append(P(w, h)).Append(" l ");
                open = true;
                break;
            case "ellipse" or "oval":
            {
                const double K = 0.5523;
                double rx = w / 2, ry = h / 2;
                path.Append($"{F(rx)} 0 m {F(rx)} {F(ry * K)} {F(rx * K)} {F(ry)} 0 {F(ry)} c {F(-rx * K)} {F(ry)} {F(-rx)} {F(ry * K)} {F(-rx)} 0 c ");
                path.Append($"{F(-rx)} {F(-ry * K)} {F(-rx * K)} {F(-ry)} 0 {F(-ry)} c {F(rx * K)} {F(-ry)} {F(rx)} {F(-ry * K)} {F(rx)} 0 c h ");
                break;
            }
            case "roundRect" or "roundrect" or "flowChartTerminator" or "flowChartAlternateProcess":
            {
                double adj = sh.Geom == "flowChartTerminator" ? 0.5 : sh.Adj ?? 0.16667;
                double r = Math.Min(w, h) * Math.Clamp(adj, 0, 0.5), K = 0.5523 * r;
                double l = -w / 2, rr = w / 2, b = -h / 2, t = h / 2;
                path.Append($"{F(l + r)} {F(t)} m {F(rr - r)} {F(t)} l {F(rr - r + K)} {F(t)} {F(rr)} {F(t - r + K)} {F(rr)} {F(t - r)} c ");
                path.Append($"{F(rr)} {F(b + r)} l {F(rr)} {F(b + r - K)} {F(rr - r + K)} {F(b)} {F(rr - r)} {F(b)} c ");
                path.Append($"{F(l + r)} {F(b)} l {F(l + r - K)} {F(b)} {F(l)} {F(b + r - K)} {F(l)} {F(b + r)} c ");
                path.Append($"{F(l)} {F(t - r)} l {F(l)} {F(t - r + K)} {F(l + r - K)} {F(t)} {F(l + r)} {F(t)} c h ");
                break;
            }
            case "flowChartConnector":
            {
                const double K = 0.5523;
                double rx = w / 2, ry = h / 2;
                path.Append($"{F(rx)} 0 m {F(rx)} {F(ry * K)} {F(rx * K)} {F(ry)} 0 {F(ry)} c {F(-rx * K)} {F(ry)} {F(-rx)} {F(ry * K)} {F(-rx)} 0 c ");
                path.Append($"{F(-rx)} {F(-ry * K)} {F(-rx * K)} {F(-ry)} 0 {F(-ry)} c {F(rx * K)} {F(-ry)} {F(rx)} {F(-ry * K)} {F(rx)} 0 c h ");
                break;
            }
            case "flowChartDecision": Poly((w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)); break;
            case "star4" or "star5" or "star6" or "star7" or "star8" or "star10" or "star12":
            {
                // Звезда: вершины на эллипсе рамки, внутренние — на уменьшенном
                int n = int.Parse(sh.Geom[4..], CultureInfo.InvariantCulture);
                double inner = sh.Adj ?? n switch { 4 => 0.125, 5 => 0.19098, 6 => 0.28868, _ => 0.375 };
                inner = Math.Clamp(inner * 2, 0.05, 1);
                var pts = new (double X, double Y)[n * 2];
                for (int i = 0; i < n * 2; i++)
                {
                    double ang = -Math.PI / 2 + i * Math.PI / n, rr = i % 2 == 0 ? 1 : inner;
                    pts[i] = (w / 2 + w / 2 * rr * Math.Cos(ang), h / 2 + h / 2 * rr * Math.Sin(ang));
                }
                if (n == 5)
                {
                    // У пятиконечной звезды Word нижние вершины касаются края рамки
                    double minY = pts.Min(p => p.Y), maxY = pts.Max(p => p.Y);
                    for (int i = 0; i < pts.Length; i++) pts[i].Y = (pts[i].Y - minY) / (maxY - minY) * h;
                }
                Poly(pts);
                break;
            }
            case "chevron":
            {
                double d = Math.Min(w, h) * (sh.Adj ?? 0.5);
                Poly((0, 0), (w - d, 0), (w, h / 2), (w - d, h), (0, h), (d, h / 2));
                break;
            }
            case "homePlate" or "pentagon":
            {
                double d = Math.Min(w, h) * (sh.Adj ?? 0.5);
                Poly((0, 0), (w - d, 0), (w, h / 2), (w - d, h), (0, h));
                break;
            }
            case "regularPentagon":
                Poly((w / 2, 0), (w, h * 0.382), (w * 0.809, h), (w * 0.191, h), (0, h * 0.382));
                break;
            case "plus" or "mathPlus" or "cross":
            {
                double d = Math.Min(w, h) * (sh.Adj ?? 0.25);
                Poly((d, 0), (w - d, 0), (w - d, d), (w, d), (w, h - d), (w - d, h - d), (w - d, h), (d, h), (d, h - d), (0, h - d), (0, d), (d, d));
                break;
            }
            case "heart":
            {
                // Сердце: две кубические кривые от нижней вершины к верхней выемке
                string C(double x1, double y1, double x2, double y2, double x3, double y3) => $"{P(x1, y1)} {P(x2, y2)} {P(x3, y3)} c ";
                path.Append(P(w / 2, h * 0.25)).Append(" m ");
                path.Append(C(w * 0.5 + w * 0.12, -h * 0.08, w * 1.25, h * 0.15, w / 2, h));
                path.Append(C(-w * 0.25, h * 0.15, w * 0.5 - w * 0.12, -h * 0.08, w / 2, h * 0.25));
                path.Append("h ");
                break;
            }
            case "triangle": Poly((w / 2, 0), (w, h), (0, h)); break;
            case "rtTriangle": Poly((0, 0), (w, h), (0, h)); break;
            case "diamond": Poly((w / 2, 0), (w, h / 2), (w / 2, h), (0, h / 2)); break;
            case "parallelogram": Poly((w * 0.25, 0), (w, 0), (w * 0.75, h), (0, h)); break;
            case "hexagon": Poly((w * 0.25, 0), (w * 0.75, 0), (w, h / 2), (w * 0.75, h), (w * 0.25, h), (0, h / 2)); break;
            case "octagon":
            {
                double d = Math.Min(w, h) * 0.29;
                Poly((d, 0), (w - d, 0), (w, d), (w, h - d), (w - d, h), (d, h), (0, h - d), (0, d));
                break;
            }
            case "rightArrow": { double hw = Math.Min(w, h); Poly((0, h * 0.25), (w - hw / 2, h * 0.25), (w - hw / 2, 0), (w, h / 2), (w - hw / 2, h), (w - hw / 2, h * 0.75), (0, h * 0.75)); break; }
            case "leftArrow": { double hw = Math.Min(w, h); Poly((w, h * 0.25), (hw / 2, h * 0.25), (hw / 2, 0), (0, h / 2), (hw / 2, h), (hw / 2, h * 0.75), (w, h * 0.75)); break; }
            case "upArrow": { double hh = Math.Min(w, h); Poly((w * 0.25, h), (w * 0.25, hh / 2), (0, hh / 2), (w / 2, 0), (w, hh / 2), (w * 0.75, hh / 2), (w * 0.75, h)); break; }
            case "downArrow": { double hh = Math.Min(w, h); Poly((w * 0.25, 0), (w * 0.75, 0), (w * 0.75, h - hh / 2), (w, h - hh / 2), (w / 2, h), (0, h - hh / 2), (w * 0.25, h - hh / 2)); break; }
            case "custom" when sh.Paths != null:
            {
                // Произвольный контур: каждый подконтур со своими признаками заливки/обводки
                bool anyFill = sh.Fill != null, anyStroke = sh.Stroke != null && sh.StrokeWidth > 0;
                if (anyFill) _sb.Append($"{Color(sh.Fill!.Value)} rg ");
                if (anyStroke) _sb.Append($"{F(sh.StrokeWidth)} w {Color(sh.Stroke!.Value)} RG ");
                foreach (var cp in sh.Paths)
                {
                    var sb = new StringBuilder();
                    foreach (var (cmd, pts) in cp.Commands)
                    {
                        switch (cmd)
                        {
                            case 'M': sb.Append(P(pts[0] * w, pts[1] * h)).Append(" m "); break;
                            case 'L': sb.Append(P(pts[0] * w, pts[1] * h)).Append(" l "); break;
                            case 'C': sb.Append(P(pts[0] * w, pts[1] * h)).Append(' ').Append(P(pts[2] * w, pts[3] * h)).Append(' ').Append(P(pts[4] * w, pts[5] * h)).Append(" c "); break;
                            case 'Z': sb.Append("h "); break;
                        }
                    }
                    bool f = anyFill && cp.Fill, st = anyStroke && cp.Stroke;
                    if (f || st) _sb.Append(sb).Append(f && st ? "B " : f ? "f " : "S ");
                }
                _sb.Append("Q\n");
                return;
            }
            default: Poly((0, 0), (w, 0), (w, h), (0, h)); break;
        }
        bool fill = sh.Fill != null && !open, stroke = sh.Stroke != null && sh.StrokeWidth > 0;
        if (fill && sh.Gradient is { Count: >= 2 } grad)
        {
            // Градиент: обрезка по контуру и полосы с интерполированным цветом поперёк направления
            _sb.Append("q ").Append(path).Append("W n ");
            double ang = -sh.GradientAngle * Math.PI / 180, ca = Math.Cos(ang), sa = Math.Sin(ang);
            double half = Math.Abs(w / 2 * ca) + Math.Abs(h / 2 * sa), across = Math.Sqrt(w * w + h * h);
            _sb.Append($"{F(ca)} {F(sa)} {F(-sa)} {F(ca)} 0 0 cm ");
            const int bands = 64;
            for (int i = 0; i < bands; i++)
            {
                double t0 = (double)i / bands, t1 = (double)(i + 1) / bands, tm = (t0 + t1) / 2;
                int j = 0;
                while (j < grad.Count - 2 && grad[j + 1].Pos < tm) j++;
                var (p0, c0) = grad[j]; var (p1, c1) = grad[j + 1];
                double u = p1 > p0 ? Math.Clamp((tm - p0) / (p1 - p0), 0, 1) : 0;
                var col = new Rgb(c0.R + (c1.R - c0.R) * u, c0.G + (c1.G - c0.G) * u, c0.B + (c1.B - c0.B) * u);
                double x0 = -half + 2 * half * t0, x1 = -half + 2 * half * t1;
                _sb.Append($"{Color(col)} rg {F(x0)} {F(-across)} {F(x1 - x0 + 0.05)} {F(2 * across)} re f ");
            }
            _sb.Append("Q ");
            fill = false;
        }
        if (fill) _sb.Append($"{Color(sh.Fill!.Value)} rg ");
        if (stroke)
        {
            _sb.Append($"{F(sh.StrokeWidth)} w {Color(sh.Stroke!.Value)} RG ");
            if (sh.Dashed) _sb.Append($"[{F(sh.StrokeWidth * 4)} {F(sh.StrokeWidth * 3)}] 0 d ");
            else if (sh.Dotted) _sb.Append($"[{F(sh.StrokeWidth)} {F(sh.StrokeWidth * 2)}] 0 d ");
        }
        if (fill || stroke) _sb.Append(path).Append(fill && stroke ? "B " : fill ? "f " : "S ");
        // Стрелки на концах линии: закрашенный треугольник
        if (open && stroke && (sh.HeadArrow || sh.TailArrow))
        {
            _sb.Append($"[] 0 d {Color(sh.Stroke!.Value)} rg ");
            void Arrow(double tx, double ty, double fx, double fy)
            {
                double dx = tx - fx, dy = ty - fy, len = Math.Sqrt(dx * dx + dy * dy);
                if (len < 0.01) return;
                dx /= len; dy /= len;
                double s = Math.Max(4, sh.StrokeWidth * 3), bx = tx - dx * s, by = ty - dy * s;
                _sb.Append(P(tx, ty)).Append(" m ").Append(P(bx - dy * s / 2, by + dx * s / 2)).Append(" l ")
                   .Append(P(bx + dy * s / 2, by - dx * s / 2)).Append(" l h f ");
            }
            if (sh.TailArrow) Arrow(w, h, 0, 0);
            if (sh.HeadArrow) Arrow(0, 0, w, h);
        }
        _sb.Append("Q\n");
    }

    public void Image(string key, byte[] bytes, double x, double y, double w, double h, double opacity = 1, bool gray = false)
    {
        if (Collecting > 0) return;
        if (SkipGraphics) return;
        var img = _doc.GetImage(key, bytes, gray);
        if (img == null)
        {
            // Неподдерживаемый формат (EMF/WMF/TIFF…) — рисуем заглушку
            FillRect(x, y, w, h, new Rgb(0.93, 0.93, 0.93));
            Rect(x, y, w, h, 0.5, new Rgb(0.6, 0.6, 0.6));
            Line(x, y, x + w, y + h, 0.5, new Rgb(0.7, 0.7, 0.7));
            Line(x + w, y, x, y + h, 0.5, new Rgb(0.7, 0.7, 0.7));
            return;
        }
        if (!Images.Contains(img)) Images.Add(img);
        string gs = opacity < 0.999 ? $"/{GState(opacity)} gs " : "";
        _sb.Append($"q {gs}{F(w)} 0 0 {F(h)} {F(x)} {F(Y(y + h))} cm /{img.Name} Do Q\n");
    }

    public void Link(double x, double y, double w, double h, string uri)
    {
        if (Collecting > 0) return;
        if (SuppressLinks > 0) return;
        var last = Links.LastOrDefault();
        // Склеиваем соседние прямоугольники одной ссылки на одной строке
        if (last != null && last.Uri == uri && Math.Abs(last.Y2 - Y(y)) < 0.5 && Math.Abs(last.X2 - x) < 1)
        {
            last.X2 = x + w;
            last.Y1 = Math.Min(last.Y1, Y(y + h));
            return;
        }
        Links.Add(new PdfLink(x, Y(y + h), x + w, Y(y), uri));
    }
}

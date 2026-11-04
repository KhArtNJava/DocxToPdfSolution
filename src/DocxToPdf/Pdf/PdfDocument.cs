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

    internal PdfImage? GetImage(string key, byte[] bytes)
    {
        if (_images.TryGetValue(key, out var img)) return img;
        var data = ImageDecoder.Decode(bytes);
        img = data == null ? null : new PdfImage(data, "Im" + (_images.Count + 1));
        _images[key] = img;
        return img;
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

        var pageIds = new List<int>();
        foreach (var page in _pages)
        {
            int pageId = w.Reserve(), contentId = w.Reserve();
            pageIds.Add(pageId);
            var c = page.Canvas;

            var res = new StringBuilder("<< /ProcSet [/PDF /Text /ImageB /ImageC]");
            if (c.Fonts.Count > 0)
            {
                res.Append(" /Font <<");
                foreach (var f in c.Fonts) res.Append($" /{f.Name} {fontIds[f]} 0 R");
                res.Append(" >>");
            }
            if (c.Images.Count > 0)
            {
                res.Append(" /XObject <<");
                foreach (var i in c.Images) res.Append($" /{i.Name} {imageIds[i]} 0 R");
                res.Append(" >>");
            }
            res.Append(" >>");

            var annots = new List<int>();
            foreach (var link in c.Links)
            {
                int id = w.Reserve();
                annots.Add(id);
                w.Obj(id, $"<< /Type /Annot /Subtype /Link /Rect [{F(link.X1)} {F(link.Y1)} {F(link.X2)} {F(link.Y2)}] /Border [0 0 0] /A << /S /URI /URI {PdfString(link.Uri)} >> >>");
            }
            string annotStr = annots.Count > 0 ? " /Annots [" + string.Join(" ", annots.Select(a => a + " 0 R")) + "]" : "";
            w.Obj(pageId, $"<< /Type /Page /Parent {pagesId} 0 R /MediaBox [0 0 {F(page.Width)} {F(page.Height)}] /Resources {res} /Contents {contentId} 0 R{annotStr} >>");
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
    readonly StringBuilder _sb = new();
    internal readonly List<PdfFont> Fonts = [];
    internal readonly List<PdfImage> Images = [];
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
        if (SkipGraphics) return;
        if (w <= 0 || h <= 0) return;
        _sb.Append($"{Color(color)} rg {F(x)} {F(Y(y + h))} {F(w)} {F(h)} re f\n");
    }

    public void Line(double x1, double y1, double x2, double y2, double width, Rgb color, bool dashed = false, bool dotted = false)
    {
        if (SkipGraphics) return;
        _sb.Append("q ");
        if (dashed) _sb.Append($"[{F(width * 3)} {F(width * 2)}] 0 d ");
        else if (dotted) _sb.Append($"[{F(width)} {F(width * 1.5)}] 0 d ");
        _sb.Append($"{F(width)} w {Color(color)} RG {F(x1)} {F(Y(y1))} m {F(x2)} {F(Y(y2))} l S Q\n");
    }

    public void Rect(double x, double y, double w, double h, double width, Rgb color)
    {
        if (SkipGraphics) return;
        _sb.Append($"q {F(width)} w {Color(color)} RG {F(x)} {F(Y(y + h))} {F(w)} {F(h)} re S Q\n");
    }

    /// <summary>Выводит глифы по базовой линии baseline.</summary>
    /// <summary>Больше нуля — текст выводится контурами (без текстового слоя).</summary>
    public int TextAsPaths;
    /// <summary>Больше нуля — выводится только невидимый текстовый слой (для поиска и копирования).</summary>
    public int InvisibleTextOnly;

    bool SkipGraphics => InvisibleTextOnly > 0;

    public void Glyphs(FontFace face, double size, double x, double baseline, IReadOnlyList<ushort> gids, IReadOnlyList<string> texts, Rgb color, IReadOnlyList<double>? extra = null)
    {
        if (gids.Count == 0) return;
        if (TextAsPaths > 0 && InvisibleTextOnly == 0)
        {
            // Контуры глифов: каждый — в своей системе координат через cm
            var font = face.Font;
            double k = size / font.UnitsPerEm, skewK = face.SynthItalic ? 0.2126 : 0;
            double gx = x;
            _sb.Append($"q {Color(color)} rg ");
            if (face.SynthBold) _sb.Append($"{Color(color)} RG {F(size * 0.03 / k)} w ");
            for (int i = 0; i < gids.Count; i++)
            {
                string path = font.GlyphPath(gids[i]);
                if (path.Length > 0)
                    _sb.Append($"q {F(k)} 0 {F(k * skewK)} {F(k)} {F(gx)} {F(Y(baseline))} cm ").Append(path).Append(face.SynthBold ? "B Q " : "f Q ");
                gx += font.Advance(gids[i]) * size + (extra != null && i < extra.Count ? extra[i] : 0);
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
        else if (face.SynthBold) _sb.Append($"2 Tr {F(size * 0.03)} w {Color(color)} RG ");
        _sb.Append($"/{pf.Name} {F(size)} Tf ");
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
                    _sb.Append("> ").Append(F(-extra[i] * 1000 / size)).Append(" <");
            }
            _sb.Append(">] TJ");
        }
        _sb.Append(" ET\n");
        if (face.SynthBold || InvisibleTextOnly > 0) _sb.Append("0 Tr\n");
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
        if (SkipGraphics) return;
        if (w <= 0 || h <= 0) return;
        _sb.Append($"{Color(color)} rg ");
        RoundedPath(x, y, w, h, radii);
        _sb.Append("f\n");
    }

    /// <summary>Обводка скруглённого прямоугольника; линия идёт по середине толщины (внутри рамки).</summary>
    public void StrokeRoundedRect(double x, double y, double w, double h, double[] radii, BorderDef b)
    {
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
        _sb.Append("q ");
        RoundedPath(x, y, w, h, radii);
        _sb.Append("W n\n");
    }

    public void EndClip() => _sb.Append("Q\n");

    /// <summary>
    /// Начинает преобразованную область: матрица задана в координатах "y вниз" относительно точки (ox, oy) страницы.
    /// </summary>
    public void BeginTransform(double ox, double oy, double a, double b, double c, double d, double e, double f)
    {
        double yo = Y(oy);
        // В системе PDF (y вверх) матрица CSS превращается в [a, -b, -c, d, e, -f]
        _sb.Append($"q 1 0 0 1 {F(ox)} {F(yo)} cm {F(a)} {F(-b)} {F(-c)} {F(d)} {F(e)} {F(-f)} cm 1 0 0 1 {F(-ox)} {F(-yo)} cm\n");
    }

    public void EndTransform() => _sb.Append("Q\n");

    /// <summary>Обрезка по многоугольнику (координаты страницы, y вниз) до EndClip().</summary>
    public void BeginClipPolygon(IReadOnlyList<(double X, double Y)> pts)
    {
        _sb.Append("q ");
        for (int i = 0; i < pts.Count; i++)
            _sb.Append($"{F(pts[i].X)} {F(Y(pts[i].Y))} {(i == 0 ? "m" : "l")} ");
        _sb.Append("h W n\n");
    }

    /// <summary>Пока больше нуля — ссылки не добавляются (повторная отрисовка того же содержимого).</summary>
    public int SuppressLinks;

    public void Image(string key, byte[] bytes, double x, double y, double w, double h)
    {
        if (SkipGraphics) return;
        var img = _doc.GetImage(key, bytes);
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
        _sb.Append($"q {F(w)} 0 0 {F(h)} {F(x)} {F(Y(y + h))} cm /{img.Name} Do Q\n");
    }

    public void Link(double x, double y, double w, double h, string uri)
    {
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

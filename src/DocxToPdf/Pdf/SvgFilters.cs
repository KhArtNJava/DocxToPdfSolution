using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocxToPdf.Docx;
using DocxToPdf.Images;

namespace DocxToPdf.Pdf;

/// <summary>
/// SVG: маски (mask) и фильтры (filter).
/// Маска — векторная: содержимое маски становится мягкой маской PDF (SMask по яркости или по альфа-каналу).
/// Фильтр работает с пикселями: элемент рисуется в растр (Raster), к нему применяются примитивы фильтра
/// (feGaussianBlur, feOffset, feFlood, feColorMatrix, feComponentTransfer, feComposite, feMerge, feBlend,
/// feDropShadow, feMorphology, feTile, feImage, feTurbulence, feConvolveMatrix, feDisplacementMap,
/// feDiffuseLighting, feSpecularLighting) и CSS-функции (blur, drop-shadow, grayscale, sepia, saturate,
/// hue-rotate, invert, opacity, brightness, contrast); результат выводится картинкой.
/// </summary>
public static partial class Svg
{
    // ================================================================
    //  Стиль и рамка элемента
    // ================================================================

    /// <summary>Вычисленный стиль элемента по цепочке предков (для содержимого mask/filter, которое наследует от своих предков).</summary>
    static Style StyleOf(SvgDocument doc, XElement? e) => e == null ? new Style() : Compute(doc, e, StyleOf(doc, e.Parent));

    /// <summary>
    /// Рамка элемента в его системе координат с учётом преобразований потомков; stroke — с толщиной обводки
    /// (видимая область), иначе геометрическая (objectBoundingBox).
    /// </summary>
    static (double X, double Y, double W, double H) EffectBBox(XElement e, Style st, Ctx ctx, bool stroke)
    {
        double x0 = double.MaxValue, y0 = double.MaxValue, x1 = double.MinValue, y1 = double.MinValue;
        void Add(double x, double y) { x0 = Math.Min(x0, x); y0 = Math.Min(y0, y); x1 = Math.Max(x1, x); y1 = Math.Max(y1, y); }
        void AddRect(M m, double x, double y, double w, double h, double grow)
        {
            foreach (var (px, py) in new[] { (x - grow, y - grow), (x + w + grow, y - grow), (x - grow, y + h + grow), (x + w + grow, y + h + grow) })
            {
                var (tx, ty) = m.Apply(px, py);
                Add(tx, ty);
            }
        }
        void Walk(XElement el, Style s, M m, int depth)
        {
            if (depth > 30) return;
            string tag = el.Name.LocalName;
            if (NonRendering.Contains(tag) || s.Get("display") == "none") return;
            switch (tag)
            {
                case "g" or "a" or "switch":
                    foreach (var ch in el.Elements()) Child(ch, s, m, depth);
                    return;
                case "svg":
                {
                    double x = LenX(el, "x", ctx, s), y = LenY(el, "y", ctx, s);
                    double w = Attr(el, "width") is { } wa && wa != "auto" ? LenX(el, "width", ctx, s) : ctx.VW;
                    double h = Attr(el, "height") is { } ha && ha != "auto" ? LenY(el, "height", ctx, s) : ctx.VH;
                    AddRect(m, x, y, w, h, 0);
                    return;
                }
                case "use":
                {
                    var target = Ref(ctx.Doc, Href(el));
                    if (target == null || target.AncestorsAndSelf().Contains(el)) return;
                    var t = m.Mul(new M(1, 0, 0, 1, LenX(el, "x", ctx, s), LenY(el, "y", ctx, s)));
                    if (target.Name.LocalName == "symbol")
                    {
                        var ss = Compute(ctx.Doc, target, s);
                        foreach (var ch in target.Elements()) Child(ch, ss, t, depth);
                    }
                    else Child(target, s, t, depth);
                    return;
                }
                case "image":
                    AddRect(m, LenX(el, "x", ctx, s), LenY(el, "y", ctx, s), LenX(el, "width", ctx, s), LenY(el, "height", ctx, s), 0);
                    return;
                case "text":
                {
                    double fs = FontSize(s);
                    double x = Numbers(Attr(el, "x")).FirstOrDefault(), y = Numbers(Attr(el, "y")).FirstOrDefault();
                    double w = TextWidth(Regex.Replace(el.Value.Trim(), @"\s+", " "), s, ctx);
                    string anchor = s.Get("text-anchor", "start");
                    if (anchor == "middle") x -= w / 2; else if (anchor == "end") x -= w;
                    AddRect(m, x, y - fs * 0.85, w, fs * 1.1, 0);
                    return;
                }
            }
            if (ShapeGeo(el, s, ctx) is { } g)
            {
                double grow = 0;
                if (stroke && s.Get("stroke", "none") != "none")
                {
                    grow = LenD(s.Get("stroke-width", "1"), ctx, s) / 2;
                    if (s.Get("stroke-linejoin") is not ("round" or "bevel")) grow *= Math.Min(4, Math.Max(1, Num(s.Get("stroke-miterlimit", "4"), 4)));
                }
                foreach (var (cmd, p) in g.Cmds)
                    for (int i = 0; i + 1 < p.Length; i += 2)
                        foreach (var (dx, dy) in grow > 0 ? new[] { (-grow, -grow), (grow, -grow), (-grow, grow), (grow, grow) } : new[] { (0.0, 0.0) })
                        {
                            var (tx, ty) = m.Apply(p[i] + dx, p[i + 1] + dy);
                            Add(tx, ty);
                        }
            }
        }
        void Child(XElement ch, Style parent, M m, int depth)
        {
            var cs = Compute(ctx.Doc, ch, parent);
            var tm = ParseTransform(Attr(ch, "transform") ?? (SpecifiedOf(ctx.Doc, ch).TryGetValue("transform", out var ct) ? ct : null));
            Walk(ch, cs, m.Mul(tm), depth + 1);
        }
        Walk(e, st, M.I, 0);
        return x0 > x1 ? (0, 0, 0, 0) : (x0, y0, x1 - x0, y1 - y0);
    }

    /// <summary>Ширина строки текста шрифтом стиля (для рамки текста).</summary>
    static double TextWidth(string text, Style s, Ctx ctx)
    {
        double fs = FontSize(s);
        if (ctx.Fonts == null) return text.Length * fs * 0.55;
        string fw = s.Get("font-weight", "normal");
        bool bold = fw is "bold" or "bolder" || int.TryParse(fw, out var wn) && wn >= 600;
        var face = ctx.Fonts.Get(FirstFamily(s.Get("font-family", "serif")), bold, s.Get("font-style") is "italic" or "oblique");
        double w = 0;
        foreach (var r in text.EnumerateRunes())
            w += face.Font.TryGetGlyph(r.Value, out var gid) && gid != 0 ? face.Font.Advance(gid) * fs : fs * 0.55;
        return w;
    }

    /// <summary>Длина/координата в единицах objectBoundingBox или пользователя.</summary>
    static double UnitLen(string? v, string def, bool obb, double bboxOrigin, double bboxSize, double viewport, bool coord)
    {
        string s = v ?? def;
        if (obb) return (coord ? bboxOrigin : 0) + Num(s) * bboxSize;
        return Len(s, viewport, 16);
    }

    // ================================================================
    //  Маска
    // ================================================================

    /// <summary>Форма-маска элемента (в его системе координат) и признак альфа-маски (mask-type: alpha).</summary>
    static (string Form, bool Alpha) MaskForm(XElement mask, XElement e, Style st, Ctx ctx)
    {
        var c = ctx.C;
        var bbox = EffectBBox(e, st, ctx, stroke: false);
        bool obb = Attr(mask, "maskUnits") != "userSpaceOnUse";
        double x = UnitLen(Attr(mask, "x"), "-10%", obb, bbox.X, bbox.W, ctx.VW, true);
        double y = UnitLen(Attr(mask, "y"), "-10%", obb, bbox.Y, bbox.H, ctx.VH, true);
        double w = UnitLen(Attr(mask, "width"), "120%", obb, bbox.X, bbox.W, ctx.VW, false);
        double h = UnitLen(Attr(mask, "height"), "120%", obb, bbox.Y, bbox.H, ctx.VH, false);
        var mst = StyleOf(ctx.Doc, mask);
        bool alpha = (Attr(mask, "mask-type") ?? mst.Get("mask-type")) == "alpha";
        bool contentObb = Attr(mask, "maskContentUnits") == "objectBoundingBox";
        string form = c.CaptureForm(() =>
        {
            if (w <= 0 || h <= 0 || contentObb && (bbox.W <= 0 || bbox.H <= 0)) return;
            c.TextAsPaths++;
            double savedAlpha = ctx.Alpha;
            ctx.Alpha = 1;
            var r = c.NewPath(); r.Rect(x, y, w, h);
            c.BeginClip(r, false);
            var cm = contentObb ? new M(bbox.W, 0, 0, bbox.H, bbox.X, bbox.Y) : M.I;
            if (!cm.IsIdentity) c.BeginTransform(0, 0, cm.A, cm.B, cm.C, cm.D, cm.E, cm.F);
            ctx.Depth += 5;   // маска внутри маски — ограничение глубины
            try { RenderChildren(mask, mst, ctx); }
            finally
            {
                ctx.Depth -= 5;
                if (!cm.IsIdentity) c.EndTransform();
                c.EndClip();
                ctx.Alpha = savedAlpha;
                c.TextAsPaths--;
            }
        }, unbounded: true);
        return (form, alpha);
    }

    // ================================================================
    //  Фильтр: подготовка, растр, вывод
    // ================================================================

    /// <summary>Растровое изображение фильтра: RGBA с предумноженной альфой, подобласть [X0, X1) × [Y0, Y1).</summary>
    sealed class FImg(float[] p, bool lin, int x0, int y0, int x1, int y1)
    {
        public float[] P = p;
        public bool Lin = lin;
        public int X0 = x0, Y0 = y0, X1 = x1, Y1 = y1;
    }

    /// <summary>Параметры сетки фильтра: размер, область в единицах пользователя, пикселей на единицу.</summary>
    sealed class FGrid
    {
        public int W, H;
        public double Rx, Ry, Rw, Rh, Sx, Sy;
        public (double X, double Y, double W, double H) BBox;
        public double Px(double ux) => (ux - Rx) * Sx;
        public double Py(double uy) => (uy - Ry) * Sy;
    }

    abstract record FItem;
    sealed record FRef(XElement Filter) : FItem;
    sealed record FFunc(string Name, string Args) : FItem;

    static List<FItem> ParseFilterList(string v, Ctx ctx)
    {
        var res = new List<FItem>();
        foreach (Match m in Regex.Matches(v, @"([a-zA-Z-]+)\(((?:[^()]|\([^()]*\))*)\)"))
        {
            string name = m.Groups[1].Value.ToLowerInvariant();
            if (name == "url")
            {
                if (Ref(ctx.Doc, "url(" + m.Groups[2].Value + ")") is { } f && f.Name.LocalName == "filter") res.Add(new FRef(f));
                else return [];   // ссылка на несуществующий фильтр — фильтр не применяется
            }
            else res.Add(new FFunc(name, m.Groups[2].Value.Trim()));
        }
        return res;
    }

    static void FilterElement(XElement e, Style st, Ctx ctx, string filterValue)
    {
        var c = ctx.C;
        var items = ParseFilterList(filterValue, ctx);
        if (items.Count == 0 || ctx.Depth > 30) { Body(e, st, ctx); return; }
        var bbox = EffectBBox(e, st, ctx, stroke: false);
        var vis = EffectBBox(e, st, ctx, stroke: true);
        // Область сетки: объединение областей всех фильтров цепочки
        double gx0 = double.MaxValue, gy0 = double.MaxValue, gx1 = double.MinValue, gy1 = double.MinValue;
        var regions = new List<(double X, double Y, double W, double H)>();
        var cur = vis;
        foreach (var it in items)
        {
            (double X, double Y, double W, double H) r;
            if (it is FRef fr)
            {
                bool obb = FilterAttr(fr.Filter, "filterUnits", ctx) != "userSpaceOnUse";
                if (obb && (bbox.W <= 0 || bbox.H <= 0)) { regions.Add((0, 0, 0, 0)); continue; }
                r = (UnitLen(FilterAttr(fr.Filter, "x", ctx), "-10%", obb, bbox.X, bbox.W, ctx.VW, true),
                     UnitLen(FilterAttr(fr.Filter, "y", ctx), "-10%", obb, bbox.Y, bbox.H, ctx.VH, true),
                     UnitLen(FilterAttr(fr.Filter, "width", ctx), "120%", obb, bbox.X, bbox.W, ctx.VW, false),
                     UnitLen(FilterAttr(fr.Filter, "height", ctx), "120%", obb, bbox.Y, bbox.H, ctx.VH, false));
            }
            else
            {
                // CSS-функция: видимая рамка с запасом на размытие и смещение тени
                var f = (FFunc)it;
                double grow = 0, dx = 0, dy = 0;
                var toks = CssTokens(f.Args);
                if (f.Name == "blur") grow = 3 * (toks.Count > 0 ? Len(toks[0], 0, 16) : 0);
                else if (f.Name == "drop-shadow")
                {
                    var lens = toks.Where(t => ParseColor(t) == null).Select(t => Len(t, 0, 16)).ToList();
                    dx = lens.ElementAtOrDefault(0); dy = lens.ElementAtOrDefault(1); grow = 3 * lens.ElementAtOrDefault(2);
                }
                r = (Math.Min(cur.X, cur.X + dx) - grow, Math.Min(cur.Y, cur.Y + dy) - grow, cur.W + Math.Abs(dx) + 2 * grow, cur.H + Math.Abs(dy) + 2 * grow);
                cur = r;
            }
            regions.Add(r);
            if (r.W <= 0 || r.H <= 0) continue;
            gx0 = Math.Min(gx0, r.X); gy0 = Math.Min(gy0, r.Y); gx1 = Math.Max(gx1, r.X + r.W); gy1 = Math.Max(gy1, r.Y + r.H);
        }
        // Пустая область фильтра — элемент не выводится
        if (gx0 >= gx1 || gy0 >= gy1) return;
        double rw = gx1 - gx0, rh = gy1 - gy0;
        // Разрешение: ~150 точек на дюйм итогового размера, не больше 4 млн пикселей
        double k = Math.Max(0.05, ctx.Scale) * 150 / 72;
        const double MaxPixels = 4e6;
        if (rw * k * rh * k > MaxPixels) k = Math.Sqrt(MaxPixels / (rw * rh));
        int W = Math.Clamp((int)Math.Ceiling(rw * k), 1, 6000), H = Math.Clamp((int)Math.Ceiling(rh * k), 1, 6000);
        var grid = new FGrid { W = W, H = H, Rx = gx0, Ry = gy0, Rw = rw, Rh = rh, Sx = W / rw, Sy = H / rh, BBox = bbox };

        // Исходный рисунок элемента (без прозрачности opacity — она применяется к результату)
        double savedAlpha = ctx.Alpha;
        ctx.Alpha = 1;
        string ops;
        try { ops = c.CaptureOps(() => Body(e, st, ctx)); }
        finally { ctx.Alpha = savedAlpha; }
        var ras = new Raster(W, H, c);
        ras.Run(ops, BaseMat(grid, c));
        if (ras.Unsupported && ops.Contains("BT"))
        {
            // Нарисовать в растр не удалось — без фильтра
            Body(e, st, ctx);
            return;
        }
        var img = new FImg(ras.Px, false, 0, 0, W, H);
        for (int i = 0; i < items.Count; i++)
        {
            var r = regions[i];
            if (r.W <= 0 || r.H <= 0) return;
            var region = PixRect(grid, r.X, r.Y, r.W, r.H);
            img = items[i] is FRef fr ? ApplyFilter(fr.Filter, img, region, grid, e, st, ctx) : ApplyCssFunction((FFunc)items[i], img, region, grid, st);
            img = ToSpace(img, false);
            Clip(img, region, W);
        }
        Emit(img, grid, savedAlpha, c);
    }

    static Raster.Mat BaseMat(FGrid g, PdfCanvas c) => new(g.Sx, 0, 0, -g.Sy, -g.Rx * g.Sx, (c.PageHeight - g.Ry) * g.Sy);

    static string? FilterAttr(XElement f, string name, Ctx ctx) => GradAttr(f, name, ctx);

    static (int X0, int Y0, int X1, int Y1) PixRect(FGrid g, double x, double y, double w, double h) =>
        (Math.Clamp((int)Math.Floor(g.Px(x) + 1e-6), 0, g.W), Math.Clamp((int)Math.Floor(g.Py(y) + 1e-6), 0, g.H),
         Math.Clamp((int)Math.Ceiling(g.Px(x + w) - 1e-6), 0, g.W), Math.Clamp((int)Math.Ceiling(g.Py(y + h) - 1e-6), 0, g.H));

    /// <summary>Всё вне прямоугольника r — прозрачно; подобласть сужается до r.</summary>
    static void Clip(FImg img, (int X0, int Y0, int X1, int Y1) r, int w)
    {
        int H = img.P.Length / 4 / w;
        for (int y = 0; y < H; y++)
            for (int x = 0; x < w; x++)
                if (x < r.X0 || x >= r.X1 || y < r.Y0 || y >= r.Y1)
                {
                    int o = (y * w + x) * 4;
                    img.P[o] = img.P[o + 1] = img.P[o + 2] = img.P[o + 3] = 0;
                }
        img.X0 = Math.Max(img.X0, r.X0); img.Y0 = Math.Max(img.Y0, r.Y0);
        img.X1 = Math.Min(img.X1, r.X1); img.Y1 = Math.Min(img.Y1, r.Y1);
    }

    /// <summary>Вывод результата: картинка RGB + альфа по непустой части сетки.</summary>
    static void Emit(FImg img, FGrid g, double opacity, PdfCanvas c)
    {
        int W = g.W, H = g.H;
        int bx0 = W, by0 = H, bx1 = -1, by1 = -1;
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
                if (img.P[(y * W + x) * 4 + 3] > 0.5f / 255)
                {
                    if (x < bx0) bx0 = x; if (x > bx1) bx1 = x;
                    if (y < by0) by0 = y; if (y > by1) by1 = y;
                }
        if (bx1 < 0) return;
        int w = bx1 - bx0 + 1, h = by1 - by0 + 1;
        var rgb = new byte[w * h * 3];
        var alpha = new byte[w * h];
        for (int y = 0; y < h; y++)
            for (int x = 0; x < w; x++)
            {
                int o = ((y + by0) * W + x + bx0) * 4, d = y * w + x;
                float a = Math.Clamp(img.P[o + 3], 0, 1);
                alpha[d] = (byte)Math.Round(a * opacity * 255);
                if (a > 0)
                    for (int ch = 0; ch < 3; ch++) rgb[d * 3 + ch] = (byte)Math.Round(Math.Clamp(img.P[o + ch] / a, 0, 1) * 255);
            }
        var data = new PdfImageData { Width = w, Height = h, Data = rgb, Alpha = alpha };
        c.RasterImage(data, g.Rx + bx0 / g.Sx, g.Ry + by0 / g.Sy, w / g.Sx, h / g.Sy);
    }

    // ================================================================
    //  Цветовые пространства (linearRGB / sRGB)
    // ================================================================

    static readonly float[] ToLinLut = Enumerable.Range(0, 4097).Select(i => { double v = i / 4096.0; return (float)(v <= 0.04045 ? v / 12.92 : Math.Pow((v + 0.055) / 1.055, 2.4)); }).ToArray();
    static readonly float[] ToSrgbLut = Enumerable.Range(0, 4097).Select(i => { double v = i / 4096.0; return (float)(v <= 0.0031308 ? v * 12.92 : 1.055 * Math.Pow(v, 1 / 2.4) - 0.055); }).ToArray();

    static float Lut(float[] lut, float v)
    {
        if (v <= 0) return 0;
        if (v >= 1) return 1;
        float f = v * 4096;
        int i = (int)f;
        return lut[i] + (lut[i + 1] - lut[i]) * (f - i);
    }

    static double SrgbToLin(double v) => Lut(ToLinLut, (float)v);

    static FImg ToSpace(FImg img, bool lin)
    {
        if (img.Lin == lin) return img;
        var lut = lin ? ToLinLut : ToSrgbLut;
        var p = new float[img.P.Length];
        for (int o = 0; o < p.Length; o += 4)
        {
            float a = img.P[o + 3];
            p[o + 3] = a;
            if (a <= 0) continue;
            for (int ch = 0; ch < 3; ch++) p[o + ch] = Lut(lut, Math.Min(1, img.P[o + ch] / a)) * a;
        }
        return new FImg(p, lin, img.X0, img.Y0, img.X1, img.Y1);
    }

    static FImg Copy(FImg i) => new((float[])i.P.Clone(), i.Lin, i.X0, i.Y0, i.X1, i.Y1);

    // ================================================================
    //  Фильтр-элемент: примитивы
    // ================================================================

    sealed class FState
    {
        public required FGrid G;
        public required Ctx Ctx;
        public required XElement Element;
        public required Style ElementStyle;
        public required FImg Source;
        public required (int X0, int Y0, int X1, int Y1) Region;
        public bool PrimObb;
        public readonly Dictionary<string, FImg> Results = new(StringComparer.Ordinal);
        public FImg? Last;
        public FImg? SourceAlphaCache;
    }

    static FImg ApplyFilter(XElement filter, FImg source, (int X0, int Y0, int X1, int Y1) region, FGrid g, XElement e, Style st, Ctx ctx)
    {
        // Примитивы могут наследоваться по href
        XElement? content = filter;
        for (int guard = 0; guard < 8 && content != null && !content.Elements().Any(x => x.Name.LocalName.StartsWith("fe")); guard++)
            content = Ref(ctx.Doc, Href(content));
        var src = Copy(source);
        Clip(src, region, g.W);
        if (content == null) return new FImg(new float[g.W * g.H * 4], false, region.X0, region.Y0, region.X1, region.Y1);
        var fst = new FState
        {
            G = g, Ctx = ctx, Element = e, ElementStyle = st, Source = src, Region = region,
            PrimObb = FilterAttr(filter, "primitiveUnits", ctx) == "objectBoundingBox",
        };
        var filterStyle = StyleOf(ctx.Doc, content);
        foreach (var prim in content.Elements())
        {
            string tag = prim.Name.LocalName;
            if (!tag.StartsWith("fe")) continue;
            var ps = Compute(ctx.Doc, prim, filterStyle);
            bool lin = ps.Get("color-interpolation-filters", "linearRGB") != "sRGB";
            FImg? res;
            try { res = Primitive(prim, tag, ps, lin, fst); }
            catch { res = null; }
            if (res == null) continue;
            // Подобласть примитива
            var sub = Subregion(prim, fst, res);
            Clip(res, sub, g.W);
            fst.Last = res;
            if (Attr(prim, "result") is { Length: > 0 } rn) fst.Results[rn] = res;
        }
        return fst.Last ?? src;
    }

    /// <summary>Подобласть примитива: x/y/width/height, по умолчанию — объединение подобластей входов (или область фильтра).</summary>
    static (int X0, int Y0, int X1, int Y1) Subregion(XElement prim, FState f, FImg res)
    {
        var g = f.G;
        var def = (res.X0, res.Y0, res.X1, res.Y1);
        if (Attr(prim, "x") == null && Attr(prim, "y") == null && Attr(prim, "width") == null && Attr(prim, "height") == null) return def;
        var b = g.BBox;
        double ux0 = g.Rx + def.X0 / g.Sx, uy0 = g.Ry + def.Y0 / g.Sy, uw = (def.X1 - def.X0) / g.Sx, uh = (def.Y1 - def.Y0) / g.Sy;
        double x = Attr(prim, "x") is { } xa ? UnitLen(xa, "0", f.PrimObb, b.X, b.W, f.Ctx.VW, true) : ux0;
        double y = Attr(prim, "y") is { } ya ? UnitLen(ya, "0", f.PrimObb, b.Y, b.H, f.Ctx.VH, true) : uy0;
        double w = Attr(prim, "width") is { } wa ? UnitLen(wa, "0", f.PrimObb, b.X, b.W, f.Ctx.VW, false) : uw;
        double h = Attr(prim, "height") is { } ha ? UnitLen(ha, "0", f.PrimObb, b.Y, b.H, f.Ctx.VH, false) : uh;
        var r = PixRect(g, x, y, w, h);
        return (Math.Max(r.X0, f.Region.X0), Math.Max(r.Y0, f.Region.Y0), Math.Min(r.X1, f.Region.X1), Math.Min(r.Y1, f.Region.Y1));
    }

    /// <summary>Вход примитива (in, in2): стандартные входы, результаты по имени, иначе — предыдущий результат.</summary>
    static FImg Input(XElement prim, string attr, FState f, bool lin)
    {
        string? name = Attr(prim, attr)?.Trim();
        FImg img;
        switch (name)
        {
            case "SourceGraphic": img = f.Source; break;
            case "SourceAlpha":
                img = f.SourceAlphaCache ??= AlphaOnly(f.Source);
                break;
            case "BackgroundImage" or "BackgroundAlpha" or "FillPaint" or "StrokePaint":
                img = new FImg(new float[f.Source.P.Length], lin, f.Region.X0, f.Region.Y0, f.Region.X1, f.Region.Y1);
                break;
            default:
                img = name != null && f.Results.TryGetValue(name, out var r) ? r : f.Last ?? f.Source;
                break;
        }
        return ToSpace(img, lin);
    }

    static FImg AlphaOnly(FImg s)
    {
        var p = new float[s.P.Length];
        for (int o = 3; o < p.Length; o += 4) p[o] = s.P[o];
        return new FImg(p, s.Lin, s.X0, s.Y0, s.X1, s.Y1);
    }

    /// <summary>Длина примитива в пикселях по x и y (единицы пользователя или objectBoundingBox).</summary>
    static (double X, double Y) PrimLen(FState f, double vx, double vy) =>
        f.PrimObb ? (vx * f.G.BBox.W * f.G.Sx, vy * f.G.BBox.H * f.G.Sy) : (vx * f.G.Sx, vy * f.G.Sy);

    static (Rgb Color, double Alpha) FloodColor(Style ps, string colorProp, string opacityProp, bool lin)
    {
        var col = ParseColor(ps.Get(colorProp, colorProp == "lighting-color" ? "white" : "black"), ps.Get("color", "black")) ?? (Rgb.Black, 0);
        double a = col.Alpha * Math.Clamp(Num(ps.Get(opacityProp, "1"), 1), 0, 1);
        var c = col.Color;
        if (lin) c = new Rgb(SrgbToLin(c.R), SrgbToLin(c.G), SrgbToLin(c.B));
        return (c, a);
    }

    static FImg Blank(FState f, bool lin, (int X0, int Y0, int X1, int Y1)? sub = null)
    {
        var r = sub ?? f.Region;
        return new FImg(new float[f.G.W * f.G.H * 4], lin, r.X0, r.Y0, r.X1, r.Y1);
    }

    static (int X0, int Y0, int X1, int Y1) Union(FImg a, FImg b) => (Math.Min(a.X0, b.X0), Math.Min(a.Y0, b.Y0), Math.Max(a.X1, b.X1), Math.Max(a.Y1, b.Y1));

    static FImg? Primitive(XElement prim, string tag, Style ps, bool lin, FState f)
    {
        var g = f.G;
        int W = g.W, H = g.H;
        switch (tag)
        {
            case "feGaussianBlur":
            {
                var input = Input(prim, "in", f, lin);
                var sd = Numbers(Attr(prim, "stdDeviation") ?? "0");
                double sx = sd.Count > 0 ? sd[0] : 0, sy = sd.Count > 1 ? sd[1] : sx;
                if (sx < 0 || sy < 0) return Blank(f, lin);
                var (px, py) = PrimLen(f, sx, sy);
                var o = Copy(input);
                Blur(o.P, W, H, px, py);
                o.X0 = f.Region.X0; o.Y0 = f.Region.Y0; o.X1 = f.Region.X1; o.Y1 = f.Region.Y1;
                return o;
            }
            case "feOffset":
            {
                var input = Input(prim, "in", f, lin);
                var (dx, dy) = PrimLen(f, Num(Attr(prim, "dx")), Num(Attr(prim, "dy")));
                return Shift(input, W, H, dx, dy);
            }
            case "feFlood":
            {
                var (c, a) = FloodColor(ps, "flood-color", "flood-opacity", lin);
                var o = Blank(f, lin);
                var sub = Subregion(prim, f, o);
                Fill(o.P, W, sub, c, a);
                return o;
            }
            case "feColorMatrix":
            {
                var input = Input(prim, "in", f, lin);
                var o = Copy(input);
                string type = Attr(prim, "type") ?? "matrix";
                var vals = Numbers(Attr(prim, "values"));
                double[]? mtx = type switch
                {
                    "saturate" => SaturateMatrix(vals.Count > 0 ? vals[0] : 1),
                    "hueRotate" => HueMatrix(vals.Count > 0 ? vals[0] : 0),
                    "luminanceToAlpha" => [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.2125, 0.7154, 0.0721, 0, 0],
                    _ => vals.Count == 20 ? [.. vals] : null,
                };
                if (mtx != null) ColorMatrix(o.P, mtx);
                return o;
            }
            case "feComponentTransfer":
            {
                var input = Input(prim, "in", f, lin);
                var o = Copy(input);
                var funcs = new Func<double, double>?[4];
                foreach (var fn in prim.Elements())
                {
                    int ch = fn.Name.LocalName switch { "feFuncR" => 0, "feFuncG" => 1, "feFuncB" => 2, "feFuncA" => 3, _ => -1 };
                    if (ch >= 0) funcs[ch] = TransferFunc(fn);
                }
                Transfer(o.P, funcs);
                return o;
            }
            case "feComposite":
            {
                var a = Input(prim, "in", f, lin);
                var b = Input(prim, "in2", f, lin);
                string op = Attr(prim, "operator") ?? "over";
                var o = new FImg(new float[a.P.Length], lin, 0, 0, 0, 0);
                (o.X0, o.Y0, o.X1, o.Y1) = Union(a, b);
                double k1 = Num(Attr(prim, "k1")), k2 = Num(Attr(prim, "k2")), k3 = Num(Attr(prim, "k3")), k4 = Num(Attr(prim, "k4"));
                for (int i = 0; i < o.P.Length; i += 4)
                {
                    float aa = a.P[i + 3], ab = b.P[i + 3];
                    float fa, fb;
                    switch (op)
                    {
                        case "in": fa = ab; fb = 0; break;
                        case "out": fa = 1 - ab; fb = 0; break;
                        case "atop": fa = ab; fb = 1 - aa; break;
                        case "xor": fa = 1 - ab; fb = 1 - aa; break;
                        case "lighter": fa = 1; fb = 1; break;
                        case "arithmetic":
                        {
                            float ra = (float)Math.Clamp(k1 * aa * ab + k2 * aa + k3 * ab + k4, 0, 1);
                            for (int ch = 0; ch < 3; ch++)
                                o.P[i + ch] = (float)Math.Clamp(k1 * a.P[i + ch] * b.P[i + ch] + k2 * a.P[i + ch] + k3 * b.P[i + ch] + k4 * ra, 0, ra);
                            o.P[i + 3] = ra;
                            continue;
                        }
                        default: fa = 1; fb = 1 - aa; break;
                    }
                    for (int ch = 0; ch < 4; ch++) o.P[i + ch] = Math.Min(1, a.P[i + ch] * fa + b.P[i + ch] * fb);
                }
                return o;
            }
            case "feMerge":
            {
                var o = Blank(f, lin);
                bool any = false;
                foreach (var node in prim.Elements().Where(x => x.Name.LocalName == "feMergeNode"))
                {
                    var src = Input(node, "in", f, lin);
                    if (!any) { (o.X0, o.Y0, o.X1, o.Y1) = (src.X0, src.Y0, src.X1, src.Y1); any = true; }
                    else (o.X0, o.Y0, o.X1, o.Y1) = Union(o, src);
                    Over(o.P, src.P);
                }
                return o;
            }
            case "feBlend":
            {
                var s = Input(prim, "in", f, lin);
                var b = Input(prim, "in2", f, lin);
                var o = new FImg(new float[s.P.Length], lin, 0, 0, 0, 0);
                (o.X0, o.Y0, o.X1, o.Y1) = Union(s, b);
                Blend(s.P, b.P, o.P, Attr(prim, "mode") ?? "normal");
                return o;
            }
            case "feDropShadow":
            {
                var input = Input(prim, "in", f, lin);
                var sd = Numbers(Attr(prim, "stdDeviation") ?? "2");
                var (bx, by) = PrimLen(f, sd.Count > 0 ? sd[0] : 2, sd.Count > 1 ? sd[1] : sd.Count > 0 ? sd[0] : 2);
                var (dx, dy) = PrimLen(f, Num(Attr(prim, "dx"), 2), Num(Attr(prim, "dy"), 2));
                var (c, a) = FloodColor(ps, "flood-color", "flood-opacity", lin);
                return DropShadow(input, W, H, bx, by, dx, dy, c, a, f.Region);
            }
            case "feMorphology":
            {
                var input = Input(prim, "in", f, lin);
                var rad = Numbers(Attr(prim, "radius") ?? "0");
                double r1 = rad.Count > 0 ? rad[0] : 0, r2 = rad.Count > 1 ? rad[1] : r1;
                if (r1 < 0 || r2 < 0) return Blank(f, lin);
                var (rx, ry) = PrimLen(f, r1, r2);
                var o = Copy(input);
                if (rx > 0 || ry > 0) Morph(o.P, W, H, (int)Math.Round(rx), (int)Math.Round(ry), Attr(prim, "operator") == "dilate");
                if (Attr(prim, "operator") == "dilate") { o.X0 = f.Region.X0; o.Y0 = f.Region.Y0; o.X1 = f.Region.X1; o.Y1 = f.Region.Y1; }
                return o;
            }
            case "feTile":
            {
                var input = Input(prim, "in", f, lin);
                var o = Blank(f, lin);
                int tw = input.X1 - input.X0, th = input.Y1 - input.Y0;
                if (tw <= 0 || th <= 0) return o;
                for (int y = 0; y < H; y++)
                    for (int x = 0; x < W; x++)
                    {
                        int sx = input.X0 + (((x - input.X0) % tw) + tw) % tw, sy = input.Y0 + (((y - input.Y0) % th) + th) % th;
                        Array.Copy(input.P, (sy * W + sx) * 4, o.P, (y * W + x) * 4, 4);
                    }
                return o;
            }
            case "feImage": return FeImage(prim, f, lin);
            case "feTurbulence":
            {
                var o = Blank(f, lin);
                var sub = Subregion(prim, f, o);
                Turbulence(prim, o.P, f, sub);
                // Шум задаёт значения в рабочем пространстве фильтра
                return o;
            }
            case "feConvolveMatrix": return Convolve(prim, Input(prim, "in", f, lin), W, H);
            case "feDisplacementMap":
            {
                var input = Input(prim, "in", f, lin);
                var map = Input(prim, "in2", f, lin);
                var (sx, sy) = PrimLen(f, Num(Attr(prim, "scale")), Num(Attr(prim, "scale")));
                int xc = Channel(Attr(prim, "xChannelSelector")), yc = Channel(Attr(prim, "yChannelSelector"));
                var o = new FImg(new float[input.P.Length], lin, input.X0, input.Y0, input.X1, input.Y1);
                for (int y = 0; y < H; y++)
                    for (int x = 0; x < W; x++)
                    {
                        int i = (y * W + x) * 4;
                        float ma = map.P[i + 3];
                        double vx = xc == 3 ? ma : ma > 0 ? map.P[i + xc] / ma : 0, vy = yc == 3 ? ma : ma > 0 ? map.P[i + yc] / ma : 0;
                        int tx = (int)Math.Floor(x + sx * (vx - 0.5) + 0.5), ty = (int)Math.Floor(y + sy * (vy - 0.5) + 0.5);
                        if (tx < 0 || ty < 0 || tx >= W || ty >= H) continue;
                        Array.Copy(input.P, (ty * W + tx) * 4, o.P, i, 4);
                    }
                return o;
            }
            case "feDiffuseLighting" or "feSpecularLighting":
                return Lighting(prim, tag == "feSpecularLighting", ps, lin, Input(prim, "in", f, lin), f);
        }
        return null;
    }

    static int Channel(string? v) => v switch { "R" => 0, "G" => 1, "B" => 2, _ => 3 };

    // ---------------- Размытие ----------------

    static void Blur(float[] p, int W, int H, double sx, double sy)
    {
        if (sx > 0.01) BlurDir(p, W, H, sx, true);
        if (sy > 0.01) BlurDir(p, W, H, sy, false);
    }

    /// <summary>Размытие по одной оси: при σ ≥ 2 — три прямоугольных (как в спецификации), иначе — точное гауссово ядро.</summary>
    static void BlurDir(float[] p, int W, int H, double sigma, bool horiz)
    {
        int len = horiz ? W : H, lines = horiz ? H : W;
        var line = new float[len * 4];
        var tmp = new float[len * 4];
        float[]? kernel = null;
        int kr = 0, d = 0;
        if (sigma < 2)
        {
            kr = Math.Max(1, (int)Math.Ceiling(sigma * 3));
            kernel = new float[2 * kr + 1];
            double sum = 0;
            for (int i = -kr; i <= kr; i++) sum += kernel[i + kr] = (float)Math.Exp(-i * i / (2 * sigma * sigma));
            for (int i = 0; i < kernel.Length; i++) kernel[i] /= (float)sum;
        }
        else d = (int)Math.Floor(sigma * 3 * Math.Sqrt(2 * Math.PI) / 4 + 0.5);
        for (int l = 0; l < lines; l++)
        {
            for (int i = 0; i < len; i++)
            {
                int src = horiz ? (l * W + i) * 4 : (i * W + l) * 4;
                for (int ch = 0; ch < 4; ch++) line[i * 4 + ch] = p[src + ch];
            }
            if (kernel != null)
            {
                for (int i = 0; i < len; i++)
                    for (int ch = 0; ch < 4; ch++)
                    {
                        float s = 0;
                        for (int k = -kr; k <= kr; k++)
                        {
                            int j = i + k;
                            if (j >= 0 && j < len) s += line[j * 4 + ch] * kernel[k + kr];
                        }
                        tmp[i * 4 + ch] = s;
                    }
                (line, tmp) = (tmp, line);
            }
            else if (d % 2 == 1)
            {
                Box(line, tmp, len, d, d / 2); Box(tmp, line, len, d, d / 2); Box(line, tmp, len, d, d / 2);
                (line, tmp) = (tmp, line);
            }
            else
            {
                Box(line, tmp, len, d, d / 2); Box(tmp, line, len, d, d / 2 - 1); Box(line, tmp, len, d + 1, d / 2);
                (line, tmp) = (tmp, line);
            }
            for (int i = 0; i < len; i++)
            {
                int dst = horiz ? (l * W + i) * 4 : (i * W + l) * 4;
                for (int ch = 0; ch < 4; ch++) p[dst + ch] = line[i * 4 + ch];
            }
        }
    }

    /// <summary>Прямоугольное усреднение: окно [i − left, i − left + size − 1], за краями — прозрачность.</summary>
    static void Box(float[] src, float[] dst, int len, int size, int left)
    {
        if (size <= 1) { Array.Copy(src, dst, len * 4); return; }
        float inv = 1f / size;
        for (int ch = 0; ch < 4; ch++)
        {
            double sum = 0;
            for (int j = -left; j < -left + size; j++) if (j >= 0 && j < len) sum += src[j * 4 + ch];
            for (int i = 0; i < len; i++)
            {
                dst[i * 4 + ch] = (float)(sum * inv);
                int add = i - left + size, rem = i - left;
                if (add >= 0 && add < len) sum += src[add * 4 + ch];
                if (rem >= 0 && rem < len) sum -= src[rem * 4 + ch];
            }
        }
    }

    static FImg Shift(FImg input, int W, int H, double dx, double dy)
    {
        int ix = (int)Math.Round(dx), iy = (int)Math.Round(dy);
        var o = new FImg(new float[input.P.Length], input.Lin, Math.Clamp(input.X0 + ix, 0, W), Math.Clamp(input.Y0 + iy, 0, H),
            Math.Clamp(input.X1 + ix, 0, W), Math.Clamp(input.Y1 + iy, 0, H));
        for (int y = 0; y < H; y++)
        {
            int sy = y - iy;
            if (sy < 0 || sy >= H) continue;
            for (int x = 0; x < W; x++)
            {
                int sx = x - ix;
                if (sx < 0 || sx >= W) continue;
                Array.Copy(input.P, (sy * W + sx) * 4, o.P, (y * W + x) * 4, 4);
            }
        }
        return o;
    }

    static void Fill(float[] p, int W, (int X0, int Y0, int X1, int Y1) r, Rgb c, double a)
    {
        for (int y = r.Y0; y < r.Y1; y++)
            for (int x = r.X0; x < r.X1; x++)
            {
                int o = (y * W + x) * 4;
                p[o] = (float)(c.R * a); p[o + 1] = (float)(c.G * a); p[o + 2] = (float)(c.B * a); p[o + 3] = (float)a;
            }
    }

    static void Over(float[] dst, float[] src)
    {
        for (int i = 0; i < dst.Length; i += 4)
        {
            float a = src[i + 3];
            if (a <= 0) continue;
            float k = 1 - a;
            for (int ch = 0; ch < 4; ch++) dst[i + ch] = src[i + ch] + dst[i + ch] * k;
        }
    }

    static FImg DropShadow(FImg input, int W, int H, double bx, double by, double dx, double dy, Rgb c, double a, (int X0, int Y0, int X1, int Y1) region)
    {
        var sh = AlphaOnly(input);
        Blur(sh.P, W, H, bx, by);
        sh = Shift(sh, W, H, dx, dy);
        for (int i = 0; i < sh.P.Length; i += 4)
        {
            float al = sh.P[i + 3] * (float)a;
            sh.P[i] = (float)c.R * al; sh.P[i + 1] = (float)c.G * al; sh.P[i + 2] = (float)c.B * al; sh.P[i + 3] = al;
        }
        Over(sh.P, input.P);
        sh.X0 = region.X0; sh.Y0 = region.Y0; sh.X1 = region.X1; sh.Y1 = region.Y1;
        return sh;
    }

    // ---------------- Цвет ----------------

    static double[] SaturateMatrix(double s) =>
    [
        0.213 + 0.787 * s, 0.715 - 0.715 * s, 0.072 - 0.072 * s, 0, 0,
        0.213 - 0.213 * s, 0.715 + 0.285 * s, 0.072 - 0.072 * s, 0, 0,
        0.213 - 0.213 * s, 0.715 - 0.715 * s, 0.072 + 0.928 * s, 0, 0,
        0, 0, 0, 1, 0,
    ];

    static double[] HueMatrix(double deg)
    {
        double r = deg * Math.PI / 180, c = Math.Cos(r), s = Math.Sin(r);
        return
        [
            0.213 + c * 0.787 - s * 0.213, 0.715 - c * 0.715 - s * 0.715, 0.072 - c * 0.072 + s * 0.928, 0, 0,
            0.213 - c * 0.213 + s * 0.143, 0.715 + c * 0.285 + s * 0.140, 0.072 - c * 0.072 - s * 0.283, 0, 0,
            0.213 - c * 0.213 - s * 0.787, 0.715 - c * 0.715 + s * 0.715, 0.072 + c * 0.928 + s * 0.072, 0, 0,
            0, 0, 0, 1, 0,
        ];
    }

    /// <summary>Матрица 4×5 к непредумноженным значениям.</summary>
    static void ColorMatrix(float[] p, double[] m)
    {
        for (int i = 0; i < p.Length; i += 4)
        {
            float a = p[i + 3];
            double r = a > 0 ? p[i] / a : 0, g = a > 0 ? p[i + 1] / a : 0, b = a > 0 ? p[i + 2] / a : 0;
            double nr = m[0] * r + m[1] * g + m[2] * b + m[3] * a + m[4];
            double ng = m[5] * r + m[6] * g + m[7] * b + m[8] * a + m[9];
            double nb = m[10] * r + m[11] * g + m[12] * b + m[13] * a + m[14];
            double na = Math.Clamp(m[15] * r + m[16] * g + m[17] * b + m[18] * a + m[19], 0, 1);
            p[i] = (float)(Math.Clamp(nr, 0, 1) * na); p[i + 1] = (float)(Math.Clamp(ng, 0, 1) * na); p[i + 2] = (float)(Math.Clamp(nb, 0, 1) * na); p[i + 3] = (float)na;
        }
    }

    static Func<double, double>? TransferFunc(XElement fn)
    {
        string type = Attr(fn, "type") ?? "identity";
        var tv = Numbers(Attr(fn, "tableValues"));
        switch (type)
        {
            case "table" when tv.Count > 0:
                return c =>
                {
                    int n = tv.Count - 1;
                    if (n == 0) return tv[0];
                    int k = Math.Min((int)Math.Floor(c * n), n - 1);
                    return tv[k] + (c - (double)k / n) * n * (tv[k + 1] - tv[k]);
                };
            case "discrete" when tv.Count > 0:
                return c => tv[Math.Min((int)Math.Floor(c * tv.Count), tv.Count - 1)];
            case "linear":
            {
                double slope = Num(Attr(fn, "slope"), 1), icp = Num(Attr(fn, "intercept"));
                return c => slope * c + icp;
            }
            case "gamma":
            {
                double amp = Num(Attr(fn, "amplitude"), 1), exp = Num(Attr(fn, "exponent"), 1), off = Num(Attr(fn, "offset"));
                return c => amp * Math.Pow(c, exp) + off;
            }
        }
        return null;
    }

    static void Transfer(float[] p, Func<double, double>?[] f)
    {
        if (f.All(x => x == null)) return;
        // Таблица на 256 значений на канал
        var lut = f.Select(fn => fn == null ? null : Enumerable.Range(0, 256).Select(i => (float)Math.Clamp(fn(i / 255.0), 0, 1)).ToArray()).ToArray();
        for (int i = 0; i < p.Length; i += 4)
        {
            float a = p[i + 3];
            float[] c = [a > 0 ? p[i] / a : 0, a > 0 ? p[i + 1] / a : 0, a > 0 ? p[i + 2] / a : 0, a];
            for (int ch = 0; ch < 4; ch++)
                if (lut[ch] is { } l)
                {
                    float v = Math.Clamp(c[ch], 0, 1) * 255;
                    int k = Math.Min(254, (int)v);
                    c[ch] = l[k] + (l[k + 1] - l[k]) * (v - k);
                }
            p[i] = c[0] * c[3]; p[i + 1] = c[1] * c[3]; p[i + 2] = c[2] * c[3]; p[i + 3] = c[3];
        }
    }

    static void Blend(float[] s, float[] b, float[] o, string mode)
    {
        mode = mode.Trim().ToLowerInvariant();
        for (int i = 0; i < o.Length; i += 4)
        {
            float sa = s[i + 3], ba = b[i + 3];
            float ra = sa + ba - sa * ba;
            if (mode is "hue" or "saturation" or "color" or "luminosity")
            {
                double[] cs = [sa > 0 ? s[i] / sa : 0, sa > 0 ? s[i + 1] / sa : 0, sa > 0 ? s[i + 2] / sa : 0];
                double[] cb = [ba > 0 ? b[i] / ba : 0, ba > 0 ? b[i + 1] / ba : 0, ba > 0 ? b[i + 2] / ba : 0];
                var mix = NonSeparable(cs, cb, mode);
                for (int ch = 0; ch < 3; ch++)
                    o[i + ch] = (float)((1 - ba) * s[i + ch] + (1 - sa) * b[i + ch] + sa * ba * mix[ch]);
                o[i + 3] = ra;
                continue;
            }
            for (int ch = 0; ch < 3; ch++)
            {
                double cs = sa > 0 ? s[i + ch] / sa : 0, cb = ba > 0 ? b[i + ch] / ba : 0;
                double m = mode switch
                {
                    "multiply" => cs * cb,
                    "screen" => cs + cb - cs * cb,
                    "darken" => Math.Min(cs, cb),
                    "lighten" => Math.Max(cs, cb),
                    "overlay" => cb <= 0.5 ? 2 * cs * cb : 1 - 2 * (1 - cs) * (1 - cb),
                    "hard-light" => cs <= 0.5 ? 2 * cs * cb : 1 - 2 * (1 - cs) * (1 - cb),
                    "color-dodge" => cb == 0 ? 0 : cs >= 1 ? 1 : Math.Min(1, cb / (1 - cs)),
                    "color-burn" => cb >= 1 ? 1 : cs <= 0 ? 0 : 1 - Math.Min(1, (1 - cb) / cs),
                    "soft-light" => cs <= 0.5 ? cb - (1 - 2 * cs) * cb * (1 - cb)
                        : cb + (2 * cs - 1) * ((cb <= 0.25 ? ((16 * cb - 12) * cb + 4) * cb : Math.Sqrt(cb)) - cb),
                    "difference" => Math.Abs(cs - cb),
                    "exclusion" => cs + cb - 2 * cs * cb,
                    _ => cs,
                };
                o[i + ch] = (float)((1 - ba) * s[i + ch] + (1 - sa) * b[i + ch] + sa * ba * m);
            }
            o[i + 3] = ra;
        }
    }

    static double[] NonSeparable(double[] cs, double[] cb, string mode)
    {
        static double Lum(double[] c) => 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2];
        static double Sat(double[] c) => c.Max() - c.Min();
        static double[] ClipC(double[] c)
        {
            double l = Lum(c), n = c.Min(), x = c.Max();
            var r = (double[])c.Clone();
            if (n < 0) for (int i = 0; i < 3; i++) r[i] = l + (r[i] - l) * l / (l - n);
            if (x > 1) for (int i = 0; i < 3; i++) r[i] = l + (r[i] - l) * (1 - l) / (x - l);
            return r;
        }
        static double[] SetLum(double[] c, double l) { double d = l - Lum(c); return ClipC([c[0] + d, c[1] + d, c[2] + d]); }
        static double[] SetSat(double[] c, double s)
        {
            double mx = c.Max(), mn = c.Min();
            var r = new double[3];
            if (mx > mn) for (int i = 0; i < 3; i++) r[i] = (c[i] - mn) * s / (mx - mn);
            return r;
        }
        return mode switch
        {
            "hue" => SetLum(SetSat(cs, Sat(cb)), Lum(cb)),
            "saturation" => SetLum(SetSat(cb, Sat(cs)), Lum(cb)),
            "color" => SetLum(cs, Lum(cb)),
            _ => SetLum(cb, Lum(cs)),
        };
    }

    // ---------------- Морфология, свёртка ----------------

    static void Morph(float[] p, int W, int H, int rx, int ry, bool dilate)
    {
        var tmp = new float[p.Length];
        // По горизонтали
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
                for (int ch = 0; ch < 4; ch++)
                {
                    float v = dilate ? 0 : 1;
                    for (int k = -rx; k <= rx; k++)
                    {
                        int xx = x + k;
                        float s = xx >= 0 && xx < W ? p[(y * W + xx) * 4 + ch] : 0;
                        v = dilate ? Math.Max(v, s) : Math.Min(v, s);
                    }
                    tmp[(y * W + x) * 4 + ch] = v;
                }
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
                for (int ch = 0; ch < 4; ch++)
                {
                    float v = dilate ? 0 : 1;
                    for (int k = -ry; k <= ry; k++)
                    {
                        int yy = y + k;
                        float s = yy >= 0 && yy < H ? tmp[(yy * W + x) * 4 + ch] : 0;
                        v = dilate ? Math.Max(v, s) : Math.Min(v, s);
                    }
                    p[(y * W + x) * 4 + ch] = v;
                }
    }

    static FImg Convolve(XElement prim, FImg input, int W, int H)
    {
        var ord = Numbers(Attr(prim, "order") ?? "3");
        int ox = Math.Max(1, (int)(ord.Count > 0 ? ord[0] : 3)), oy = Math.Max(1, (int)(ord.Count > 1 ? ord[1] : ox));
        var km = Numbers(Attr(prim, "kernelMatrix"));
        if (km.Count != ox * oy) return input;
        double divisor = Num(Attr(prim, "divisor"), 0);
        if (divisor == 0) { divisor = km.Sum(); if (divisor == 0) divisor = 1; }
        double bias = Num(Attr(prim, "bias"));
        int tx = Attr(prim, "targetX") != null ? (int)Num(Attr(prim, "targetX")) : ox / 2, ty = Attr(prim, "targetY") != null ? (int)Num(Attr(prim, "targetY")) : oy / 2;
        string edge = Attr(prim, "edgeMode") ?? "duplicate";
        bool preserveAlpha = Attr(prim, "preserveAlpha") == "true";
        var o = new FImg(new float[input.P.Length], input.Lin, input.X0, input.Y0, input.X1, input.Y1);
        float Sample(int x, int y, int ch)
        {
            if (x < 0 || y < 0 || x >= W || y >= H)
            {
                if (edge == "none") return 0;
                if (edge == "wrap") { x = ((x % W) + W) % W; y = ((y % H) + H) % H; }
                else { x = Math.Clamp(x, 0, W - 1); y = Math.Clamp(y, 0, H - 1); }
            }
            int i = (y * W + x) * 4;
            if (preserveAlpha && ch < 3) { float a = input.P[i + 3]; return a > 0 ? input.P[i + ch] / a : 0; }
            return input.P[i + ch];
        }
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                int i = (y * W + x) * 4;
                double[] acc = new double[4];
                for (int j = 0; j < oy; j++)
                    for (int k = 0; k < ox; k++)
                    {
                        double kv = km[(oy - 1 - j) * ox + (ox - 1 - k)];
                        for (int ch = 0; ch < (preserveAlpha ? 3 : 4); ch++) acc[ch] += Sample(x - tx + k, y - ty + j, ch) * kv;
                    }
                if (preserveAlpha)
                {
                    float a = input.P[i + 3];
                    for (int ch = 0; ch < 3; ch++) o.P[i + ch] = (float)Math.Clamp(acc[ch] / divisor + bias, 0, 1) * a;
                    o.P[i + 3] = a;
                }
                else
                {
                    float a = (float)Math.Clamp(acc[3] / divisor + bias, 0, 1);
                    for (int ch = 0; ch < 3; ch++) o.P[i + ch] = (float)Math.Clamp(acc[ch] / divisor + bias * a, 0, a);
                    o.P[i + 3] = a;
                }
            }
        return o;
    }

    // ---------------- feImage ----------------

    static FImg FeImage(XElement prim, FState f, bool lin)
    {
        var g = f.G;
        var ctx = f.Ctx;
        var c = ctx.C;
        var o = Blank(f, false);
        var sub = Subregion(prim, f, o);
        string? href = Href(prim)?.Trim();
        if (string.IsNullOrEmpty(href)) return o;
        string ops;
        if (href.StartsWith('#'))
        {
            var target = Ref(ctx.Doc, href);
            if (target == null || target.AncestorsAndSelf().Contains(f.Element) || target == prim) return o;
            ops = c.CaptureOps(() => RenderElement(target, StyleOf(ctx.Doc, target.Parent), ctx));
        }
        else
        {
            if (!href.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return o;
            byte[] bytes;
            try
            {
                int comma = href.IndexOf(',');
                string meta = href[5..comma], data = href[(comma + 1)..];
                bytes = meta.Contains(";base64", StringComparison.OrdinalIgnoreCase)
                    ? Convert.FromBase64String(Regex.Replace(Uri.UnescapeDataString(data), @"\s", ""))
                    : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(data));
            }
            catch { return o; }
            double iw, ih;
            if (IsSvg(bytes) && Size(bytes) is { } ss) { iw = ss.W / 0.75; ih = ss.H / 0.75; }
            else if (ImageDecoder.Decode(bytes) is { } info) { iw = info.Width; ih = info.Height; }
            else return o;
            double ux = g.Rx + sub.X0 / g.Sx, uy = g.Ry + sub.Y0 / g.Sy, uw = (sub.X1 - sub.X0) / g.Sx, uh = (sub.Y1 - sub.Y0) / g.Sy;
            if (uw <= 0 || uh <= 0) return o;
            var vm = ViewMatrix((0, 0, iw, ih), ux, uy, uw, uh, Attr(prim, "preserveAspectRatio"));
            var (dx, dy) = vm.Apply(0, 0);
            ops = c.CaptureOps(() => c.Image("svg-feimage-" + bytes.Length + "-" + href.GetHashCode(), bytes, dx, dy, iw * vm.A, ih * vm.D));
        }
        var ras = new Raster(g.W, g.H, c);
        ras.Run(ops, BaseMat(g, c));
        Array.Copy(ras.Px, o.P, o.P.Length);
        (o.X0, o.Y0, o.X1, o.Y1) = sub;
        return ToSpace(o, lin);
    }

    // ---------------- feTurbulence (алгоритм из спецификации) ----------------

    sealed class Perlin
    {
        const int BSize = 0x100, BM = 0xff, PerlinN = 0x1000;
        readonly int[] _lattice = new int[BSize + BSize + 2];
        readonly double[,,] _gradient = new double[4, BSize + BSize + 2, 2];

        public Perlin(long seed)
        {
            seed = Setup(seed);
            for (int k = 0; k < 4; k++)
                for (int i = 0; i < BSize; i++)
                {
                    _lattice[i] = i;
                    for (int j = 0; j < 2; j++) _gradient[k, i, j] = (double)(((seed = Random(seed)) % (BSize + BSize)) - BSize) / BSize;
                    double s = Math.Sqrt(_gradient[k, i, 0] * _gradient[k, i, 0] + _gradient[k, i, 1] * _gradient[k, i, 1]);
                    if (s > 0) { _gradient[k, i, 0] /= s; _gradient[k, i, 1] /= s; }
                }
            for (int i = BSize - 1; i > 0; i--)
            {
                int k = _lattice[i];
                int j = (int)((seed = Random(seed)) % BSize);
                _lattice[i] = _lattice[j]; _lattice[j] = k;
            }
            for (int i = 0; i < BSize + 2; i++)
            {
                _lattice[BSize + i] = _lattice[i];
                for (int k = 0; k < 4; k++) for (int j = 0; j < 2; j++) _gradient[k, BSize + i, j] = _gradient[k, i, j];
            }
        }

        const long RandM = 2147483647, RandA = 16807, RandQ = 127773, RandR = 2836;
        static long Setup(long s) { if (s <= 0) s = -(s % (RandM - 1)) + 1; if (s > RandM - 1) s = RandM - 1; return s; }
        static long Random(long s) { s = RandA * (s % RandQ) - RandR * (s / RandQ); if (s <= 0) s += RandM; return s; }

        static double SCurve(double t) => t * t * (3 - 2 * t);
        static double Lerp(double t, double a, double b) => a + t * (b - a);

        public sealed class Stitch { public int Width, Height, WrapX, WrapY; }

        double Noise2(int ch, double vx, double vy, Stitch? st)
        {
            double t = vx + PerlinN;
            int bx0 = (int)t & BM, bx1 = (bx0 + 1) & BM;
            double rx0 = t - (int)t, rx1 = rx0 - 1;
            t = vy + PerlinN;
            int by0 = (int)t & BM, by1 = (by0 + 1) & BM;
            double ry0 = t - (int)t, ry1 = ry0 - 1;
            if (st != null)
            {
                if (bx0 >= st.WrapX) bx0 -= st.Width;
                if (bx1 >= st.WrapX) bx1 -= st.Width;
                if (by0 >= st.WrapY) by0 -= st.Height;
                if (by1 >= st.WrapY) by1 -= st.Height;
            }
            bx0 &= BM; bx1 &= BM; by0 &= BM; by1 &= BM;
            int i = _lattice[bx0], j = _lattice[bx1];
            int b00 = _lattice[i + by0], b10 = _lattice[j + by0], b01 = _lattice[i + by1], b11 = _lattice[j + by1];
            double sx = SCurve(rx0), sy = SCurve(ry0);
            double u = rx0 * _gradient[ch, b00, 0] + ry0 * _gradient[ch, b00, 1];
            double v = rx1 * _gradient[ch, b10, 0] + ry0 * _gradient[ch, b10, 1];
            double a = Lerp(sx, u, v);
            u = rx0 * _gradient[ch, b01, 0] + ry1 * _gradient[ch, b01, 1];
            v = rx1 * _gradient[ch, b11, 0] + ry1 * _gradient[ch, b11, 1];
            double b = Lerp(sx, u, v);
            return Lerp(sy, a, b);
        }

        public double Turbulence(int ch, double px, double py, double bfx, double bfy, int octaves, bool fractal, bool stitch, double tx, double ty, double tw, double th)
        {
            Stitch? st = null;
            if (stitch)
            {
                if (bfx != 0) { double lo = Math.Floor(tw * bfx) / tw, hi = Math.Ceiling(tw * bfx) / tw; bfx = bfx / lo < hi / bfx ? lo : hi; }
                if (bfy != 0) { double lo = Math.Floor(th * bfy) / th, hi = Math.Ceiling(th * bfy) / th; bfy = bfy / lo < hi / bfy ? lo : hi; }
                st = new Stitch { Width = (int)(tw * bfx + 0.5), Height = (int)(th * bfy + 0.5) };
                st.WrapX = (int)(tx * bfx + PerlinN + st.Width); st.WrapY = (int)(ty * bfy + PerlinN + st.Height);
            }
            double sum = 0, vx = px * bfx, vy = py * bfy, ratio = 1;
            for (int n = 0; n < octaves; n++)
            {
                double nv = Noise2(ch, vx, vy, st);
                sum += (fractal ? nv : Math.Abs(nv)) / ratio;
                vx *= 2; vy *= 2; ratio *= 2;
                if (st != null)
                {
                    st.Width *= 2; st.WrapX = 2 * st.WrapX - PerlinN;
                    st.Height *= 2; st.WrapY = 2 * st.WrapY - PerlinN;
                }
            }
            return sum;
        }
    }

    static void Turbulence(XElement prim, float[] p, FState f, (int X0, int Y0, int X1, int Y1) sub)
    {
        var g = f.G;
        var bf = Numbers(Attr(prim, "baseFrequency") ?? "0");
        double bfx = bf.Count > 0 ? bf[0] : 0, bfy = bf.Count > 1 ? bf[1] : bfx;
        if (bfx < 0 || bfy < 0) return;
        int oct = Math.Clamp((int)Num(Attr(prim, "numOctaves"), 1), 0, 24);
        long seed = (long)Math.Round(Num(Attr(prim, "seed"), 0));
        bool fractal = Attr(prim, "type") == "fractalNoise";
        bool stitch = Attr(prim, "stitchTiles") == "stitch";
        var noise = new Perlin(seed);
        double tx = g.Rx + sub.X0 / g.Sx, ty = g.Ry + sub.Y0 / g.Sy, tw = (sub.X1 - sub.X0) / g.Sx, th = (sub.Y1 - sub.Y0) / g.Sy;
        for (int y = sub.Y0; y < sub.Y1; y++)
            for (int x = sub.X0; x < sub.X1; x++)
            {
                // Точка в единицах пользователя (шум задан в системе координат элемента)
                double ux = g.Rx + (x + 0.5) / g.Sx, uy = g.Ry + (y + 0.5) / g.Sy;
                if (f.PrimObb) { ux = (ux - g.BBox.X) / Math.Max(1e-9, g.BBox.W); uy = (uy - g.BBox.Y) / Math.Max(1e-9, g.BBox.H); }
                var v = new double[4];
                for (int ch = 0; ch < 4; ch++)
                {
                    double t = noise.Turbulence(ch, ux, uy, bfx, bfy, oct, fractal, stitch, tx, ty, tw, th);
                    v[ch] = Math.Clamp(fractal ? (t + 1) / 2 : t, 0, 1);
                }
                int o = (y * g.W + x) * 4;
                p[o] = (float)(v[0] * v[3]); p[o + 1] = (float)(v[1] * v[3]); p[o + 2] = (float)(v[2] * v[3]); p[o + 3] = (float)v[3];
            }
    }

    // ---------------- Освещение ----------------

    static FImg Lighting(XElement prim, bool specular, Style ps, bool lin, FImg input, FState f)
    {
        var g = f.G;
        int W = g.W, H = g.H;
        double surface = Num(Attr(prim, "surfaceScale"), 1);
        double kd = Num(Attr(prim, "diffuseConstant"), 1), ks = Num(Attr(prim, "specularConstant"), 1), se = Num(Attr(prim, "specularExponent"), 1);
        var (lc, _) = FloodColor(ps, "lighting-color", "flood-opacity-none", lin);
        var light = prim.Elements().FirstOrDefault(x => x.Name.LocalName is "feDistantLight" or "fePointLight" or "feSpotLight");
        var o = new FImg(new float[W * H * 4], lin, f.Region.X0, f.Region.Y0, f.Region.X1, f.Region.Y1);
        if (light == null) return o;
        double zScale = (g.Sx + g.Sy) / 2;
        // Источник в пикселях
        double LX(string a) => g.Px(Num(Attr(light, a)));
        double LY(string a) => g.Py(Num(Attr(light, a)));
        string kind = light.Name.LocalName;
        double az = Num(Attr(light, "azimuth")) * Math.PI / 180, el = Num(Attr(light, "elevation")) * Math.PI / 180;
        double lx = 0, ly = 0, lz = 0, px = 0, py = 0, pz = 0, sexp = 1, cone = -1;
        if (kind == "feDistantLight") { lx = Math.Cos(az) * Math.Cos(el); ly = Math.Sin(az) * Math.Cos(el); lz = Math.Sin(el); }
        else
        {
            lx = LX("x"); ly = LY("y"); lz = Num(Attr(light, "z")) * zScale;
            if (kind == "feSpotLight")
            {
                px = LX("pointsAtX"); py = LY("pointsAtY"); pz = Num(Attr(light, "pointsAtZ")) * zScale;
                sexp = Num(Attr(light, "specularExponent"), 1);
                if (Attr(light, "limitingConeAngle") is { } lca) cone = Math.Cos(Math.Abs(Num(lca)) * Math.PI / 180);
            }
        }
        float A(int x, int y) => input.P[(Math.Clamp(y, 0, H - 1) * W + Math.Clamp(x, 0, W - 1)) * 4 + 3];
        for (int y = 0; y < H; y++)
            for (int x = 0; x < W; x++)
            {
                // Нормаль по оператору Собеля
                double nx = -surface / 4 * ((A(x + 1, y - 1) + 2 * A(x + 1, y) + A(x + 1, y + 1)) - (A(x - 1, y - 1) + 2 * A(x - 1, y) + A(x - 1, y + 1)));
                double ny = -surface / 4 * ((A(x - 1, y + 1) + 2 * A(x, y + 1) + A(x + 1, y + 1)) - (A(x - 1, y - 1) + 2 * A(x, y - 1) + A(x + 1, y - 1)));
                double nl = Math.Sqrt(nx * nx + ny * ny + 1);
                nx /= nl; ny /= nl; double nz = 1 / nl;
                double Lx, Ly, Lz;
                double cr = lc.R, cg = lc.G, cb = lc.B;
                if (kind == "feDistantLight") { Lx = lx; Ly = ly; Lz = lz; }
                else
                {
                    double z = surface * A(x, y);
                    Lx = lx - x; Ly = ly - y; Lz = lz - z;
                    double ll = Math.Sqrt(Lx * Lx + Ly * Ly + Lz * Lz);
                    if (ll > 0) { Lx /= ll; Ly /= ll; Lz /= ll; }
                    if (kind == "feSpotLight")
                    {
                        double sx = px - lx, sy = py - ly, sz = pz - lz, sl = Math.Sqrt(sx * sx + sy * sy + sz * sz);
                        if (sl > 0) { sx /= sl; sy /= sl; sz /= sl; }
                        double minus = -(Lx * sx + Ly * sy + Lz * sz);
                        double att = minus <= 0 || cone >= 0 && minus < cone ? 0 : Math.Pow(minus, sexp);
                        cr *= att; cg *= att; cb *= att;
                    }
                }
                int i = (y * W + x) * 4;
                if (!specular)
                {
                    double d = kd * Math.Max(0, nx * Lx + ny * Ly + nz * Lz);
                    o.P[i] = (float)Math.Clamp(d * cr, 0, 1); o.P[i + 1] = (float)Math.Clamp(d * cg, 0, 1); o.P[i + 2] = (float)Math.Clamp(d * cb, 0, 1); o.P[i + 3] = 1;
                }
                else
                {
                    double hx = Lx, hy = Ly, hz = Lz + 1, hl = Math.Sqrt(hx * hx + hy * hy + hz * hz);
                    if (hl > 0) { hx /= hl; hy /= hl; hz /= hl; }
                    double s = ks * Math.Pow(Math.Max(0, nx * hx + ny * hy + nz * hz), se);
                    float r = (float)Math.Clamp(s * cr, 0, 1), gg = (float)Math.Clamp(s * cg, 0, 1), b = (float)Math.Clamp(s * cb, 0, 1);
                    o.P[i] = r; o.P[i + 1] = gg; o.P[i + 2] = b; o.P[i + 3] = Math.Max(r, Math.Max(gg, b));
                }
            }
        return o;
    }

    // ================================================================
    //  CSS-функции фильтра
    // ================================================================

    static List<string> CssTokens(string s)
    {
        var res = new List<string>();
        foreach (Match m in Regex.Matches(s, @"[a-zA-Z]+\([^)]*\)|#[0-9a-fA-F]+|[^\s,()]+")) res.Add(m.Value);
        return res;
    }

    static double Amount(string s, double def)
    {
        s = s.Trim();
        if (s.Length == 0) return def;
        return s.EndsWith('%') ? Num(s) : Num(s, def);
    }

    static double Angle(string s)
    {
        s = s.Trim().ToLowerInvariant();
        double v = Num(Regex.Replace(s, "[a-z]+$", ""));
        return s.EndsWith("grad") ? v * 0.9 : s.EndsWith("rad") ? v * 180 / Math.PI : s.EndsWith("turn") ? v * 360 : v;
    }

    static FImg ApplyCssFunction(FFunc fn, FImg img, (int X0, int Y0, int X1, int Y1) region, FGrid g, Style st)
    {
        int W = g.W, H = g.H;
        var o = Copy(ToSpace(img, false));
        var t = CssTokens(fn.Args);
        string a0 = t.Count > 0 ? t[0] : "";
        double fs = FontSize(st);
        switch (fn.Name)
        {
            case "blur":
            {
                double r = Len(a0.Length > 0 ? a0 : "0", 0, fs);
                Blur(o.P, W, H, r * g.Sx, r * g.Sy);
                (o.X0, o.Y0, o.X1, o.Y1) = region;
                return o;
            }
            case "drop-shadow":
            {
                var colTok = t.FirstOrDefault(x => ParseColor(x, st.Get("color", "black")) != null);
                var lens = t.Where(x => x != colTok).Select(x => Len(x, 0, fs)).ToList();
                var col = colTok != null ? ParseColor(colTok, st.Get("color", "black"))!.Value : ParseColor(st.Get("color", "black")) ?? (Rgb.Black, 1);
                double dx = lens.ElementAtOrDefault(0), dy = lens.ElementAtOrDefault(1), sd = lens.ElementAtOrDefault(2);
                return DropShadow(o, W, H, sd * g.Sx, sd * g.Sy, dx * g.Sx, dy * g.Sy, col.Color, col.Alpha, region);
            }
            case "grayscale":
            {
                double k = 1 - Math.Clamp(Amount(a0, 1), 0, 1);
                ColorMatrix(o.P,
                [
                    0.2126 + 0.7874 * k, 0.7152 - 0.7152 * k, 0.0722 - 0.0722 * k, 0, 0,
                    0.2126 - 0.2126 * k, 0.7152 + 0.2848 * k, 0.0722 - 0.0722 * k, 0, 0,
                    0.2126 - 0.2126 * k, 0.7152 - 0.7152 * k, 0.0722 + 0.9278 * k, 0, 0,
                    0, 0, 0, 1, 0,
                ]);
                return o;
            }
            case "sepia":
            {
                double k = 1 - Math.Clamp(Amount(a0, 1), 0, 1);
                ColorMatrix(o.P,
                [
                    0.393 + 0.607 * k, 0.769 - 0.769 * k, 0.189 - 0.189 * k, 0, 0,
                    0.349 - 0.349 * k, 0.686 + 0.314 * k, 0.168 - 0.168 * k, 0, 0,
                    0.272 - 0.272 * k, 0.534 - 0.534 * k, 0.131 + 0.869 * k, 0, 0,
                    0, 0, 0, 1, 0,
                ]);
                return o;
            }
            case "saturate": ColorMatrix(o.P, SaturateMatrix(Amount(a0, 1))); return o;
            case "hue-rotate": ColorMatrix(o.P, HueMatrix(a0.Length > 0 ? Angle(a0) : 0)); return o;
            case "invert":
            {
                double k = Math.Clamp(Amount(a0, 1), 0, 1);
                Transfer(o.P, [c => k + c * (1 - 2 * k), c => k + c * (1 - 2 * k), c => k + c * (1 - 2 * k), null]);
                return o;
            }
            case "opacity":
            {
                double k = Math.Clamp(Amount(a0, 1), 0, 1);
                for (int i = 0; i < o.P.Length; i++) o.P[i] *= (float)k;
                return o;
            }
            case "brightness":
            {
                double k = Math.Max(0, Amount(a0, 1));
                Transfer(o.P, [c => c * k, c => c * k, c => c * k, null]);
                return o;
            }
            case "contrast":
            {
                double k = Math.Max(0, Amount(a0, 1));
                Transfer(o.P, [c => c * k - 0.5 * k + 0.5, c => c * k - 0.5 * k + 0.5, c => c * k - 0.5 * k + 0.5, null]);
                return o;
            }
        }
        return o;
    }
}

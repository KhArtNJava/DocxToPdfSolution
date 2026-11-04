using System.Globalization;
using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>Сноски, плавающие объекты (картинки, надписи, WordArt), буквицы.</summary>
public sealed partial class DocxReader
{
    static readonly XNamespace WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    static readonly XNamespace PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    static readonly XNamespace WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    static readonly XNamespace W10 = "urn:schemas-microsoft-com:office:word";

    /// <summary>Блоки, которые абзац порождает перед собой (плавающие объекты).</summary>
    List<Block>? _pending;

    readonly Dictionary<string, (XElement Note, PartContext Ctx)> _footnotes = [];
    readonly Dictionary<string, (XElement Note, PartContext Ctx)> _endnotes = [];
    readonly List<(string Mark, List<Block> Blocks)> _endnoteList = [];
    int _footCount, _endCount;
    string? _noteMark;

    // ================================================================
    //  Сноски
    // ================================================================

    void LoadNotes(string? part, Dictionary<string, (XElement, PartContext)> target)
    {
        if (part == null) return;
        var x = LoadXml(part);
        if (x?.Root == null) return;
        var ctx = new PartContext { Rels = LoadRels(part), HeaderFooter = false };
        foreach (var n in x.Root.Elements())
        {
            string? id = (string?)n.Attribute(W + "id");
            string? type = (string?)n.Attribute(W + "type");
            if (id != null && type is null or "normal") target[id] = (n, ctx);
        }
    }

    List<Block> ParseNote((XElement Note, PartContext Ctx) note, string mark)
    {
        var saved = _noteMark;
        _noteMark = mark;
        var blocks = new List<Block>();
        ParseBlockContainer(note.Note, note.Ctx, blocks);
        _noteMark = saved;
        return blocks;
    }

    /// <summary>Ссылки на сноски и номер сноски внутри её текста.</summary>
    bool ParseNoteElement(XElement el, TextStyle style, List<Inline> output)
    {
        var n = el.Name;
        if (n == W + "footnoteReference")
        {
            string? id = (string?)el.Attribute(W + "id");
            bool custom = el.Attribute(W + "customMarkFollows") is { } cm && IsOn(cm);
            string mark = custom ? "" : NoteMark(_footStart - 1 + ++_footCount, _footFmt);
            var blocks = id != null && _footnotes.TryGetValue(id, out var note) ? ParseNote(note, mark) : [];
            output.Add(new FootnoteInline(mark, style, blocks));
            return true;
        }
        if (n == W + "endnoteReference")
        {
            string? id = (string?)el.Attribute(W + "id");
            string mark = NoteMark(_endStart - 1 + ++_endCount, _endFmt);
            if (id != null && _endnotes.TryGetValue(id, out var note)) _endnoteList.Add((mark, ParseNote(note, mark)));
            output.Add(new TextInline(mark, style));
            return true;
        }
        if (n == W + "footnoteRef" || n == W + "endnoteRef")
        {
            if (!string.IsNullOrEmpty(_noteMark)) output.Add(new TextInline(_noteMark, style));
            return true;
        }
        if (n == W + "separator" || n == W + "continuationSeparator") return true;
        return false;
    }

    /// <summary>Концевые сноски — в конце документа после короткой линии.</summary>
    void AppendEndnotes(List<Block> blocks)
    {
        if (_endnoteList.Count == 0) return;
        var line = new Paragraph
        {
            SpaceBefore = 12, IndRight = 360, LineRule = "exact", LineValue = 2,
            BorderBottom = new BorderDef { Width = 0.5, Color = Rgb.Black }, MarkStyle = new TextStyle { Size = 1 },
        };
        // Последний абзац документа может закрывать раздел — сноски вставляем перед ним не нужно: добавляем в конец
        blocks.Add(line);
        foreach (var (_, b) in _endnoteList) blocks.AddRange(b);
    }

    // ================================================================
    //  DrawingML: картинки и надписи (inline / anchor)
    // ================================================================

    static double Emu(string? v) => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var e) ? e / 12700.0 : 0;

    /// <summary>Цвета схемы темы (заполняются при чтении theme1.xml).</summary>
    [ThreadStatic] internal static Dictionary<string, Rgb>? ThemeColors;

    /// <summary>
    /// Цвет DrawingML (srgbClr, sysClr, schemeClr, prstClr) с модификаторами lumMod/lumOff, shade, tint.
    /// Принимает элемент-контейнер (solidFill, fillRef, gs …).
    /// </summary>
    static Rgb? DrawingColor(XElement? fill)
    {
        if (fill == null) return null;
        var c = fill.Elements().FirstOrDefault(e => e.Name.LocalName is "srgbClr" or "sysClr" or "schemeClr" or "prstClr" or "scrgbClr");
        if (c == null) return null;
        Rgb? baseColor = c.Name.LocalName switch
        {
            "srgbClr" => Rgb.Parse((string?)c.Attribute("val")),
            "sysClr" => Rgb.Parse((string?)c.Attribute("lastClr")) ?? ((string?)c.Attribute("val") == "window" ? Rgb.White : Rgb.Black),
            "schemeClr" => SchemeColor((string?)c.Attribute("val")),
            "prstClr" => (string?)c.Attribute("val") switch
            {
                "black" => Rgb.Black, "white" => Rgb.White, "red" => new Rgb(1, 0, 0), "green" => new Rgb(0, 0.5, 0), "blue" => new Rgb(0, 0, 1),
                "yellow" => new Rgb(1, 1, 0), "gray" => new Rgb(0.5, 0.5, 0.5), _ => Rgb.Black
            },
            "scrgbClr" => new Rgb(EmuRaw(c.Attribute("r")) / 100000, EmuRaw(c.Attribute("g")) / 100000, EmuRaw(c.Attribute("b")) / 100000),
            _ => null
        };
        if (baseColor is not { } col) return null;
        return ApplyColorMods(col, c);
    }

    static Rgb? SchemeColor(string? name)
    {
        name = name switch { "tx1" => "dk1", "bg1" => "lt1", "tx2" => "dk2", "bg2" => "lt2", "phClr" => "accent1", _ => name };
        if (name != null && ThemeColors != null && ThemeColors.TryGetValue(name, out var t)) return t;
        return name switch
        {
            "lt1" => Rgb.White, "dk1" => Rgb.Black, "dk2" => Rgb.Parse("44546A"), "lt2" => Rgb.Parse("E7E6E6"),
            "accent1" => Rgb.Parse("4472C4"), "accent2" => Rgb.Parse("ED7D31"), "accent3" => Rgb.Parse("A5A5A5"),
            "accent4" => Rgb.Parse("FFC000"), "accent5" => Rgb.Parse("5B9BD5"), "accent6" => Rgb.Parse("70AD47"),
            "hlink" => Rgb.Parse("0563C1"), _ => null
        };
    }

    /// <summary>Модификаторы цвета DrawingML: lumMod/lumOff (в HSL), shade, tint.</summary>
    static Rgb ApplyColorMods(Rgb col, XElement clr)
    {
        double V(XElement e) => EmuRaw(e.Attribute("val")) / 100000.0;
        foreach (var m in clr.Elements())
        {
            switch (m.Name.LocalName)
            {
                case "shade": { double k = V(m); col = new Rgb(col.R * k, col.G * k, col.B * k); break; }
                case "tint": { double k = V(m); col = new Rgb(col.R + (1 - col.R) * (1 - k), col.G + (1 - col.G) * (1 - k), col.B + (1 - col.B) * (1 - k)); break; }
                case "lumMod" or "lumOff":
                {
                    var (hh, ss, ll) = ToHsl(col);
                    ll = m.Name.LocalName == "lumMod" ? ll * V(m) : ll + V(m);
                    col = FromHsl(hh, ss, Math.Clamp(ll, 0, 1));
                    break;
                }
            }
        }
        return col;
    }

    static (double H, double S, double L) ToHsl(Rgb c)
    {
        double max = Math.Max(c.R, Math.Max(c.G, c.B)), min = Math.Min(c.R, Math.Min(c.G, c.B)), l = (max + min) / 2;
        if (max - min < 1e-9) return (0, 0, l);
        double d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        double h = max == c.R ? (c.G - c.B) / d + (c.G < c.B ? 6 : 0) : max == c.G ? (c.B - c.R) / d + 2 : (c.R - c.G) / d + 4;
        return (h / 6, s, l);
    }

    static Rgb FromHsl(double h, double s, double l)
    {
        if (s < 1e-9) return new Rgb(l, l, l);
        double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
        double Hue(double t) { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p; }
        return new Rgb(Hue(h + 1.0 / 3), Hue(h), Hue(h - 1.0 / 3));
    }

    /// <summary>Обрезка (a:srcRect), поворот и отражение (a:xfrm) картинки DrawingML.</summary>
    static void ApplyPictureGeometry(XElement container, ImageInline img, bool inline)
    {
        static double Pct(XElement? e, string a) => double.TryParse((string?)e?.Attribute(a), NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v / 100000.0 : 0;
        var src = container.Descendants(A + "srcRect").FirstOrDefault();
        if (src != null)
        {
            img.CropL = Pct(src, "l"); img.CropT = Pct(src, "t"); img.CropR = Pct(src, "r"); img.CropB = Pct(src, "b");
            // Отрицательные значения — поля вокруг картинки; доля видимой части должна быть положительной
            if (1 - img.CropL - img.CropR < 0.01 || 1 - img.CropT - img.CropB < 0.01) img.CropL = img.CropT = img.CropR = img.CropB = 0;
        }
        var xfrm = container.Elements().FirstOrDefault(e => e.Name.LocalName == "spPr")?.Element(A + "xfrm")
                   ?? container.Descendants(PIC + "spPr").FirstOrDefault()?.Element(A + "xfrm")
                   ?? container.Descendants(A + "xfrm").FirstOrDefault();
        if (xfrm != null)
        {
            if (double.TryParse((string?)xfrm.Attribute("rot"), NumberStyles.Float, CultureInfo.InvariantCulture, out var rot)) img.Rotation = rot / 60000.0 % 360;
            img.FlipH = (string?)xfrm.Attribute("flipH") is "1" or "true";
            img.FlipV = (string?)xfrm.Attribute("flipV") is "1" or "true";
        }
        if (inline && img.Rotation != 0)
        {
            // В строке картинка занимает описанный прямоугольник повёрнутой рамки
            double a = img.Rotation * Math.PI / 180, c = Math.Abs(Math.Cos(a)), s = Math.Abs(Math.Sin(a));
            img.Width = img.FrameW * c + img.FrameH * s;
            img.Height = img.FrameW * s + img.FrameH * c;
        }
    }

    /// <summary>Геометрия, заливка и контур фигуры DrawingML (spPr, при отсутствии — ссылки стиля wps:style).</summary>
    static ShapeSpec ParseShapeSpec(XElement spPr, XElement? shapeStyle)
    {
        var sh = new ShapeSpec { Geom = (string?)spPr.Element(A + "prstGeom")?.Attribute("prst") ?? "rect" };
        if (spPr.Element(A + "custGeom")?.Element(A + "pathLst") is { } pathLst)
        {
            // Произвольный контур: точки пути в его системе (w×h) → доли рамки
            sh.Geom = "custom";
            sh.Paths = [];
            foreach (var path in pathLst.Elements(A + "path"))
            {
                double pw = EmuRaw(path.Attribute("w")), ph = EmuRaw(path.Attribute("h"));
                if (pw <= 0) pw = 1; if (ph <= 0) ph = 1;
                var cp = new CustomPath
                {
                    Fill = (string?)path.Attribute("fill") != "none",
                    Stroke = (string?)path.Attribute("stroke") is not ("0" or "false"),
                };
                double[] Pts(XElement e) => e.Elements(A + "pt").SelectMany(pt => new[] { EmuRaw(pt.Attribute("x")) / pw, EmuRaw(pt.Attribute("y")) / ph }).ToArray();
                double lastX = 0, lastY = 0;
                foreach (var cmd in path.Elements())
                {
                    switch (cmd.Name.LocalName)
                    {
                        case "moveTo": { var a = Pts(cmd); if (a.Length >= 2) { cp.Commands.Add(('M', a)); (lastX, lastY) = (a[0], a[1]); } break; }
                        case "lnTo": { var a = Pts(cmd); if (a.Length >= 2) { cp.Commands.Add(('L', a)); (lastX, lastY) = (a[0], a[1]); } break; }
                        case "cubicBezTo": { var a = Pts(cmd); if (a.Length >= 6) { cp.Commands.Add(('C', a)); (lastX, lastY) = (a[4], a[5]); } break; }
                        case "quadBezTo":
                        {
                            var a = Pts(cmd);
                            if (a.Length >= 4)
                            {
                                // Квадратичная кривая → кубическая
                                double c1x = lastX + 2.0 / 3 * (a[0] - lastX), c1y = lastY + 2.0 / 3 * (a[1] - lastY);
                                double c2x = a[2] + 2.0 / 3 * (a[0] - a[2]), c2y = a[3] + 2.0 / 3 * (a[1] - a[3]);
                                cp.Commands.Add(('C', [c1x, c1y, c2x, c2y, a[2], a[3]]));
                                (lastX, lastY) = (a[2], a[3]);
                            }
                            break;
                        }
                        case "arcTo":
                        {
                            // Дуга эллипса от текущей точки: радиусы wR/hR, начальный угол и размах (1/60000°)
                            double wr = EmuRaw(cmd.Attribute("wR")) / pw, hr = EmuRaw(cmd.Attribute("hR")) / ph;
                            double st0 = EmuRaw(cmd.Attribute("stAng")) / 60000 * Math.PI / 180, sw = EmuRaw(cmd.Attribute("swAng")) / 60000 * Math.PI / 180;
                            double ecx = lastX - wr * Math.Cos(st0), ecy = lastY - hr * Math.Sin(st0);
                            int segs = Math.Max(1, (int)Math.Ceiling(Math.Abs(sw) / (Math.PI / 2)));
                            double step = sw / segs, a0 = st0;
                            for (int k = 0; k < segs; k++)
                            {
                                double a1 = a0 + step, kk = 4.0 / 3 * Math.Tan(step / 4);
                                double x0 = ecx + wr * Math.Cos(a0), y0 = ecy + hr * Math.Sin(a0);
                                double x3 = ecx + wr * Math.Cos(a1), y3 = ecy + hr * Math.Sin(a1);
                                double x1 = x0 - kk * wr * Math.Sin(a0), y1 = y0 + kk * hr * Math.Cos(a0);
                                double x2 = x3 + kk * wr * Math.Sin(a1), y2 = y3 - kk * hr * Math.Cos(a1);
                                cp.Commands.Add(('C', [x1, y1, x2, y2, x3, y3]));
                                a0 = a1;
                                (lastX, lastY) = (x3, y3);
                            }
                            break;
                        }
                        case "close": cp.Commands.Add(('Z', [])); break;
                    }
                }
                sh.Paths.Add(cp);
            }
        }
        if (spPr.Element(A + "prstGeom")?.Element(A + "avLst")?.Elements(A + "gd").FirstOrDefault() is { } gd
            && ((string?)gd.Attribute("fmla"))?.Split(' ') is ["val", var av] && double.TryParse(av, NumberStyles.Float, CultureInfo.InvariantCulture, out var adj))
            sh.Adj = adj / 100000;
        if (spPr.Element(A + "noFill") != null) sh.Fill = null;
        else if (spPr.Element(A + "gradFill") is { } gf && gf.Element(A + "gsLst") is { } gs)
        {
            // Градиент: опорные цвета и угол (радиальный выводится линейным)
            sh.Gradient = gs.Elements(A + "gs").Select(g => (EmuRaw(g.Attribute("pos")) / 100000, DrawingColor(g) ?? Rgb.White))
                .OrderBy(t => t.Item1).ToList();
            sh.GradientAngle = EmuRaw(gf.Element(A + "lin")?.Attribute("ang")) / 60000;
            if (gf.Element(A + "path") != null) sh.GradientAngle = 90;
            sh.Fill = sh.Gradient.Count > 0 ? sh.Gradient[0].Color : null;
            if (sh.Gradient.Count < 2) sh.Gradient = null;
        }
        else sh.Fill = DrawingColor(spPr.Element(A + "solidFill")) ?? DrawingColor(shapeStyle?.Element(A + "fillRef"));
        var ln = spPr.Element(A + "ln");
        if (ln?.Element(A + "noFill") != null) sh.Stroke = null;
        else sh.Stroke = DrawingColor(ln?.Element(A + "solidFill")) ?? DrawingColor(shapeStyle?.Element(A + "lnRef"))
                         ?? (ln != null ? Rgb.Black : null);   // без a:ln и стиля контура нет
        if (ln?.Attribute("w") != null) sh.StrokeWidth = Math.Max(0.25, Emu((string?)ln.Attribute("w")));
        string dash = (string?)ln?.Element(A + "prstDash")?.Attribute("val") ?? "solid";
        sh.Dotted = dash is "dot" or "sysDot";
        sh.Dashed = !sh.Dotted && dash != "solid";
        sh.HeadArrow = (string?)ln?.Element(A + "headEnd")?.Attribute("type") is { } he && he != "none";
        sh.TailArrow = (string?)ln?.Element(A + "tailEnd")?.Attribute("type") is { } te && te != "none";
        if (sh.Geom.Contains("Connector") || sh.Geom == "line") sh.Fill = null;
        return sh;
    }

    void ParseDrawing(XElement drawing, RunState st, TextStyle style, List<Inline> output)
    {
        foreach (var container in drawing.Elements())
        {
            bool anchor = container.Name == WP + "anchor";
            if (!anchor && container.Name != WP + "inline") continue;
            var extent = container.Element(WP + "extent");
            double w = Emu((string?)extent?.Attribute("cx")), h = Emu((string?)extent?.Attribute("cy"));

            // Ссылка на объекте (wp:docPr/a:hlinkClick)
            var objStyle = style;
            if ((string?)container.Element(WP + "docPr")?.Element(A + "hlinkClick")?.Attribute(R + "id") is { } hrid
                && st.Ctx.Rels.TryGetValue(hrid, out var hrel))
                objStyle = style with { Link = hrel.External ? hrel.Target : "#" + hrel.Target.TrimStart('#') };

            var gobj = container.Element(A + "graphic")?.Element(A + "graphicData")?.Elements().FirstOrDefault();
            var item = gobj != null ? ParseGraphicObject(gobj, w, h, st, objStyle, inline: !anchor) : null;
            if (item == null) continue;
            if (!anchor) { output.Add(item); continue; }

            List<Block> content;
            BoxStyle box = new();
            if (item.Content != null && item.Rotation == 0 && item.Children == null && item.Shape == null)
            {
                // Обычная надпись с привязкой: содержимое в рамке (обтекание, разрыв и т. п. как раньше)
                content = item.Content;
                box = item.Box ?? new BoxStyle();
            }
            else
            {
                var p = new Paragraph { MarkStyle = style with { Size = 1 }, LineRule = "auto", SnapToGrid = false };
                p.Inlines.Add(item);
                content = [p];
            }
            EmitAnchor(container, w, h, content, box);
        }
    }

    /// <summary>
    /// Объект DrawingML внутри рамки w×h: картинка (pic:pic), фигура или надпись (wps:wsp), группа (wpg:wgp/grpSp).
    /// </summary>
    ImageInline? ParseGraphicObject(XElement obj, double w, double h, RunState st, TextStyle style, bool inline)
    {
        string ln = obj.Name.LocalName;
        if (ln == "pic")
        {
            var blip = obj.Descendants(A + "blip").FirstOrDefault();
            string? rid = (string?)blip?.Attribute(R + "embed") ?? (string?)blip?.Attribute(R + "link");
            if (rid == null) return null;
            var tmp = new List<Inline>();
            AddImage(rid, w, h, st, style, tmp);
            if (tmp.Count == 0) return null;
            var img = (ImageInline)tmp[0];
            ApplyPictureGeometry(obj, img, inline);
            // Прозрачность и оттенки серого (эффекты a:blip)
            if (blip!.Element(A + "alphaModFix") is { } amf) img.Opacity = Math.Clamp(EmuRaw(amf.Attribute("amt")) / 100000, 0, 1);
            img.Grayscale = blip.Element(A + "grayscl") != null;
            // Рамка картинки
            var pln = obj.Element(PIC + "spPr")?.Element(A + "ln");
            if (pln != null && pln.Element(A + "noFill") == null && DrawingColor(pln.Element(A + "solidFill")) is { } pc)
                img.Border = new BorderDef { Width = pln.Attribute("w") != null ? Math.Max(0.25, Emu((string?)pln.Attribute("w"))) : 0.75, Color = pc };
            return img;
        }
        if (ln is "wgp" or "grpSp")
        {
            var gx = obj.Element(WPG + "grpSpPr")?.Element(A + "xfrm");
            double chOffX = EmuRaw(gx?.Element(A + "chOff")?.Attribute("x")), chOffY = EmuRaw(gx?.Element(A + "chOff")?.Attribute("y"));
            double chExtX = EmuRaw(gx?.Element(A + "chExt")?.Attribute("cx")), chExtY = EmuRaw(gx?.Element(A + "chExt")?.Attribute("cy"));
            double sx = chExtX > 0 ? w / chExtX : 1 / 12700.0, sy = chExtY > 0 ? h / chExtY : 1 / 12700.0;
            var group = new ImageInline("", w, h, style) { Children = [], FrameW = w, FrameH = h };
            foreach (var child in obj.Elements())
            {
                if (child.Name.LocalName is not ("wsp" or "pic" or "grpSp")) continue;
                var cx = child.Name.LocalName == "grpSp"
                    ? child.Element(WPG + "grpSpPr")?.Element(A + "xfrm")
                    : child.Elements().FirstOrDefault(e => e.Name.LocalName == "spPr")?.Element(A + "xfrm");
                double ox = (EmuRaw(cx?.Element(A + "off")?.Attribute("x")) - chOffX) * sx;
                double oy = (EmuRaw(cx?.Element(A + "off")?.Attribute("y")) - chOffY) * sy;
                double cw = EmuRaw(cx?.Element(A + "ext")?.Attribute("cx")) * sx, ch = EmuRaw(cx?.Element(A + "ext")?.Attribute("cy")) * sy;
                var ci = ParseGraphicObject(child, cw, ch, st, style, inline: false);
                if (ci != null) group.Children.Add((ci, ox, oy));
            }
            ApplyGroupRotation(obj, group, inline);
            return group.Children.Count > 0 ? group : null;
        }
        if (ln == "chart" && (string?)obj.Attribute(R + "id") is { } chartRid && st.Ctx.Rels.TryGetValue(chartRid, out var chartRel)
            && ParseChart(chartRel.Target) is { } chart)
            return new ImageInline("", w, h, style) { Chart = chart, FrameW = w, FrameH = h };
        if (ln != "wsp") return null;
        var spPr = obj.Element(WPS + "spPr");
        var txbx = obj.Element(WPS + "txbx")?.Element(W + "txbxContent");
        if (txbx != null)
        {
            var content = new List<Block>();
            ParseBlockContainer(txbx, st.Ctx, content);
            var box = new BoxStyle();
            if (spPr != null && spPr.Element(A + "noFill") == null)
                box.Background = DrawingColor(spPr.Element(A + "solidFill")) ?? (spPr.Element(A + "solidFill") == null ? DrawingColor(obj.Element(WPS + "style")?.Element(A + "fillRef")) : null);
            var lnEl = spPr?.Element(A + "ln");
            if (lnEl != null && lnEl.Element(A + "noFill") == null)
            {
                var lc = DrawingColor(lnEl.Element(A + "solidFill")) ?? Rgb.Black;
                double lw = lnEl.Attribute("w") != null ? Math.Max(0.25, Emu((string?)lnEl.Attribute("w"))) : 0.75;
                var b = new BorderDef { Width = lw, Color = lc };
                box.BorderTop = box.BorderBottom = box.BorderLeft = box.BorderRight = b;
            }
            var bodyPr = obj.Element(WPS + "bodyPr");
            box.PadLeft = bodyPr?.Attribute("lIns") != null ? Emu((string?)bodyPr.Attribute("lIns")) : 7.2;
            box.PadRight = bodyPr?.Attribute("rIns") != null ? Emu((string?)bodyPr.Attribute("rIns")) : 7.2;
            box.PadTop = bodyPr?.Attribute("tIns") != null ? Emu((string?)bodyPr.Attribute("tIns")) : 3.6;
            box.PadBottom = bodyPr?.Attribute("bIns") != null ? Emu((string?)bodyPr.Attribute("bIns")) : 3.6;
            box.MinHeight = h;
            // Вертикальный текст: содержимое в ячейке с поворотом строк (как textDirection у ячейки таблицы)
            string? vert = (string?)bodyPr?.Attribute("vert");
            if (vert is "vert" or "vert270" or "eaVert" or "wordArtVert" or "mongolianVert" or "wordArtVertRtl")
            {
                var t = BoxTable(content, box, w, h, null);
                t.Rows[0].Cells[0].TextDirection = vert == "vert270" ? "btLr" : "tbRl";
                t.Rows[0].HeightRule = "exact";
                content = [t];
                box = new BoxStyle { MinHeight = h };
            }
            var tb = new ImageInline("", w, h, style) { Content = content, Box = box, FrameW = w, FrameH = h };
            // Надпись в фигуре (эллипс, стрелка …), с градиентом или с выравниванием текста по вертикали:
            // фон и контур рисуются фигурой, текст — поверх неё
            string anchorV = (string?)bodyPr?.Attribute("anchor") ?? "t";
            string prst = (string?)spPr?.Element(A + "prstGeom")?.Attribute("prst") ?? "rect";
            if (spPr != null && vert is null && (prst != "rect" || anchorV != "t" || spPr.Element(A + "gradFill") != null))
            {
                var tShape = ParseShapeSpec(spPr, obj.Element(WPS + "style"));
                if (box.Background is { } bg && tShape.Gradient == null) tShape.Fill = bg;
                if (box.BorderTop is { } bt) { tShape.Stroke = bt.Color; tShape.StrokeWidth = bt.Width; }
                box.Background = null;
                box.BorderTop = box.BorderBottom = box.BorderLeft = box.BorderRight = null;
                box.MinHeight = 0;
                tb.Shape = tShape;
                tb.VAnchor = anchorV is "ctr" or "b" ? anchorV : "t";
            }
            ApplyPictureGeometry(obj, tb, inline);
            return tb;
        }
        if (spPr == null) return null;
        // Фигура без текста: прямоугольник, эллипс, линия, стрелка, произвольный контур …
        var shape = ParseShapeSpec(spPr, obj.Element(WPS + "style"));
        var img2 = new ImageInline("", Math.Max(w, shape.StrokeWidth), Math.Max(h, shape.StrokeWidth), style) { Shape = shape, FrameW = w, FrameH = h };
        ApplyPictureGeometry(obj, img2, inline);
        return img2;
    }

    static readonly XNamespace CH = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>Диаграмма (word/charts/chartN.xml): тип, ряды с кэшем значений, категории, заголовок, легенда.</summary>
    ChartSpec? ParseChart(string part)
    {
        var x = LoadXml(part);
        var chart = x?.Root?.Element(CH + "chart");
        var plot = chart?.Element(CH + "plotArea");
        if (plot == null) return null;
        var kind = plot.Elements().FirstOrDefault(e => e.Name.LocalName.EndsWith("Chart"));
        if (kind == null) return null;
        string n = kind.Name.LocalName;
        var spec = new ChartSpec
        {
            Type = n switch
            {
                "barChart" or "bar3DChart" => (string?)kind.Element(CH + "barDir")?.Attribute("val") == "bar" ? "bar" : "col",
                "lineChart" or "line3DChart" or "stockChart" or "radarChart" => "line",
                "areaChart" or "area3DChart" => "area",
                "pieChart" or "pie3DChart" or "ofPieChart" => "pie",
                "doughnutChart" => "doughnut",
                "scatterChart" or "bubbleChart" => "scatter",
                _ => "col"
            },
            Grouping = (string?)kind.Element(CH + "grouping")?.Attribute("val") ?? "clustered",
        };
        // Заголовок: текст или (если один ряд и заголовок не удалён) имя ряда
        var title = chart!.Element(CH + "title");
        if (title != null)
            spec.Title = string.Concat(title.Descendants(A + "t").Select(t => t.Value)) is { Length: > 0 } tt ? tt : null;
        spec.Legend = chart.Element(CH + "legend") != null;
        spec.LegendPos = (string?)chart.Element(CH + "legend")?.Element(CH + "legendPos")?.Attribute("val") ?? "r";
        spec.Gridlines = plot.Elements().Any(e => e.Name.LocalName == "valAx" && e.Element(CH + "majorGridlines") != null)
                         || (spec.Type is not ("pie" or "doughnut") && !plot.Elements().Any(e => e.Name.LocalName == "valAx"));
        spec.DataLabels = (string?)kind.Element(CH + "dLbls")?.Element(CH + "showVal")?.Attribute("val") is "1" or "true";

        static List<string> Cache(XElement? e) =>
            e?.Descendants().FirstOrDefault(d => d.Name.LocalName is "strCache" or "numCache")?.Elements(CH + "pt")
                .OrderBy(pt => (int)EmuRaw(pt.Attribute("idx"))).Select(pt => (string?)pt.Element(CH + "v") ?? "").ToList() ?? [];

        foreach (var ser in kind.Elements(CH + "ser"))
        {
            var cs = new ChartSeries
            {
                Name = string.Concat(ser.Element(CH + "tx")?.Descendants(CH + "v").Select(v => v.Value) ?? []),
                Color = DrawingColor(ser.Element(CH + "spPr")?.Element(A + "solidFill"))
                        ?? DrawingColor(ser.Element(CH + "spPr")?.Element(A + "ln")?.Element(A + "solidFill")),
            };
            var vals = Cache(ser.Element(CH + "val") ?? ser.Element(CH + "yVal"));
            cs.Values = vals.Select(v => double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0).ToList();
            foreach (var dpt in ser.Elements(CH + "dPt"))
                if (DrawingColor(dpt.Element(CH + "spPr")?.Element(A + "solidFill")) is { } pc)
                    cs.PointColors[(int)EmuRaw(dpt.Element(CH + "idx")?.Attribute("val"))] = pc;
            if (spec.Categories.Count == 0) spec.Categories = Cache(ser.Element(CH + "cat") ?? ser.Element(CH + "xVal"));
            spec.Series.Add(cs);
        }
        if (spec.Series.Count == 0) return null;
        int count = spec.Series.Max(se => se.Values.Count);
        while (spec.Categories.Count < count) spec.Categories.Add((spec.Categories.Count + 1).ToString(CultureInfo.InvariantCulture));
        // Цвета по умолчанию — акценты темы по порядку
        string[] acc = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"];
        for (int i = 0; i < spec.Series.Count; i++) spec.Series[i].Color ??= SchemeColor(acc[i % 6]);
        // Круговая: у каждой точки свой цвет (по умолчанию — акценты по порядку)
        if (spec.Type is "pie" or "doughnut")
            foreach (var se in spec.Series)
                for (int i = 0; i < se.Values.Count; i++) se.PointColors.TryAdd(i, AccentColor(i));
        return spec;
    }

    /// <summary>Цвет акцента темы по номеру (для точек круговой диаграммы).</summary>
    internal static Rgb AccentColor(int i) => SchemeColor("accent" + (i % 6 + 1)) ?? Rgb.Black;

    static double EmuRaw(XAttribute? a) => double.TryParse((string?)a, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0;

    static void ApplyGroupRotation(XElement grp, ImageInline img, bool inline)
    {
        var x = grp.Element(WPG + "grpSpPr")?.Element(A + "xfrm");
        if (x == null) return;
        if (double.TryParse((string?)x.Attribute("rot"), NumberStyles.Float, CultureInfo.InvariantCulture, out var rot)) img.Rotation = rot / 60000.0 % 360;
        img.FlipH = (string?)x.Attribute("flipH") is "1" or "true";
        img.FlipV = (string?)x.Attribute("flipV") is "1" or "true";
        if (inline && img.Rotation != 0)
        {
            double a = img.Rotation * Math.PI / 180, c = Math.Abs(Math.Cos(a)), s = Math.Abs(Math.Sin(a));
            img.Width = img.FrameW * c + img.FrameH * s;
            img.Height = img.FrameW * s + img.FrameH * c;
        }
    }

    void AddFloat(string side, double w, double h, List<Block> content, BoxStyle box, double mt, double mb, double ml, double mr)
    {
        (_pending ?? []).Add(new FloatBlock
        {
            Side = side, Width = w, Box = box, Blocks = content,
            MarginTop = mt, MarginBottom = mb, MarginLeft = ml, MarginRight = mr,
        });
    }

    /// <summary>Якорь DrawingML: обтекание, положение, расстояния до текста.</summary>
    void EmitAnchor(XElement a, double w, double h, List<Block> content, BoxStyle box)
    {
        double dL = Emu((string?)a.Attribute("distL")), dR = Emu((string?)a.Attribute("distR"));
        double dT = Emu((string?)a.Attribute("distT")), dB = Emu((string?)a.Attribute("distB"));
        bool behind = a.Attribute("behindDoc") is { } bd && bd.Value is "1" or "true";
        var posH = a.Element(WP + "positionH"); var posV = a.Element(WP + "positionV");
        string hRel = (string?)posH?.Attribute("relativeFrom") ?? "column";
        string vRel = (string?)posV?.Attribute("relativeFrom") ?? "paragraph";
        string? hAlign = posH?.Element(WP + "align")?.Value; string? vAlign = posV?.Element(WP + "align")?.Value;
        double hOff = Emu(posH?.Element(WP + "posOffset")?.Value), vOff = Emu(posV?.Element(WP + "posOffset")?.Value);

        string wrap = a.Elements().FirstOrDefault(e => e.Name.LocalName.StartsWith("wrap"))?.Name.LocalName ?? "wrapNone";
        // Обтекание с вертикальной позицией от страницы/полей: объект ставится на своё место на странице
        // (обтекание в потоке возможно только для привязки к абзацу)
        bool absoluteV = vRel is "page" or "margin" or "topMargin" or "bottomMargin" or "insideMargin" or "outsideMargin";
        if (wrap is "wrapSquare" or "wrapTight" or "wrapThrough" && !behind && !absoluteV)
        {
            // Обтекание: плавающий блок у левого или правого края
            string side = hAlign switch
            {
                "right" or "outside" => "right",
                "left" or "inside" => "left",
                "center" => "left",
                _ => hOff > 200 ? "right" : "left"
            };
            double ml = side == "left" && hAlign == null && hRel is "margin" or "column" ? Math.Max(0, hOff) : 0;
            double mt = vRel is "paragraph" or "line" && vAlign == null ? Math.Max(0, vOff) : 0;
            AddFloat(side, w, h, content, box, mt, dB, ml + (side == "right" ? dL : 0), side == "left" ? dR : 0);
            return;
        }
        if (wrap == "wrapTopAndBottom")
        {
            // Отдельной строкой между абзацами
            if (content.Count == 1 && content[0] is Paragraph ip && box.Background == null && box.BorderTop == null)
            {
                ip.Align = hAlign switch { "center" => "center", "right" => "right", _ => "left" };
                if (hAlign == null && hOff > 0) ip.IndLeft = hOff;
                ip.SpaceBefore = dT + Math.Max(0, vOff); ip.SpaceAfter = dB;
                _pending?.Add(ip);
            }
            else _pending?.Add(BoxTable(content, box, w, h, hAlign));
            return;
        }
        // Без обтекания / за текстом — абсолютное положение
        var pb = new PositionedBlock
        {
            Width = w, Box = box, Blocks = content, ZIndex = behind ? -1 : 1,
            HRel = hRel switch { "page" => "page", "character" or "paragraph" => "flow", _ => "margin" },
            VRel = vRel switch { "page" => "page", "margin" or "topMargin" or "bottomMargin" => "margin", _ => "flow" },
            AlignX = hAlign, AlignY = vAlign, Left = hAlign == null ? hOff : null, Top = vAlign == null ? vOff : null,
        };
        _pending?.Add(pb);
    }

    /// <summary>Рамка-таблица из одной ячейки (надпись «сверху и снизу»).</summary>
    static Table BoxTable(List<Block> content, BoxStyle box, double w, double h, string? align)
    {
        var cell = new TableCell
        {
            Blocks = content, Shading = box.Background,
            Borders = new CellBorders
            {
                Top = box.BorderTop ?? new BorderDef { Val = "nil" }, Bottom = box.BorderBottom ?? new BorderDef { Val = "nil" },
                Left = box.BorderLeft ?? new BorderDef { Val = "nil" }, Right = box.BorderRight ?? new BorderDef { Val = "nil" },
            },
            Margins = new Margins { Top = box.PadTop, Bottom = box.PadBottom, Left = box.PadLeft, Right = box.PadRight },
        };
        var t = new Table { Grid = [w], Align = align is "center" or "right" ? align : "left", Borders = new CellBorders() };
        t.Rows.Add(new TableRow { Height = h, Cells = [cell] });
        return t;
    }

    // ================================================================
    //  VML: картинки, надписи, WordArt (водяные знаки)
    // ================================================================

    static Dictionary<string, string> CssMap(string css) =>
        css.Split(';').Select(p => p.Split(':', 2)).Where(kv => kv.Length == 2)
           .GroupBy(kv => kv[0].Trim().ToLowerInvariant()).ToDictionary(g => g.Key, g => g.Last()[1].Trim());

    static Rgb? VmlColor(string? v)
    {
        if (string.IsNullOrEmpty(v)) return null;
        v = v.Split(' ')[0].Trim();
        if (v.StartsWith('#')) v = v[1..];
        if (v.Length == 3) v = string.Concat(v.Select(c => new string(c, 2)));
        return v.ToLowerInvariant() switch
        {
            "black" => Rgb.Black, "white" => Rgb.White, "silver" => new Rgb(0.75, 0.75, 0.75), "gray" or "grey" => new Rgb(0.5, 0.5, 0.5),
            "red" => new Rgb(1, 0, 0), "blue" => new Rgb(0, 0, 1), "green" => new Rgb(0, 0.5, 0), "yellow" => new Rgb(1, 1, 0),
            _ => Rgb.Parse(v)
        };
    }

    void ParseVml(XElement pict, RunState st, TextStyle style, List<Inline> output)
    {
        foreach (var shape in pict.Elements().Where(e => e.Name.Namespace == V && e.Name.LocalName is "shape" or "rect" or "roundrect" or "oval" or "line"))
        {
            var css = CssMap((string?)shape.Attribute("style") ?? "");
            double w = CssLength((string?)shape.Attribute("style") ?? "", "width"), h = CssLength((string?)shape.Attribute("style") ?? "", "height");
            if (shape.Name.LocalName == "line")
            {
                // v:line from="x,y" to="x,y"
                static (double, double) Pt(string? v)
                {
                    var parts = (v ?? "0,0").Split(',');
                    double X(int i) => i < parts.Length ? CssLength("v:" + parts[i].Trim(), "v") : 0;
                    return (X(0), X(1));
                }
                var (x1, y1) = Pt((string?)shape.Attribute("from")); var (x2, y2) = Pt((string?)shape.Attribute("to"));
                w = Math.Abs(x2 - x1); h = Math.Abs(y2 - y1);
            }
            bool absolute = css.TryGetValue("position", out var pos) && pos == "absolute";

            List<Block>? content = null;
            var box = new BoxStyle();
            TransformResolver? rot = null;

            if (shape.Element(V + "textpath") is { } tp && (string?)tp.Attribute("string") is { Length: > 0 } wordArt)
            {
                // WordArt / водяной знак: текст по размеру фигуры
                var fill = VmlColor((string?)shape.Attribute("fillcolor")) ?? new Rgb(0.75, 0.75, 0.75);
                if (double.TryParse(((string?)shape.Element(V + "fill")?.Attribute("opacity"))?.TrimStart('.') is { } op && op.Length > 0
                        ? "0." + op : "1", NumberStyles.Float, CultureInfo.InvariantCulture, out var opacity) && opacity < 1)
                    fill = new Rgb(1 - (1 - fill.R) * opacity, 1 - (1 - fill.G) * opacity, 1 - (1 - fill.B) * opacity);
                var tpCss = CssMap((string?)tp.Attribute("style") ?? "");
                string font = tpCss.TryGetValue("font-family", out var ff) ? ff.Trim('"', '\'', '&', ';').Replace("quot", "") : "Calibri";
                double fs = Math.Min(h * 0.85, w / Math.Max(1, wordArt.Length * 0.62));
                var ts = new TextStyle { FontAscii = font, FontHAnsi = font, FontEastAsia = font, Size = fs, Color = fill, Bold = true };
                var p = new Paragraph { Align = "center", MarkStyle = ts, LineRule = "auto", SnapToGrid = false };
                p.Inlines.Add(new TextInline(wordArt, ts));
                content = [p];
                box.MinHeight = h;
                box.PadTop = Math.Max(0, (h - fs * 1.2) / 2);
            }
            else if (shape.Element(V + "textbox")?.Element(W + "txbxContent") is { } tx)
            {
                content = [];
                ParseBlockContainer(tx, st.Ctx, content);
                if ((string?)shape.Attribute("filled") is not ("f" or "false"))
                    box.Background = VmlColor((string?)shape.Attribute("fillcolor")) ?? Rgb.White;
                if ((string?)shape.Attribute("stroked") is not ("f" or "false"))
                {
                    var b = new BorderDef { Width = CssLength("w:" + ((string?)shape.Attribute("strokeweight") ?? "0.75pt"), "w") is > 0 and var sw ? sw : 0.75,
                        Color = VmlColor((string?)shape.Attribute("strokecolor")) ?? Rgb.Black };
                    box.BorderTop = box.BorderBottom = box.BorderLeft = box.BorderRight = b;
                }
                box.PadLeft = box.PadRight = 7.2; box.PadTop = box.PadBottom = 3.6;
                box.MinHeight = h;
            }
            else if (shape.Element(V + "imagedata") is { } img && (string?)img.Attribute(R + "id") is { } rid)
            {
                var tmp = new List<Inline>();
                AddImage(rid, w, h, st, style, tmp);
                if (tmp.Count == 0) continue;
                if (!absolute) { output.AddRange(tmp); continue; }
                var p = new Paragraph { MarkStyle = style with { Size = 1 }, SnapToGrid = false };
                p.Inlines.AddRange(tmp);
                content = [p];
            }
            else if (shape.Name.LocalName is "rect" or "roundrect" or "oval" or "line"
                     || (shape.Name.LocalName == "shape" && ((string?)shape.Attribute("type") ?? "").Replace("#", "") is "_x0000_t32" or "_x0000_t202" or "_x0000_t1" or "" && shape.Element(V + "path") == null))
            {
                // Фигура VML без текста
                var sh = new ShapeSpec
                {
                    Geom = shape.Name.LocalName switch
                    {
                        "oval" => "ellipse", "roundrect" => "roundRect", "line" => "line",
                        _ => ((string?)shape.Attribute("type") ?? "").EndsWith("t32") ? "line" : "rect"
                    },
                };
                bool filled = (string?)shape.Attribute("filled") is not ("f" or "false") && sh.Geom != "line";
                if (filled) sh.Fill = VmlColor((string?)shape.Attribute("fillcolor")) ?? Rgb.White;
                if ((string?)shape.Attribute("stroked") is not ("f" or "false"))
                {
                    sh.Stroke = VmlColor((string?)shape.Attribute("strokecolor")) ?? Rgb.Black;
                    double swv = CssLength("w:" + ((string?)shape.Attribute("strokeweight") ?? "0.75pt"), "w");
                    sh.StrokeWidth = swv > 0 ? swv : 0.75;
                    var stroke = shape.Element(V + "stroke");
                    string ds = (string?)stroke?.Attribute("dashstyle") ?? "solid";
                    sh.Dotted = ds.Contains("dot", StringComparison.OrdinalIgnoreCase) && !ds.Contains("dash", StringComparison.OrdinalIgnoreCase);
                    sh.Dashed = !sh.Dotted && ds != "solid";
                    sh.TailArrow = (string?)stroke?.Attribute("endarrow") is { } ea && ea != "none";
                    sh.HeadArrow = (string?)stroke?.Attribute("startarrow") is { } sa && sa != "none";
                }
                var shapeImg = new ImageInline("", Math.Max(w, sh.StrokeWidth), Math.Max(h, sh.StrokeWidth), style) { Shape = sh, FrameW = w, FrameH = h };
                if (!absolute) { output.Add(shapeImg); continue; }
                var p = new Paragraph { MarkStyle = style with { Size = 1 }, SnapToGrid = false };
                p.Inlines.Add(shapeImg);
                content = [p];
            }
            if (content == null) continue;

            if (css.TryGetValue("rotation", out var rotS) && double.TryParse(rotS.Replace("fd", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out var deg))
            {
                if (rotS.EndsWith("fd")) deg /= 65536;
                double rad = deg * Math.PI / 180, cos = Math.Cos(rad), sin = Math.Sin(rad);
                rot = (bw, bh) =>
                {
                    double cx = bw / 2, cy = bh / 2;
                    return new BoxTransform { H = [cos, -sin, cx - cos * cx + sin * cy, sin, cos, cy - sin * cx - cos * cy, 0, 0, 1] };
                };
            }

            string wrapType = (string?)shape.Element(W10 + "wrap")?.Attribute("type") ?? (absolute ? "none" : "inline");
            if (!absolute || wrapType is "square" or "tight" or "through")
            {
                string side = css.TryGetValue("mso-position-horizontal", out var mh) && mh is "right" or "outside" ? "right" : "left";
                (_pending ?? []).Add(new FloatBlock { Side = side, Width = w, Box = box, Blocks = content, Transform = rot, MarginRight = side == "left" ? 9 : 0, MarginLeft = side == "right" ? 9 : 0 });
                continue;
            }
            double Css(string k) => css.TryGetValue(k, out var v) ? CssLength(k + ":" + v, k) : 0;
            css.TryGetValue("mso-position-horizontal", out var ah);
            css.TryGetValue("mso-position-vertical", out var av);
            css.TryGetValue("mso-position-horizontal-relative", out var rh);
            css.TryGetValue("mso-position-vertical-relative", out var rv);
            bool behind = css.TryGetValue("z-index", out var z) && z.StartsWith('-');
            _pending?.Add(new PositionedBlock
            {
                Width = w, Box = box, Blocks = content, Transform = rot, ZIndex = behind ? -1 : 1,
                HRel = rh switch { "page" => "page", "text" or "char" => "flow", _ => rh == null ? "flow" : "margin" },
                VRel = rv switch { "page" => "page", "margin" or "top-margin-area" => "margin", _ => rv == null ? "flow" : "flow" },
                AlignX = ah is "absolute" ? null : ah, AlignY = av is "absolute" ? null : av,
                Left = ah is null or "absolute" ? Css("margin-left") + Css("left") : null,
                Top = av is null or "absolute" ? Css("margin-top") + Css("top") : null,
            });
        }
    }
}

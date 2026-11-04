using System.Globalization;
using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>Сноски, плавающие объекты (картинки, надписи, WordArt), буквицы.</summary>
public sealed partial class DocxReader
{
    static readonly XNamespace WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    static readonly XNamespace PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
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

    static Rgb? DrawingColor(XElement? fill)
    {
        if (fill == null) return null;
        var c = fill.Element(A + "srgbClr");
        if (c != null) return Rgb.Parse((string?)c.Attribute("val"));
        var sys = fill.Element(A + "sysClr");
        if (sys != null) return Rgb.Parse((string?)sys.Attribute("lastClr"));
        if (fill.Element(A + "schemeClr") is { } sc)
            return ((string?)sc.Attribute("val")) switch
            {
                "bg1" or "lt1" => Rgb.White, "tx1" or "dk1" => Rgb.Black,
                "accent1" => Rgb.Parse("4472C4"), "accent2" => Rgb.Parse("ED7D31"), "accent3" => Rgb.Parse("A5A5A5"),
                "accent4" => Rgb.Parse("FFC000"), "accent5" => Rgb.Parse("5B9BD5"), "accent6" => Rgb.Parse("70AD47"),
                _ => null
            };
        return null;
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
        var xfrm = container.Descendants(PIC + "spPr").FirstOrDefault()?.Element(A + "xfrm")
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

    void ParseDrawing(XElement drawing, RunState st, TextStyle style, List<Inline> output)
    {
        foreach (var container in drawing.Elements())
        {
            bool anchor = container.Name == WP + "anchor";
            if (!anchor && container.Name != WP + "inline") continue;
            var extent = container.Element(WP + "extent");
            double w = Emu((string?)extent?.Attribute("cx")), h = Emu((string?)extent?.Attribute("cy"));

            // Содержимое: картинка или надпись
            List<Block>? content = null;
            BoxStyle box = new();
            var blip = container.Descendants(A + "blip").FirstOrDefault();
            var txbx = container.Descendants(WPS + "txbx").FirstOrDefault()?.Element(W + "txbxContent")
                       ?? container.Descendants(W + "txbxContent").FirstOrDefault();
            if (txbx != null)
            {
                content = [];
                ParseBlockContainer(txbx, st.Ctx, content);
                var spPr = container.Descendants(WPS + "spPr").FirstOrDefault();
                if (spPr != null && spPr.Element(A + "noFill") == null) box.Background = DrawingColor(spPr.Element(A + "solidFill"));
                var ln = spPr?.Element(A + "ln");
                if (ln != null && ln.Element(A + "noFill") == null)
                {
                    var lc = DrawingColor(ln.Element(A + "solidFill")) ?? Rgb.Black;
                    double lw = ln.Attribute("w") != null ? Math.Max(0.25, Emu((string?)ln.Attribute("w"))) : 0.75;
                    var b = new BorderDef { Width = lw, Color = lc };
                    box.BorderTop = box.BorderBottom = box.BorderLeft = box.BorderRight = b;
                }
                var bodyPr = container.Descendants(WPS + "bodyPr").FirstOrDefault();
                box.PadLeft = bodyPr?.Attribute("lIns") != null ? Emu((string?)bodyPr.Attribute("lIns")) : 7.2;
                box.PadRight = bodyPr?.Attribute("rIns") != null ? Emu((string?)bodyPr.Attribute("rIns")) : 7.2;
                box.PadTop = bodyPr?.Attribute("tIns") != null ? Emu((string?)bodyPr.Attribute("tIns")) : 3.6;
                box.PadBottom = bodyPr?.Attribute("bIns") != null ? Emu((string?)bodyPr.Attribute("bIns")) : 3.6;
                box.MinHeight = h;
            }
            else if (blip != null)
            {
                string? rid = (string?)blip.Attribute(R + "embed") ?? (string?)blip.Attribute(R + "link");
                if (rid == null) continue;
                var tmp = new List<Inline>();
                AddImage(rid, w, h, st, style, tmp);
                if (tmp.Count == 0) continue;
                ApplyPictureGeometry(container, (ImageInline)tmp[0], !anchor);
                if (!anchor) { output.AddRange(tmp); continue; }
                var p = new Paragraph { MarkStyle = style with { Size = 1 }, LineRule = "auto", SnapToGrid = false };
                p.Inlines.AddRange(tmp);
                content = [p];
            }
            else continue;

            if (!anchor)
            {
                // Надпись в строке — плавающий блок слева (упрощение)
                AddFloat("left", w, h, content, box, 0, 0, 0, 0);
                continue;
            }
            EmitAnchor(container, w, h, content, box);
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
        if (wrap is "wrapSquare" or "wrapTight" or "wrapThrough" && !behind)
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
        foreach (var shape in pict.Elements().Where(e => e.Name.Namespace == V && e.Name.LocalName is "shape" or "rect" or "roundrect" or "oval"))
        {
            var css = CssMap((string?)shape.Attribute("style") ?? "");
            double w = CssLength((string?)shape.Attribute("style") ?? "", "width"), h = CssLength((string?)shape.Attribute("style") ?? "", "height");
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

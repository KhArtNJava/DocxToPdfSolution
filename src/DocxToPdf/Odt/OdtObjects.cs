using System.Globalization;
using System.Xml.Linq;
using DocxToPdf.Docx;
using DocxToPdf.Images;
using DocxToPdf.Pdf;

namespace DocxToPdf.Odt;

public sealed partial class OdtReader
{
    // ================================================================
    //  Таблицы
    // ================================================================

    Table ParseTable(XElement tbl, Ctx ctx)
    {
        var tstyle = FindStyle((string?)tbl.Attribute(Tb + "style-name"), "table", ctx);
        var tp = new Dictionary<string, string>();
        foreach (var s in Chain(tstyle, ctx)) foreach (var (k, v) in s.Table) tp[k] = v;

        // Колонки (с повторами)
        var colWidths = new List<double?>();
        var colRel = new List<double?>();
        var colStyles = new List<OStyle?>();
        foreach (var c in tbl.Descendants().Where(e => e.Name == Tb + "table-column" && Owner(e) == tbl))
        {
            int rep = int.TryParse((string?)c.Attribute(Tb + "number-columns-repeated"), out var r) ? Math.Min(r, 1024) : 1;
            var cs = FindStyle((string?)c.Attribute(Tb + "style-name"), "table-column", ctx);
            double? w = cs != null ? Length(cs.Column.GetValueOrDefault("style:column-width")) : null;
            double? rel = cs?.Column.GetValueOrDefault("style:rel-column-width") is { } rw ? RelWidth(rw) : null;
            var dcs = FindStyle((string?)c.Attribute(Tb + "default-cell-style-name"), "table-cell", ctx);
            for (int i = 0; i < rep; i++) { colWidths.Add(w); colRel.Add(rel); colStyles.Add(dcs); }
        }

        // Строки: заголовочные (table-header-rows), группы, повторы
        var rows = new List<(XElement Row, bool Header)>();
        void CollectRows(XElement parent, bool header)
        {
            foreach (var e in parent.Elements())
            {
                if (e.Name == Tb + "table-row")
                {
                    int rep = int.TryParse((string?)e.Attribute(Tb + "number-rows-repeated"), out var r) ? Math.Min(r, 1000) : 1;
                    for (int i = 0; i < rep; i++) rows.Add((e, header));
                }
                else if (e.Name == Tb + "table-header-rows") CollectRows(e, true);
                else if (e.Name == Tb + "table-rows" || e.Name == Tb + "table-row-group") CollectRows(e, header);
            }
        }
        CollectRows(tbl, false);

        // Сетка: ячейки (с повторами) по строкам
        var grid = new List<List<XElement>>();
        foreach (var (row, _) in rows)
        {
            var cells = new List<XElement>();
            foreach (var c in row.Elements())
            {
                if (c.Name != Tb + "table-cell" && c.Name != Tb + "covered-table-cell") continue;
                int rep = int.TryParse((string?)c.Attribute(Tb + "number-columns-repeated"), out var r) ? Math.Min(r, 1024) : 1;
                for (int i = 0; i < rep; i++) cells.Add(c);
            }
            grid.Add(cells);
        }
        int cols = Math.Max(colWidths.Count, grid.Count > 0 ? grid.Max(g => g.Count) : 0);
        if (cols == 0) cols = 1;
        while (colWidths.Count < cols) { colWidths.Add(null); colRel.Add(null); colStyles.Add(null); }

        // Ширина таблицы и колонок
        double avail = _curProps.PageWidth - _curProps.MarginLeft - _curProps.MarginRight;
        double ml = Len(tp, "fo:margin-left"), mr = Len(tp, "fo:margin-right");
        double? tw = Length(tp.GetValueOrDefault("style:width"));
        if (tp.GetValueOrDefault("style:rel-width") is { } relw && relw.EndsWith('%') && double.TryParse(relw[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var rwp))
            tw = avail * rwp / 100;
        string align = tp.GetValueOrDefault("table:align") ?? "margins";
        if (align == "margins") tw = Math.Max(10, (ctx.TableDepth > 0 ? tw ?? avail : avail) - ml - mr);
        double sumKnown = colWidths.Sum(w => w ?? 0);
        var widths = new double[cols];
        double total = tw ?? (sumKnown > 0 ? sumKnown : avail);
        if (colWidths.All(w => w != null) && sumKnown > 0) for (int i = 0; i < cols; i++) widths[i] = colWidths[i]!.Value * total / sumKnown;
        else if (colRel.All(r => r != null)) { double sr = colRel.Sum(r => r!.Value); for (int i = 0; i < cols; i++) widths[i] = total * colRel[i]!.Value / sr; }
        else
        {
            int unknown = colWidths.Count(w => w == null);
            double rest = Math.Max(0, total - sumKnown);
            for (int i = 0; i < cols; i++) widths[i] = colWidths[i] ?? (unknown > 0 ? rest / unknown : total / cols);
        }

        var table = new Table
        {
            Grid = widths.ToList(),
            Borders = new CellBorders(),
            CellMargins = new Margins { Left = 0, Right = 0, Top = 0, Bottom = 0 },
            Align = align switch { "center" => "center", "right" => "right", _ => "left" },
            Indent = align is "left" or "margins" ? ml : 0,
            SpaceBefore = Math.Max(0, Len(tp, "fo:margin-top")),
            SpaceAfter = Math.Max(0, Len(tp, "fo:margin-bottom")),
            PageBreakBefore = tp.GetValueOrDefault("fo:break-before") == "page" && ctx.TableDepth == 0 && !ctx.HeaderFooter,
        };
        if (tp.GetValueOrDefault("fo:break-after") == "page" && ctx.TableDepth == 0 && !ctx.HeaderFooter) _breakAfterTable = true;
        if (align == "right" && tw != null) table.Indent = 0;
        // Таблица справа налево: колонки в обратном порядке
        if (tp.GetValueOrDefault("style:writing-mode") is "rl-tb" or "rl") table.Rtl = true;
        var tableBg = Color(tp.GetValueOrDefault("fo:background-color"));

        // Строки модели
        var vspan = new int[cols];            // оставшиеся строки объединения по вертикали
        var vspanW = new int[cols];           // ширина объединения (колонок)
        var ictx = new Ctx { HeaderFooter = ctx.HeaderFooter, InNote = ctx.InNote, TableDepth = ctx.TableDepth + 1 };
        for (int r = 0; r < rows.Count; r++)
        {
            var (rowEl, header) = rows[r];
            var rs = FindStyle((string?)rowEl.Attribute(Tb + "style-name"), "table-row", ctx);
            var row = new TableRow { IsHeader = header };
            if (rs != null)
            {
                if (Length(rs.Row.GetValueOrDefault("style:row-height")) is double rh) { row.Height = rh; row.HeightRule = "exact"; }
                else if (Length(rs.Row.GetValueOrDefault("style:min-row-height")) is double mrh) { row.Height = mrh; row.HeightRule = "atLeast"; }
                row.CantSplit = rs.Row.GetValueOrDefault("fo:keep-together") == "always";
            }
            var cells = grid[r];
            var rowDefault = FindStyle((string?)rowEl.Attribute(Tb + "default-cell-style-name"), "table-cell", ctx);
            int c = 0;
            while (c < cols)
            {
                if (vspan[c] > 0)
                {
                    // Продолжение объединения по вертикали
                    int span = Math.Max(1, vspanW[c]);
                    row.Cells.Add(new TableCell { GridSpan = span, VMerge = VMerge.Continue, Borders = new CellBorders() });
                    for (int k = c; k < c + span && k < cols; k++) vspan[k]--;
                    c += span;
                    continue;
                }
                if (c >= cells.Count)
                {
                    row.Cells.Add(new TableCell { GridSpan = 1, Borders = new CellBorders() });
                    c++;
                    continue;
                }
                var ce = cells[c];
                if (ce.Name == Tb + "covered-table-cell")
                {
                    // Покрытая ячейка без объединения выше (нестандартно) — пустая
                    row.Cells.Add(new TableCell { GridSpan = 1, Borders = new CellBorders() });
                    c++;
                    continue;
                }
                int colSpan = int.TryParse((string?)ce.Attribute(Tb + "number-columns-spanned"), out var csn) ? Math.Clamp(csn, 1, cols - c) : 1;
                int rowSpan = int.TryParse((string?)ce.Attribute(Tb + "number-rows-spanned"), out var rsn) ? Math.Max(1, rsn) : 1;
                var cstyle = FindStyle((string?)ce.Attribute(Tb + "style-name"), "table-cell", ctx) ?? rowDefault ?? colStyles[c];
                var cp = new Dictionary<string, string>();
                foreach (var s in Chain(cstyle, ctx)) foreach (var (k, v) in s.Cell) cp[k] = v;
                var cellParaDefaults = cstyle != null ? Chain(cstyle, ctx).SelectMany(s => s.Para).ToList() : [];
                var tc = new TableCell
                {
                    GridSpan = colSpan,
                    VMerge = rowSpan > 1 ? VMerge.Restart : VMerge.None,
                    Shading = Color(cp.GetValueOrDefault("fo:background-color")) ?? tableBg,
                    VAlign = cp.GetValueOrDefault("style:vertical-align") switch { "middle" => "center", "bottom" => "bottom", _ => "top" },
                    Borders = new CellBorders
                    {
                        Top = SideBorder(cp, "top") ?? new BorderDef { Val = "nil" },
                        Bottom = SideBorder(cp, "bottom") ?? new BorderDef { Val = "nil" },
                        Left = SideBorder(cp, "left") ?? new BorderDef { Val = "nil" },
                        Right = SideBorder(cp, "right") ?? new BorderDef { Val = "nil" },
                    },
                    Margins = new Margins { Top = Padding(cp, "top"), Bottom = Padding(cp, "bottom"), Left = Padding(cp, "left"), Right = Padding(cp, "right") },
                };
                string? wm = cp.GetValueOrDefault("style:writing-mode") ?? cp.GetValueOrDefault("loext:writing-mode");
                if (wm is "tb-rl" or "tb") tc.TextDirection = "tbRl";
                else if (wm is "bt-lr") tc.TextDirection = "btLr";
                if (cp.GetValueOrDefault("style:direction") == "ttb") tc.TextDirection = "tbRl";
                if (cp.GetValueOrDefault("style:rotation-angle") == "90") tc.TextDirection = "btLr";
                else if (cp.GetValueOrDefault("style:rotation-angle") == "270") tc.TextDirection = "tbRl";
                // Отступ рамки от текста хранится в padding
                foreach (var b in new[] { tc.Borders.Top, tc.Borders.Bottom, tc.Borders.Left, tc.Borders.Right }) if (b != null) b.Space = 0;
                var saved = _pendingAnchors;
                ParseBlocks(ce, ictx, tc.Blocks, top: false);
                _pendingAnchors = saved;
                if (tc.Blocks.Count == 0) tc.Blocks.Add(new Paragraph { MarkStyle = DefaultTextStyle(ctx) });
                row.Cells.Add(tc);
                if (rowSpan > 1) for (int k = c; k < c + colSpan; k++) { vspan[k] = rowSpan - 1; vspanW[k] = k == c ? colSpan : 0; }
                c += colSpan;
            }
            table.Rows.Add(row);
        }
        if (_breakAfterTable) { _pendingBreak = true; _breakAfterTable = false; }
        return table;
    }

    bool _breakAfterTable;

    /// <summary>Ближайшая таблица-владелец элемента (колонки вложенных таблиц не относятся к внешней).</summary>
    static XElement? Owner(XElement e)
    {
        for (var p = e.Parent; p != null; p = p.Parent) if (p.Name == Tb + "table") return p;
        return null;
    }

    // ================================================================
    //  Рамки, картинки, фигуры
    // ================================================================

    Dictionary<string, string> GraphicProps(XElement el, Ctx ctx)
    {
        var gp = new Dictionary<string, string>();
        if (_defaults.TryGetValue("graphic", out var d)) foreach (var (k, v) in d.Graphic) gp[k] = v;
        var st = FindStyle((string?)el.Attribute(Dr + "style-name"), "graphic", ctx);
        foreach (var s in Chain(st, ctx)) foreach (var (k, v) in s.Graphic) gp[k] = v;
        return gp;
    }

    /// <summary>Объект draw:* (рамка, фигура, группа) — строчный или привязанный (плавающий/абсолютный).</summary>
    void ParseDrawing(XElement el, Ctx ctx, TextStyle style, List<Inline> output, bool blockLevel)
    {
        var n = el.Name;
        if (n == Dr + "a")
        {
            string? href = (string?)el.Attribute(XLink + "href");
            foreach (var c in el.Elements()) ParseDrawing(c, ctx, style with { Link = href ?? style.Link }, output, blockLevel);
            return;
        }
        if (n == Dr + "g")
        {
            // Группа: один объект с вложенными фигурами в их собственных координатах
            var kids = new List<(ImageInline Item, double X, double Y)>();
            foreach (var c in el.Elements())
            {
                if (c.Name.Namespace != Dr) continue;
                var tmp = new List<Inline>();
                var savedP = _pendingAnchors; _pendingAnchors = null;
                var cc = new XElement(c);
                cc.SetAttributeValue(Tx + "anchor-type", "as-char");
                ParseDrawing(cc, ctx, style, tmp, blockLevel: false);
                _pendingAnchors = savedP;
                double cx0 = Length((string?)c.Attribute(Svg + "x")) ?? 0, cy0 = Length((string?)c.Attribute(Svg + "y")) ?? 0;
                if (c.Name == Dr + "line")
                {
                    cx0 = Math.Min(Length((string?)c.Attribute(Svg + "x1")) ?? 0, Length((string?)c.Attribute(Svg + "x2")) ?? 0);
                    cy0 = Math.Min(Length((string?)c.Attribute(Svg + "y1")) ?? 0, Length((string?)c.Attribute(Svg + "y2")) ?? 0);
                }
                foreach (var it in tmp.OfType<ImageInline>()) { it.Width = it.FrameW; kids.Add((it, cx0, cy0)); }
            }
            if (kids.Count == 0) return;
            double gx = kids.Min(k => k.X), gy = kids.Min(k => k.Y);
            double gw = kids.Max(k => k.X + k.Item.FrameW) - gx, gh = kids.Max(k => k.Y + k.Item.FrameH) - gy;
            var group = new ImageInline("", Math.Max(1, gw), Math.Max(1, gh), style)
            {
                FrameW = Math.Max(1, gw), FrameH = Math.Max(1, gh),
                Children = kids.Select(k => (k.Item, k.X - gx, k.Y - gy)).ToList(),
            };
            string gAnchor = (string?)el.Attribute(Tx + "anchor-type") ?? (blockLevel ? "paragraph" : "as-char");
            var ggp = GraphicProps(el, ctx);
            if (gAnchor == "as-char")
            {
                double hm2 = Len(ggp, "fo:margin-left") + Len(ggp, "fo:margin-right");
                group.Width += hm2;
                output.Add(group);
            }
            else EmitAnchored(group, gAnchor, ggp, gx, gy, ctx, output);
            return;
        }
        string anchor = (string?)el.Attribute(Tx + "anchor-type") ?? (blockLevel ? "paragraph" : "as-char");
        var gp = GraphicProps(el, ctx);
        if (gp.GetValueOrDefault("draw:visible") == "false") return;

        // Размер и положение
        double w = Length((string?)el.Attribute(Svg + "width")) ?? 0, h = Length((string?)el.Attribute(Svg + "height")) ?? 0;
        double x = Length((string?)el.Attribute(Svg + "x")) ?? 0, y = Length((string?)el.Attribute(Svg + "y")) ?? 0;
        if (n == Dr + "line")
        {
            double x1 = Length((string?)el.Attribute(Svg + "x1")) ?? 0, y1 = Length((string?)el.Attribute(Svg + "y1")) ?? 0;
            double x2 = Length((string?)el.Attribute(Svg + "x2")) ?? 0, y2 = Length((string?)el.Attribute(Svg + "y2")) ?? 0;
            x = Math.Min(x1, x2); y = Math.Min(y1, y2); w = Math.Abs(x2 - x1); h = Math.Abs(y2 - y1);
        }
        double rot = 0;
        if ((string?)el.Attribute(Dr + "transform") is { } tr) (rot, x, y) = ParseTransform(tr, x, y, w, h);

        ImageInline? item = n == Dr + "frame" ? BuildFrame(el, ctx, style, gp, ref w, ref h) : BuildShape(el, ctx, style, gp, w, h);
        if (item == null) return;
        if (rot != 0) { item.Rotation = rot; }
        if (gp.GetValueOrDefault("style:mirror") is { } mir) { item.FlipH = mir.Contains("horizontal"); item.FlipV = mir.Contains("vertical"); }

        if (anchor is "as-char")
        {
            // Относительно базовой линии: «сверху» — объект стоит на базовой линии; относительно строки/символов — по их верху/низу
            string vposA = gp.GetValueOrDefault("style:vertical-pos") ?? "top";
            string vrelA = gp.GetValueOrDefault("style:vertical-rel") ?? "baseline";
            item.VAlign = vrelA == "baseline"
                ? vposA is "middle" or "center" ? "middle" : null
                : vposA switch { "middle" or "center" => "middle", "top" => "top", "bottom" => "bottom", _ => null };
            if (gp.GetValueOrDefault("style:vertical-pos") == "from-top" && y != 0) item.Style = item.Style with { Position = -y };
            if (item.Rotation != 0 && item.Content == null)
            {
                double a = item.Rotation * Math.PI / 180, cc = Math.Abs(Math.Cos(a)), ss = Math.Abs(Math.Sin(a));
                item.Width = item.FrameW * cc + item.FrameH * ss;
                item.Height = item.FrameW * ss + item.FrameH * cc;
            }
            // Поля рамки слева/справа (LibreOffice учитывает их и у объектов «как символ»): объект — по центру расширенного места
            double hm = Len(gp, "fo:margin-left") + Len(gp, "fo:margin-right");
            if (hm > 0.01 && item.Math == null)
            {
                item.Width += hm;
                if (item.Content == null && item.Shape == null && item.Chart == null && !item.HasGeometry) item.CropL = 1e-9;
            }
            output.Add(item);
            return;
        }
        EmitAnchored(item, anchor, gp, x, y, ctx, output);
    }

    /// <summary>draw:transform «rotate (a) translate (x y)» → угол (по часовой, градусы) и левый верхний угол рамки.</summary>
    static (double Rot, double X, double Y) ParseTransform(string tr, double x, double y, double w, double h)
    {
        double rot = 0, tx = x, ty = y;
        bool hasTranslate = false;
        foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(tr, @"(rotate|translate|skewX|scale)\s*\(([^)]*)\)"))
        {
            var args = m.Groups[2].Value.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);
            if (m.Groups[1].Value == "rotate" && args.Length > 0 && double.TryParse(args[0], NumberStyles.Float, CultureInfo.InvariantCulture, out var a)) rot = -a * 180 / Math.PI;
            else if (m.Groups[1].Value == "translate" && args.Length > 0) { tx = Length(args[0]) ?? 0; ty = args.Length > 1 ? Length(args[1]) ?? 0 : 0; hasTranslate = true; }
        }
        if (hasTranslate && rot != 0)
        {
            // translate задаёт положение повёрнутого левого верхнего угла; переводим к центру и обратно к неповёрнутой рамке
            double ar = -rot * Math.PI / 180;
            double cx = tx + (w / 2) * Math.Cos(ar) + (h / 2) * Math.Sin(ar);
            double cy = ty - (w / 2) * Math.Sin(ar) + (h / 2) * Math.Cos(ar);
            tx = cx - w / 2; ty = cy - h / 2;
        }
        else if (hasTranslate) { }
        rot = ((rot % 360) + 360) % 360;
        return (rot, tx, ty);
    }

    ImageInline? BuildFrame(XElement el, Ctx ctx, TextStyle style, Dictionary<string, string> gp, ref double w, ref double h)
    {
        var box = FrameBox(gp);
        // Ширина/высота в %: style:rel-width, fo:min-height
        if (w <= 0 && (string?)el.Attribute(Fo + "min-width") is { } mw) w = Length(mw) ?? 0;
        double minH = Length((string?)el.Attribute(Fo + "min-height")) ?? 0;
        if (h <= 0) h = minH;
        if ((string?)el.Attribute(St + "rel-width") is { } relW && relW.EndsWith('%') && double.TryParse(relW[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var rwp))
            w = (_curProps.PageWidth - _curProps.MarginLeft - _curProps.MarginRight) * rwp / 100;

        foreach (var c in el.Elements())
        {
            if (c.Name == Dr + "image")
            {
                // Несколько вариантов (SVG + PNG): первый, который удаётся показать
                var bytes = ImageBytes(c);
                if (bytes == null) continue;
                bool meta = Metafile.IsMetafile(bytes) || DocxToPdf.Pdf.Svg.IsSvg(bytes);
                if (!meta && ImageDecoder.Decode(bytes) == null) continue;
                string name = "odt-img-" + (_doc.Media.Count + 1);
                _doc.Media[name] = bytes;
                var img = new ImageInline(name, Math.Max(1, w), Math.Max(1, h), style) { FrameW = Math.Max(1, w), FrameH = Math.Max(1, h) };
                ApplyClip(img, gp, bytes);
                if (box.BorderTop != null || box.Background != null) img.Border = box.BorderTop;
                if (gp.TryGetValue("draw:image-opacity", out var op) && op.EndsWith('%') && double.TryParse(op[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var opv)) img.Opacity = opv / 100;
                if (gp.GetValueOrDefault("draw:color-mode") == "greyscale") img.Grayscale = true;
                return img;
            }
            if (c.Name == Dr + "text-box")
            {
                var blocks = new List<Block>();
                var saved = _pendingAnchors;
                var savedBreak = (_pendingBreak, _breakAfter);
                ParseBlocks(c, new Ctx { HeaderFooter = ctx.HeaderFooter, InNote = ctx.InNote, TableDepth = ctx.TableDepth + 1 }, blocks, top: false);
                _pendingAnchors = saved;
                (_pendingBreak, _breakAfter) = savedBreak;
                double tbMin = Length((string?)c.Attribute(Fo + "min-height")) ?? 0;
                if (h <= 0) h = tbMin;
                if (w <= 0) w = Length((string?)c.Attribute(Fo + "min-width")) ?? 100;
                box.MinHeight = Math.Max(h, tbMin);
                if (gp.GetValueOrDefault("draw:textarea-vertical-align") is "middle" or "bottom") box.MinHeight = h;
                var item = new ImageInline("", Math.Max(1, w), Math.Max(1, h), style) { Content = blocks, Box = box, FrameW = Math.Max(1, w), FrameH = Math.Max(1, h) };
                item.VAnchor = gp.GetValueOrDefault("draw:textarea-vertical-align") switch { "middle" => "ctr", "bottom" => "b", _ => "t" };
                VerticalText(item, gp, w, h);
                if (item.VAnchor != "t" && item.Content is { Count: > 0 } && item.Content[0] is not Table { Rows.Count: 1 } )
                {
                    // Текст по центру/снизу рамки: фон и рамку рисует фигура, текст выравнивается в ней
                    item.Shape = new ShapeSpec { Geom = "rect", Fill = box.Background, Stroke = box.BorderTop?.Color, StrokeWidth = box.BorderTop?.Width ?? 0 };
                    if (box.BorderTop == null) item.Shape.Stroke = null;
                    box.Background = null; box.BorderTop = box.BorderBottom = box.BorderLeft = box.BorderRight = null; box.MinHeight = 0;
                }
                return item;
            }
            if (c.Name == Dr + "object" || c.Name == Dr + "object-ole")
            {
                if (BuildObject(c, style, w, h) is { } obj) return obj;
                // Замещающее изображение объекта
                var rep = el.Elements(Dr + "image").FirstOrDefault();
                if (rep != null) continue;
            }
            if (c.Name == Dr + "plugin" || c.Name == Dr + "applet" || c.Name == Dr + "floating-frame") continue;
        }
        // Только замещающее изображение (ObjectReplacements) после неудачи с объектом
        foreach (var c in el.Elements(Dr + "image"))
        {
            var bytes = ImageBytes(c);
            if (bytes == null) continue;
            if (!Metafile.IsMetafile(bytes) && !DocxToPdf.Pdf.Svg.IsSvg(bytes) && ImageDecoder.Decode(bytes) == null) continue;
            string name = "odt-img-" + (_doc.Media.Count + 1);
            _doc.Media[name] = bytes;
            return new ImageInline(name, Math.Max(1, w), Math.Max(1, h), style) { FrameW = Math.Max(1, w), FrameH = Math.Max(1, h) };
        }
        return null;
    }

    byte[]? ImageBytes(XElement img)
    {
        if (img.Element(Office + "binary-data") is { } bd)
        {
            try { return Convert.FromBase64String(new string(bd.Value.Where(ch => !char.IsWhiteSpace(ch)).ToArray())); } catch { return null; }
        }
        string? href = (string?)img.Attribute(XLink + "href");
        if (string.IsNullOrEmpty(href)) return null;
        if (href.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            int comma = href.IndexOf(',');
            try { return comma > 0 ? Convert.FromBase64String(href[(comma + 1)..]) : null; } catch { return null; }
        }
        var inner = LoadBytes(href);
        if (inner != null) return inner;
        // Внешний файл (относительно документа)
        try
        {
            string path = href.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(href).LocalPath : Uri.UnescapeDataString(href);
            if (path.StartsWith("../")) path = path[3..];   // ссылки в пакете считаются от его корня как «папки»
            if (!Path.IsPathRooted(path) && _baseDir != null) path = Path.Combine(_baseDir, path);
            return File.Exists(path) ? File.ReadAllBytes(path) : null;
        }
        catch { return null; }
    }

    /// <summary>Обрезка fo:clip="rect(top, right, bottom, left)" — в долях исходной картинки.</summary>
    static void ApplyClip(ImageInline img, Dictionary<string, string> gp, byte[] bytes)
    {
        if (!gp.TryGetValue("fo:clip", out var clip) || !clip.StartsWith("rect(")) return;
        var parts = clip[5..].TrimEnd(')').Split([',', ' '], StringSplitOptions.RemoveEmptyEntries).Select(p => Length(p) ?? 0).ToArray();
        if (parts.Length != 4 || parts.All(p => p == 0)) return;
        var info = ImageDecoder.Decode(bytes);
        if (info == null) return;
        // Размер картинки в pt по её разрешению (по умолчанию 96 dpi)
        double iw = info.Width * 0.75, ih = info.Height * 0.75;
        var (dx, dy) = ImageDpi(bytes);
        if (dx > 0) iw = info.Width * 72.0 / dx;
        if (dy > 0) ih = info.Height * 72.0 / dy;
        img.CropT = parts[0] / ih; img.CropR = parts[1] / iw; img.CropB = parts[2] / ih; img.CropL = parts[3] / iw;
    }

    /// <summary>Разрешение картинки: PNG pHYs, JPEG JFIF (APP0); 0 — не задано.</summary>
    static (double X, double Y) ImageDpi(byte[] b)
    {
        if (b.Length > 33 && b[0] == 0x89 && b[1] == 'P')
        {
            int p = 8;
            while (p + 12 <= b.Length)
            {
                int len = b[p] << 24 | b[p + 1] << 16 | b[p + 2] << 8 | b[p + 3];
                string type = System.Text.Encoding.ASCII.GetString(b, p + 4, 4);
                if (type == "pHYs" && len >= 9 && b[p + 16] == 1)
                {
                    double px = (uint)(b[p + 8] << 24 | b[p + 9] << 16 | b[p + 10] << 8 | b[p + 11]);
                    double py = (uint)(b[p + 12] << 24 | b[p + 13] << 16 | b[p + 14] << 8 | b[p + 15]);
                    return (px * 0.0254, py * 0.0254);
                }
                if (type is "IDAT" or "IEND" || len < 0) break;
                p += 12 + len;
            }
        }
        else if (b.Length > 18 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF && b[3] == 0xE0 && b[6] == 'J' && b[7] == 'F')
        {
            int unit = b[13];
            double x = b[14] << 8 | b[15], y = b[16] << 8 | b[17];
            if (unit == 1) return (x, y);
            if (unit == 2) return (x * 2.54, y * 2.54);
        }
        return (0, 0);
    }

    BoxStyle FrameBox(Dictionary<string, string> gp)
    {
        var box = new BoxStyle
        {
            BorderTop = SideBorder(gp, "top"), BorderBottom = SideBorder(gp, "bottom"),
            BorderLeft = SideBorder(gp, "left"), BorderRight = SideBorder(gp, "right"),
            PadTop = Padding(gp, "top"), PadBottom = Padding(gp, "bottom"), PadLeft = Padding(gp, "left"), PadRight = Padding(gp, "right"),
        };
        foreach (var b in new[] { box.BorderTop, box.BorderBottom, box.BorderLeft, box.BorderRight }) if (b != null) b.Space = 0;
        string fill = gp.GetValueOrDefault("draw:fill") ?? (gp.ContainsKey("fo:background-color") ? "bg" : "none");
        if (fill == "solid") box.Background = Color(gp.GetValueOrDefault("draw:fill-color"));
        else if (fill == "bg" || fill == "none" && gp.ContainsKey("fo:background-color")) box.Background = Color(gp.GetValueOrDefault("fo:background-color"));
        return box;
    }

    /// <summary>Привязанный объект: обтекание (плавающий блок) или абсолютное положение.</summary>
    void EmitAnchored(ImageInline item, string anchor, Dictionary<string, string> gp, double x, double y, Ctx ctx, List<Inline> output)
    {
        if (_pendingAnchors == null) { output.Add(item); return; }
        string wrap = gp.GetValueOrDefault("style:wrap") ?? "none";
        string hpos = gp.GetValueOrDefault("style:horizontal-pos") ?? "from-left";
        string hrel = gp.GetValueOrDefault("style:horizontal-rel") ?? (anchor == "page" ? "page" : "paragraph");
        string vpos = gp.GetValueOrDefault("style:vertical-pos") ?? "from-top";
        string vrel = gp.GetValueOrDefault("style:vertical-rel") ?? (anchor == "page" ? "page" : "paragraph");
        bool background = gp.GetValueOrDefault("style:run-through") == "background";
        double mL = Len(gp, "fo:margin-left"), mR = Len(gp, "fo:margin-right"), mT = Len(gp, "fo:margin-top"), mB = Len(gp, "fo:margin-bottom");
        double w = item.Rotation != 0 && item.Content == null ? RotatedW(item) : item.FrameW, h = item.FrameH;
        if (item.Rotation != 0 && item.Content == null) { item.Width = w; item.Height = RotatedH(item); }

        // Содержимое плавающего/абсолютного блока: абзац с объектом
        var p = new Paragraph { MarkStyle = item.Style with { Size = 1 }, LineRule = "auto" };
        p.Inlines.Add(item);
        var content = new List<Block> { p };
        var box = new BoxStyle();
        double contentW = item.Width;

        string? alignX = hpos switch { "left" or "inside" => "left", "center" => "center", "right" or "outside" => "right", _ => null };
        string? alignY = vpos switch { "top" => "top", "middle" or "center" => "center", "bottom" => "bottom", _ => null };
        bool pageV = vrel is "page" or "page-content" || anchor == "page";
        bool pageH = hrel is "page" or "page-content" or "page-start-margin" or "page-end-margin" || anchor == "page";
        double bodyW = _curProps.PageWidth - _curProps.MarginLeft - _curProps.MarginRight;

        if (wrap is "none" && !pageV && !background && anchor is "paragraph" or "char")
        {
            // «Без обтекания»: объект на своём месте, строки текста, задевающие его по вертикали, — ниже объекта.
            // Объект — абсолютный блок от абзаца; занятая полоса — общий «распорный» плавающий блок на всю ширину
            double top = (alignY == null ? y : 0) - mT;
            double bottom = (alignY == null ? y : 0) + h + mB;
            var posNo = new PositionedBlock
            {
                Width = contentW, Box = box, Blocks = content, ZIndex = 1,
                HRel = pageH ? "page" : "flow", VRel = "flow",
                AlignX = alignX, Left = alignX == null ? x : null, Top = alignY == null ? y : 0,
            };
            _pendingAnchors.Add(posNo);
            if (_noWrapSpacer is { } sp0 && _pendingAnchors.Contains(sp0.Float))
            {
                double t0 = Math.Min(sp0.Top, top), b0 = Math.Max(sp0.Bottom, bottom);
                sp0.Float.MarginTop = Math.Max(0, t0);
                ((Paragraph)sp0.Float.Blocks[0]).LineValue = Math.Max(0.1, b0 - Math.Max(0, t0));
                _noWrapSpacer = (sp0.Float, t0, b0);
            }
            else
            {
                var spacer = new Paragraph { MarkStyle = new TextStyle { Size = 1 }, LineRule = "exact", LineValue = Math.Max(0.1, bottom - Math.Max(0, top)), SnapToGrid = false };
                var fl = new FloatBlock { Side = "left", Width = bodyW, MarginTop = Math.Max(0, top), Blocks = [spacer] };
                _pendingAnchors.Insert(0, fl);
                _noWrapSpacer = (fl, top, bottom);
            }
            return;
        }
        if (wrap is not ("none" or "run-through") && !background)
        {
            // Обтекание: сторона — где объект (wrap=left — текст слева, объект справа)
            double xRel = pageH ? x - _curProps.MarginLeft : x;
            string side = wrap switch
            {
                "left" => "right",
                "right" => "left",
                _ => alignX == "right" ? "right" : alignX == "left" ? "left" : alignX == "center" ? "left" : xRel + w / 2 > bodyW / 2 ? "right" : "left",
            };
            double ml = side == "left" ? (alignX == null ? Math.Max(0, xRel) : alignX == "center" ? Math.Max(0, (bodyW - w) / 2) : 0) : mL;
            double mr = side == "left" ? mR : (alignX == null ? Math.Max(0, bodyW - xRel - w) : 0);
            if (pageV)
            {
                var fallback = new PositionedBlock
                {
                    Width = contentW, Box = box, Blocks = content, ZIndex = 1,
                    HRel = pageH ? "page" : "margin", VRel = vrel == "page-content" ? "margin" : "page",
                    AlignX = alignX, AlignY = alignY, Left = alignX == null ? x : null, Top = alignY == null ? y : null,
                };
                _pendingAnchors.Add(new FloatBlock
                {
                    Side = side, Width = contentW, Box = box, Blocks = content,
                    MarginBottom = mB, MarginLeft = ml, MarginRight = mr,
                    AbsVRel = vrel == "page-content" ? "margin" : "page", AbsVAlign = alignY, AbsVOffset = y, AbsFallback = fallback,
                    WrapGap = Math.Max(mL, mR),
                });
                return;
            }
            double mt = alignY == null ? Math.Max(0, y) : 0;
            _pendingAnchors.Add(new FloatBlock
            {
                Side = side, Width = contentW, Box = box, Blocks = content,
                MarginTop = mt, MarginBottom = mB, MarginLeft = ml, MarginRight = mr,
            });
            return;
        }
        // Сквозное обтекание / фон / без обтекания на странице: абсолютное положение
        var pb = new PositionedBlock
        {
            Width = contentW, Box = box, Blocks = content, ZIndex = background ? -1 : 1,
            HRel = pageH ? "page" : hrel is "paragraph" or "paragraph-content" or "char" ? "flow" : "margin",
            VRel = pageV ? (vrel == "page-content" ? "margin" : "page") : "flow",
            AlignX = alignX, AlignY = alignY,
            Left = alignX == null ? x : null, Top = alignY == null ? y : null,
        };
        if (pageH && hrel == "page-content" && alignX == null) { pb.HRel = "margin"; }
        _pendingAnchors.Add(pb);
    }

    (FloatBlock Float, double Top, double Bottom)? _noWrapSpacer;

    static double RotatedW(ImageInline i) { double a = i.Rotation * Math.PI / 180; return i.FrameW * Math.Abs(Math.Cos(a)) + i.FrameH * Math.Abs(Math.Sin(a)); }
    static double RotatedH(ImageInline i) { double a = i.Rotation * Math.PI / 180; return i.FrameW * Math.Abs(Math.Sin(a)) + i.FrameH * Math.Abs(Math.Cos(a)); }

    // ---------------- Фигуры ----------------

    ImageInline? BuildShape(XElement el, Ctx ctx, TextStyle style, Dictionary<string, string> gp, double w, double h)
    {
        var n = el.Name;
        var shape = new ShapeSpec();
        string fill = gp.GetValueOrDefault("draw:fill") ?? "solid";
        if (fill == "solid") shape.Fill = Color(gp.GetValueOrDefault("draw:fill-color")) ?? new Rgb(0x72 / 255.0, 0x9f / 255.0, 0xcf / 255.0);
        else if (fill == "gradient")
        {
            var g = _drawStyles.GetValueOrDefault("gradient:" + (gp.GetValueOrDefault("draw:fill-gradient-name") ?? ""));
            var c1 = Color((string?)g?.Attribute(Dr + "start-color")) ?? Rgb.Black;
            var c2 = Color((string?)g?.Attribute(Dr + "end-color")) ?? Rgb.White;
            string gs = (string?)g?.Attribute(Dr + "style") ?? "linear";
            double ang = AngleDeg((string?)g?.Attribute(Dr + "angle"));
            shape.Fill = c1;
            shape.Gradient = gs == "axial" ? [(0, c2), (0.5, c1), (1, c2)] : [(0, c1), (1, c2)];
            shape.GradientRadial = gs is "radial" or "ellipsoid" or "square" or "rectangular";
            // ODF: угол против часовой от вертикали «сверху вниз»; модель: угол направления градиента (0 — слева направо, по часовой)
            shape.GradientAngle = ((90 - ang) % 360 + 360) % 360;
            if (shape.GradientRadial)
            {
                shape.Gradient = [(0, c2), (1, c1)];
                shape.GradientCx = (double.TryParse(((string?)g?.Attribute(Dr + "cx"))?.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var cx) ? cx : 50) / 100;
                shape.GradientCy = (double.TryParse(((string?)g?.Attribute(Dr + "cy"))?.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var cy) ? cy : 50) / 100;
            }
        }
        else if (fill is "hatch" or "bitmap") shape.Fill = Color(gp.GetValueOrDefault("draw:fill-color")) ?? new Rgb(0.85, 0.85, 0.85);
        string stroke = gp.GetValueOrDefault("draw:stroke") ?? "solid";
        if (stroke != "none")
        {
            shape.Stroke = Color(gp.GetValueOrDefault("svg:stroke-color")) ?? new Rgb(0x34 / 255.0, 0x65 / 255.0, 0xa4 / 255.0);
            shape.StrokeWidth = Math.Max(0.5, Length(gp.GetValueOrDefault("svg:stroke-width")) ?? 0.75);
            if (stroke == "dash")
            {
                var dash = _drawStyles.GetValueOrDefault("stroke-dash:" + (gp.GetValueOrDefault("draw:stroke-dash") ?? ""));
                if ((string?)dash?.Attribute(Dr + "dots1-length") is { } d1 && d1.Contains('%') == false && (Length(d1) ?? 1) < 0.6 * shape.StrokeWidth + 0.5) shape.Dotted = true;
                else shape.Dashed = true;
            }
        }
        shape.HeadArrow = gp.ContainsKey("draw:marker-start") && gp["draw:marker-start"].Length > 0;
        shape.TailArrow = gp.ContainsKey("draw:marker-end") && gp["draw:marker-end"].Length > 0;

        if (n == Dr + "rect") shape.Geom = (string?)el.Attribute(Dr + "corner-radius") is { } cr && (Length(cr) ?? 0) > 0 ? "roundRect" : "rect";
        else if (n == Dr + "ellipse" || n == Dr + "circle") shape.Geom = "ellipse";
        else if (n == Dr + "line" || n == Dr + "connector" || n == Dr + "measure")
        {
            double x1 = Length((string?)el.Attribute(Svg + "x1")) ?? 0, y1 = Length((string?)el.Attribute(Svg + "y1")) ?? 0;
            double x2 = Length((string?)el.Attribute(Svg + "x2")) ?? 0, y2 = Length((string?)el.Attribute(Svg + "y2")) ?? 0;
            shape.Geom = "custom";
            shape.Fill = null;
            double ww = Math.Max(1e-6, Math.Abs(x2 - x1)), hh = Math.Max(1e-6, Math.Abs(y2 - y1));
            double sx = x1 <= x2 ? 0 : 1, sy = y1 <= y2 ? 0 : 1;
            shape.Paths = [new CustomPath { Fill = false, Commands = [('M', [sx, sy]), ('L', [1 - sx, 1 - sy])] }];
            if (w <= 0) w = ww; if (h <= 0) h = hh;
            w = Math.Max(w, 0.1); h = Math.Max(h, 0.1);
        }
        else if (n == Dr + "polygon" || n == Dr + "polyline")
        {
            var vb = ViewBox(el, w, h);
            var pts = ((string?)el.Attribute(Dr + "points") ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Split(',')).Where(a => a.Length == 2)
                .Select(a => (double.Parse(a[0], CultureInfo.InvariantCulture), double.Parse(a[1], CultureInfo.InvariantCulture))).ToList();
            if (pts.Count < 2) return null;
            var cp = new CustomPath { Fill = n == Dr + "polygon" };
            cp.Commands.Add(('M', [(pts[0].Item1 - vb.X) / vb.W, (pts[0].Item2 - vb.Y) / vb.H]));
            foreach (var (px, py) in pts.Skip(1)) cp.Commands.Add(('L', [(px - vb.X) / vb.W, (py - vb.Y) / vb.H]));
            if (n == Dr + "polygon") cp.Commands.Add(('Z', []));
            shape.Geom = "custom"; shape.Paths = [cp];
            if (n == Dr + "polyline") shape.Fill = null;
        }
        else if (n == Dr + "path")
        {
            var vb = ViewBox(el, w, h);
            var paths = SvgPath((string?)el.Attribute(Svg + "d") ?? "", vb);
            if (paths.Count == 0) return null;
            shape.Geom = "custom"; shape.Paths = paths;
        }
        else if (n == Dr + "custom-shape" && el.Element(Dr + "enhanced-geometry") is { } fw && (string?)fw.Attribute(Dr + "text-path") == "true")
        {
            // Fontwork (водяной знак «ЧЕРНОВИК» и т.п.): текст по размеру фигуры, цвет — заливка
            string text = string.Join(" ", el.Elements(Tx + "p").Select(p => p.Value)).Trim();
            if (text.Length == 0) return null;
            var col = shape.Fill ?? new Rgb(0.75, 0.75, 0.75);
            if (gp.TryGetValue("draw:opacity", out var op) && op.EndsWith('%') && double.TryParse(op[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var opv) && opv < 100)
                col = new Rgb(1 - (1 - col.R) * opv / 100, 1 - (1 - col.G) * opv / 100, 1 - (1 - col.B) * opv / 100);
            var tps = FindStyle((string?)el.Elements(Tx + "p").FirstOrDefault()?.Attribute(Tx + "style-name"), "paragraph", ctx);
            var tprops = ParaPropsOf(tps, ctx).Text;
            var baseTs = ToTextStyle(tprops);
            double fs = Math.Min(h * 0.85, w / Math.Max(1, text.Length * 0.62));
            var ts = baseTs with { Size = fs, Color = col, Link = null, Underline = false };
            var para = new Paragraph { Align = "center", MarkStyle = ts, LineRule = "auto", SnapToGrid = false };
            para.Inlines.Add(new TextInline(text, ts));
            return new ImageInline("", Math.Max(1, w), Math.Max(1, h), style)
            {
                Content = [para], Box = new BoxStyle { MinHeight = h, PadTop = Math.Max(0, (h - fs * 1.2) / 2) }, FrameW = Math.Max(1, w), FrameH = Math.Max(1, h),
            };
        }
        else if (n == Dr + "custom-shape")
        {
            var geo = el.Element(Dr + "enhanced-geometry");
            string type = (string?)geo?.Attribute(Dr + "type") ?? "rectangle";
            shape.Geom = type switch
            {
                "rectangle" or "ooxml-rect" or "flowchart-process" => "rect",
                "round-rectangle" or "ooxml-roundRect" or "flowchart-alternate-process" => "roundRect",
                "ellipse" or "circle" or "ooxml-ellipse" or "flowchart-connector" => "ellipse",
                "isosceles-triangle" or "ooxml-triangle" => "triangle",
                "right-triangle" or "ooxml-rtTriangle" => "rtTriangle",
                "diamond" or "ooxml-diamond" or "flowchart-decision" => "diamond",
                "parallelogram" or "ooxml-parallelogram" or "flowchart-data" => "parallelogram",
                "hexagon" or "ooxml-hexagon" or "flowchart-preparation" => "hexagon",
                "octagon" or "ooxml-octagon" => "octagon",
                "pentagon" or "ooxml-pentagon" or "ooxml-homePlate" => "homePlate",
                "star4" or "ooxml-star4" => "star4", "star5" or "ooxml-star5" => "star5", "star6" or "ooxml-star6" => "star6",
                "star8" or "ooxml-star8" => "star8", "star12" or "ooxml-star12" => "star12",
                "right-arrow" or "ooxml-rightArrow" => "rightArrow", "left-arrow" or "ooxml-leftArrow" => "leftArrow",
                "up-arrow" or "ooxml-upArrow" => "upArrow", "down-arrow" or "ooxml-downArrow" => "downArrow",
                "heart" or "ooxml-heart" => "heart", "cross" or "ooxml-plus" => "plus", "chevron" or "ooxml-chevron" => "chevron",
                "flowchart-terminator" or "ooxml-flowChartTerminator" => "flowChartTerminator",
                "ooxml-line" or "line" => "line",
                _ => "",
            };
            if (shape.Geom is "" or "star4" or "star6" or "star8" or "star12" or "chevron" or "heart" or "plus" or "homePlate"
                || shape.Geom.EndsWith("Arrow") || type.StartsWith("ooxml-") && shape.Geom is not ("rect" or "ellipse" or "line"))
            {
                // Контур фигуры по её формулам (draw:enhanced-path); встроенная форма — если контура нет
                var ps = geo != null ? EnhancedGeometry.Build(geo) : null;
                if (ps is { Count: > 0 })
                {
                    shape.Geom = "custom"; shape.Paths = ps;
                    if (shape.Fill == null) foreach (var pth in ps) pth.Fill = false;
                }
                else if (shape.Geom == "") shape.Geom = "rect";
            }
            if (geo != null && (string?)geo.Attribute(Dr + "mirror-horizontal") == "true") _shapeFlipH = true;
            if (geo != null && (string?)geo.Attribute(Dr + "mirror-vertical") == "true") _shapeFlipV = true;
        }
        else return null;

        // Текст внутри фигуры
        List<Block>? content = null;
        if (el.Elements().Any(e => e.Name == Tx + "p" || e.Name == Tx + "list"))
        {
            content = [];
            var saved = _pendingAnchors;
            ParseBlocks(el, new Ctx { HeaderFooter = ctx.HeaderFooter, InNote = ctx.InNote, TableDepth = ctx.TableDepth + 1,
                ShapeText = FindStyle((string?)el.Attribute(Dr + "text-style-name"), "paragraph", ctx) }, content, top: false);
            _pendingAnchors = saved;
            if (content.All(b => b is Paragraph { Inlines.Count: 0 })) content = null;
        }
        var item = new ImageInline("", Math.Max(0.5, w), Math.Max(0.5, h), style) { Shape = shape, FrameW = Math.Max(0.5, w), FrameH = Math.Max(0.5, h), FlipH = _shapeFlipH, FlipV = _shapeFlipV };
        _shapeFlipH = _shapeFlipV = false;
        if (content != null)
        {
            item.Content = content;
            item.Box = new BoxStyle
            {
                PadLeft = Len(gp, "fo:padding-left", 7), PadRight = Len(gp, "fo:padding-right", 7),
                PadTop = Len(gp, "fo:padding-top", 3.6), PadBottom = Len(gp, "fo:padding-bottom", 3.6), MinHeight = 0,
            };
            item.VAnchor = gp.GetValueOrDefault("draw:textarea-vertical-align") switch { "top" => "t", "bottom" => "b", _ => "ctr" };
            VerticalText(item, gp, w, h);
            // По умолчанию текст в фигуре LibreOffice — по центру
            foreach (var b in content.OfType<Paragraph>())
                if (gp.GetValueOrDefault("draw:textarea-horizontal-align") is "center" || b.Align == "left" && b.StyleId is "" or "Standard") b.Align = b.Align == "left" ? "center" : b.Align;
        }
        return item;
    }

    bool _shapeFlipH, _shapeFlipV;

    /// <summary>Вертикальный текст (writing-mode tb-rl / bt-lr): содержимое — в ячейке с поворотом текста.</summary>
    static void VerticalText(ImageInline item, Dictionary<string, string> gp, double w, double h)
    {
        string? wm = gp.GetValueOrDefault("style:writing-mode") ?? gp.GetValueOrDefault("loext:writing-mode");
        if (wm is not ("tb-rl" or "bt-lr" or "tb") || item.Content == null || item.Box == null) return;
        var box = item.Box;
        var cell = new TableCell
        {
            Blocks = item.Content, TextDirection = wm == "bt-lr" ? "btLr" : "tbRl",
            Borders = new CellBorders { Top = new BorderDef { Val = "nil" }, Bottom = new BorderDef { Val = "nil" }, Left = new BorderDef { Val = "nil" }, Right = new BorderDef { Val = "nil" } },
            Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
        };
        double innerW = Math.Max(1, w - box.PadLeft - box.PadRight), innerH = Math.Max(1, h - box.PadTop - box.PadBottom);
        var t = new Table { Grid = [innerW], Borders = new CellBorders(), CellMargins = new Margins { Left = 0, Right = 0, Top = 0, Bottom = 0 } };
        t.Rows.Add(new TableRow { Height = innerH, HeightRule = "exact", Cells = [cell] });
        item.Content = [t];
        item.VAnchor = "t";
    }

    static double AngleDeg(string? v)
    {
        if (string.IsNullOrEmpty(v)) return 0;
        if (v.EndsWith("deg")) return double.TryParse(v[..^3], NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? d : 0;
        if (v.EndsWith("rad")) return double.TryParse(v[..^3], NumberStyles.Float, CultureInfo.InvariantCulture, out var r) ? r * 180 / Math.PI : 0;
        if (v.EndsWith("grad")) return double.TryParse(v[..^4], NumberStyles.Float, CultureInfo.InvariantCulture, out var g) ? g * 0.9 : 0;
        return double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var t) ? t / 10 : 0;   // десятые доли градуса
    }

    static (double X, double Y, double W, double H) ViewBox(XElement el, double w, double h)
    {
        var p = ((string?)el.Attribute(Svg + "viewBox") ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(t => double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0).ToArray();
        return p.Length == 4 && p[2] > 0 && p[3] > 0 ? (p[0], p[1], p[2], p[3]) : (0, 0, Math.Max(1, w), Math.Max(1, h));
    }

    /// <summary>Контур SVG (M L H V C S Q T Z, относительные команды) → нормализованные пути.</summary>
    static List<CustomPath> SvgPath(string d, (double X, double Y, double W, double H) vb)
    {
        var res = new List<CustomPath>();
        var toks = System.Text.RegularExpressions.Regex.Matches(d, @"[A-Za-z]|-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?").Select(m => m.Value).ToList();
        int i = 0;
        double cx = 0, cy = 0, sx = 0, sy = 0, lcx = 0, lcy = 0;
        char cmd = 'M';
        CustomPath? cur = null;
        double Num() => i < toks.Count && !char.IsLetter(toks[i][0]) ? double.Parse(toks[i++], CultureInfo.InvariantCulture) : 0;
        bool HasNum() => i < toks.Count && !char.IsLetter(toks[i][0]);
        double[] N(params double[] pts) { var r = new double[pts.Length]; for (int k = 0; k < pts.Length; k += 2) { r[k] = (pts[k] - vb.X) / vb.W; r[k + 1] = (pts[k + 1] - vb.Y) / vb.H; } return r; }
        while (i < toks.Count)
        {
            if (char.IsLetter(toks[i][0])) cmd = toks[i++][0];
            bool rel = char.IsLower(cmd);
            double ox = rel ? cx : 0, oy = rel ? cy : 0;
            switch (char.ToUpperInvariant(cmd))
            {
                case 'M':
                    cx = ox + Num(); cy = oy + Num(); sx = cx; sy = cy;
                    cur = new CustomPath(); res.Add(cur);
                    cur.Commands.Add(('M', N(cx, cy)));
                    cmd = rel ? 'l' : 'L';
                    break;
                case 'L': cx = ox + Num(); cy = oy + Num(); cur?.Commands.Add(('L', N(cx, cy))); break;
                case 'H': cx = ox + Num(); cur?.Commands.Add(('L', N(cx, cy))); break;
                case 'V': cy = oy + Num(); cur?.Commands.Add(('L', N(cx, cy))); break;
                case 'C':
                {
                    double x1 = ox + Num(), y1 = oy + Num(), x2 = ox + Num(), y2 = oy + Num();
                    cx = ox + Num(); cy = oy + Num(); lcx = x2; lcy = y2;
                    cur?.Commands.Add(('C', N(x1, y1, x2, y2, cx, cy)));
                    break;
                }
                case 'S':
                {
                    double x1 = 2 * cx - lcx, y1 = 2 * cy - lcy, x2 = ox + Num(), y2 = oy + Num();
                    cx = ox + Num(); cy = oy + Num(); lcx = x2; lcy = y2;
                    cur?.Commands.Add(('C', N(x1, y1, x2, y2, cx, cy)));
                    break;
                }
                case 'Q':
                {
                    double qx = ox + Num(), qy = oy + Num(), ex = ox + Num(), ey = oy + Num();
                    cur?.Commands.Add(('C', N(cx + 2.0 / 3 * (qx - cx), cy + 2.0 / 3 * (qy - cy), ex + 2.0 / 3 * (qx - ex), ey + 2.0 / 3 * (qy - ey), ex, ey)));
                    cx = ex; cy = ey; lcx = qx; lcy = qy;
                    break;
                }
                case 'A':
                {
                    // Дуга — отрезком (упрощение)
                    Num(); Num(); Num(); Num(); Num();
                    cx = ox + Num(); cy = oy + Num();
                    cur?.Commands.Add(('L', N(cx, cy)));
                    break;
                }
                case 'Z': cur?.Commands.Add(('Z', [])); cx = sx; cy = sy; break;
                default: i++; break;
            }
            if (char.ToUpperInvariant(cmd) == 'Z' && HasNum()) cmd = 'L';
        }
        return res;
    }

    // ================================================================
    //  Встроенные объекты: формулы (MathML) и диаграммы
    // ================================================================

    ImageInline? BuildObject(XElement obj, TextStyle style, double w, double h)
    {
        XElement? root = null;
        if (obj.Element(Mml + "math") is { } inlineMath) root = inlineMath;
        else if (obj.Element(Office + "document") is { } od) root = od;
        else if ((string?)obj.Attribute(XLink + "href") is { } href)
        {
            string dir = href.TrimStart('.', '/').TrimEnd('/');
            root = LoadXml(dir + "/content.xml")?.Root;
        }
        if (root == null) return null;
        var math = root.Name == Mml + "math" ? root : root.Descendants(Mml + "math").FirstOrDefault();
        if (math != null)
        {
            var node = MathFrom(math);
            if (node == null) return null;
            // Размер шрифта формулы — из настроек объекта (по умолчанию 12 pt)
            double size = style.Size;
            if (root.Descendants().FirstOrDefault(e => e.Name.LocalName == "config-item" && (string?)e.Attribute(Cfg + "name") == "BaseFontHeight") is { } bfh
                && double.TryParse(bfh.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var bf)) size = bf;
            else if (math.Attribute("mathsize") != null && Length((string?)math.Attribute("mathsize")) is double ms) size = ms;
            return new ImageInline("", 0, 0, style with { Size = size, Link = null }) { Math = node, MathDisplay = false };
        }
        var chartEl = root.Descendants(Ch + "chart").FirstOrDefault();
        if (chartEl != null)
        {
            var spec = ChartFrom(chartEl);
            if (spec == null) return null;
            return new ImageInline("", Math.Max(1, w), Math.Max(1, h), style) { Chart = spec, FrameW = Math.Max(1, w), FrameH = Math.Max(1, h) };
        }
        return null;
    }

    // ---------------- MathML ----------------

    static MathNode? MathFrom(XElement e)
    {
        var items = MathChildren(e);
        return items.Count == 1 ? items[0] : new MRow(items);
    }

    static List<MathNode> MathChildren(XElement e)
    {
        var list = new List<MathNode>();
        foreach (var c in e.Elements())
        {
            if (c.Name.LocalName is "annotation" or "annotation-xml") continue;
            if (c.Name.LocalName == "semantics") { list.AddRange(MathChildren(c)); continue; }
            var m = MathNodeOf(c);
            if (m != null) list.Add(m);
        }
        // Оператор суммы/интеграла без тела: телом становится следующий элемент строки
        for (int i = 0; i < list.Count; i++)
        {
            if (list[i] is MText { Text: var op } && NaryOps.Contains(op) && i + 1 < list.Count)
                list[i] = new MNary(op, null, null, new MRow([]), false);
            if (list[i] is MNary { E: MRow { Items.Count: 0 } } nary && i + 1 < list.Count)
            {
                nary.E = list[i + 1];
                list.RemoveAt(i + 1);
            }
        }
        return list;
    }

    static MathNode Row(XElement e)
    {
        var l = MathChildren(e);
        return l.Count == 1 ? l[0] : new MRow(l);
    }

    static MathNode Arg(XElement e, int i)
    {
        var els = e.Elements().Where(x => x.Name.LocalName is not ("annotation" or "annotation-xml")).ToList();
        return i < els.Count ? MathNodeOf(els[i]) ?? new MRow([]) : new MRow([]);
    }

    static readonly HashSet<string> NaryOps = ["∑", "∏", "∐", "∫", "∬", "∭", "∮", "∯", "∰", "⋃", "⋂", "⋁", "⋀", "⨁", "⨂"];

    static MathNode? MathNodeOf(XElement e)
    {
        string ln = e.Name.LocalName;
        string variant = (string?)e.Attribute("mathvariant") ?? "";
        bool bold = variant.Contains("bold");
        switch (ln)
        {
            case "math" or "mrow" or "mstyle" or "mpadded" or "mphantom" or "menclose" or "merror" or "semantics":
            {
                var r = Row(e);
                if (ln == "menclose") return new MBox(r, true);
                if (ln == "mphantom") return new MText(" ", true, false);
                return r;
            }
            case "mi":
            {
                string t = e.Value.Trim();
                bool plain = variant is "normal" or "bold" || t.Length > 1 && variant == "";
                if (variant is "italic" or "bold-italic") plain = false;
                return new MText(t, plain, bold);
            }
            case "mn": return new MText(e.Value.Trim(), true, bold);
            case "mtext" or "ms": return new MText(e.Value, true, bold);
            case "mspace": return new MText(" ", true, false);
            case "mo": return new MText(e.Value.Trim(), true, bold);
            case "mfrac":
                return new MFrac(Arg(e, 0), Arg(e, 1), (string?)e.Attribute("linethickness") is "0" or "0pt" or "0px" ? "noBar" : (string?)e.Attribute("bevelled") == "true" ? "skw" : "bar");
            case "msqrt": return new MRad(null, Row(e));
            case "mroot": return new MRad(Arg(e, 1), Arg(e, 0));
            case "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover":
            {
                var b = Arg(e, 0);
                MathNode? sub = ln is "msub" or "msubsup" or "munder" or "munderover" ? Arg(e, 1) : null;
                MathNode? sup = ln is "msup" ? Arg(e, 1) : ln is "msubsup" or "munderover" ? Arg(e, 2) : ln is "mover" ? Arg(e, 1) : null;
                bool undOvr = ln is "munder" or "mover" or "munderover";
                if (b is MText { Text: var op } && NaryOps.Contains(op))
                {
                    // Оператор с пределами: тело — следующий элемент строки (подставляется при сборке строки)
                    return new MNary(op, sub, sup, new MRow([]), undOvr || op is "∑" or "∏" or "∐" or "⋃" or "⋂");
                }
                if (ln == "mover" && sup is MText { Text: var acc } && acc.Length == 1 && acc[0] is '¯' or '‾' or '_' or '―') return new MBar(true, b);
                if (ln == "munder" && sub is MText { Text: var uacc } && uacc.Length == 1 && uacc[0] is '_' or '‾' or '¯') return new MBar(false, b);
                if (ln == "mover" && sup is MText { Text: var acc2 } && acc2.Length == 1 && "^˜~→⃗˙¨ˇ´`˘⌢".Contains(acc2[0])) return new MAcc(acc2, b);
                if (ln == "mover" && sup is MText { Text: var gc } && gc.Length == 1 && gc[0] is '⏞' or '︷') return new MGroupChr(gc, true, b);
                if (ln == "munder" && sub is MText { Text: var gc2 } && gc2.Length == 1 && gc2[0] is '⏟' or '︸') return new MGroupChr(gc2, false, b);
                if (undOvr && b is MText { Text: "lim" or "max" or "min" or "sup" or "inf" } lt) return new MLim(new MText(lt.Text, true, false), (sub ?? sup)!, sup != null && sub == null);
                if (undOvr) return new MLim(b, (sub ?? sup)!, sup != null && sub == null);
                return new MScript(b, sub, sup, false);
            }
            case "mmultiscripts":
            {
                var els = e.Elements().ToList();
                var b = els.Count > 0 ? MathNodeOf(els[0]) ?? new MRow([]) : new MRow([]);
                MathNode? sub = els.Count > 1 && els[1].Name.LocalName != "none" ? MathNodeOf(els[1]) : null;
                MathNode? sup = els.Count > 2 && els[2].Name.LocalName != "none" ? MathNodeOf(els[2]) : null;
                return new MScript(b, sub, sup, false);
            }
            case "mfenced":
            {
                string open = (string?)e.Attribute("open") ?? "(", close = (string?)e.Attribute("close") ?? ")", sep = (string?)e.Attribute("separators") ?? ",";
                return new MDelim(open, close, sep.Length > 0 ? sep[..1] : "", MathChildren(e));
            }
            case "mtable":
            {
                var rows = e.Elements().Where(r => r.Name.LocalName is "mtr" or "mlabeledtr")
                    .Select(r => r.Elements().Where(c => c.Name.LocalName == "mtd").Select(c => Row(c)).ToList()).ToList();
                if (rows.Count > 0 && rows.All(r => r.Count == 1)) return new MEqArr(rows.Select(r => r[0]).ToList());
                return new MMatrix(rows);
            }
            case "none": return null;
            default: return e.HasElements ? Row(e) : new MText(e.Value, true, false);
        }
    }

    // ---------------- Диаграммы ----------------

    static ChartSpec? ChartFrom(XElement chart)
    {
        string cls = ((string?)chart.Attribute(Ch + "class") ?? "chart:bar").Replace("chart:", "");
        var plot = chart.Element(Ch + "plot-area");
        var table = chart.Descendants(Tb + "table").FirstOrDefault();
        if (plot == null || table == null) return null;
        // Локальная таблица данных: первая строка — имена рядов, первая колонка — категории
        var rows = new List<List<string>>();
        foreach (var r in table.Descendants(Tb + "table-row"))
        {
            var cells = new List<string>();
            foreach (var c in r.Elements(Tb + "table-cell"))
            {
                string v = (string?)c.Attribute(Office + "value") ?? string.Concat(c.Elements(Tx + "p").Select(p => p.Value));
                int rep = int.TryParse((string?)c.Attribute(Tb + "number-columns-repeated"), out var rr) ? rr : 1;
                for (int i = 0; i < rep; i++) cells.Add(v);
            }
            rows.Add(cells);
        }
        if (rows.Count < 2) return null;
        var spec = new ChartSpec
        {
            Title = chart.Element(Ch + "title")?.Value.Trim() is { Length: > 0 } t ? t : null,
            Legend = chart.Element(Ch + "legend") != null,
            LegendPos = ((string?)chart.Element(Ch + "legend")?.Attribute(Ch + "legend-position")) switch { "bottom" => "b", "top" => "t", "start" => "l", _ => "r" },
        };
        // Тип и группировка (по стилю plot-area: chart:stacked / percentage / vertical)
        var style = chart.Document?.Root?.Descendants(St + "style").FirstOrDefault(s => (string?)s.Attribute(St + "name") == (string?)plot.Attribute(Ch + "style-name"));
        var cp = style?.Element(St + "chart-properties");
        bool stacked = (string?)cp?.Attribute(Ch + "stacked") == "true", percent = (string?)cp?.Attribute(Ch + "percentage") == "true";
        bool horizontal = (string?)cp?.Attribute(Ch + "vertical") == "true";
        spec.ThreeD = (string?)cp?.Attribute(Ch + "three-dimensional") == "true" || (string?)plot.Attribute(Dr + "projection") != null && (string?)cp?.Attribute(Ch + "three-dimensional") != "false" && plot.Attribute(Dr + "shade-mode") != null;
        spec.Grouping = percent ? "percentStacked" : stacked ? "stacked" : "clustered";
        string baseType = cls switch { "bar" => horizontal ? "bar" : "col", "line" => "line", "area" => "area", "circle" or "ring" => "pie", "scatter" => "scatter", "radar" or "filled-radar" => "radar", _ => "col" };
        if (cls == "ring") baseType = "doughnut";
        spec.Type = baseType;
        spec.Categories = rows.Skip(1).Select(r => r.Count > 0 ? r[0] : "").ToList();
        int nSeries = rows[0].Count - 1;
        var seriesEls = plot.Elements(Ch + "series").ToList();
        var axisY = plot.Elements(Ch + "axis").FirstOrDefault(a => (string?)a.Attribute(Ch + "dimension") == "y");
        spec.Gridlines = axisY?.Element(Ch + "grid") != null;
        for (int s = 0; s < nSeries; s++)
        {
            var ser = new ChartSeries { Name = rows[0][s + 1], Type = baseType, Grouping = spec.Grouping };
            foreach (var r in rows.Skip(1)) ser.Values.Add(r.Count > s + 1 && double.TryParse(r[s + 1], NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : 0);
            if (s < seriesEls.Count)
            {
                var se = seriesEls[s];
                if ((string?)se.Attribute(Ch + "class") is { } sc)
                    ser.Type = sc.Replace("chart:", "") switch { "line" => "line", "area" => "area", "bar" => baseType is "bar" ? "bar" : "col", _ => ser.Type };
                var sst = chart.Document?.Root?.Descendants(St + "style").FirstOrDefault(x => (string?)x.Attribute(St + "name") == (string?)se.Attribute(Ch + "style-name"));
                var gpe = sst?.Element(St + "graphic-properties");
                ser.Color = Color((string?)gpe?.Attribute(Dr + "fill-color")) ?? (ser.Type == "line" ? Color((string?)gpe?.Attribute(Svg + "stroke-color")) : null);
                if (ser.Type == "line") ser.Color = Color((string?)gpe?.Attribute(Svg + "stroke-color")) ?? ser.Color;
                ser.Secondary = (string?)se.Attribute(Ch + "attached-axis") == "secondary-y";
                if ((string?)sst?.Element(St + "chart-properties")?.Attribute(Ch + "symbol-type") is "none") ser.Marker = "none";
                // Цвета точек круговой диаграммы
                int pi = 0;
                foreach (var dp in se.Elements(Ch + "data-point"))
                {
                    int rep = int.TryParse((string?)dp.Attribute(Ch + "repeated"), out var rr) ? rr : 1;
                    var dps = chart.Document?.Root?.Descendants(St + "style").FirstOrDefault(x => (string?)x.Attribute(St + "name") == (string?)dp.Attribute(Ch + "style-name"));
                    var dc = Color((string?)dps?.Element(St + "graphic-properties")?.Attribute(Dr + "fill-color"));
                    for (int k = 0; k < rep; k++) { if (dc != null) ser.PointColors[pi] = dc.Value; pi++; }
                }
                var dls = sst?.Element(St + "chart-properties");
                if ((string?)dls?.Attribute(Ch + "data-label-number") is "value" or "percentage" || (string?)dls?.Attribute(Ch + "data-label-text") == "true") spec.DataLabels = true;
            }
            spec.Series.Add(ser);
        }
        return spec.Series.Count > 0 ? spec : null;
    }
}

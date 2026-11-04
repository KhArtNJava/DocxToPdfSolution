using System.Globalization;
using DocxToPdf.Docx;
using DocxToPdf.Images;

namespace DocxToPdf.Html;

/// <summary>Режимы раскладки CSS: float, position, flexbox, grid.</summary>
public sealed partial class HtmlReader
{
    readonly List<PositionedBlock> _fixedBlocks = [];

    // =================================================================
    //  Общие помощники
    // =================================================================

    static BoxStyle BoxOf(CssStyle s, bool withBackground = true) => new()
    {
        MinHeight = s.Height == null || s.Height.EndsWith('%') ? 0
            : (CssValues.Length(s.Height, s.FontSize) ?? 0) + (s.BoxSizing == "border-box" ? 0
                : s.PaddingTop + s.PaddingBottom + (s.BorderTop?.ToBorderDef()?.Width ?? 0) + (s.BorderBottom?.ToBorderDef()?.Width ?? 0)),
        Background = withBackground ? s.Background : null,
        BorderTop = s.BorderTop?.ToBorderDef(),
        BorderBottom = s.BorderBottom?.ToBorderDef(),
        BorderLeft = s.BorderLeft?.ToBorderDef(),
        BorderRight = s.BorderRight?.ToBorderDef(),
        PadTop = s.PaddingTop, PadBottom = s.PaddingBottom, PadLeft = s.PaddingLeft, PadRight = s.PaddingRight,
    };

    static double HExtra(CssStyle s) =>
        s.PaddingLeft + s.PaddingRight + (s.BorderLeft?.ToBorderDef()?.Width ?? 0) + (s.BorderRight?.ToBorderDef()?.Width ?? 0);

    /// <summary>Ширина рамки элемента из CSS width (или null).</summary>
    static double? BorderBoxWidth(CssStyle s, double percentBase)
    {
        var w = CssValues.Length(s.Width, s.FontSize, percentBase);
        if (w == null) return null;
        return s.BoxSizing == "border-box" ? w : w + HExtra(s);
    }

    /// <summary>Вёрстка содержимого элемента в отдельный список блоков заданной ширины.</summary>
    List<Block> WalkInto(Node n, CssStyle s, double contentWidth, bool anchor)
    {
        var blocks = new List<Block>();
        var flow = new Flow(blocks);
        var ctx = new BlockCtx { IndL = 0, IndR = 0, Owner = s, Width = Math.Max(5, contentWidth) };
        AnchorBlock? ab = null;
        if (anchor)
        {
            ab = new AnchorBlock
            {
                X = -s.PaddingLeft, Width = contentWidth + s.PaddingLeft + s.PaddingRight, TopOffset = -s.PaddingTop,
                Height = PaddingBoxHeight(s),
            };
            blocks.Add(ab);
            ctx.Anchor = ab;
        }
        double saved = _contentWidth;
        _contentWidth = ctx.Width;
        if (n.Tag is "pre" or "listing" or "xmp") flow.FirstNewlineInPre = true;
        if (n.IsText) AddText(n.Text, s, flow, ctx);
        else if (n.Tag == "img")
        {
            var imgStyle = n.Style!;
            AddImage(n, imgStyle, flow, ctx);
        }
        else WalkChildren(n, flow, ctx);
        Flush(flow);
        if (!n.IsText) ApplyFirstLine(n, s, blocks);
        _contentWidth = saved;
        if (ab != null && ab.Children.Count == 0) blocks.Remove(ab);
        return blocks;
    }

    /// <summary>Размер картинки по атрибутам/CSS/собственному размеру (в пунктах).</summary>
    (double W, double H)? ImageSize(Node img, CssStyle s, double avail)
    {
        var bytes = LoadResource(img.Attr("src"));
        var info = bytes != null ? ImageDecoder.Decode(bytes) : null;
        if (info == null) return null;
        double iw = info.Width * 0.75, ih = info.Height * 0.75;
        double? w = CssValues.Length(s.Width, s.FontSize, avail);
        double? h = s.Height?.EndsWith('%') == true ? null : CssValues.Length(s.Height, s.FontSize);
        if (w != null && h == null) h = w * ih / iw;
        else if (h != null && w == null) w = h * iw / ih;
        return (w ?? iw, h ?? ih);
    }

    /// <summary>Минимальная и предпочтительная ширина рамки элемента (min-content / max-content).</summary>
    (double Min, double Pref) MeasureItem(Node n, CssStyle s, double avail)
    {
        if (n.IsText)
        {
            var tmp = new Node { Tag = "span", Children = [n] };
            var (mn0, pf0) = Measure(tmp, avail);
            return (mn0, pf0);
        }
        double extra = HExtra(s);
        if (n.Tag == "img")
        {
            var sz = ImageSize(n, s, avail);
            double w = (sz?.W ?? 20) + extra;
            return (w, w);
        }
        var (mn, pf) = Measure(n, avail);
        mn += extra; pf += extra;
        if (BorderBoxWidth(s, avail) is double fw) { pf = fw; mn = Math.Min(mn, fw); }
        return (mn, Math.Max(mn, pf));
    }

    // =================================================================
    //  Блок с боковыми рамками / заданной высотой → рамка-ячейка
    // =================================================================

    static bool NeedsBox(Node n, CssStyle s) =>
        s.Display == "block" && n.Tag is not ("html" or "body" or "ul" or "ol")
        && (s.BorderLeft?.ToBorderDef() != null || s.BorderRight?.ToBorderDef() != null
            || (s.Height != null && !s.Height.EndsWith('%') && CssValues.Length(s.Height, s.FontSize) is > 0)
            // фон с внутренними отступами должен заливать и отступы
            || (s.Background != null && (s.PaddingLeft > 0 || s.PaddingRight > 0 || s.PaddingTop > 0 || s.PaddingBottom > 0)));

    /// <summary>Таблица из одной ячейки, изображающая рамку элемента (фон, рамки, padding, height).</summary>
    Table MakeBoxTable(Node n, CssStyle s, double width)
    {
        var cell = new TableCell
        {
            Shading = s.Background,
            Borders = new CellBorders
            {
                Top = s.BorderTop?.ToBorderDef() ?? new BorderDef { Val = "nil" },
                Bottom = s.BorderBottom?.ToBorderDef() ?? new BorderDef { Val = "nil" },
                Left = s.BorderLeft?.ToBorderDef() ?? new BorderDef { Val = "nil" },
                Right = s.BorderRight?.ToBorderDef() ?? new BorderDef { Val = "nil" },
            },
            VAlign = "top",
        };
        cell.Margins = new Margins
        {
            Top = s.PaddingTop + (cell.Borders.Top.IsNone ? 0 : cell.Borders.Top.Width),
            Bottom = s.PaddingBottom + (cell.Borders.Bottom.IsNone ? 0 : cell.Borders.Bottom.Width),
            Left = s.PaddingLeft + (cell.Borders.Left.IsNone ? 0 : cell.Borders.Left.Width),
            Right = s.PaddingRight + (cell.Borders.Right.IsNone ? 0 : cell.Borders.Right.Width),
        };
        cell.Blocks = WalkInto(n, s, width - (cell.Margins.Left ?? 0) - (cell.Margins.Right ?? 0), anchor: s.Position == "relative");
        var row = new TableRow { CantSplit = s.AvoidBreakInside };
        if (BoxHeight(s) is double h) row.Height = h;
        row.Cells.Add(cell);
        var t = new Table
        {
            Grid = [width], Borders = new CellBorders(),
            CellMargins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
        };
        t.Rows.Add(row);
        return t;
    }

    static double? PaddingBoxHeight(CssStyle s) =>
        s.Height == null || s.Height.EndsWith('%') || CssValues.Length(s.Height, s.FontSize) is not double h ? null
            : s.BoxSizing == "border-box" ? h - (s.BorderTop?.ToBorderDef()?.Width ?? 0) - (s.BorderBottom?.ToBorderDef()?.Width ?? 0)
            : h + s.PaddingTop + s.PaddingBottom;

    /// <summary>Высота рамки из CSS height (с учётом box-sizing), null — авто.</summary>
    static double? BoxHeight(CssStyle s)
    {
        if (s.Height == null || s.Height.EndsWith('%') || CssValues.Length(s.Height, s.FontSize) is not double h) return null;
        return s.BoxSizing == "border-box" ? h
            : h + s.PaddingTop + s.PaddingBottom + (s.BorderTop?.ToBorderDef()?.Width ?? 0) + (s.BorderBottom?.ToBorderDef()?.Width ?? 0);
    }

    void AddBoxTable(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        Flush(f);
        f.PendingMargin = Math.Max(f.PendingMargin, s.MarginTop);
        if (s.PageBreakBefore) f.PendingBreak = true;
        if (s.Clear != null) f.PendingClear = s.Clear;

        double avail = Math.Max(20, _contentWidth - b.IndL - b.IndR - s.MarginLeft - s.MarginRight);
        double width = BorderBoxWidth(s, avail) ?? avail;
        width = Math.Min(width, avail);
        double indent = b.IndL + s.MarginLeft;
        if (s.MarginLeftAuto && s.MarginRightAuto) indent += (avail - width) / 2;
        else if (s.MarginLeftAuto) indent += avail - width;

        var t = MakeBoxTable(n, s, width);
        t.Indent = indent;
        t.SpaceBefore = f.PendingMargin; t.PageBreakBefore = f.PendingBreak; t.Clear = f.PendingClear;
        f.PendingMargin = 0; f.PendingBreak = false; f.PendingClear = null;
        if (s.Position == "relative")
        {
            t.Indent += CssValues.Length(s.Left, s.FontSize, _contentWidth) ?? -(CssValues.Length(s.Right, s.FontSize, _contentWidth) ?? 0);
        }
        f.Target.Add(t);
        f.PendingMargin = Math.Max(f.PendingMargin, s.MarginBottom);
        if (s.PageBreakAfter) f.PendingBreak = true;
    }

    // =================================================================
    //  float
    // =================================================================

    void AddFloat(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        string side = s.Float!;
        s.Float = null;   // внутри блока элемент уже не плавающий
        double avail = Math.Max(20, b.Width);
        double extra = HExtra(s);
        double width;
        if (BorderBoxWidth(s, avail) is double fw) width = fw;
        else
        {
            var (mn, pf) = MeasureItem(n, s, avail);
            width = Math.Min(Math.Max(pf, mn), avail);
        }
        width = Math.Min(width, avail);

        var fb = new FloatBlock
        {
            Side = side,
            Width = width,
            MarginTop = s.MarginTop, MarginBottom = s.MarginBottom,
            MarginLeft = Math.Max(0, s.MarginLeft) + (side == "left" ? b.IndL : 0),
            MarginRight = Math.Max(0, s.MarginRight) + (side == "right" ? b.IndR : 0),
            Box = BoxOf(s),
        };
        // У картинки рамка/отступы уже учтены в блоке — сама картинка без отступов
        if (n.Tag == "img")
        {
            s.MarginTop = s.MarginBottom = s.MarginLeft = s.MarginRight = 0;
            s.Display = "inline";
            s.Width = (width - extra).ToString(CultureInfo.InvariantCulture) + "pt";
        }
        fb.Transform = TransformOf(s);
        if (n.Tag == "img") { s.Transform = s.Translate = s.Rotate = s.Scale = null; }
        fb.Blocks = WalkInto(n, s, width - extra, anchor: s.Position == "relative");
        f.Target.Add(fb);
        f.LastSpace = f.Cur == null || f.LastSpace;
    }

    // =================================================================
    //  position: absolute / fixed
    // =================================================================

    void AddPositioned(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        bool isFixed = s.Position == "fixed";
        string mode = s.Position;
        s.Position = "static";
        s.Float = null;
        double cbW = isFixed ? _opt.PageWidth : b.Anchor?.Width ?? _contentWidth;
        double cbH = isFixed ? _opt.PageHeight : b.Anchor?.Height ?? 700;
        double fs = s.FontSize;
        double? left = CssValues.Length(s.Left, fs, cbW), right = CssValues.Length(s.Right, fs, cbW);
        double? top = CssValues.Length(s.Top, fs, cbH), bottom = CssValues.Length(s.Bottom, fs, cbH);
        double extra = HExtra(s);

        double width;
        if (BorderBoxWidth(s, cbW) is double fw) width = fw;
        else if (left != null && right != null) width = Math.Max(10, cbW - left.Value - right.Value);
        else
        {
            var (mn, pf) = MeasureItem(n, s, cbW);
            width = Math.Min(Math.Max(pf, mn), cbW);
        }
        var transform = TransformOf(s);
        if (n.Tag == "img")
        {
            s.Display = "inline";
            s.MarginTop = s.MarginBottom = s.MarginLeft = s.MarginRight = 0;
            s.Width = (width - extra).ToString(CultureInfo.InvariantCulture) + "pt";
            s.Transform = s.Translate = s.Rotate = s.Scale = null;
        }
        var pb = new PositionedBlock
        {
            Mode = mode,
            Left = left + (left != null ? s.MarginLeft : 0), Right = right + (right != null ? s.MarginRight : 0),
            Top = top + (top != null ? s.MarginTop : 0), Bottom = bottom + (bottom != null ? s.MarginBottom : 0),
            Width = width, ZIndex = s.ZIndex, Box = BoxOf(s), Transform = transform,
        };
        pb.Blocks = WalkInto(n, s, width - extra, anchor: true);

        if (isFixed) _fixedBlocks.Add(pb);
        else if (b.Anchor != null) b.Anchor.Children.Add(pb);
        else
        {
            // Без позиционированного предка — относительно страницы; статическая позиция — по потоку
            if (pb.Left == null && pb.Right == null) pb.Left = b.IndL;
            f.Target.Add(pb);
        }
    }

    // =================================================================
    //  Общая модель "ряд ячеек" для flex/grid
    // =================================================================

    static string VAlignFor(CssStyle container, CssStyle? item)
    {
        string a = item?.AlignSelf ?? container.AlignItems;
        return a switch
        {
            "center" => "center",
            "flex-end" or "end" or "self-end" => "bottom",
            _ => "top"
        };
    }

    TableCell ItemCell(Node item, CssStyle itemStyle, CssStyle container, double width, int span = 1)
    {
        bool isText = item.IsText;
        // Элемент с transform: рамка/фон элемента поворачиваются вместе с содержимым
        if (!isText && item.Tag != "img" && TransformOf(itemStyle) is { } trBox)
        {
            var boxTable = MakeBoxTable(item, itemStyle, width);
            return new TableCell
            {
                GridSpan = span, Shading = container.Background, VAlign = VAlignFor(container, itemStyle),
                Borders = new CellBorders(), Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
                Blocks = [new TransformBlock { Blocks = [boxTable], BoxX = 0, BoxWidth = width, Resolve = trBox }],
            };
        }
        var s = isText ? container : itemStyle;
        var cell = new TableCell
        {
            GridSpan = span,
            Shading = (!isText ? itemStyle.Background : null) ?? container.Background,
            VAlign = VAlignFor(container, isText ? null : itemStyle),
            Borders = new CellBorders(),
            Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
        };
        if (!isText)
        {
            cell.Borders = new CellBorders
            {
                Top = itemStyle.BorderTop?.ToBorderDef(), Bottom = itemStyle.BorderBottom?.ToBorderDef(),
                Left = itemStyle.BorderLeft?.ToBorderDef(), Right = itemStyle.BorderRight?.ToBorderDef(),
            };
            cell.Margins = new Margins
            {
                Top = itemStyle.PaddingTop + (cell.Borders.Top?.Width ?? 0),
                Bottom = itemStyle.PaddingBottom + (cell.Borders.Bottom?.Width ?? 0),
                Left = itemStyle.PaddingLeft + (cell.Borders.Left?.Width ?? 0),
                Right = itemStyle.PaddingRight + (cell.Borders.Right?.Width ?? 0),
            };
        }
        double contentW = width - (cell.Margins.Left ?? 0) - (cell.Margins.Right ?? 0);
        if (!isText && item.Tag == "img")
        {
            itemStyle.Display = "inline";
            itemStyle.MarginLeft = itemStyle.MarginRight = 0;
            if (itemStyle.Width == null || itemStyle.Width.EndsWith('%')) itemStyle.Width = null;
        }
        cell.Blocks = WalkInto(item, s, contentW, anchor: !isText && s.Position == "relative");
        return cell;
    }

    TableCell SpacerCell(CssStyle container, int span = 1) => new()
    {
        GridSpan = span,
        Shading = container.Background,
        Borders = new CellBorders(),
        Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
    };

    Table NewLayoutTable(BlockCtx inner, List<double> widths) => new()
    {
        Grid = widths,
        Indent = inner.IndL,
        Borders = new CellBorders(),
        CellMargins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
    };

    /// <summary>Элементы контейнера: дочерние элементы и непустой текст (анонимные элементы).</summary>
    List<Node> ContainerItems(Node n, Flow f, BlockCtx inner)
    {
        var items = new List<Node>();
        foreach (var c in n.Children)
        {
            if (c.IsText) { if (c.Text.Trim().Length > 0) items.Add(c); continue; }
            if (!c.IsElement || c.Style?.Display == "none") continue;
            if (c.Style!.Position is "absolute" or "fixed") { AddPositioned(c, c.Style, f, inner); continue; }
            c.Style.Float = null;   // float у элементов flex/grid не действует
            items.Add(c);
        }
        return items.OrderBy(x => x.IsText ? 0 : x.Style!.Order).ToList();
    }

    void ApplyContainerBorders(List<Table> tables, CssStyle s)
    {
        if (tables.Count == 0) return;
        foreach (var t in tables)
        {
            t.Borders.Left = s.BorderLeft?.ToBorderDef();
            t.Borders.Right = s.BorderRight?.ToBorderDef();
        }
        tables[0].Borders.Top = s.BorderTop?.ToBorderDef();
        tables[^1].Borders.Bottom = s.BorderBottom?.ToBorderDef();
    }

    // =================================================================
    //  Flexbox
    // =================================================================

    sealed class FlexItem
    {
        public required Node Node;
        public required CssStyle Style;
        public double Base, Min, Size, MarginL, MarginR;
        public bool AutoL, AutoR;
        public double Grow, Shrink;
    }

    void AddFlex(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        var inner = BeginBox(n, s, f, b, out int startIdx);
        double avail = inner.Width;
        double fs = s.FontSize;
        double colGap = CssValues.Length(s.ColumnGap, fs, avail) ?? 0;
        double rowGap = CssValues.Length(s.RowGap, fs, avail) ?? 0;
        var items = ContainerItems(n, f, inner);

        if (s.FlexDirection.StartsWith("column"))
        {
            // Колонка — обычный вертикальный поток с промежутками
            if (s.FlexDirection == "column-reverse") items.Reverse();
            bool first = true;
            foreach (var it in items)
            {
                if (!first) f.PendingMargin += rowGap;
                first = false;
                if (it.IsText) { AddText(it.Text, s, f, inner); Flush(f); continue; }
                var st = it.Style!;
                if (st.Display == "inline") st.Display = "block";
                if (s.AlignItems == "center" && st.Width != null) st.MarginLeftAuto = st.MarginRightAuto = true;
                Walk(it, f, inner);
                Flush(f);
            }
            EndBox(n, s, f, startIdx);
            return;
        }
        if (s.FlexDirection == "row-reverse") items.Reverse();

        // Базовые размеры
        var flex = new List<FlexItem>();
        foreach (var it in items)
        {
            var st = it.IsText ? s : it.Style!;
            var (mn, pf) = MeasureItem(it, st, avail);
            double basis = pf;
            if (!it.IsText)
            {
                double? fb = st.FlexBasis is null or "auto" or "content" ? null : CssValues.Length(st.FlexBasis, st.FontSize, avail);
                if (fb != null) basis = st.BoxSizing == "border-box" ? fb.Value : fb.Value + HExtra(st);
                else if (BorderBoxWidth(st, avail) is double w) basis = w;
            }
            flex.Add(new FlexItem
            {
                Node = it, Style = st, Base = basis, Min = Math.Min(mn, Math.Max(basis, mn)),
                MarginL = it.IsText ? 0 : Math.Max(0, st.MarginLeft), MarginR = it.IsText ? 0 : Math.Max(0, st.MarginRight),
                AutoL = !it.IsText && st.MarginLeftAuto, AutoR = !it.IsText && st.MarginRightAuto,
                Grow = it.IsText ? 0 : st.FlexGrow, Shrink = it.IsText ? 1 : st.FlexShrink,
            });
        }

        // Разбиение на строки (flex-wrap)
        var lines = new List<List<FlexItem>>();
        var curLine = new List<FlexItem>();
        double used = 0;
        foreach (var it in flex)
        {
            double need = it.Base + it.MarginL + it.MarginR + (curLine.Count > 0 ? colGap : 0);
            if (s.FlexWrap && curLine.Count > 0 && used + need > avail + 0.5)
            {
                lines.Add(curLine); curLine = []; used = 0;
                need -= colGap;
            }
            curLine.Add(it);
            used += need;
        }
        if (curLine.Count > 0) lines.Add(curLine);

        var tables = new List<Table>();
        bool firstLine = true;
        foreach (var line in lines)
        {
            double gaps = colGap * (line.Count - 1);
            double margins = line.Sum(i => i.MarginL + i.MarginR);
            foreach (var i in line) i.Size = i.Base;
            double free = avail - gaps - margins - line.Sum(i => i.Size);

            if (free > 0 && line.Any(i => i.Grow > 0))
            {
                double g = line.Sum(i => i.Grow);
                foreach (var i in line) i.Size += free * i.Grow / g;
                free = 0;
            }
            else if (free < 0)
            {
                // Сжатие пропорционально flex-shrink * базовому размеру, но не меньше min-content
                for (int iter = 0; iter < 4 && free < -0.01; iter++)
                {
                    var shrinkable = line.Where(i => i.Shrink > 0 && i.Size > i.Min + 0.01).ToList();
                    double weight = shrinkable.Sum(i => i.Shrink * i.Base);
                    if (weight <= 0) break;
                    double deficit = -free;
                    foreach (var i in shrinkable)
                    {
                        double cut = Math.Min(deficit * i.Shrink * i.Base / weight, i.Size - i.Min);
                        i.Size -= cut;
                        free += cut;
                    }
                }
                if (free < -0.01)
                {
                    // Даже min-content не помещается — масштабируем
                    double total = line.Sum(i => i.Size);
                    double k = Math.Max(0.1, (avail - gaps - margins) / total);
                    foreach (var i in line) i.Size *= k;
                }
                free = 0;
            }

            // justify-content и auto-отступы → промежутки между ячейками
            int k2 = line.Count;
            var before = new double[k2 + 1];   // before[i] — промежуток перед i-м элементом; before[k] — в конце
            for (int i = 1; i < k2; i++) before[i] = colGap;
            int autos = line.Count(i => i.AutoL) + line.Count(i => i.AutoR);
            if (free > 0 && autos > 0)
            {
                double share = free / autos;
                for (int i = 0; i < k2; i++)
                {
                    if (line[i].AutoL) before[i] += share;
                    if (line[i].AutoR) before[i + 1] += share;
                }
            }
            else if (free > 0)
            {
                switch (s.JustifyContent)
                {
                    case "flex-end" or "end" or "right": before[0] += free; break;
                    case "center": before[0] += free / 2; before[k2] += free / 2; break;
                    case "space-between":
                        if (k2 > 1) for (int i = 1; i < k2; i++) before[i] += free / (k2 - 1);
                        else before[k2] += free;
                        break;
                    case "space-around":
                        for (int i = 0; i < k2; i++) { before[i] += free / k2 / 2; before[i + 1] += free / k2 / 2; }
                        break;
                    case "space-evenly":
                        for (int i = 0; i <= k2; i++) before[i] += free / (k2 + 1);
                        break;
                    default: before[k2] += free; break;
                }
            }

            // Колонки таблицы: [промежуток] элемент [промежуток] …
            var widths = new List<double>();
            var row = new TableRow { CantSplit = s.AvoidBreakInside };
            void Spacer(double w)
            {
                if (w < 0.05) return;
                widths.Add(w);
                row.Cells.Add(SpacerCell(s));
            }
            for (int i = 0; i < k2; i++)
            {
                Spacer(before[i] + line[i].MarginL + (i > 0 ? line[i - 1].MarginR : 0));
                widths.Add(Math.Max(1, line[i].Size));
                row.Cells.Add(ItemCell(line[i].Node, line[i].Style, s, Math.Max(1, line[i].Size)));
            }
            Spacer(before[k2] + line[^1].MarginR);
            foreach (var it in line)
                if (!it.Node.IsText && BoxHeight(it.Style) is double ih) row.Height = Math.Max(row.Height, ih);
            if (CssValues.Length(s.Height, fs) is double h && lines.Count == 1) row.Height = h - s.PaddingTop - s.PaddingBottom;

            var t = NewLayoutTable(inner, widths);
            t.Rows.Add(row);
            t.SpaceBefore = firstLine ? f.PendingMargin : rowGap;
            if (firstLine) { t.PageBreakBefore = f.PendingBreak; t.Clear = f.PendingClear; }
            firstLine = false;
            tables.Add(t);
        }
        f.PendingMargin = 0; f.PendingBreak = false; f.PendingClear = null;
        ApplyContainerBorders(tables, s);
        f.Target.AddRange(tables);
        EndBox(n, s, f, startIdx);
    }

    // =================================================================
    //  Grid
    // =================================================================

    sealed class Track
    {
        public string Kind = "auto";   // fixed | fr | auto
        public double Value;           // фиксированный размер или доля fr
        public double? MinFixed;       // minmax(min, …)
        public double? MaxFixed;       // minmax(…, max)
        public double Size;
    }

    List<Track> ParseTracks(string? v, double avail, double gap, double fs, out bool autoFit, out bool autoRepeat)
    {
        autoFit = false; autoRepeat = false;
        var result = new List<Track>();
        if (string.IsNullOrWhiteSpace(v) || v == "none") return result;
        // убираем имена линий [a b]
        var sb = new System.Text.StringBuilder();
        int depth = 0;
        foreach (char c in v) { if (c == '[') depth++; else if (c == ']') depth--; else if (depth == 0) sb.Append(c); }

        Track? One(string tok)
        {
            tok = tok.Trim();
            if (tok.StartsWith("minmax(") && tok.EndsWith(')'))
            {
                var parts = tok[7..^1].Split(',');
                if (parts.Length != 2) return null;
                var lo = One(parts[0]); var hi = One(parts[1]);
                if (lo == null || hi == null) return null;
                var t = hi.Kind == "fr" ? new Track { Kind = "fr", Value = hi.Value } : hi.Kind == "fixed" ? new Track { Kind = "auto", MaxFixed = hi.Value } : new Track { Kind = "auto" };
                if (lo.Kind == "fixed") t.MinFixed = lo.Value;
                return t;
            }
            if (tok.StartsWith("fit-content(") && tok.EndsWith(')'))
                return new Track { Kind = "auto", MaxFixed = CssValues.Length(tok[12..^1], fs, avail) };
            if (tok.EndsWith("fr") && double.TryParse(tok[..^2], NumberStyles.Float, CultureInfo.InvariantCulture, out var fr))
                return new Track { Kind = "fr", Value = fr };
            if (tok is "auto" or "min-content" or "max-content") return new Track { Kind = "auto" };
            var len = CssValues.Length(tok, fs, avail);
            return len != null ? new Track { Kind = "fixed", Value = len.Value } : null;
        }

        foreach (var tok in CssValues.Tokens(sb.ToString()))
        {
            if (tok.StartsWith("repeat(") && tok.EndsWith(')'))
            {
                string body = tok[7..^1];
                int comma = body.IndexOf(',');
                if (comma < 0) continue;
                string count = body[..comma].Trim();
                var pattern = CssValues.Tokens(body[(comma + 1)..]).Select(One).Where(t => t != null).Cast<Track>().ToList();
                if (pattern.Count == 0) continue;
                int times;
                if (count is "auto-fill" or "auto-fit")
                {
                    autoRepeat = true;
                    autoFit = count == "auto-fit";
                    double unit = pattern.Sum(p => p.Kind == "fixed" ? p.Value : p.MinFixed ?? p.MaxFixed ?? 100);
                    times = Math.Max(1, (int)Math.Floor((avail + gap) / (unit + gap * pattern.Count)));
                }
                else times = Math.Clamp(int.TryParse(count, out var tm) ? tm : 1, 1, 200);
                for (int i = 0; i < times; i++)
                    foreach (var p in pattern) result.Add(new Track { Kind = p.Kind, Value = p.Value, MinFixed = p.MinFixed, MaxFixed = p.MaxFixed });
                continue;
            }
            var t1 = One(tok);
            if (t1 != null) result.Add(t1);
        }
        return result;
    }

    /// <summary>Разбор grid-column/grid-row: начальная линия (0-based) и размах.</summary>
    static (int? Start, int Span) ParsePlacement(string? start, string? end, int lines)
    {
        int? Line(string? v)
        {
            if (v == null || v == "auto" || v.StartsWith("span")) return null;
            if (!int.TryParse(v.Trim(), out var n) || n == 0) return null;
            return n > 0 ? n - 1 : lines + 1 + n;   // -1 — последняя линия
        }
        int SpanOf(string? v) => v != null && v.StartsWith("span") && int.TryParse(v[4..].Trim(), out var sp) ? Math.Max(1, sp) : 0;

        int? s = Line(start), e = Line(end);
        int span = Math.Max(SpanOf(start), SpanOf(end));
        if (s != null && e != null)
        {
            if (e < s) (s, e) = (e, s);
            return (s, Math.Max(1, e.Value - s.Value));
        }
        if (s == null && e != null) return (Math.Max(0, e.Value - Math.Max(1, span)), Math.Max(1, span));
        return (s, Math.Max(1, span));
    }

    sealed class GridItem
    {
        public required Node Node;
        public required CssStyle Style;
        public int Row, Col, RowSpan, ColSpan;
        public double Min, Pref;
    }

    void AddGrid(Node n, CssStyle s, Flow f, BlockCtx b)
    {
        var inner = BeginBox(n, s, f, b, out int startIdx);
        double avail = inner.Width;
        double fs = s.FontSize;
        double colGap = CssValues.Length(s.ColumnGap, fs, avail) ?? 0;
        double rowGap = CssValues.Length(s.RowGap, fs, avail) ?? 0;
        var items = ContainerItems(n, f, inner);
        if (items.Count == 0) { EndBox(n, s, f, startIdx); return; }

        var tracks = ParseTracks(s.GridTemplateColumns, avail, colGap, fs, out bool autoFit, out _);
        int cols = Math.Max(1, tracks.Count);

        // Размещение элементов (row-major авторазмещение)
        var occupied = new HashSet<(int, int)>();
        var gitems = new List<GridItem>();
        int curR = 0, curC = 0;
        foreach (var it in items)
        {
            var st = it.IsText ? s : it.Style!;
            var (cs0, cspan) = it.IsText ? (null, 1) : ParsePlacement(st.GridColumnStart, st.GridColumnEnd, cols);
            var (rs0, rspan) = it.IsText ? (null, 1) : ParsePlacement(st.GridRowStart, st.GridRowEnd, 1000);
            if (cs0 != null && cs0 + cspan > cols) { cols = cs0.Value + cspan; }
            cspan = Math.Min(cspan, Math.Max(cols, cspan));
            if (cspan > cols) cols = cspan;
            bool Free(int r, int c)
            {
                if (c + cspan > cols) return false;
                for (int rr = r; rr < r + rspan; rr++)
                    for (int cc = c; cc < c + cspan; cc++)
                        if (occupied.Contains((rr, cc))) return false;
                return true;
            }
            int row, col;
            if (cs0 != null && rs0 != null) { row = rs0.Value; col = cs0.Value; }
            else if (rs0 != null) { row = rs0.Value; col = 0; while (!Free(row, col) && col < cols) col++; if (col >= cols) col = 0; }
            else if (cs0 != null) { col = cs0.Value; row = 0; while (!Free(row, col)) row++; }
            else
            {
                row = curR; col = curC;
                while (!Free(row, col)) { col++; if (col + cspan > cols) { col = 0; row++; } }
                curR = row; curC = col + cspan;
                if (curC >= cols) { curC = 0; curR++; }
            }
            for (int rr = row; rr < row + rspan; rr++)
                for (int cc = col; cc < col + cspan; cc++) occupied.Add((rr, cc));
            var (mn, pf) = MeasureItem(it, st, avail);
            gitems.Add(new GridItem { Node = it, Style = st, Row = row, Col = col, RowSpan = rspan, ColSpan = cspan, Min = mn, Pref = pf });
        }
        int rows = gitems.Max(g => g.Row + g.RowSpan);

        // auto-fit: пустые дорожки схлопываются
        if (autoFit)
        {
            int usedCols = gitems.Max(g => g.Col + g.ColSpan);
            if (usedCols < cols) cols = usedCols;
        }
        while (tracks.Count < cols) tracks.Add(new Track { Kind = tracks.Count > 0 && s.GridTemplateColumns != null ? "auto" : "fr", Value = 1 });
        if (tracks.Count > cols) tracks.RemoveRange(cols, tracks.Count - cols);

        // Размеры дорожек
        double space = avail - colGap * (cols - 1);
        var cMin = new double[cols];
        var cPref = new double[cols];
        foreach (var g in gitems.Where(g => g.ColSpan == 1))
        {
            cMin[g.Col] = Math.Max(cMin[g.Col], g.Min);
            cPref[g.Col] = Math.Max(cPref[g.Col], g.Pref);
        }
        double fixedSum = 0;
        foreach (var (t, i) in tracks.Select((t, i) => (t, i)))
        {
            switch (t.Kind)
            {
                case "fixed": t.Size = t.Value; break;
                case "auto":
                    double pref = cPref[i];
                    if (t.MaxFixed != null) pref = Math.Min(pref, t.MaxFixed.Value);
                    t.Size = Math.Max(t.MinFixed ?? cMin[i], pref);
                    break;
                case "fr": t.Size = t.MinFixed ?? 0; break;
            }
            if (t.Kind != "fr") fixedSum += t.Size;
        }
        var frTracks = tracks.Where(t => t.Kind == "fr").ToList();
        double remain = space - fixedSum;
        if (frTracks.Count > 0)
        {
            double totalFr = frTracks.Sum(t => t.Value);
            if (totalFr < 1) totalFr = 1;   // сумма fr < 1 занимает лишь часть места
            for (int iter = 0; iter < 3; iter++)
            {
                foreach (var t in frTracks)
                {
                    int i = tracks.IndexOf(t);
                    double minC = Math.Max(t.MinFixed ?? 0, cMin[i]);
                    t.Size = Math.Max(minC, Math.Max(0, remain) * t.Value / totalFr);
                }
                double over = frTracks.Sum(t => t.Size) - Math.Max(0, remain);
                if (over <= 0.5) break;
                // Минимумы не дали распределить поровну — уменьшаем долю остальных
                remain -= over / Math.Max(1, frTracks.Count);
            }
        }
        double sum = tracks.Sum(t => t.Size);
        if (sum > space + 0.5)
        {
            // Не помещается — сжимаем auto-дорожки к их минимуму, затем всё пропорционально
            double over = sum - space;
            var autos = tracks.Select((t, i) => (t, i)).Where(x => x.t.Kind == "auto").ToList();
            double slack = autos.Sum(x => Math.Max(0, x.t.Size - cMin[x.i]));
            if (slack > 0)
                foreach (var (t, i) in autos) t.Size -= Math.Min(over, slack) * Math.Max(0, t.Size - cMin[i]) / slack;
            sum = tracks.Sum(t => t.Size);
            if (sum > space) foreach (var t in tracks) t.Size *= space / sum;
        }

        // Высоты строк из grid-template-rows
        var rowTracks = ParseTracks(s.GridTemplateRows, 700, rowGap, fs, out _, out _);

        // Модель: дорожки чередуются с промежутками (если gap > 0)
        int gc = colGap > 0.01 ? 2 : 1, gr = rowGap > 0.01 ? 2 : 1;
        int mCols = cols * gc - (gc - 1), mRows = rows * gr - (gr - 1);
        var widths = new List<double>();
        for (int c = 0; c < cols; c++) { widths.Add(Math.Max(0.5, tracks[c].Size)); if (gc == 2 && c < cols - 1) widths.Add(colGap); }

        var at = new Dictionary<(int, int), GridItem>();
        foreach (var g in gitems)
            for (int r = g.Row * gr; r <= (g.Row + g.RowSpan - 1) * gr; r++)
                for (int c = g.Col * gc; c <= (g.Col + g.ColSpan - 1) * gc; c++)
                    at[(r, c)] = g;

        var table = NewLayoutTable(inner, widths);
        table.SpaceBefore = f.PendingMargin;
        table.PageBreakBefore = f.PendingBreak;
        table.Clear = f.PendingClear;
        f.PendingMargin = 0; f.PendingBreak = false; f.PendingClear = null;
        for (int mr = 0; mr < mRows; mr++)
        {
            bool gapRow = gr == 2 && mr % 2 == 1;
            var row = new TableRow();
            if (gapRow) { row.Height = rowGap; row.HeightRule = "exact"; }
            else if (mr / gr < rowTracks.Count && rowTracks[mr / gr].Kind == "fixed") row.Height = rowTracks[mr / gr].Value;
            int mc = 0;
            while (mc < mCols)
            {
                if (at.TryGetValue((mr, mc), out var g) && mc == g.Col * gc)
                {
                    int span = (g.Col + g.ColSpan - 1) * gc - g.Col * gc + 1;
                    double w = widths.Skip(mc).Take(span).Sum();
                    TableCell cell;
                    if (mr == g.Row * gr)
                    {
                        cell = ItemCell(g.Node, g.Style, s, w, span);
                        if (g.RowSpan > 1) cell.VMerge = VMerge.Restart;
                    }
                    else
                    {
                        cell = SpacerCell(s, span);
                        cell.VMerge = VMerge.Continue;
                        cell.Shading = g.Node.IsText ? s.Background : g.Style.Background ?? s.Background;
                        if (!g.Node.IsText)
                        {
                            bool lastRow = mr == (g.Row + g.RowSpan - 1) * gr;
                            cell.Borders = new CellBorders
                            {
                                Left = g.Style.BorderLeft?.ToBorderDef(), Right = g.Style.BorderRight?.ToBorderDef(),
                                Bottom = lastRow ? g.Style.BorderBottom?.ToBorderDef() : null,
                            };
                        }
                    }
                    row.Cells.Add(cell);
                    mc += span;
                }
                else
                {
                    row.Cells.Add(SpacerCell(s));
                    mc++;
                }
            }
            table.Rows.Add(row);
        }
        ApplyContainerBorders([table], s);
        f.Target.Add(table);
        EndBox(n, s, f, startIdx);
    }

    // =================================================================
    //  transform
    // =================================================================

    readonly HashSet<Node> _transforming = [];

    static double Angle(string v)
    {
        v = v.Trim();
        double Num(string x) => double.Parse(x, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (v.EndsWith("deg")) return Num(v[..^3]) * Math.PI / 180;
        if (v.EndsWith("grad")) return Num(v[..^4]) * Math.PI / 200;
        if (v.EndsWith("rad")) return Num(v[..^3]);
        if (v.EndsWith("turn")) return Num(v[..^4]) * 2 * Math.PI;
        return Num(v);
    }

    static Transform2D Mul(Transform2D m, Transform2D n) => new(
        m.A * n.A + m.C * n.B, m.B * n.A + m.D * n.B,
        m.A * n.C + m.C * n.D, m.B * n.C + m.D * n.D,
        m.A * n.E + m.C * n.F + m.E, m.B * n.E + m.D * n.F + m.F, 0, 0);

    /// <summary>Строит функцию вычисления матрицы преобразования элемента (null — преобразования нет).</summary>
    static TransformResolver? TransformOf(CssStyle s)
    {
        if (s.Transform == null && s.Translate == null && s.Rotate == null && s.Scale == null) return null;
        // Порядок CSS: translate → rotate → scale → transform
        var fns = new List<(string Name, List<string> Args)>();
        void FromIndividual(string? v, string fn)
        {
            if (v == null) return;
            var t = CssValues.Tokens(v);
            if (fn == "rotate" && t.Count > 1) t = [t[^1]];   // rotate: z 45deg → только угол
            fns.Add((fn, t));
        }
        FromIndividual(s.Translate, "translate");
        FromIndividual(s.Rotate, "rotate");
        FromIndividual(s.Scale, "scale");
        if (s.Transform != null)
        {
            string v = s.Transform;
            int i = 0;
            while (i < v.Length)
            {
                while (i < v.Length && char.IsWhiteSpace(v[i])) i++;
                int st = i;
                while (i < v.Length && v[i] != '(') i++;
                if (i >= v.Length) break;
                string name = v[st..i].Trim();
                int depth = 0, a = i + 1;
                for (; i < v.Length; i++)
                {
                    if (v[i] == '(') depth++;
                    else if (v[i] == ')' && --depth == 0) break;
                }
                string args = v[a..Math.Min(i, v.Length)];
                i++;
                fns.Add((name, args.Split(',').SelectMany(x => CssValues.Tokens(x)).ToList()));
            }
        }
        if (fns.Count == 0) return null;
        double fs = s.FontSize;
        string? origin = s.TransformOrigin;

        return (w, h) =>
        {
            var m = new Transform2D(1, 0, 0, 1, 0, 0, 0, 0);
            double L(string x, double basis) => CssValues.Length(x, fs, basis) ?? 0;
            double N(string x) => double.TryParse(x.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var d) ? (x.EndsWith('%') ? d / 100 : d) : 1;
            foreach (var (name, args) in fns)
            {
                string a0 = args.ElementAtOrDefault(0) ?? "0", a1 = args.ElementAtOrDefault(1) ?? "";
                Transform2D t;
                try
                {
                    t = name switch
                    {
                        "translate" or "translate3d" => new(1, 0, 0, 1, L(a0, w), a1.Length > 0 ? L(a1, h) : 0, 0, 0),
                        "translatex" => new(1, 0, 0, 1, L(a0, w), 0, 0, 0),
                        "translatey" => new(1, 0, 0, 1, 0, L(a0, h), 0, 0),
                        "scale" or "scale3d" => new(N(a0), 0, 0, a1.Length > 0 ? N(a1) : N(a0), 0, 0, 0, 0),
                        "scalex" => new(N(a0), 0, 0, 1, 0, 0, 0, 0),
                        "scaley" => new(1, 0, 0, N(a0), 0, 0, 0, 0),
                        "rotate" or "rotatez" => Rot(Angle(a0)),
                        "skew" => new(1, Math.Tan(a1.Length > 0 ? Angle(a1) : 0), Math.Tan(Angle(a0)), 1, 0, 0, 0, 0),
                        "skewx" => new(1, 0, Math.Tan(Angle(a0)), 1, 0, 0, 0, 0),
                        "skewy" => new(1, Math.Tan(Angle(a0)), 0, 1, 0, 0, 0, 0),
                        "matrix" when args.Count == 6 => new(N(args[0]), N(args[1]), N(args[2]), N(args[3]), N(args[4]) * 0.75, N(args[5]) * 0.75, 0, 0),
                        _ => new(1, 0, 0, 1, 0, 0, 0, 0)   // 3D-функции (rotateX, perspective …) игнорируются
                    };
                }
                catch { continue; }
                m = Mul(m, t);
            }
            var (ox, oy) = Origin(origin, w, h, fs);
            return m with { Ox = ox, Oy = oy };
        };

        static Transform2D Rot(double a) => new(Math.Cos(a), Math.Sin(a), -Math.Sin(a), Math.Cos(a), 0, 0, 0, 0);
    }

    /// <summary>transform-origin: ключевые слова, длины, проценты (по умолчанию центр).</summary>
    static (double X, double Y) Origin(string? v, double w, double h, double fs)
    {
        double x = w / 2, y = h / 2;
        if (string.IsNullOrWhiteSpace(v)) return (x, y);
        var t = CssValues.Tokens(v);
        for (int i = 0; i < Math.Min(2, t.Count); i++)
        {
            string tok = t[i];
            switch (tok)
            {
                case "left": x = 0; break;
                case "right": x = w; break;
                case "top": y = 0; break;
                case "bottom": y = h; break;
                case "center": if (i == 0) x = w / 2; else y = h / 2; break;
                default:
                    if (i == 0) x = CssValues.Length(tok, fs, w) ?? x;
                    else y = CssValues.Length(tok, fs, h) ?? y;
                    break;
            }
        }
        return (x, y);
    }
}

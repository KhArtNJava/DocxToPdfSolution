using System.Globalization;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;

namespace DocxToPdf.Layout;

/// <summary>Вёрстка блоков (абзацев и таблиц) в колонку заданной ширины.</summary>
public sealed class BlockLayout(LayoutContext ctx)
{
    readonly LayoutContext _ctx = ctx;

    /// <summary>Область, занятая плавающими блоками.</summary>
    public interface IExclusion
    {
        /// <summary>Сколько занято слева и справа в полосе [y0, y1).</summary>
        (double Left, double Right) At(double y0, double y1);
        /// <summary>Ближайшая нижняя граница плавающего блока ниже y.</summary>
        double? NextBottom(double y);
    }

    /// <summary>Плавающий блок в группе; Polygon — контур обтекания (координаты группы), если обтекание «по контуру».</summary>
    sealed record FloatBox(string Side, double X0, double X1, double Top, double Bottom, List<(double X, double Y)>? Polygon = null, double Gap = 0)
    {
        /// <summary>Горизонтальные границы контура в полосе [y0, y1]; null — контур полосу не задевает.</summary>
        public (double Min, double Max)? Extent(double y0, double y1)
        {
            if (Polygon == null) return (X0, X1);
            double min = double.MaxValue, max = double.MinValue;
            int n = Polygon.Count;
            for (int i = 0; i < n; i++)
            {
                var (ax, ay) = Polygon[i]; var (bx, by) = Polygon[(i + 1) % n];
                if (ay >= y0 && ay <= y1) { min = Math.Min(min, ax); max = Math.Max(max, ax); }
                // Пересечения ребра с границами полосы
                foreach (double yy in new[] { y0, y1 })
                    if ((ay - yy) * (by - yy) < 0)
                    {
                        double x = ax + (bx - ax) * (yy - ay) / (by - ay);
                        min = Math.Min(min, x); max = Math.Max(max, x);
                    }
            }
            return min <= max ? (min - Gap, max + Gap) : null;
        }
    }

    /// <summary>Плавающие блоки текущей группы; координаты относительно верха группы.</summary>
    sealed class FloatArea(double width, double offsetY) : IExclusion
    {
        public readonly List<FloatBox> Boxes = [];
        public double OffsetY = offsetY;   // сдвиг: y абзаца → y группы
        readonly double _width = width;

        public (double Left, double Right) At(double y0, double y1)
        {
            double l = 0, r = 0;
            y0 += OffsetY; y1 += OffsetY;
            foreach (var b in Boxes)
            {
                if (b.Top >= y1 - 0.01 || b.Bottom <= y0 + 0.01) continue;
                if (b.Extent(y0, y1) is not { } ext) continue;
                if (b.Side == "left") l = Math.Max(l, ext.Max);
                else r = Math.Max(r, _width - ext.Min);
            }
            return (l, r);
        }

        public double? NextBottom(double y)
        {
            y += OffsetY;
            double? best = null;
            foreach (var b in Boxes)
                if (b.Bottom > y + 0.01 && (best == null || b.Bottom < best)) best = b.Bottom;
            return best - OffsetY;
        }

        public double MaxBottom(string side = "both") =>
            Boxes.Where(b => side == "both" || b.Side == side).Select(b => b.Bottom).DefaultIfEmpty(0).Max();
    }

    /// <summary>Отступ сверху под верхнюю (или межабзацную) границу.</summary>
    static double TopBorderPad(Paragraph p)
    {
        var b = p.GroupPrev ? p.BorderBetween : p.BorderTop;
        return b != null ? b.Width + b.Space : 0;
    }

    static bool SameBorder(BorderDef? a, BorderDef? b) =>
        (a == null && b == null) || (a != null && b != null && a.Val == b.Val && Math.Abs(a.Width - b.Width) < 0.01 && a.Color == b.Color);

    /// <summary>
    /// Соседние абзацы с одинаковыми рамками и отступами образуют группу с общей рамкой (как в Word):
    /// между ними нет нижней/верхней границы (только «между», если задана), боковые идут непрерывно.
    /// </summary>
    static void GroupBorders(IReadOnlyList<Block> blocks)
    {
        Paragraph? prev = null;
        foreach (var b in blocks)
        {
            if (b is not Paragraph p) { prev = null; continue; }
            p.GroupPrev = false; p.GroupNext = false;
            bool hasBox = p.BorderTop != null || p.BorderBottom != null || p.BorderLeft != null || p.BorderRight != null || p.BorderBetween != null;
            if (prev != null && hasBox && SameBorder(prev.BorderTop, p.BorderTop) && SameBorder(prev.BorderBottom, p.BorderBottom)
                && SameBorder(prev.BorderLeft, p.BorderLeft) && SameBorder(prev.BorderRight, p.BorderRight) && SameBorder(prev.BorderBetween, p.BorderBetween)
                && Math.Abs(prev.IndLeft - p.IndLeft) < 0.01 && Math.Abs(prev.IndRight - p.IndRight) < 0.01)
            {
                prev.GroupNext = true;
                p.GroupPrev = true;
            }
            prev = p;
        }
    }

    public List<Frag> LayoutBlocks(IReadOnlyList<Block> blocks, double width)
    {
        GroupBorders(blocks);
        var frags = new List<Frag>();
        Paragraph? prev = null;
        LineFrag? prevLast = null;

        GroupFrag? group = null;
        FloatArea? area = null;
        double gy = 0;   // текущая позиция внутри группы
        int groupStart = -1, curBlock = -1;
        var blockFrags = new Dictionary<int, int>();   // число фрагментов блока вне групп
        int fragsAtStart = -1;
        bool groupTouched = false;
        bool curInGroup = false;          // у текущего блока уже есть строки в группе
        GroupFrag? tailOf = null;         // группа, после которой идут остальные строки её последнего абзаца
        int tailBlock = -1;

        void EndGroup()
        {
            groupTouched = true;
            if (group == null) return;
            group.Height = Math.Max(gy, area!.MaxBottom());
            frags.Add(group);
            // Последний блок группы: текущий, если его строки начались в группе
            int last = curInGroup ? curBlock : curBlock - 1;
            if (groupStart >= 0 && last >= groupStart && group.AbsFloats.Count == 0)
            {
                var src = blocks.Skip(groupStart).Take(last - groupStart + 1).ToList();
                double w0 = width;
                int after = last + 1;
                // Следующие блоки (их фрагменты вне групп) — для обтекания отложенного плавающего блока
                group.Rebuild = room =>
                {
                    var follow = new List<(Block B, int Frags)>();
                    for (int k = after; k < blocks.Count && follow.Count < 8; k++)
                    {
                        if (!blockFrags.TryGetValue(k, out var n) || blocks[k] is not Paragraph) break;
                        follow.Add((blocks[k], n));
                    }
                    return RebuildGroup(src, follow, w0, room);
                };
                if (curInGroup) { tailOf = group; tailBlock = curBlock; }
            }
            group = null; area = null; gy = 0; groupStart = -1;
        }

        void Emit(Frag f)
        {
            if (group != null && gy >= area!.MaxBottom() - 0.01) EndGroup();
            if (group == null)
            {
                if (tailOf != null && tailBlock == curBlock) tailOf.TailFrags++;
                frags.Add(f);
                return;
            }
            curInGroup = true;
            gy += f.SpaceBefore;
            group.Items.Add((f, 0, gy));
            gy += f.Height + f.SpaceAfter;
        }

        void Clear(string? side)
        {
            if (group == null || side == null) return;
            gy = Math.Max(gy, area!.MaxBottom(side == "both" ? "both" : side));
        }

        for (int bi = 0; bi < blocks.Count; bi++)
        {
            if (bi > 0 && fragsAtStart >= 0 && group == null && !groupTouched) blockFrags[bi - 1] = frags.Count - fragsAtStart;
            fragsAtStart = group == null ? frags.Count : -1;
            groupTouched = false;
            var block = blocks[bi];
            if (block is MasonryBlock mb) BuildMasonry(mb);
            curBlock = bi;
            curInGroup = false;
            if (tailBlock != bi) tailOf = null;
            switch (block)
            {
                case Paragraph p:
                {
                    Clear(p.Clear);
                    if (group != null && gy >= area!.MaxBottom() - 0.01) EndGroup();
                    double sb = p.SpaceBefore;
                    bool contextual = prev != null && prev.StyleId == p.StyleId;
                    if (contextual && p.ContextualSpacing) sb = 0;
                    IExclusion? excl = null;
                    if (group != null) { area!.OffsetY = gy + sb; excl = area; }
                    var lines = LayoutParagraph(p, width, excl);
                    if (lines.Count > 0)
                    {
                        lines[0].SpaceBefore = sb;
                        if (contextual && prev!.ContextualSpacing && prevLast != null && group == null) prevLast.SpaceAfter = 0;
                    }
                    foreach (var l in lines) Emit(l);
                    prev = p;
                    prevLast = lines.Count > 0 ? lines[^1] : null;
                    break;
                }
                case Table t:
                {
                    if (t.Clear != null) Clear(t.Clear);
                    if (group != null && gy >= area!.MaxBottom() - 0.01) EndGroup();
                    if (group != null)
                    {
                        // Таблица рядом с плавающим блоком: сужается до свободной ширины и входит в группу
                        area!.OffsetY = 0;
                        var (fl, fr) = area.At(gy, gy + 14);
                        double avail = width - fl - fr;
                        if (avail >= Math.Min(144, width * 0.35))
                        {
                            var rows = LayoutTable(t, avail);
                            if (rows.Count > 0) gy += rows[0].SpaceBefore;
                            foreach (var row in rows) { group.Items.Add((row, fl, gy)); gy += row.Height; }
                            curInGroup = true;
                            if (rows.Count > 0) gy += rows[^1].SpaceAfter;
                            prev = null; prevLast = null;
                            break;
                        }
                        // Слишком узко — под плавающими блоками
                        Clear("both"); EndGroup();
                    }
                    frags.AddRange(LayoutTable(t, width));
                    prev = null; prevLast = null;
                    break;
                }
                case FloatBlock fb when fb.UseFallback && fb.AbsFallback != null:
                    Emit(MakeAbs(fb.AbsFallback, 0, width, relative: false));
                    break;
                case FloatBlock fb:
                {
                    if (fb.DropLines > 0) PrepareDropCap(fb, bi + 1 < blocks.Count ? blocks[bi + 1] as Paragraph : null, width);
                    if (fb.Width <= 0) fb.Width = MeasureContentWidth(fb.Blocks) + fb.Box.HorizontalExtra;
                    double w = fb.Width + fb.MarginLeft + fb.MarginRight;
                    var box = Transformed(BuildBox(fb.Blocks, fb.Width, fb.Box), fb.Transform);
                    double h = box.Height + fb.MarginTop + fb.MarginBottom;
                    // Выше страницы — тоже плавающий: группа и сам блок разрываются между страницами
                    if (w > width + 0.5)
                    {
                        // Шире колонки — выводим в потоке
                        if (group != null) { Clear("both"); EndGroup(); }
                        frags.Add(box);
                        break;
                    }
                    if (group == null) { group = new GroupFrag(); area = new FloatArea(width, 0); gy = 0; groupStart = bi; }
                    groupTouched = true;
                    curInGroup = true;
                    double top = gy;
                    if (bi + 1 < blocks.Count && blocks[bi + 1] is Paragraph np && np.Clear == null) top += np.SpaceBefore;
                    double x;
                    while (true)
                    {
                        area!.OffsetY = 0;
                        var (l, r) = area.At(top, top + h);
                        if (l + w + r <= width + 0.5) { x = fb.Side == "right" ? width - r - w : l; break; }
                        var nb = area.NextBottom(top);
                        if (nb == null) { x = 0; break; }
                        top = nb.Value;
                    }
                    group.Items.Add((box, x + fb.MarginLeft, top + fb.MarginTop));
                    // Контур обтекания (доли рамки) → координаты группы
                    var poly = fb.WrapPolygon?.Select(pt => (x + fb.MarginLeft + pt.X * fb.Width, top + fb.MarginTop + pt.Y * box.Height)).ToList();
                    // Объект, сдвинутый к позиции на странице: промежуток над ним не занят (текст идёт во всю ширину)
                    double boxTop = fb.AbsVRel != null ? top + fb.MarginTop : top;
                    area.Boxes.Add(new FloatBox(fb.Side, x, x + w, boxTop, top + h, poly, fb.WrapGap));
                    if (fb.AbsVRel != null) group.AbsFloats.Add((fb, top, box.Height));
                    break;
                }
                case PositionedBlock pb:
                    Emit(MakeAbs(pb, 0, width, relative: false));
                    break;
                case TransformBlock tb:
                {
                    var inner = LayoutBlocks(tb.Blocks, width);
                    if (inner.Count == 0) break;
                    var tf = new TransformFrag { BoxX = tb.BoxX, BoxWidth = tb.BoxWidth, T = null! };
                    double y = 0;
                    for (int i = 0; i < inner.Count; i++)
                    {
                        if (i > 0) y += inner[i].SpaceBefore;
                        tf.Items.Add((inner[i], y));
                        y += inner[i].Height + (i < inner.Count - 1 ? inner[i].SpaceAfter : 0);
                    }
                    tf.Height = y;
                    if (y > _ctx.MaxImageHeight)
                    {
                        // Выше страницы — выводим без преобразования, чтобы не потерять содержимое
                        foreach (var f in inner) Emit(f);
                        break;
                    }
                    tf.SpaceBefore = inner[0].SpaceBefore;
                    tf.SpaceAfter = inner[^1].SpaceAfter;
                    tf.PageBreakBefore = inner[0].PageBreakBefore;
                    tf.KeepWithNext = inner[^1].KeepWithNext;
                    tf.T = tb.Resolve(tb.BoxWidth, y);
                    Emit(tf);
                    break;
                }
                case AnchorBlock ab:
                    foreach (var child in ab.Children)
                    {
                        var af = MakeAbs(child, ab.X, ab.Width, relative: true);
                        af.TopOffset = ab.TopOffset;
                        af.Anchor = ab;
                        Emit(af);
                    }
                    break;
                case AnchorEndBlock ae:
                    Emit(new AnchorEndFrag { Anchor = ae.Anchor, PadBottom = ae.PadBottom });
                    break;
                case FlexLinesBlock fl:
                {
                    // Строки flex-контейнера: высота контейнера распределяется между ними (align-content)
                    var lines = fl.Lines.Select(t => LayoutTable(t, width)).ToList();
                    double total = lines.Sum(l => l.Sum(f => f.Height)) + lines.Skip(1).Sum(l => l.Count > 0 ? l[0].SpaceBefore : 0);
                    double extra = fl.MinHeight - total;
                    int n = lines.Count;
                    if (extra > 0.5 && n > 0)
                    {
                        if (fl.AlignContent is "normal" or "stretch")
                        {
                            // Растяжение: каждая строка выше на равную долю — повторная вёрстка с заданной высотой
                            for (int i = 0; i < n; i++)
                            {
                                var t = fl.Lines[i];
                                if (t.Rows.Count == 0) continue;
                                t.Rows[0].Height = lines[i].Sum(f => f.Height) + extra / n;
                                lines[i] = LayoutTable(t, width);
                            }
                        }
                        else
                        {
                            var before = new double[n + 1];   // before[i] — перед строкой i; before[n] — после последней
                            switch (fl.AlignContent)
                            {
                                case "flex-end" or "end": before[0] = extra; break;
                                case "center": before[0] = extra / 2; before[n] = extra / 2; break;
                                case "space-between": if (n > 1) for (int i = 1; i < n; i++) before[i] = extra / (n - 1); else before[n] = extra; break;
                                case "space-around": for (int i = 0; i < n; i++) before[i] += extra / (2 * n); for (int i = 1; i <= n; i++) before[i] += extra / (2 * n); break;
                                case "space-evenly": for (int i = 0; i <= n; i++) before[i] = extra / (n + 1); break;
                                default: before[n] = extra; break;
                            }
                            for (int i = 0; i < n; i++) if (lines[i].Count > 0) lines[i][0].SpaceBefore += before[i];
                            if (lines[^1].Count > 0) lines[^1][^1].SpaceAfter += before[n];
                        }
                    }
                    foreach (var l in lines) foreach (var f in l) Emit(f);
                    prev = null; prevLast = null;
                    break;
                }
            }
        }
        EndGroup();
        ResolveAnchorBottoms(frags);
        return frags;
    }

    /// <summary>
    /// absolute с bottom внутри relative-предка: когда конец предка в этом же списке фрагментов, его высота известна —
    /// bottom переводится в top (относительно верха предка). Иначе решение откладывается до разбивки на страницы.
    /// </summary>
    static void ResolveAnchorBottoms(List<Frag> frags)
    {
        if (!frags.Any(f => f is AnchorEndFrag)) return;
        var start = new Dictionary<AbsFrag, double>();
        var ends = new Dictionary<AnchorBlock, (double Y, double Pad)>();
        double y = 0;
        foreach (var f in frags)
        {
            if (f is AbsFrag a) { start[a] = y; continue; }
            if (f is AnchorEndFrag e) { ends[e.Anchor] = (y + e.SpaceBefore, e.PadBottom); continue; }
            y += f.SpaceBefore + f.Height + f.SpaceAfter;
        }
        foreach (var (a, sy) in start)
        {
            if (a.Anchor == null || a.Top != null || a.Bottom == null || !ends.TryGetValue(a.Anchor, out var end)) continue;
            double boxH = end.Y + end.Pad - (sy + a.TopOffset);
            a.Top = boxH - a.Bottom.Value - a.Content.Height;
            a.Bottom = null;
        }
    }

    /// <summary>Вёрстка содержимого надписи в строке и элементов группы фигур.</summary>
    void BuildObject(ImageInline img)
    {
        if (img.Content != null)
        {
            var built = BuildBox(img.Content, Math.Max(1, img.FrameW), img.Box ?? new BoxStyle());
            img.Built = built;
            // Автоподбор высоты: содержимое выше рамки увеличивает её
            if (built.Height > img.FrameH + 0.5)
            {
                if (img.Rotation == 0) img.Height = built.Height;
                img.FrameH = built.Height;
            }
        }
        if (img.Chart != null) img.Built = new ChartLayout(this).Build(img.Chart, Math.Max(20, img.FrameW), Math.Max(20, img.FrameH));
        if (img.Math != null)
        {
            // Формула: размеры по вёрстке, базовая линия формулы совпадает с базовой линией строки
            var mf = new MathLayout(_ctx, img.Style).Build(img.Math, img.MathDisplay);
            img.Built = mf;
            img.Width = img.FrameW = Math.Max(0.1, mf.Width);
            img.Height = img.FrameH = Math.Max(0.1, mf.Ascent + mf.Descent);
            img.MathDescent = mf.Descent;
        }
        if (img.Children != null)
            foreach (var (child, _, _) in img.Children) BuildObject(child);
    }

    /// <summary>Ширина содержимого без переносов (для блоков «по содержимому»).</summary>
    double MeasureContentWidth(List<Block> blocks)
    {
        double max = 0;
        foreach (var f in LayoutBlocks(blocks, 5000))
        {
            if (f is LineFrag lf && lf.Pieces.Count > 0)
                max = Math.Max(max, lf.Left + lf.Pieces.Max(p => p.X + p.Width));
            else if (f is RowFrag rf && rf.Cells.Count > 0)
                max = Math.Max(max, rf.Cells.Max(c => c.X + c.Width));
        }
        return Math.Max(1, max);
    }

    /// <summary>
    /// Буквица Word (framePr dropCap): высота ровно N строк следующего абзаца,
    /// базовая линия буквы — внизу этих строк.
    /// </summary>
    void PrepareDropCap(FloatBlock fb, Paragraph? next, double width)
    {
        if (fb.Blocks.Count == 0 || fb.Blocks[0] is not Paragraph dp) return;
        double lh = 14;
        if (next != null)
        {
            var lines = LayoutParagraph(next, width);
            if (lines.Count > 0) lh = lines[0].Height - lines[0].TopPad;
        }
        dp.LineRule = "exact";
        dp.LineValue = fb.DropLines * lh;
        dp.SpaceBefore = dp.SpaceAfter = 0;
        dp.IndLeft = dp.IndRight = dp.IndFirst = 0;
        dp.Shading = null; dp.BorderTop = dp.BorderBottom = dp.BorderLeft = dp.BorderRight = null;
        fb.MarginBottom = 0;
        if (fb.MarginRight <= 0) fb.MarginRight = 2;
        fb.Width = 0;
    }

    /// <summary>Вёрстка содержимого блока с фоном/рамкой/отступами.</summary>
    public BoxFrag BuildBox(List<Block> blocks, double outerWidth, BoxStyle box)
    {
        double inner = Math.Max(5, outerWidth - box.HorizontalExtra);
        var frags = LayoutBlocks(blocks, inner);
        var bf = new BoxFrag { Box = box, Width = outerWidth };
        double y = 0;
        for (int i = 0; i < frags.Count; i++)
        {
            var f = frags[i];
            if (i > 0) y += f.SpaceBefore;
            bf.Items.Add((f, y));
            y += f.Height + (i < frags.Count - 1 ? f.SpaceAfter : 0);
        }
        bf.Height = Math.Max(box.MinHeight, y + box.PadTop + box.PadBottom + (box.BorderTop?.Width ?? 0) + (box.BorderBottom?.Width ?? 0));
        return bf;
    }

    public BoxFrag BuildBoxTransformed(PositionedBlock pb) => Transformed(BuildBox(pb.Blocks, pb.Width, pb.Box), pb.Transform);

    /// <summary>Оборачивает блок в преобразование (рамка блока сохраняет размеры для раскладки).</summary>
    static BoxFrag Transformed(BoxFrag box, TransformResolver? tr)
    {
        if (tr == null) return box;
        var tf = new TransformFrag { BoxX = 0, BoxWidth = box.Width, Height = box.Height, T = tr(box.Width, box.Height) };
        tf.Items.Add((box, 0));
        var wrapper = new BoxFrag { Box = new BoxStyle(), Width = box.Width, Height = box.Height };
        wrapper.Items.Add((tf, 0));
        return wrapper;
    }

    AbsFrag MakeAbs(PositionedBlock pb, double anchorX, double anchorWidth, bool relative) => new()
    {
        HRel = pb.HRel, VRel = pb.VRel, AlignX = pb.AlignX, AlignY = pb.AlignY,
        Content = Transformed(BuildBox(pb.Blocks, pb.Width, pb.Box), pb.Transform),
        Left = pb.Left, Top = pb.Top, Right = pb.Right, Bottom = pb.Bottom,
        Fixed = pb.Mode == "fixed", RelativeToFlow = relative, AnchorX = anchorX, AnchorWidth = anchorWidth,
        ZIndex = pb.ZIndex, Height = 0,
    };

    // =================================================================
    //  Абзац
    // =================================================================

    static bool IsCjk(int cp) =>
        cp is >= 0x2E80 and <= 0x9FFF or >= 0xAC00 and <= 0xD7AF or >= 0xF900 and <= 0xFAFF or >= 0xFF00 and <= 0xFFEF or >= 0x20000 and <= 0x2FFFF;

    static bool IsSpace(int cp) => cp is ' ' or 0x2002 or 0x2003 or 0x2009 or 0x3000;

    FontFace PrimaryFace(int cp, TextStyle st) =>
        _ctx.Fonts.Get(cp < 0x80 ? st.FontAscii : IsCjk(cp) ? st.FontEastAsia : st.FontHAnsi, st.Bold, st.Italic, st.Weight);

    (FontFace Face, ushort Gid, int Cp) ResolveGlyph(int cp, TextStyle st)
    {
        string family = cp < 0x80 ? st.FontAscii : IsCjk(cp) ? st.FontEastAsia : st.FontHAnsi;
        var face = _ctx.Fonts.Get(family, st.Bold, st.Italic, st.Weight);
        if (face.Font.TryGetGlyph(cp, out var gid)) return (face, gid, cp);

        // Символьные шрифты (Symbol, Wingdings) без самого шрифта
        int mapped = FontManager.MapSymbolChar(family, cp);
        if (mapped != cp)
        {
            if (face.Font.TryGetGlyph(mapped, out gid)) return (face, gid, mapped);
            var fbm = _ctx.Fonts.FallbackFor(mapped, family, st.Bold, st.Italic, st.Weight);
            if (fbm != null && fbm.Font.TryGetGlyph(mapped, out gid)) return (fbm, gid, mapped);
            cp = mapped;
        }
        // Неразрывный пробел, если его нет в шрифте
        if (cp == 0xA0 && face.Font.TryGetGlyph(' ', out gid)) return (face, gid, cp);
        var fb = _ctx.Fonts.FallbackFor(cp, family, st.Bold, st.Italic, st.Weight);
        if (fb != null && fb.Font.TryGetGlyph(cp, out gid)) return (fb, gid, cp);
        return (face, 0, cp);
    }

    /// <summary>Вертикальные метрики берутся из заявленного шрифта (как в Word), а не из запасного.</summary>
    static void SetMetrics(Piece p, TextStyle st, FontFace metricsFace)
    {
        var f = metricsFace.Font;
        double em = f.UnitsPerEm;
        double asc = f.Ascent / em * st.Size, desc = f.Descent / em * st.Size;
        p.Ascent = Math.Max(asc, asc * (p.Size / st.Size) + p.Rise);
        p.Descent = Math.Max(desc, desc * (p.Size / st.Size) - p.Rise);
        // Знаки ударения (w:em) увеличивают высоту строки, чтобы не наезжать на соседние строки
        if (st.Emphasis == "underDot") p.Descent = Math.Max(p.Descent, st.Size * 0.4 - p.Rise);
        else if (st.Emphasis != null) p.Ascent = Math.Max(p.Ascent, Math.Max(asc, st.Size * 0.95) + st.Size * 0.25 + p.Rise);
    }

    (double Size, double Rise) EffectiveSize(TextStyle st)
    {
        double size = st.Size, rise = st.Position;
        if (st.VertAlign == 1) { size *= 0.65; rise += st.Size * 0.33; }
        else if (st.VertAlign == -1) { size *= 0.65; rise -= st.Size * 0.14; }
        return (size, rise);
    }

    /// <summary>Текст инлайна для алгоритма bidi (те же строки, что выводит Itemize; объекты — U+FFFC).</summary>
    string BidiTextOf(Inline inl, Func<TextStyle, TextStyle>? overlay)
    {
        TextStyle St(TextStyle s0) => overlay != null ? overlay(s0) : s0;
        string Caps(string t, TextStyle s0) => St(s0).Caps ? t.ToUpperInvariant() : t;
        return inl switch
        {
            TextInline t => Caps(t.Text, t.Style),
            FieldInline f => Caps(f.Kind == FieldKind.PageRef ? (string.IsNullOrEmpty(f.Cached) ? "99" : f.Cached) : _ctx.FieldText?.Invoke(f.Kind) ?? "99", f.Style),
            FootnoteInline fn => Caps(fn.Mark, fn.Style),
            TabInline => "\t",
            BreakInline => "\u2028",
            _ => "\uFFFC",
        };
    }

    /// <summary>Превращает инлайны в последовательность кусков для переноса строк.</summary>
    List<Piece> Itemize(Paragraph para, Func<TextStyle, TextStyle>? overlay = null)
    {
        var pieces = new List<Piece>();
        bool breakNext = false;
        int src = -1;

        // Уровни встраивания (UAX #9) для текста всего абзаца
        var bidiText = new StringBuilder();
        var inlineStart = new int[para.Inlines.Count];
        for (int k = 0; k < para.Inlines.Count; k++)
        {
            inlineStart[k] = bidiText.Length;
            bidiText.Append(BidiTextOf(para.Inlines[k], overlay));
        }
        var levels = Bidi.ResolveLevels(bidiText.ToString(), para.Rtl);
        byte paraLevel = (byte)(para.Rtl ? 1 : 0);
        byte LevelAt(int i) => i >= 0 && i < levels.Length ? levels[i] : paraLevel;
        int textBase = 0;

        void AddText(string text, TextStyle st)
        {
            if (st.Caps) text = text.ToUpperInvariant();
            var (size, rise) = EffectiveSize(st);
            Piece? cur = null;
            Piece? kernPrev = null;     // кусок с предыдущим глифом прогона (для кернинга пары)
            int pos = 0;
            var shaped = Bidi.ShapeArabic(text);
            foreach (var rune in text.EnumerateRunes())
            {
                int cp = rune.Value;
                int at = pos;
                pos += rune.Utf16SequenceLength;
                if (cp == 0xAD)
                {
                    // Мягкий перенос не отображается; это допустимое место переноса (с дефисом)
                    if (cur != null && cur.Kind == PieceKind.Text && st.Hyphens > 0) (cur.Soft ??= []).Add(cur.Gids.Count);
                    continue;
                }
                if (Bidi.IsControl(cp)) continue;  // управляющие символы направления не выводятся
                byte level = LevelAt(textBase + at);
                if (shaped != null && shaped[at] == -1)
                {
                    // Алеф вошёл в лигатуру лам-алеф: глифа нет, буква добавляется к тексту лигатуры (для копирования)
                    if (cur != null && cur.Texts.Count > 0) cur.Texts[^1] += char.ConvertFromUtf32(cp);
                    continue;
                }
                if (cp == '\t') { var tp = MakeTab(st); tp.Pos.Add(at); pieces.Add(tp); cur = null; kernPrev = null; breakNext = true; continue; }
                if (cp is '\n' or '\r') { pieces.Add(new Piece { Kind = PieceKind.LineBreak, Style = st, Src = src, Pos = [at] }); cur = null; continue; }
                if (cp == 0x200B) { cur = null; breakNext = true; continue; } // нулевой пробел — точка переноса

                bool space = IsSpace(cp);
                double gsize = size;
                if (st.SmallCaps && !st.Caps && Rune.IsLower(rune))
                {
                    cp = Rune.ToUpperInvariant(rune).Value;
                    gsize = size * 0.75;
                }
                // Арабские буквы — контекстная форма для глифа, исходная буква для копирования;
                // на нечётном уровне (справа налево) скобки и т. п. зеркалируются
                int glyphCp = shaped != null && shaped[at] != 0 ? shaped[at] : level % 2 == 1 ? Bidi.Mirror(cp) : cp;
                var (face, gid, outCp) = ResolveGlyph(glyphCp, st);
                if (glyphCp != cp) outCp = cp;
                bool cjk = IsCjk(cp);
                var kind = space ? PieceKind.Space : PieceKind.Text;
                // Кусок не смешивает символы разных уровней встраивания
                bool newLevel = cur != null && cur.Level != level && kind == PieceKind.Text && cur.Kind == PieceKind.Text;
                if (cur == null || cur.Kind != kind || cur.Face != face || cur.Size != gsize || cjk || cur.Level != level)
                {
                    cur = new Piece { Kind = kind, Style = st, Face = face, Size = gsize, Rise = rise, Src = src, Level = level };
                    cur.BreakBefore = kind == PieceKind.Text && (breakNext || cjk) && !newLevel;
                    SetMetrics(cur, st, PrimaryFace(cp, st));
                    pieces.Add(cur);
                    breakNext = false;
                }
                double adv = face.Font.Advance(gid) * gsize * st.HScale + (space ? st.WordSpacing : st.LetterSpacing);
                // Парный кернинг с предыдущим глифом прогона (в т.ч. с пробелом; не для текста справа налево)
                if (st.Kerning && kernPrev is { } kp && kp.Face == face && kp.Size == gsize && kp.Gids.Count > 0 && kp.Adv.Count == kp.Gids.Count && level % 2 == 0 && kp.Level % 2 == 0)
                {
                    double kern = face.Font.Kerning(kp.Gids[^1], gid) * gsize * st.HScale;
                    if (kern != 0) { kp.Adv[^1] += kern; kp.Width += kern; }
                }
                cur.Gids.Add(gid);
                cur.Pos.Add(at);
                cur.Texts.Add(char.ConvertFromUtf32(outCp == 0xA0 ? ' ' : outCp));
                cur.Adv.Add(adv);
                cur.Width += adv;
                kernPrev = cur;
                if (space) { breakNext = true; }
                else if ((cp is '-' or 0x2013 or 0x2014 or '/' && cur.Gids.Count > 1) || cjk)
                {
                    // перенос допускается после дефиса/тире
                    cur = null;
                    breakNext = true;
                }
            }
        }

        Piece MakeTab(TextStyle st)
        {
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic, st.Weight);
            var p = new Piece { Kind = PieceKind.Tab, Style = st, Face = face, Size = st.Size, Src = src };
            SetMetrics(p, st, face);
            return p;
        }

        // Края строчных блоков (span с рамкой/отступами) — «распорки» шириной margin+border+padding
        var openBoxes = new List<InlineBox>();
        void SyncBoxes(TextStyle st)
        {
            var chain = InlineBox.Chain(st.Box);
            int common = 0;
            while (common < openBoxes.Count && common < chain.Count && openBoxes[common] == chain[common]) common++;
            for (int k = openBoxes.Count - 1; k >= common; k--) AddSpacer(openBoxes[k], start: false, st);
            openBoxes.RemoveRange(common, openBoxes.Count - common);
            for (int k = common; k < chain.Count; k++) { AddSpacer(chain[k], start: true, st); openBoxes.Add(chain[k]); }
        }
        void AddSpacer(InlineBox box, bool start, TextStyle st)
        {
            double w = start ? box.StartWidth : box.EndWidth;
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic, st.Weight);
            var p = new Piece { Kind = PieceKind.Spacer, Style = st, Width = w, SpacerBox = box, SpacerStart = start, Size = st.Size };
            SetMetrics(p, st, face);
            if (start) { p.BreakBefore = breakNext; breakNext = false; }
            pieces.Add(p);
        }

        for (src = 0; src < para.Inlines.Count; src++)
        {
            var inl = para.Inlines[src];
            if (overlay != null)
                inl = inl switch
                {
                    TextInline ti => new TextInline(ti.Text, overlay(ti.Style)),
                    FieldInline fi => new FieldInline(fi.Kind, overlay(fi.Style)) { Arg = fi.Arg, Cached = fi.Cached },
                    TabInline tb => new TabInline(overlay(tb.Style)),
                    BreakInline bi => new BreakInline(bi.Kind, overlay(bi.Style)),
                    _ => inl
                };
            var inlStyle = inl switch
            {
                TextInline ti => ti.Style, FieldInline fi => fi.Style, TabInline tb => tb.Style, FootnoteInline fn2 => fn2.Style,
                BreakInline bi => bi.Style, ImageInline ii => ii.Style, _ => null
            };
            if (inlStyle != null) SyncBoxes(inlStyle);
            textBase = inlineStart[src];
            int piecesBefore = pieces.Count;
            switch (inl)
            {
                case TextInline t: AddText(t.Text, t.Style); break;
                case FieldInline f:
                {
                    // В основном тексте номер страницы известен только при отрисовке — ширина по оценке
                    int before = pieces.Count;
                    string est = f.Kind == FieldKind.PageRef ? (string.IsNullOrEmpty(f.Cached) ? "99" : f.Cached) : _ctx.FieldText?.Invoke(f.Kind) ?? "99";
                    AddText(est, f.Style);
                    for (int k = before; k < pieces.Count; k++) { pieces[k].Field = f.Kind; pieces[k].FieldArg = f.Arg; }
                }
                    break;
                case CommentInline cm:
                {
                    // Невидимая метка места примечания (нулевой ширины)
                    var face = _ctx.Fonts.Get(cm.Style.FontHAnsi, cm.Style.Bold, cm.Style.Italic, cm.Style.Weight);
                    var mp = new Piece { Kind = PieceKind.Text, Style = cm.Style, Face = face, Size = cm.Style.Size, Src = src, Comment = cm.Info };
                    SetMetrics(mp, cm.Style, face);
                    pieces.Add(mp);
                    break;
                }
                case FootnoteInline fn:
                {
                    int before = pieces.Count;
                    AddText(fn.Mark, fn.Style);
                    if (pieces.Count > before) pieces[before].Note = fn.Blocks;
                    break;
                }
                case TabInline tab:
                {
                    var tp = MakeTab(tab.Style);
                    if (tab.PtabAlign != null) tp.PTab = tab;
                    pieces.Add(tp);
                    breakNext = true;
                }
                    break;
                case BreakInline br:
                    pieces.Add(new Piece
                    {
                        Kind = br.Kind == BreakKind.Line ? PieceKind.LineBreak : PieceKind.PageBreak, Style = br.Style, Src = src,
                        ColumnBreak = br.Kind == BreakKind.Column
                    });
                    breakNext = true;
                    break;
                case ImageInline img:
                {
                    BuildObject(img);
                    double w = img.Width, h = img.Height;
                    if ((w <= 0 || h <= 0) && img.Shape == null) { w = 100; h = 75; }
                    w = Math.Max(w, 0.5); h = Math.Max(h, 0.5);
                    double maxW = _contentWidthHint - para.IndLeft - para.IndRight;
                    if (maxW > 10 && w > maxW) { h *= maxW / w; w = maxW; }
                    if (h > _ctx.MaxImageHeight) { w *= _ctx.MaxImageHeight / h; h = _ctx.MaxImageHeight; }
                    // vertical-align картинки относительно текста строки
                    double fsz = img.Style.Size, iasc = h, idesc = 0;
                    var mf = _ctx.Fonts.Get(img.Style.FontHAnsi, img.Style.Bold, img.Style.Italic, img.Style.Weight).Font;
                    double fAsc = mf.Ascent / (double)mf.UnitsPerEm * fsz, fDesc = mf.Descent / (double)mf.UnitsPerEm * fsz;
                    if (img.Math != null) { idesc = img.MathDescent * h / Math.Max(0.1, img.Height); iasc = h - idesc; }
                    else switch (img.VAlign)
                    {
                        case "middle": iasc = h / 2 + fsz * 0.25; idesc = h - iasc; break;   // центр на высоте x/2
                        case "text-top" or "top": iasc = fAsc; idesc = h - fAsc; break;
                        case "text-bottom" or "bottom": idesc = fDesc; iasc = h - fDesc; break;
                    }
                    var ip = new Piece
                    {
                        Kind = PieceKind.Image, Style = img.Style, Image = img, Width = w, Ascent = iasc, Descent = Math.Max(0, idesc), Src = src,
                        BreakBefore = true, Size = img.Style.Size
                    };
                    pieces.Add(ip);
                    breakNext = true;
                    break;
                }
            }
            // Нетекстовые куски (картинки, табуляции, распорки) — уровень начала инлайна
            for (int k = piecesBefore; k < pieces.Count; k++)
                if (pieces[k].Kind is not (PieceKind.Text or PieceKind.Space)) pieces[k].Level = LevelAt(textBase);
        }
        src = -1;
        if (openBoxes.Count > 0) SyncBoxes(para.MarkStyle with { Box = null });
        return pieces;
    }

    double _contentWidthHint = 500;

    sealed class RawLine
    {
        public List<Piece> Pieces = [];
        public double OffL, OffR;      // сужение строки плавающими блоками
        public double ExtraTop;        // сдвиг вниз (строка не поместилась рядом с float)
        public double StartX;          // начало строки (номер списка, выровненный вправо/по центру, выступает влево)
        public bool First;
        public PieceKind? EndBreak;   // LineBreak / PageBreak, если строка закончилась разрывом
        public bool ColumnBreak;
    }

    /// <summary>Позиция w:ptab: относительно полей или отступов абзаца, независимо от позиций табуляции.</summary>
    static (double Pos, TabStop Stop) PositionalTab(Paragraph p, TabInline pt, double width)
    {
        bool ind = pt.PtabRelativeTo == "indent";
        double left = ind ? p.IndLeft : 0, right = ind ? width - p.IndRight : width;
        string align = pt.PtabAlign switch { "center" => "center", "right" => "right", _ => "left" };
        double pos = align switch { "center" => (left + right) / 2, "right" => right, _ => left };
        return (pos, new TabStop(pos, align, pt.PtabLeader));
    }

    (double Pos, TabStop Stop) NextTab(Paragraph p, double absPos)
    {
        TabStop? best = null;
        foreach (var t in p.Tabs)
        {
            if (t.Align is "clear" or "bar") continue;
            if (t.Pos > absPos + 0.01) { best = t; break; }
        }
        // Неявная позиция табуляции на отступе слева при выступе первой строки
        if (p.IndFirst < 0 && absPos < p.IndLeft - 0.01 && (best == null || best.Pos > p.IndLeft))
            best = new TabStop(p.IndLeft, "left", "none");
        if (best != null) return (best.Pos, best);
        double step = _ctx.Doc.DefaultTabStop;
        double pos = (Math.Floor(absPos / step + 0.0001) + 1) * step;
        return (pos, new TabStop(pos, "left", "none"));
    }

    public List<LineFrag> LayoutParagraph(Paragraph p, double width, IExclusion? excl = null)
    {
        _contentWidthHint = width;
        var result = p.FirstLine != null ? LayoutWithFirstLine(p, width, excl) : null;
        if (result == null)
        {
            var pieces = Itemize(p);
            var raw = BreakLines(p, pieces, width, excl);
            result = new List<LineFrag>();
            for (int i = 0; i < raw.Count; i++)
                result.Add(BuildLine(p, raw[i], width, i == raw.Count - 1));
        }

        if (result.Count > 0)
        {
            result[0].SpaceBefore = p.SpaceBefore;
            result[^1].SpaceAfter = p.SpaceAfter;
            result[0].PageBreakBefore = p.PageBreakBefore;
            int n = result.Count;
            if (p.KeepLines) for (int i = 0; i < n - 1; i++) result[i].KeepWithNext = true;
            if (p.WidowControl && n >= 2)
            {
                result[0].KeepWithNext = true;
                result[n - 2].KeepWithNext = true;
            }
            if (p.KeepNext) result[^1].KeepWithNext = true;
        }
        return result;
    }

    /// <summary>
    /// ::first-line: первая строка набирается в своём стиле, затем остаток абзаца
    /// (с точки разрыва) переносится заново обычным стилем.
    /// </summary>
    List<LineFrag>? LayoutWithFirstLine(Paragraph p, double width, IExclusion? excl)
    {
        var piecesA = Itemize(p, p.FirstLine);
        var rawA = BreakLines(p, piecesA, width, excl);
        if (rawA.Count == 0) return null;
        var first = rawA[0];
        if (rawA.Count == 1) return [BuildLine(p, first, width, true)];

        // Точка разрыва: после последнего символа первой строки
        var last = first.Pieces.LastOrDefault(q => q.Src >= 0);
        if (last == null) return null;
        int cutInline, cutOffset;
        if (p.Inlines[last.Src] is TextInline lt && last.Pos.Count > 0)
        {
            int at = last.Pos[^1];
            cutInline = last.Src;
            cutOffset = at + (at < lt.Text.Length && char.IsHighSurrogate(lt.Text[at]) ? 2 : 1);
        }
        else { cutInline = last.Src + 1; cutOffset = 0; }

        var rest = new List<Inline>();
        for (int i = cutInline; i < p.Inlines.Count; i++)
        {
            if (i == cutInline && cutOffset > 0 && p.Inlines[i] is TextInline t)
            {
                if (cutOffset < t.Text.Length) rest.Add(new TextInline(t.Text[cutOffset..], t.Style));
                continue;
            }
            rest.Add(p.Inlines[i]);
        }
        var firstFrag = BuildLine(p, first, width, rest.Count == 0);
        if (rest.Count == 0) return [firstFrag];

        var pr = p.ShallowCopy();
        pr.Inlines = rest;
        pr.FirstLine = null;
        var raw = BreakLines(p, Itemize(pr), width, excl, startFirst: false, startY: firstFrag.Height);
        var result = new List<LineFrag> { firstFrag };
        for (int i = 0; i < raw.Count; i++) result.Add(BuildLine(p, raw[i], width, i == raw.Count - 1));
        return result;
    }

    List<RawLine> BreakLines(Paragraph p, List<Piece> pieces, double width, IExclusion? excl, bool startFirst = true, double? startY = null)
    {
        var lines = new List<RawLine>();
        var cur = new RawLine { First = startFirst };
        double x = 0;
        double extraRoom = 0;   // правая/центральная табуляция за правым полем позволяет вылезти за край
        if (startFirst && p.LabelAlign != null)
        {
            // lvlJc right/center: номер заканчивается (или центрирован) на позиции начала строки
            double lw = pieces.Where(q => q.Src == 0 && q.Kind is PieceKind.Text or PieceKind.Space).Sum(q => q.Width);
            cur.StartX = x = p.LabelAlign == "center" ? -lw / 2 : -lw;
        }

        // Вертикальная позиция текущей строки (для обтекания плавающих блоков)
        double curY = startY ?? TopBorderPad(p);
        double estH = excl != null ? LineMetrics(p, [], cur).H : 0;
        void UpdateExcl()
        {
            if (excl == null) return;
            (cur.OffL, cur.OffR) = excl.At(curY, curY + estH);
        }
        UpdateExcl();

        // Сдвиг плавающими блоками действует, только если они есть (отрицательные отступы сохраняются)
        double LineStart(bool first) => cur.OffL > 0 ? Math.Max(p.IndLeft + (first ? p.IndFirst : 0), cur.OffL) : p.IndLeft + (first ? p.IndFirst : 0);
        double Avail(bool first) => width - (cur.OffR > 0 ? Math.Max(p.IndRight, cur.OffR) : p.IndRight) - LineStart(first);
        bool HasContent() => cur.Pieces.Any(q => q.Kind is PieceKind.Text or PieceKind.Image);

        void Finish(PieceKind? endBreak)
        {
            cur.EndBreak = endBreak;
            lines.Add(cur);
            if (excl != null)
                curY += cur.ExtraTop + LineMetrics(p, cur.Pieces.Where(q => q.Kind is not (PieceKind.LineBreak or PieceKind.PageBreak)).ToList(), cur).H;
            cur = new RawLine { First = false };
            x = 0;
            extraRoom = 0;
            UpdateExcl();
        }

        int i = 0;
        int hyphenRun = 0;      // строк подряд, заканчивающихся переносом
        while (i < pieces.Count)
        {
            var pc = pieces[i];
            switch (pc.Kind)
            {
                case PieceKind.LineBreak:
                case PieceKind.PageBreak:
                    cur.Pieces.Add(pc);
                    cur.ColumnBreak = pc.ColumnBreak;
                    Finish(pc.Kind);
                    i++;
                    continue;
                case PieceKind.Space:
                    cur.Pieces.Add(pc);
                    x += pc.Width;
                    i++;
                    continue;
                case PieceKind.Tab:
                {
                    double abs = LineStart(cur.First) + x;
                    var (pos, stop) = pc.PTab is { } pt ? PositionalTab(p, pt, width) : NextTab(p, abs);
                    double w = pos - abs;
                    if (stop.Align is "right" or "end" or "center" or "decimal")
                    {
                        double seg = 0, tail = 0;
                        bool foundSep = false;
                        for (int j = i + 1; j < pieces.Count && pieces[j].Kind is PieceKind.Text or PieceKind.Space or PieceKind.Image or PieceKind.Spacer && !foundSep; j++)
                        {
                            if (stop.Align == "decimal" && pieces[j].Kind == PieceKind.Text)
                            {
                                // По разделителю: на позиции табуляции стоит первая запятая/точка числа
                                int k = pieces[j].Texts.FindIndex(t => t is "," or ".");
                                if (k >= 0)
                                {
                                    seg += pieces[j].Adv.Take(k).Sum();
                                    tail = pieces[j].Width - pieces[j].Adv.Take(k).Sum();
                                    foundSep = true; break;
                                }
                            }
                            seg += pieces[j].Width;
                        }
                        if (stop.Align == "center") seg /= 2;
                        w = Math.Max(0, pos - abs - seg);
                        extraRoom = Math.Max(0, pos + tail - LineStart(cur.First) - Avail(cur.First));
                    }
                    else extraRoom = 0;
                    // Позиция табуляции за правым краем — переносим на новую строку (как Word)
                    if (x + w > Avail(cur.First) + 0.5 && HasContent() && stop.Align is "left" or "start")
                    {
                        Finish(null);
                        continue;
                    }
                    pc.Tab = stop;
                    pc.Width = w;
                    cur.Pieces.Add(pc);
                    x += w;
                    i++;
                    continue;
                }
            }

            // Слово = группа кусков без точек переноса между ними
            int end = i + 1;
            double gw = pc.Width;
            while (end < pieces.Count && pieces[end].Kind is PieceKind.Text or PieceKind.Image or PieceKind.Spacer && !pieces[end].BreakBefore)
            {
                gw += pieces[end].Width;
                end++;
            }

            double avail = Avail(cur.First) + extraRoom;
            if (x + gw <= avail + 0.01)
            {
                for (int j = i; j < end; j++) cur.Pieces.Add(pieces[j]);
                x += gw;
                i = end;
                continue;
            }
            // Перенос слова по слогам (мягкие переносы и словарь): начало слова — в конце строки с дефисом
            if ((HasContent() ? avail - x >= p.HyphenZone : true) && (p.HyphenLimit <= 0 || hyphenRun < p.HyphenLimit)
                && TryHyphenate(p, pieces, i, end, avail - x, cur.Pieces) is int cut)
            {
                i = cut;
                Finish(null);
                hyphenRun++;
                continue;
            }
            if (HasContent())
            {
                Finish(null);
                hyphenRun = 0;
                continue;
            }
            // Рядом с плавающим блоком не помещается — опускаем строку ниже него
            if (excl != null && (cur.OffL > 0 || cur.OffR > 0) && excl.NextBottom(curY) is double nb && nb > curY)
            {
                cur.ExtraTop += nb - curY;
                curY = nb;
                UpdateExcl();
                continue;
            }
            // Слово длиннее строки — режем по глифам
            double room = avail - x;
            int j2 = i;
            bool placed = false;
            for (; j2 < end; j2++)
            {
                var q = pieces[j2];
                if (q.Kind is PieceKind.Image or PieceKind.Spacer || q.Width <= room)
                {
                    if (q.Kind is PieceKind.Image or PieceKind.Spacer && q.Width > room && placed) break;
                    cur.Pieces.Add(q);
                    room -= q.Width;
                    placed = true;
                    continue;
                }
                int k = 0;
                double acc = 0;
                while (k < q.Gids.Count && acc + q.Adv[k] <= room) acc += q.Adv[k++];
                if (k == 0 && !placed) { k = 1; acc = q.Adv[0]; }
                if (k > 0)
                {
                    var head = q.CloneEmpty();
                    head.Gids.AddRange(q.Gids.Take(k)); head.Texts.AddRange(q.Texts.Take(k)); head.Adv.AddRange(q.Adv.Take(k)); head.Pos.AddRange(q.Pos.Take(k));
                    head.Width = acc;
                    head.BreakBefore = q.BreakBefore;
                    cur.Pieces.Add(head);
                    var tail = q.CloneEmpty();
                    tail.Gids.AddRange(q.Gids.Skip(k)); tail.Texts.AddRange(q.Texts.Skip(k)); tail.Adv.AddRange(q.Adv.Skip(k)); tail.Pos.AddRange(q.Pos.Skip(k));
                    tail.Width = q.Width - acc;
                    tail.BreakBefore = true;
                    pieces[j2] = tail;
                    placed = true;
                }
                break;
            }
            i = j2;
            if (i < pieces.Count && i < end) pieces[i].BreakBefore = true;
            Finish(null);
        }
        if (cur.Pieces.Count > 0 || lines.Count == 0 || lines[^1].EndBreak == PieceKind.LineBreak)
            lines.Add(cur);
        return lines;
    }

    /// <summary>
    /// Пересборка группы «float + текст», когда плавающий блок не помещается в остаток страницы (room):
    /// текстовые блоки группы набираются во всю ширину, пока помещаются (абзац может разрезаться по строке);
    /// плавающие блоки с остальным текстом набираются заново и идут дальше (на следующей странице текст снова их обтекает).
    /// </summary>
    (List<Frag> First, List<Frag> Remainder, int Consumed)? RebuildGroup(List<Block> src, List<(Block B, int Frags)> follow, double width, double room)
    {
        var first = new List<Frag>();
        var rest = new List<Block>();
        double y = 0;
        bool full = true;          // текст ещё помещается
        foreach (var b in src)
        {
            if (b is FloatBlock) { rest.Add(b); continue; }
            if (!full) { rest.Add(b); continue; }
            var fr = LayoutBlocks([b], width);
            if (fr.Any(f => f is GroupFrag)) return null;      // вложенная группа — не пересобираем
            double h = 0;
            int take = 0;
            for (int k = 0; k < fr.Count; k++)
            {
                double sb = k == 0 && y == 0 ? 0 : fr[k].SpaceBefore;
                if (y + h + sb + fr[k].Height > room + 0.01) break;
                h += sb + fr[k].Height + (k < fr.Count - 1 ? 0 : fr[k].SpaceAfter);
                take++;
            }
            if (take == fr.Count) { first.AddRange(fr); y += h; continue; }
            full = false;
            // Абзац частично: строки по разрез, остаток — новым абзацем без отступа первой строки
            if (b is Paragraph p && fr.All(f => f is LineFrag))
            {
                int n = take;
                if (p.WidowControl && fr.Count > 1)
                {
                    if (n == 1) n = 0;                          // висячая строка
                    if (fr.Count - n == 1 && n > 1) n--;        // одинокая последняя строка
                }
                if (n > 0 && CutParagraph(p, (LineFrag)fr[n - 1]) is { } restP)
                {
                    first.AddRange(fr.Take(n));
                    rest.Add(restP);
                    continue;
                }
            }
            rest.Add(b);
        }
        if (first.Count == 0) return null;
        // Следующие абзацы обтекают отложенный блок: добавляем их, пока текст не дойдёт до низа блока
        int consumed = 0;
        double floatH = 0, textH = 0;
        foreach (var b in rest)
        {
            if (b is FloatBlock fb) floatH = Math.Max(floatH, LayoutBlocks([fb], width).Sum(f => f.Height));
            else textH += LayoutBlocks([b], width * 0.6).Sum(f => f.Height + f.SpaceBefore);
        }
        foreach (var (b, n) in follow)
        {
            if (textH >= floatH) break;
            rest.Add(b);
            consumed += n;
            textH += LayoutBlocks([b], width * 0.6).Sum(f => f.Height + f.SpaceBefore);
        }
        var restFrags = LayoutBlocks(rest, width);
        if (restFrags.Count > 0) restFrags[0].SpaceBefore = 0;
        return (first, restFrags, consumed);
    }

    /// <summary>Остаток абзаца после строки line (с того же места текста); null — строка последняя.</summary>
    static Paragraph? CutParagraph(Paragraph p, LineFrag line)
    {
        var last = line.Pieces.LastOrDefault(q => q.Src >= 0 && q.Pos.Count > 0);
        if (last == null) return null;
        int cutInline = last.Src, cutOffset;
        if (p.Inlines[last.Src] is TextInline lt)
        {
            int at = last.Pos[^1];
            cutOffset = at + (at < lt.Text.Length && char.IsHighSurrogate(lt.Text[at]) ? 2 : 1);
            if (cutOffset >= lt.Text.Length) { cutInline++; cutOffset = 0; }
        }
        else { cutInline++; cutOffset = 0; }
        var restInl = new List<Inline>();
        for (int i = cutInline; i < p.Inlines.Count; i++)
        {
            if (i == cutInline && cutOffset > 0 && p.Inlines[i] is TextInline t)
            {
                string tail = t.Text[cutOffset..].TrimStart(' ');
                if (tail.Length > 0) restInl.Add(new TextInline(tail, t.Style));
                continue;
            }
            restInl.Add(p.Inlines[i]);
        }
        if (restInl.Count == 0) return null;
        var r = p.ShallowCopy();
        r.Inlines = restInl;
        r.SpaceBefore = 0;
        r.IndFirst = 0;
        r.PageBreakBefore = false;
        r.DropLines = 0;
        r.Bookmarks = null;
        r.FirstLine = null;
        r.LabelAlign = null;
        return r;
    }

    /// <summary>
    /// Перенос слова (куски i..end−1) так, чтобы начало с дефисом поместилось в room. Места переноса — мягкие переносы
    /// и (при автоматических переносах) шаблоны словаря языка. Возвращает индекс куска-остатка либо null.
    /// </summary>
    int? TryHyphenate(Paragraph p, List<Piece> pieces, int i, int end, double room, List<Piece> line)
    {
        if (room <= 0) return null;
        // Слово: только текстовые куски (картинки и распорки внутри слова не переносим)
        for (int j = i; j < end; j++) if (pieces[j].Kind != PieceKind.Text || pieces[j].Face == null || pieces[j].Field != null) return null;
        // Текст слова и соответствие символов глифам
        var sb = new StringBuilder();
        var map = new List<(int Piece, int Glyph)>();     // для каждого символа текста — глиф
        for (int j = i; j < end; j++)
        {
            var q = pieces[j];
            for (int g = 0; g < q.Texts.Count; g++)
                foreach (char ch in q.Texts[g]) { sb.Append(ch); map.Add((j, g)); }
        }
        string word = sb.ToString();
        if (word.Length < 2) return null;
        // Кандидаты: (кусок, индекс глифа, перед которым ставится перенос)
        var cands = new SortedSet<(int Piece, int Glyph)>();
        bool anySoft = false;
        for (int j = i; j < end; j++)
            if (pieces[j].Soft is { } soft)
                foreach (int g in soft) if (g > 0 && g < pieces[j].Gids.Count) { cands.Add((j, g)); anySoft = true; }
        var st0 = pieces[i].Style;
        if (!anySoft && st0.Hyphens == 2 && !p.NoHyphenation && Hyphenator.For(st0.Lang) is { } hy)
        {
            // Буквы слова без начальной и конечной пунктуации
            int a = 0, b = word.Length;
            while (a < b && !char.IsLetter(word[a])) a++;
            while (b > a && !char.IsLetter(word[b - 1]) && word[b - 1] != '\'') b--;
            string core = word[a..b];
            bool caps = core.Length > 1 && core.All(c => !char.IsLetter(c) || char.IsUpper(c));
            int minWord = p.HyphenMinWord > 0 ? p.HyphenMinWord : 5;
            if (core.Length >= minWord && core.All(c => char.IsLetter(c) || c is '\'' or '’') && !(caps && p.HyphenNoCaps))
            {
                foreach (int L in hy.Points(core, p.HyphenMinLeft, p.HyphenMinRight))
                {
                    int at = a + L;                         // символ, перед которым перенос
                    if (at <= 0 || at >= map.Count) continue;
                    var (pj, pg) = map[at];
                    if (map[at - 1] == map[at]) continue;   // внутри лигатуры
                    cands.Add((pj, pg));
                }
            }
        }
        if (cands.Count == 0) return null;
        // Наибольшее начало слова, которое помещается вместе с дефисом
        foreach (var (pj, pg) in cands.Reverse())
        {
            double w = 0;
            for (int j = i; j < pj; j++) w += pieces[j].Width;
            var q = pieces[pj];
            for (int g = 0; g < pg; g++) w += q.Adv[g];
            var face = q.Face!;
            face.Font.TryGetGlyph('-', out var hg);
            double hw = face.Font.Advance(hg) * q.Size * q.Style.HScale;
            if (w + hw > room + 0.01) continue;
            for (int j = i; j < pj; j++) line.Add(pieces[j]);
            var head = q.CloneEmpty();
            head.Gids.AddRange(q.Gids.Take(pg)); head.Texts.AddRange(q.Texts.Take(pg)); head.Adv.AddRange(q.Adv.Take(pg)); head.Pos.AddRange(q.Pos.Take(pg));
            // Кернинг последней буквы начала рассчитан на следующую букву — заменяем на ширину самой буквы
            if (head.Adv.Count > 0) head.Adv[^1] = face.Font.Advance(head.Gids[^1]) * q.Size * q.Style.HScale + q.Style.LetterSpacing;
            head.Gids.Add(hg); head.Texts.Add("-"); head.Adv.Add(hw); head.Pos.Add(q.Pos.Count > pg ? q.Pos[pg] : q.Pos.LastOrDefault());
            head.Width = head.Adv.Sum();
            head.BreakBefore = q.BreakBefore;
            line.Add(head);
            var tail = q.CloneEmpty();
            tail.Gids.AddRange(q.Gids.Skip(pg)); tail.Texts.AddRange(q.Texts.Skip(pg)); tail.Adv.AddRange(q.Adv.Skip(pg)); tail.Pos.AddRange(q.Pos.Skip(pg));
            if (q.Soft != null) tail.Soft = q.Soft.Where(g => g > pg).Select(g => g - pg).ToList();
            tail.Width = tail.Adv.Sum();
            tail.BreakBefore = true;
            pieces[pj] = tail;
            return pj;
        }
        return null;
    }

    (double H, double Baseline) LineMetrics(Paragraph p, List<Piece> pieces, RawLine raw)
    {
        double asc = 0, desc = 0;
        foreach (var q in pieces)
        {
            asc = Math.Max(asc, q.Ascent);
            desc = Math.Max(desc, q.Descent);
        }
        if (asc == 0 && desc == 0)
        {
            var st = raw.Pieces.FirstOrDefault()?.Style ?? p.MarkStyle;
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic, st.Weight);
            asc = face.Font.Ascent / (double)face.Font.UnitsPerEm * st.Size;
            desc = face.Font.Descent / (double)face.Font.UnitsPerEm * st.Size;
        }
        // Высота строки определяется текстом; картинка только увеличивает её
        double natural = asc + desc;
        double h, baseline;
        switch (p.LineRule)
        {
            case "exact":
                h = p.LineValue;
                baseline = h - desc;
                break;
            case "atLeast":
                h = Math.Max(p.LineValue, natural);
                baseline = h - desc;
                break;
            case "css":
            {
                // CSS line-height: «половинный интерлиньяж» — лишнее (или недостающее) место делится
                // поровну над и под текстом; высокие картинки увеличивают строку
                double tAsc = pieces.Where(q => q.Kind != PieceKind.Image).Select(q => q.Ascent).DefaultIfEmpty(0).Max();
                double tDesc = pieces.Where(q => q.Kind != PieceKind.Image).Select(q => q.Descent).DefaultIfEmpty(0).Max();
                if (tAsc == 0 && tDesc == 0) { tAsc = asc; tDesc = desc; }
                h = p.LineValue;
                baseline = tAsc + (p.LineValue - (tAsc + tDesc)) / 2;
                if (asc > baseline) { h += asc - baseline; baseline = asc; }
                if (desc > h - baseline) h = baseline + desc;
                break;
            }
            default:
                double factor = p.LineValue / 240.0;
                bool hasImage = pieces.Any(q => q.Kind == PieceKind.Image);
                double textAsc = pieces.Where(q => q.Kind != PieceKind.Image).Select(q => q.Ascent).DefaultIfEmpty(0).Max();
                if (hasImage)
                {
                    double tn = Math.Max(textAsc, 0) + desc;
                    h = Math.Max(natural, tn * factor);
                    baseline = asc;
                }
                else
                {
                    h = natural * factor;
                    baseline = factor >= 1 ? asc : asc * factor;
                }
                break;
        }

        // Сетка документа (docGrid): высота строки кратна шагу сетки, текст по центру ячейки сетки
        double pitch = _ctx.GridPitch;
        if (pitch > 1 && p.SnapToGrid && p.LineRule == "auto")
        {
            // Как в Word: число шагов сетки по высоте шрифта, умноженное на множитель интервала
            double nh = Math.Max(1, Math.Ceiling((asc + desc) / pitch - 0.02)) * pitch * (p.LineValue / 240.0);
            baseline += (nh - h) / 2;
            h = nh;
        }
        return (h, baseline);
    }

    LineFrag BuildLine(Paragraph p, RawLine raw, double width, bool lastOfPara)
    {
        var pieces = raw.Pieces.Where(q => q.Kind is not (PieceKind.LineBreak or PieceKind.PageBreak)).ToList();
        double lineStart = raw.OffL > 0 ? Math.Max(p.IndLeft + (raw.First ? p.IndFirst : 0), raw.OffL) : p.IndLeft + (raw.First ? p.IndFirst : 0);
        double avail = width - (raw.OffR > 0 ? Math.Max(p.IndRight, raw.OffR) : p.IndRight) - lineStart;
        var (h, baseline) = LineMetrics(p, pieces, raw);

        // Визуальный порядок для текста справа налево (после переноса строк, как в UAX #9)
        Bidi.Reorder(pieces, p.Rtl);

        // Позиции и выравнивание
        int lastVisible = pieces.FindLastIndex(q => q.Kind is not PieceKind.Space);
        double contentW = 0;
        for (int i = 0; i <= lastVisible; i++) contentW += pieces[i].Width;

        double offset = 0;
        string align = p.Rtl ? p.Align switch { "left" => "right", "right" => "left", _ => p.Align } : p.Align;
        bool justify = align == "both" && !lastOfPara && raw.EndBreak == null;
        if (align == "center") offset = (avail - contentW) / 2;
        else if (align == "right") offset = avail - contentW;

        double extraPerSpace = 0;
        int lastTab = pieces.FindLastIndex(q => q.Kind == PieceKind.Tab);
        if (justify && lastVisible >= 0)
        {
            int spaces = 0;
            for (int i = lastTab + 1; i < lastVisible; i++) if (pieces[i].Kind == PieceKind.Space) spaces++;
            double extra = avail - contentW;
            if (spaces > 0 && extra > 0) extraPerSpace = extra / spaces;
        }

        // Распределённое выравнивание: все строки (и последняя) растягиваются за счёт промежутков между символами
        double perGlyph = 0;
        if (align == "distribute" && lastVisible >= 0)
        {
            int glyphs = 0;
            for (int i = lastTab + 1; i <= lastVisible; i++) if (pieces[i].Kind is PieceKind.Text or PieceKind.Space) glyphs += pieces[i].Gids.Count;
            double extra = avail - contentW;
            if (glyphs > 1 && extra > 0) perGlyph = extra / (glyphs - 1);
        }

        double xx = offset + raw.StartX;
        for (int i = 0; i < pieces.Count; i++)
        {
            var q = pieces[i];
            q.X = xx;
            q.GlyphExtra = perGlyph > 0 && i > lastTab && i <= lastVisible && q.Kind is PieceKind.Text or PieceKind.Space ? perGlyph : 0;
            xx += q.Width + q.GlyphExtra * q.Gids.Count;
            if (extraPerSpace > 0 && q.Kind == PieceKind.Space && i > lastTab && i < lastVisible) xx += extraPerSpace;
        }

        var lf = new LineFrag
        {
            Pieces = pieces,
            Para = p,
            Doc = _ctx.Doc,
            Left = lineStart,
            Baseline = baseline,
            ParaLeft = raw.OffL > 0 ? Math.Max(p.IndLeft + Math.Min(0, p.IndFirst), raw.OffL) : p.IndLeft + Math.Min(0, p.IndFirst),
            ParaRight = width - (raw.OffR > 0 ? Math.Max(p.IndRight, raw.OffR) : p.IndRight),
            IsFirst = raw.First,
            IsLast = lastOfPara,
            PageBreakAfter = raw.EndBreak == PieceKind.PageBreak && !raw.ColumnBreak,
            ColumnBreakAfter = raw.ColumnBreak,
            NoteBlocks = pieces.Any(q => q.Note != null) ? pieces.Where(q => q.Note != null).Select(q => q.Note!).ToList() : null,
        };
        if (raw.First) lf.TopPad = TopBorderPad(p);
        lf.TopPad += raw.ExtraTop;
        if (lastOfPara && p.BorderBottom != null && !p.GroupNext) lf.BottomPad = p.BorderBottom.Width + p.BorderBottom.Space;
        lf.Height = lf.TopPad + h + lf.BottomPad;
        return lf;
    }

    // =================================================================
    //  Таблица
    // =================================================================

    /// <summary>
    /// Каменная кладка: высоты элементов — по фактической вёрстке; каждый элемент встаёт в самую короткую колонку
    /// (при равенстве — левую), затем по границам элементов строится таблица с объединёнными по вертикали ячейками.
    /// </summary>
    void BuildMasonry(MasonryBlock mb)
    {
        mb.Rows.Clear();
        int n = Math.Max(1, mb.Columns), gc = Math.Max(1, mb.GapCols);
        double SpanW(int c, int k) { double w = 0; for (int i = c * gc; i <= (c + k - 1) * gc && i < mb.Grid.Count; i++) w += mb.Grid[i]; return w; }
        var run = new double[n];
        var placed = new List<(TableCell Cell, int C, int K, double Top, double Bottom)>();
        foreach (var (byStart, span, col, minH) in mb.Items)
        {
            int k = Math.Clamp(span, 1, n);
            int c = 0;
            if (col is int ec) c = Math.Clamp(ec, 0, n - k);
            else
            {
                double best = double.MaxValue;
                for (int i = 0; i + k <= n; i++)
                {
                    double m = 0;
                    for (int j = i; j < i + k; j++) m = Math.Max(m, run[j]);
                    if (m < best - 0.01) { best = m; c = i; }
                }
            }
            double top = 0;
            for (int j = c; j < c + k; j++) top = Math.Max(top, run[j]);
            var cell = byStart[Math.Min(c, byStart.Length - 1)];
            double w = SpanW(c, k);
            cell.GridSpan = 1; cell.VMerge = VMerge.None;
            var mt = new Table { Grid = [w], Borders = new CellBorders(), CellMargins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 } };
            var mrow = new TableRow();
            mrow.Cells.Add(cell);
            mt.Rows.Add(mrow);
            double h = Math.Max(Math.Max(0.5, minH), LayoutTable(mt, w).Sum(fr => fr.Height));
            placed.Add((cell, c, k, top, top + h));
            for (int j = c; j < c + k; j++) run[j] = top + h + mb.RowGap;
        }
        if (placed.Count == 0) return;
        // align-tracks: выравнивание элементов колонки по высоте контейнера (колонки с элементами на несколько колонок — как есть)
        if (mb.AlignTracks.Count > 0)
        {
            double total = Math.Max(mb.ContainerHeight, placed.Max(p => p.Bottom));
            for (int c = 0; c < n; c++)
            {
                if (placed.Any(p => p.K > 1 && p.C <= c && c < p.C + p.K)) continue;
                var idx = Enumerable.Range(0, placed.Count).Where(i => placed[i].C == c).OrderBy(i => placed[i].Top).ToList();
                var segs = idx.Select(i => (placed[i].Top, placed[i].Bottom - placed[i].Top)).ToList();
                DocxToPdf.Html.HtmlReader.AlignTrack(mb.AlignTracks[Math.Min(c, mb.AlignTracks.Count - 1)], segs, total);
                for (int j = 0; j < idx.Count; j++) placed[idx[j]] = placed[idx[j]] with { Top = segs[j].Item1, Bottom = segs[j].Item1 + segs[j].Item2 };
            }
        }
        // Границы строк — верх и низ каждого элемента
        var ys = new List<double>();
        foreach (var y in placed.SelectMany(p => new[] { p.Top, p.Bottom }).OrderBy(y => y))
            if (ys.Count == 0 || y - ys[^1] > 0.05) ys.Add(y);
        int Idx(double y)
        {
            int best = 0;
            for (int i = 1; i < ys.Count; i++) if (Math.Abs(ys[i] - y) < Math.Abs(ys[best] - y)) best = i;
            return best;
        }
        var spans = placed.Select(p => (p.Cell, p.C, p.K, R0: Idx(p.Top), R1: Idx(p.Bottom))).ToList();
        int mcols = mb.Grid.Count;
        for (int ri = 0; ri + 1 < ys.Count; ri++)
        {
            var row = new TableRow { Height = ys[ri + 1] - ys[ri] };
            int mc = 0;
            while (mc < mcols)
            {
                int pi = spans.FindIndex(p => p.C * gc == mc && p.R0 <= ri && ri < p.R1);
                if (pi >= 0)
                {
                    var p = spans[pi];
                    int ms = (p.K - 1) * gc + 1;
                    if (p.R0 == ri)
                    {
                        p.Cell.GridSpan = ms;
                        p.Cell.VMerge = p.R1 - p.R0 > 1 ? VMerge.Restart : VMerge.None;
                        row.Cells.Add(p.Cell);
                    }
                    else
                    {
                        row.Cells.Add(new TableCell
                        {
                            GridSpan = ms, VMerge = VMerge.Continue, Shading = p.Cell.Shading, Radius = p.Cell.Radius,
                            Borders = new CellBorders { Left = p.Cell.Borders?.Left, Right = p.Cell.Borders?.Right, Bottom = ri + 1 == p.R1 ? p.Cell.Borders?.Bottom : null },
                            Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 },
                        });
                    }
                    mc += ms;
                }
                else
                {
                    row.Cells.Add(new TableCell { Borders = new CellBorders(), Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 } });
                    mc++;
                }
            }
            mb.Rows.Add(row);
        }
    }

    public List<Frag> LayoutTable(Table t, double avail)
    {
        var info = new TableInfo();
        var result = new List<Frag>();
        if (t.Rows.Count == 0) return result;

        int cols = Math.Max(t.Grid.Count, t.Rows.Max(r => r.Cells.Sum(c => c.GridSpan)));
        var grid = new List<double>(t.Grid);
        while (grid.Count < cols) grid.Add(grid.Count > 0 ? grid.Average() : 72);
        double total = grid.Sum();
        // Предпочтительная ширина таблицы (tblW): в процентах — от ширины текста, в пунктах — как задана
        double? target = t.WidthPct is { } wp ? wp * avail : t.Width is > 0 ? t.Width : null;
        if (target is { } tw && total > 0 && Math.Abs(tw - total) > 1)
        {
            double k = tw / total;
            for (int i = 0; i < grid.Count; i++) grid[i] *= k;
            total = tw;
        }
        double maxW = avail - Math.Max(0, t.Indent);
        if (total > maxW + 1 && maxW > 20)
        {
            double k = maxW / total;
            for (int i = 0; i < grid.Count; i++) grid[i] *= k;
            total = maxW;
        }
        double x0 = t.Align switch
        {
            "center" => (avail - total) / 2,
            "right" => avail - total,
            _ => t.Indent
        };
        var colX = new double[cols + 1];
        for (int i = 0; i < cols; i++) colX[i + 1] = colX[i] + grid[i];

        var rowsCells = new List<List<CellBox>>();
        var rowHeights = new List<double>();
        for (int ri = 0; ri < t.Rows.Count; ri++)
        {
            var row = t.Rows[ri];
            var boxes = new List<CellBox>();
            int g = 0;
            double rowH = 0;
            foreach (var cell in row.Cells)
            {
                int span = Math.Min(cell.GridSpan, Math.Max(1, cols - g));
                if (g >= cols) break;
                var m = new Margins();
                m.Overlay(t.CellMargins);
                m.Overlay(cell.Margins);
                var box = new CellBox
                {
                    X = x0 + colX[g], Width = colX[g + span] - colX[g],
                    MarginLeft = m.Left ?? 5.4, MarginRight = m.Right ?? 5.4, MarginTop = m.Top ?? 0, MarginBottom = m.Bottom ?? 0,
                    Shading = cell.Shading, VAlign = cell.VAlign, VMerge = cell.VMerge, GridIndex = g, Span = span, Radius = cell.Radius,
                };
                var b = cell.Borders;
                if (t.CellSpacing > 0)
                {
                    // Интервалы между ячейками: у каждой ячейки своя рамка (внутренние границы таблицы),
                    // внешняя рамка таблицы рисуется отдельно вокруг всех ячеек
                    box.X += t.CellSpacing; box.Width = Math.Max(1, box.Width - 2 * t.CellSpacing);
                    box.Top = b?.Top ?? t.Borders.InsideH ?? t.Borders.Top;
                    box.Bottom = b?.Bottom ?? t.Borders.InsideH ?? t.Borders.Bottom;
                    box.Left = b?.Left ?? t.Borders.InsideV ?? t.Borders.Left;
                    box.Right = b?.Right ?? t.Borders.InsideV ?? t.Borders.Right;
                }
                else
                {
                    box.Top = b?.Top ?? (ri == 0 ? t.Borders.Top : t.Borders.InsideH);
                    box.Bottom = b?.Bottom ?? (ri == t.Rows.Count - 1 ? t.Borders.Bottom : t.Borders.InsideH);
                    box.Left = b?.Left ?? (g == 0 ? t.Borders.Left : t.Borders.InsideV);
                    box.Right = b?.Right ?? (g + span >= cols ? t.Borders.Right : t.Borders.InsideV);
                }

                if (cell.VMerge != VMerge.Continue && cell.TextDirection != null)
                {
                    // Повёрнутый текст: строки идут вдоль высоты строки таблицы
                    double cw0 = Math.Max(1, box.Width - box.MarginLeft - box.MarginRight);
                    double len = row.Height > 0 ? row.Height - box.MarginTop - box.MarginBottom : MeasureContentWidth(cell.Blocks) + 1;
                    var rf = LayoutBlocks(cell.Blocks, Math.Max(10, len));
                    var rot = new RotatedFrag { Length = Math.Max(10, len), Direction = cell.TextDirection, CellWidth = cw0 };
                    double ry = 0;
                    foreach (var f in rf) { ry += f.SpaceBefore; rot.Items.Add((f, ry)); ry += f.Height + f.SpaceAfter; }
                    rot.Thickness = ry;
                    rot.Height = rot.Length;
                    box.Items.Add((rot, 0));
                    box.ContentHeight = rot.Height;
                    if (cell.VMerge == VMerge.None) rowH = Math.Max(rowH, rot.Height + box.MarginTop + box.MarginBottom);
                }
                else if (cell.VMerge != VMerge.Continue)
                {
                    double cw = Math.Max(1, box.Width - box.MarginLeft - box.MarginRight);
                    var frags = LayoutBlocks(cell.Blocks, cw);
                    double y = 0;
                    foreach (var f in frags)
                    {
                        y += f.SpaceBefore;
                        box.Items.Add((f, y));
                        y += f.Height + f.SpaceAfter;
                    }
                    box.ContentHeight = y;
                    if (cell.VMerge == VMerge.None)
                        rowH = Math.Max(rowH, y + box.MarginTop + box.MarginBottom);
                }
                boxes.Add(box);
                g += span;
            }
            if (row.HeightRule == "exact" && (row.Height > 0 || row.Role is 1 or 2 or 3 or 4)) rowH = row.Height;
            else rowH = Math.Max(rowH, row.Height);
            // Интервал сверху и снизу ячеек: между строками — двойной, у края таблицы — одинарный
            if (t.CellSpacing > 0) rowH += 2 * t.CellSpacing;
            if (rowH <= 0 && !(row.HeightRule == "exact" && row.Role is 1 or 2 or 3 or 4)) rowH = 1;   // служебные строки сетки могут быть нулевой высоты
            if (t.Rtl)
            {
                // Таблица справа налево: колонки зеркально, левая и правая границы ячеек меняются местами
                foreach (var bx in boxes)
                {
                    bx.X = 2 * x0 + total - bx.X - bx.Width;
                    (bx.Left, bx.Right) = (bx.Right, bx.Left);
                }
            }
            rowsCells.Add(boxes);
            rowHeights.Add(rowH);
        }

        // align-content: лишняя высота контейнера сетки распределяется между строками, промежутками и служебными строками
        double contentSum = rowHeights.Sum();
        if (t.MinContentHeight > contentSum + 0.5)
        {
            double extra = t.MinContentHeight - contentSum;
            var content = Enumerable.Range(0, t.Rows.Count).Where(i => t.Rows[i].Role == 0).ToList();
            var gaps = Enumerable.Range(0, t.Rows.Count).Where(i => t.Rows[i].Role == 1).ToList();
            int top = t.Rows.FindIndex(r => r.Role == 2), bottom = t.Rows.FindIndex(r => r.Role == 3);
            void Add(int i, double v) { if (i >= 0 && i < rowHeights.Count) rowHeights[i] += v; }
            int n = content.Count;
            switch (t.AlignContent)
            {
                case "start" or "flex-start": Add(bottom, extra); break;
                case "end" or "flex-end": Add(top, extra); break;
                case "center": Add(top, extra / 2); Add(bottom, extra / 2); break;
                case "space-between":
                    if (gaps.Count > 0) foreach (var gi in gaps) Add(gi, extra / gaps.Count); else Add(bottom, extra);
                    break;
                case "space-around":
                    Add(top, extra / (2 * Math.Max(1, n))); Add(bottom, extra / (2 * Math.Max(1, n)));
                    foreach (var gi in gaps) Add(gi, extra / Math.Max(1, n));
                    break;
                case "space-evenly":
                    Add(top, extra / (n + 1)); Add(bottom, extra / (n + 1));
                    foreach (var gi in gaps) Add(gi, extra / (n + 1));
                    break;
                default:
                {
                    // normal/stretch: растягиваются строки без явной высоты
                    var auto = content.Where(i => t.Rows[i].AutoHeight).ToList();
                    if (auto.Count == 0) Add(bottom, extra);
                    else foreach (var i in auto) Add(i, extra / auto.Count);
                    break;
                }
            }
        }

        // Вертикальные объединения: убираем внутренние границы, растягиваем последнюю строку
        var merged = new List<(CellBox Box, int First, int Last)>();
        for (int ri = 0; ri < rowsCells.Count; ri++)
        {
            foreach (var box in rowsCells[ri])
            {
                if (box.VMerge != VMerge.Restart) continue;
                int last = ri;
                while (last + 1 < rowsCells.Count && rowsCells[last + 1].Any(c => c.GridIndex == box.GridIndex && c.VMerge == VMerge.Continue))
                    last++;
                if (last == ri)
                {
                    rowHeights[ri] = Math.Max(rowHeights[ri], box.ContentHeight + box.MarginTop + box.MarginBottom);
                    box.VMerge = VMerge.None;
                    continue;
                }
                double need = box.ContentHeight + box.MarginTop + box.MarginBottom;
                double have = 0;
                for (int k = ri; k <= last; k++) have += rowHeights[k];
                if (need > have)
                {
                    // Нехватка высоты — в последнюю строку содержимого (промежутки сетки и их части не растягиваются)
                    int grow = last;
                    while (grow > ri && t.Rows[grow].Role is 1 or 4 && t.Rows[grow].HeightRule == "exact") grow--;
                    rowHeights[grow] += need - have;
                }
                merged.Add((box, ri, last));
                box.Bottom = null;
                for (int k = ri + 1; k <= last; k++)
                {
                    var cont = rowsCells[k].First(c => c.GridIndex == box.GridIndex && c.VMerge == VMerge.Continue);
                    cont.Top = null;
                    if (k < last) cont.Bottom = null;
                    cont.Shading ??= box.Shading;
                }
            }
        }
        // Высота объединения — после всех растяжений (соседнее объединение могло увеличить общую строку)
        foreach (var (box, first, last) in merged)
        {
            double h = 0;
            for (int k = first; k <= last; k++) h += rowHeights[k];
            box.MergedHeight = h;
        }

        for (int ri = 0; ri < rowsCells.Count; ri++)
        {
            var row = t.Rows[ri];
            var rf = new RowFrag
            {
                Table = info, Cells = rowsCells[ri], Height = rowHeights[ri], IsHeader = row.IsHeader,
                CantSplit = row.CantSplit || row.HeightRule == "exact" || rowsCells[ri].Any(c => c.VMerge == VMerge.Restart),
                ClipContent = row.HeightRule == "exact" && row.Height > 0,
            };
            if (t.CellSpacing > 0)
            {
                rf.CantSplit = true;
                rf.InsetTop = t.CellSpacing;
                rf.InsetBottom = t.CellSpacing;
                rf.OuterX = x0; rf.OuterWidth = total;
                rf.OuterTop = ri == 0 ? t.Borders.Top : null;
                rf.OuterBottom = ri == rowsCells.Count - 1 ? t.Borders.Bottom : null;
                rf.OuterLeft = t.Borders.Left; rf.OuterRight = t.Borders.Right;
            }
            if (row.IsHeader && info.HeaderRows.Count == ri) info.HeaderRows.Add(rf);
            result.Add(rf);
        }
        if (result.Count > 0)
        {
            result[0].PageBreakBefore = t.PageBreakBefore;
            result[0].SpaceBefore = t.SpaceBefore;
            result[^1].SpaceAfter = t.SpaceAfter;
        }
        // Строки заголовка не отрываем от первой строки данных
        foreach (var h in info.HeaderRows) h.KeepWithNext = true;
        return result;
    }
}

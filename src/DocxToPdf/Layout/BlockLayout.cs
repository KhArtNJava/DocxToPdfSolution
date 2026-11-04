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

    sealed record FloatBox(string Side, double X0, double X1, double Top, double Bottom);

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
                if (b.Side == "left") l = Math.Max(l, b.X1);
                else r = Math.Max(r, _width - b.X0);
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

    public List<Frag> LayoutBlocks(IReadOnlyList<Block> blocks, double width)
    {
        var frags = new List<Frag>();
        Paragraph? prev = null;
        LineFrag? prevLast = null;

        GroupFrag? group = null;
        FloatArea? area = null;
        double gy = 0;   // текущая позиция внутри группы

        void EndGroup()
        {
            if (group == null) return;
            group.Height = Math.Max(gy, area!.MaxBottom());
            frags.Add(group);
            group = null; area = null; gy = 0;
        }

        void Emit(Frag f)
        {
            if (group != null && gy >= area!.MaxBottom() - 0.01) EndGroup();
            if (group == null) { frags.Add(f); return; }
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
            var block = blocks[bi];
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
                    // Таблица рядом с float не сужается — размещаем ниже плавающих блоков
                    if (group != null) { Clear("both"); EndGroup(); }
                    frags.AddRange(LayoutTable(t, width));
                    prev = null; prevLast = null;
                    break;
                }
                case FloatBlock fb:
                {
                    double w = fb.Width + fb.MarginLeft + fb.MarginRight;
                    var box = Transformed(BuildBox(fb.Blocks, fb.Width, fb.Box), fb.Transform);
                    double h = box.Height + fb.MarginTop + fb.MarginBottom;
                    if (w > width + 0.5 || h > _ctx.MaxImageHeight * 0.9)
                    {
                        // Слишком большой для обтекания — выводим в потоке
                        if (group != null) { Clear("both"); EndGroup(); }
                        frags.Add(box);
                        break;
                    }
                    if (group == null) { group = new GroupFrag(); area = new FloatArea(width, 0); gy = 0; }
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
                    area.Boxes.Add(new FloatBox(fb.Side, x, x + w, top, top + h));
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
                        Emit(af);
                    }
                    break;
            }
        }
        EndGroup();
        return frags;
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
        _ctx.Fonts.Get(cp < 0x80 ? st.FontAscii : IsCjk(cp) ? st.FontEastAsia : st.FontHAnsi, st.Bold, st.Italic);

    (FontFace Face, ushort Gid, int Cp) ResolveGlyph(int cp, TextStyle st)
    {
        string family = cp < 0x80 ? st.FontAscii : IsCjk(cp) ? st.FontEastAsia : st.FontHAnsi;
        var face = _ctx.Fonts.Get(family, st.Bold, st.Italic);
        if (face.Font.TryGetGlyph(cp, out var gid)) return (face, gid, cp);

        // Символьные шрифты (Symbol, Wingdings) без самого шрифта
        int mapped = FontManager.MapSymbolChar(family, cp);
        if (mapped != cp)
        {
            if (face.Font.TryGetGlyph(mapped, out gid)) return (face, gid, mapped);
            var fbm = _ctx.Fonts.FallbackFor(mapped, family, st.Bold, st.Italic);
            if (fbm != null && fbm.Font.TryGetGlyph(mapped, out gid)) return (fbm, gid, mapped);
            cp = mapped;
        }
        // Неразрывный пробел, если его нет в шрифте
        if (cp == 0xA0 && face.Font.TryGetGlyph(' ', out gid)) return (face, gid, cp);
        var fb = _ctx.Fonts.FallbackFor(cp, family, st.Bold, st.Italic);
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
    }

    (double Size, double Rise) EffectiveSize(TextStyle st)
    {
        double size = st.Size, rise = st.Position;
        if (st.VertAlign == 1) { size *= 0.65; rise += st.Size * 0.33; }
        else if (st.VertAlign == -1) { size *= 0.65; rise -= st.Size * 0.14; }
        return (size, rise);
    }

    /// <summary>Превращает инлайны в последовательность кусков для переноса строк.</summary>
    List<Piece> Itemize(Paragraph para, Func<TextStyle, TextStyle>? overlay = null)
    {
        var pieces = new List<Piece>();
        bool breakNext = false;
        int src = -1;

        void AddText(string text, TextStyle st)
        {
            if (st.Caps) text = text.ToUpperInvariant();
            var (size, rise) = EffectiveSize(st);
            Piece? cur = null;
            int pos = 0;
            foreach (var rune in text.EnumerateRunes())
            {
                int cp = rune.Value;
                int at = pos;
                pos += rune.Utf16SequenceLength;
                if (cp == 0xAD) continue;          // мягкий перенос не отображаем
                if (cp == '\t') { var tp = MakeTab(st); tp.Pos.Add(at); pieces.Add(tp); cur = null; breakNext = true; continue; }
                if (cp is '\n' or '\r') { pieces.Add(new Piece { Kind = PieceKind.LineBreak, Style = st, Src = src, Pos = [at] }); cur = null; continue; }
                if (cp == 0x200B) { cur = null; breakNext = true; continue; } // нулевой пробел — точка переноса

                bool space = IsSpace(cp);
                double gsize = size;
                if (st.SmallCaps && !st.Caps && Rune.IsLower(rune))
                {
                    cp = Rune.ToUpperInvariant(rune).Value;
                    gsize = size * 0.75;
                }
                var (face, gid, outCp) = ResolveGlyph(cp, st);
                bool cjk = IsCjk(cp);
                var kind = space ? PieceKind.Space : PieceKind.Text;
                if (cur == null || cur.Kind != kind || cur.Face != face || cur.Size != gsize || cjk)
                {
                    cur = new Piece { Kind = kind, Style = st, Face = face, Size = gsize, Rise = rise, Src = src };
                    cur.BreakBefore = kind == PieceKind.Text && (breakNext || cjk);
                    SetMetrics(cur, st, PrimaryFace(cp, st));
                    pieces.Add(cur);
                    breakNext = false;
                }
                double adv = face.Font.Advance(gid) * gsize + (space ? st.WordSpacing : st.LetterSpacing);
                cur.Gids.Add(gid);
                cur.Pos.Add(at);
                cur.Texts.Add(char.ConvertFromUtf32(outCp == 0xA0 ? ' ' : outCp));
                cur.Adv.Add(adv);
                cur.Width += adv;
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
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic);
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
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic);
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
                    FieldInline fi => new FieldInline(fi.Kind, overlay(fi.Style)),
                    TabInline tb => new TabInline(overlay(tb.Style)),
                    BreakInline bi => new BreakInline(bi.Kind, overlay(bi.Style)),
                    _ => inl
                };
            var inlStyle = inl switch
            {
                TextInline ti => ti.Style, FieldInline fi => fi.Style, TabInline tb => tb.Style,
                BreakInline bi => bi.Style, ImageInline ii => ii.Style, _ => null
            };
            if (inlStyle != null) SyncBoxes(inlStyle);
            switch (inl)
            {
                case TextInline t: AddText(t.Text, t.Style); break;
                case FieldInline f:
                    AddText(_ctx.FieldText?.Invoke(f.Kind) ?? "1", f.Style);
                    break;
                case TabInline tab:
                    pieces.Add(MakeTab(tab.Style));
                    breakNext = true;
                    break;
                case BreakInline br:
                    pieces.Add(new Piece { Kind = br.Kind == BreakKind.Line ? PieceKind.LineBreak : PieceKind.PageBreak, Style = br.Style, Src = src });
                    breakNext = true;
                    break;
                case ImageInline img:
                {
                    double w = img.Width, h = img.Height;
                    if (w <= 0 || h <= 0) { w = 100; h = 75; }
                    double maxW = _contentWidthHint - para.IndLeft - para.IndRight;
                    if (maxW > 10 && w > maxW) { h *= maxW / w; w = maxW; }
                    if (h > _ctx.MaxImageHeight) { w *= _ctx.MaxImageHeight / h; h = _ctx.MaxImageHeight; }
                    // vertical-align картинки относительно текста строки
                    double fsz = img.Style.Size, iasc = h, idesc = 0;
                    var mf = _ctx.Fonts.Get(img.Style.FontHAnsi, img.Style.Bold, img.Style.Italic).Font;
                    double fAsc = mf.Ascent / (double)mf.UnitsPerEm * fsz, fDesc = mf.Descent / (double)mf.UnitsPerEm * fsz;
                    switch (img.VAlign)
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
        public bool First;
        public PieceKind? EndBreak;   // LineBreak / PageBreak, если строка закончилась разрывом
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

        // Вертикальная позиция текущей строки (для обтекания плавающих блоков)
        double curY = startY ?? (p.BorderTop != null ? p.BorderTop.Width + p.BorderTop.Space : 0);
        double estH = excl != null ? LineMetrics(p, [], cur).H : 0;
        void UpdateExcl()
        {
            if (excl == null) return;
            (cur.OffL, cur.OffR) = excl.At(curY, curY + estH);
        }
        UpdateExcl();

        double LineStart(bool first) => Math.Max(p.IndLeft + (first ? p.IndFirst : 0), cur.OffL);
        double Avail(bool first) => width - Math.Max(p.IndRight, cur.OffR) - LineStart(first);
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
        while (i < pieces.Count)
        {
            var pc = pieces[i];
            switch (pc.Kind)
            {
                case PieceKind.LineBreak:
                case PieceKind.PageBreak:
                    cur.Pieces.Add(pc);
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
                    var (pos, stop) = NextTab(p, abs);
                    double w = pos - abs;
                    if (stop.Align is "right" or "end" or "center" or "decimal")
                    {
                        double seg = 0;
                        for (int j = i + 1; j < pieces.Count && pieces[j].Kind is PieceKind.Text or PieceKind.Space or PieceKind.Image; j++)
                            seg += pieces[j].Width;
                        if (stop.Align == "center") seg /= 2;
                        w = Math.Max(0, pos - abs - seg);
                        extraRoom = Math.Max(0, pos - LineStart(cur.First) - Avail(cur.First));
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
            if (HasContent())
            {
                Finish(null);
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
            var face = _ctx.Fonts.Get(st.FontHAnsi, st.Bold, st.Italic);
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

        return (h, baseline);
    }

    LineFrag BuildLine(Paragraph p, RawLine raw, double width, bool lastOfPara)
    {
        var pieces = raw.Pieces.Where(q => q.Kind is not (PieceKind.LineBreak or PieceKind.PageBreak)).ToList();
        double lineStart = Math.Max(p.IndLeft + (raw.First ? p.IndFirst : 0), raw.OffL);
        double avail = width - Math.Max(p.IndRight, raw.OffR) - lineStart;
        var (h, baseline) = LineMetrics(p, pieces, raw);

        // Позиции и выравнивание
        int lastVisible = pieces.FindLastIndex(q => q.Kind is not PieceKind.Space);
        double contentW = 0;
        for (int i = 0; i <= lastVisible; i++) contentW += pieces[i].Width;

        double offset = 0;
        string align = p.Align;
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

        double xx = offset;
        for (int i = 0; i < pieces.Count; i++)
        {
            var q = pieces[i];
            q.X = xx;
            xx += q.Width;
            if (extraPerSpace > 0 && q.Kind == PieceKind.Space && i > lastTab && i < lastVisible) xx += extraPerSpace;
        }

        var lf = new LineFrag
        {
            Pieces = pieces,
            Para = p,
            Doc = _ctx.Doc,
            Left = lineStart,
            Baseline = baseline,
            ParaLeft = Math.Max(p.IndLeft + Math.Min(0, p.IndFirst), raw.OffL),
            ParaRight = width - Math.Max(p.IndRight, raw.OffR),
            IsFirst = raw.First,
            IsLast = lastOfPara,
            PageBreakAfter = raw.EndBreak == PieceKind.PageBreak,
        };
        if (raw.First && p.BorderTop != null) lf.TopPad = p.BorderTop.Width + p.BorderTop.Space;
        lf.TopPad += raw.ExtraTop;
        if (lastOfPara && p.BorderBottom != null) lf.BottomPad = p.BorderBottom.Width + p.BorderBottom.Space;
        lf.Height = lf.TopPad + h + lf.BottomPad;
        return lf;
    }

    // =================================================================
    //  Таблица
    // =================================================================

    public List<Frag> LayoutTable(Table t, double avail)
    {
        var info = new TableInfo();
        var result = new List<Frag>();
        if (t.Rows.Count == 0) return result;

        int cols = Math.Max(t.Grid.Count, t.Rows.Max(r => r.Cells.Sum(c => c.GridSpan)));
        var grid = new List<double>(t.Grid);
        while (grid.Count < cols) grid.Add(grid.Count > 0 ? grid.Average() : 72);
        double total = grid.Sum();
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
                box.Top = b?.Top ?? (ri == 0 ? t.Borders.Top : t.Borders.InsideH);
                box.Bottom = b?.Bottom ?? (ri == t.Rows.Count - 1 ? t.Borders.Bottom : t.Borders.InsideH);
                box.Left = b?.Left ?? (g == 0 ? t.Borders.Left : t.Borders.InsideV);
                box.Right = b?.Right ?? (g + span >= cols ? t.Borders.Right : t.Borders.InsideV);

                if (cell.VMerge != VMerge.Continue)
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
            if (row.HeightRule == "exact" && row.Height > 0) rowH = row.Height;
            else rowH = Math.Max(rowH, row.Height);
            if (rowH <= 0) rowH = 1;
            rowsCells.Add(boxes);
            rowHeights.Add(rowH);
        }

        // Вертикальные объединения: убираем внутренние границы, растягиваем последнюю строку
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
                if (need > have) rowHeights[last] += need - have;
                have = 0;
                for (int k = ri; k <= last; k++) have += rowHeights[k];
                box.MergedHeight = have;
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

        for (int ri = 0; ri < rowsCells.Count; ri++)
        {
            var row = t.Rows[ri];
            var rf = new RowFrag
            {
                Table = info, Cells = rowsCells[ri], Height = rowHeights[ri], IsHeader = row.IsHeader,
                CantSplit = row.CantSplit || row.HeightRule == "exact" || rowsCells[ri].Any(c => c.VMerge == VMerge.Restart)
            };
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

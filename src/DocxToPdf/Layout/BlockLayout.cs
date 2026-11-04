using System.Globalization;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;

namespace DocxToPdf.Layout;

/// <summary>Вёрстка блоков (абзацев и таблиц) в колонку заданной ширины.</summary>
public sealed class BlockLayout(LayoutContext ctx)
{
    readonly LayoutContext _ctx = ctx;

    public List<Frag> LayoutBlocks(IReadOnlyList<Block> blocks, double width)
    {
        var frags = new List<Frag>();
        Paragraph? prev = null;
        LineFrag? prevLast = null;
        foreach (var block in blocks)
        {
            if (block is Paragraph p)
            {
                var lines = LayoutParagraph(p, width);
                if (prev != null && prev.StyleId == p.StyleId && lines.Count > 0)
                {
                    if (p.ContextualSpacing) lines[0].SpaceBefore = 0;
                    if (prev.ContextualSpacing && prevLast != null) prevLast.SpaceAfter = 0;
                }
                frags.AddRange(lines);
                prev = p;
                prevLast = lines.Count > 0 ? lines[^1] : null;
            }
            else if (block is Table t)
            {
                var rows = LayoutTable(t, width);
                // Абзац с "не отрывать от следующего" перед таблицей — держим с первой строкой
                frags.AddRange(rows);
                prev = null; prevLast = null;
            }
        }
        return frags;
    }

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
    List<Piece> Itemize(Paragraph para)
    {
        var pieces = new List<Piece>();
        bool breakNext = false;

        void AddText(string text, TextStyle st)
        {
            if (st.Caps) text = text.ToUpperInvariant();
            var (size, rise) = EffectiveSize(st);
            Piece? cur = null;
            foreach (var rune in text.EnumerateRunes())
            {
                int cp = rune.Value;
                if (cp == 0xAD) continue;          // мягкий перенос не отображаем
                if (cp == '\t') { pieces.Add(MakeTab(st)); cur = null; breakNext = true; continue; }
                if (cp is '\n' or '\r') { pieces.Add(new Piece { Kind = PieceKind.LineBreak, Style = st }); cur = null; continue; }
                if (cp == 0x200B) { cur = null; breakNext = true; continue; } // нулевой пробел — точка переноса

                bool space = IsSpace(cp);
                var (face, gid, outCp) = ResolveGlyph(cp, st);
                bool cjk = IsCjk(cp);
                var kind = space ? PieceKind.Space : PieceKind.Text;
                if (cur == null || cur.Kind != kind || cur.Face != face || cjk)
                {
                    cur = new Piece { Kind = kind, Style = st, Face = face, Size = size, Rise = rise };
                    cur.BreakBefore = kind == PieceKind.Text && (breakNext || cjk);
                    SetMetrics(cur, st, PrimaryFace(cp, st));
                    pieces.Add(cur);
                    breakNext = false;
                }
                double adv = face.Font.Advance(gid) * size;
                cur.Gids.Add(gid);
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
            var p = new Piece { Kind = PieceKind.Tab, Style = st, Face = face, Size = st.Size };
            SetMetrics(p, st, face);
            return p;
        }

        foreach (var inl in para.Inlines)
        {
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
                    pieces.Add(new Piece { Kind = br.Kind == BreakKind.Line ? PieceKind.LineBreak : PieceKind.PageBreak, Style = br.Style });
                    breakNext = true;
                    break;
                case ImageInline img:
                {
                    double w = img.Width, h = img.Height;
                    if (w <= 0 || h <= 0) { w = 100; h = 75; }
                    double maxW = _contentWidthHint - para.IndLeft - para.IndRight;
                    if (maxW > 10 && w > maxW) { h *= maxW / w; w = maxW; }
                    if (h > _ctx.MaxImageHeight) { w *= _ctx.MaxImageHeight / h; h = _ctx.MaxImageHeight; }
                    pieces.Add(new Piece
                    {
                        Kind = PieceKind.Image, Style = img.Style, Image = img, Width = w, Ascent = h, Descent = 0,
                        BreakBefore = true, Size = img.Style.Size
                    });
                    breakNext = true;
                    break;
                }
            }
        }
        return pieces;
    }

    double _contentWidthHint = 500;

    sealed class RawLine
    {
        public List<Piece> Pieces = [];
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

    public List<LineFrag> LayoutParagraph(Paragraph p, double width)
    {
        _contentWidthHint = width;
        var pieces = Itemize(p);
        var raw = BreakLines(p, pieces, width);
        var result = new List<LineFrag>();
        for (int i = 0; i < raw.Count; i++)
            result.Add(BuildLine(p, raw[i], width, i == raw.Count - 1));

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

    List<RawLine> BreakLines(Paragraph p, List<Piece> pieces, double width)
    {
        var lines = new List<RawLine>();
        var cur = new RawLine { First = true };
        double x = 0;
        double extraRoom = 0;   // правая/центральная табуляция за правым полем позволяет вылезти за край

        double LineStart(bool first) => p.IndLeft + (first ? p.IndFirst : 0);
        double Avail(bool first) => width - p.IndRight - LineStart(first);
        bool HasContent() => cur.Pieces.Any(q => q.Kind is PieceKind.Text or PieceKind.Image);

        void Finish(PieceKind? endBreak)
        {
            cur.EndBreak = endBreak;
            lines.Add(cur);
            cur = new RawLine { First = false };
            x = 0;
            extraRoom = 0;
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
            while (end < pieces.Count && pieces[end].Kind is PieceKind.Text or PieceKind.Image && !pieces[end].BreakBefore)
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
            // Слово длиннее строки — режем по глифам
            double room = avail - x;
            int j2 = i;
            bool placed = false;
            for (; j2 < end; j2++)
            {
                var q = pieces[j2];
                if (q.Kind == PieceKind.Image || q.Width <= room)
                {
                    if (q.Kind == PieceKind.Image && q.Width > room && placed) break;
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
                    head.Gids.AddRange(q.Gids.Take(k)); head.Texts.AddRange(q.Texts.Take(k)); head.Adv.AddRange(q.Adv.Take(k));
                    head.Width = acc;
                    head.BreakBefore = q.BreakBefore;
                    cur.Pieces.Add(head);
                    var tail = q.CloneEmpty();
                    tail.Gids.AddRange(q.Gids.Skip(k)); tail.Texts.AddRange(q.Texts.Skip(k)); tail.Adv.AddRange(q.Adv.Skip(k));
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

    LineFrag BuildLine(Paragraph p, RawLine raw, double width, bool lastOfPara)
    {
        var pieces = raw.Pieces.Where(q => q.Kind is not (PieceKind.LineBreak or PieceKind.PageBreak)).ToList();
        double lineStart = p.IndLeft + (raw.First ? p.IndFirst : 0);
        double avail = width - p.IndRight - lineStart;

        // Метрики строки
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
            ParaLeft = p.IndLeft + Math.Min(0, p.IndFirst),
            ParaRight = width - p.IndRight,
            IsFirst = raw.First,
            IsLast = lastOfPara,
            PageBreakAfter = raw.EndBreak == PieceKind.PageBreak,
        };
        if (raw.First && p.BorderTop != null) lf.TopPad = p.BorderTop.Width + p.BorderTop.Space;
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
                    Shading = cell.Shading, VAlign = cell.VAlign, VMerge = cell.VMerge, GridIndex = g, Span = span,
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

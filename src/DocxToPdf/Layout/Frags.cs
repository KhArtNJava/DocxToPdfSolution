using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Pdf;

namespace DocxToPdf.Layout;

/// <summary>Общий контекст вёрстки.</summary>
public sealed class LayoutContext(DocxDocument doc, FontManager fonts)
{
    public readonly DocxDocument Doc = doc;
    public readonly FontManager Fonts = fonts;
    /// <summary>Текст полей PAGE/NUMPAGES (только для колонтитулов).</summary>
    public Func<FieldKind, string>? FieldText;
    /// <summary>Максимальная высота изображения (высота области текста).</summary>
    public double MaxImageHeight = 700;
}

/// <summary>Неделимый (или делимый, как строка таблицы) фрагмент потока, размещаемый на странице.</summary>
public abstract class Frag
{
    public double Height;
    public double SpaceBefore, SpaceAfter;
    public bool KeepWithNext, PageBreakBefore, PageBreakAfter;

    /// <summary>Рисует фрагмент; (x, y) — левый верхний угол (без SpaceBefore).</summary>
    public abstract void Draw(PdfCanvas c, double x, double y);

    /// <summary>Пытается разделить фрагмент так, чтобы первая часть уместилась в avail.</summary>
    public virtual (Frag First, Frag Remainder)? Split(double avail) => null;
}

public enum PieceKind { Text, Space, Tab, LineBreak, PageBreak, Image, Spacer }

/// <summary>Элемент строки: однородный по шрифту кусок текста, пробел, табуляция, картинка или разрыв.</summary>
public sealed class Piece
{
    public PieceKind Kind;
    public TextStyle Style = null!;
    public FontFace? Face;
    public double Size;
    public double Rise;
    public List<ushort> Gids = [];
    public List<string> Texts = [];
    public List<double> Adv = [];
    public double Width;
    public double Ascent, Descent;
    public bool BreakBefore;
    public ImageInline? Image;
    public TabStop? Tab;
    /// <summary>Для Spacer: чей это край (начало или конец строчного блока).</summary>
    public InlineBox? SpacerBox;
    public bool SpacerStart;
    /// <summary>Источник: индекс инлайна в абзаце и позиции символов каждого глифа.</summary>
    public int Src = -1;
    public List<int> Pos = [];
    public double X;

    public Piece CloneEmpty() => new()
    {
        Kind = Kind, Style = Style, Face = Face, Size = Size, Rise = Rise, Ascent = Ascent, Descent = Descent, Src = Src
    };
}

/// <summary>Одна строка абзаца.</summary>
public sealed class LineFrag : Frag
{
    public required List<Piece> Pieces;
    public required Paragraph Para;
    public required DocxDocument Doc;
    public double Left;          // смещение начала строки от левого края колонки
    public double Baseline;      // от верха строки (с учётом TopPad)
    public double ParaLeft, ParaRight;
    public double TopPad, BottomPad;
    public bool IsFirst, IsLast;

    public override void Draw(PdfCanvas c, double x, double y)
    {
        x += Para.OffsetX;
        y += Para.OffsetY;
        if (Para.Shading is { } shd)
            c.FillRect(x + ParaLeft, y, ParaRight - ParaLeft, Height, shd);
        if (IsFirst && Para.BorderTop is { } bt)
            DrawBorder(c, bt, x + ParaLeft, x + ParaRight, y + bt.Width / 2);
        if (IsLast && Para.BorderBottom is { } bb)
            DrawBorder(c, bb, x + ParaLeft, x + ParaRight, y + Height - bb.Width / 2);

        double baseY = y + TopPad + Baseline;
        double lx = x + Left;

        // Конечные пробелы не подчёркиваются
        int lastContent = Pieces.FindLastIndex(p => p.Kind is PieceKind.Text or PieceKind.Image or PieceKind.Tab);

        // 0. Строчные блоки (span с рамкой/фоном/отступами): сначала внешние
        DrawInlineBoxes(c, lx, baseY, lastContent);

        // 1. Фоны (выделение цветом)
        for (int i = 0; i < Pieces.Count; i++)
        {
            var p = Pieces[i];
            if (p.Style.Background is { } bg && p.Kind is PieceKind.Text or PieceKind.Space or PieceKind.Tab && i <= lastContent)
            {
                double w = NextX(i) - p.X;
                c.FillRect(lx + p.X, baseY - p.Ascent, w, p.Ascent + p.Descent, bg);
            }
        }

        // 2. Текст, сгруппированный по шрифту/цвету
        int k = 0;
        while (k < Pieces.Count)
        {
            var p = Pieces[k];
            if (p.Kind is not (PieceKind.Text or PieceKind.Space) || p.Face == null || p.Gids.Count == 0) { k++; continue; }
            var gids = new List<ushort>(); var texts = new List<string>(); var extra = new List<double>();
            bool anyExtra = false;
            int m = k;
            while (m < Pieces.Count)
            {
                var q = Pieces[m];
                if (q.Kind is not (PieceKind.Text or PieceKind.Space) || q.Face != p.Face || q.Size != p.Size || q.Rise != p.Rise || q.Style.Color != p.Style.Color || q.Gids.Count == 0)
                    break;
                gids.AddRange(q.Gids); texts.AddRange(q.Texts);
                double ls = q.Kind == PieceKind.Text ? q.Style.LetterSpacing : q.Style.WordSpacing;
                for (int g = 0; g < q.Gids.Count; g++) extra.Add(ls);
                if (ls != 0) anyExtra = true;
                // Зазор до следующего куска (выравнивание по ширине)
                double gap = m + 1 < Pieces.Count ? Pieces[m + 1].X - (q.X + q.Width) : 0;
                if (Math.Abs(gap) > 0.001) { extra[^1] += gap; anyExtra = true; }
                m++;
            }
            c.Glyphs(p.Face, p.Size, lx + p.X, baseY - p.Rise, gids, texts, p.Style.Color, anyExtra ? extra : null);
            k = m;
        }

        // 3. Картинки, заполнители табуляции, подчёркивания, ссылки
        for (int i = 0; i < Pieces.Count; i++)
        {
            var p = Pieces[i];
            double px = lx + p.X;
            if (p.Kind == PieceKind.Image && p.Image != null && Doc.Media.TryGetValue(p.Image.PartName, out var bytes))
            {
                double ih = p.Ascent + p.Descent;
                var tr = p.Image.Transform?.Invoke(p.Width, ih);
                var img = p.Image;
                double iy = baseY - p.Ascent, iw = p.Width;
                void DrawImg()
                {
                    var ir = img.Radius?.Resolve(iw, ih);
                    bool clip = PdfCanvas.HasRadius(ir);
                    if (clip) c.BeginRoundedClip(px, iy, iw, ih, ir!);
                    c.Image(img.PartName, bytes, px, iy, iw, ih);
                    if (clip) c.EndClip();
                }
                if (tr != null) TransformPainter.Paint(c, px, iy, iw, ih, tr, DrawImg);
                else DrawImg();
            }
            if (p.Kind == PieceKind.Tab && p.Tab != null && p.Tab.Leader is not ("none" or "") && p.Width > 4)
                DrawLeader(c, p, px, baseY);

            if (i > lastContent) continue;
            var st = p.Style;
            double w = NextX(i) - p.X;
            if (p.Kind == PieceKind.Image) w = p.Width;
            if ((st.Underline || st.Strike) && p.Face != null && p.Kind != PieceKind.Image)
            {
                var f = p.Face.Font;
                double size = st.Size;
                double thick = Math.Max(0.4, f.UnderlineThickness / (double)f.UnitsPerEm * size);
                if (st.Underline)
                {
                    double uy = baseY - p.Rise - f.UnderlinePosition / (double)f.UnitsPerEm * size + thick / 2;
                    c.Line(px, uy, px + w, uy, thick, st.Color);
                    if (st.DoubleUnderline) c.Line(px, uy + thick * 2, px + w, uy + thick * 2, thick, st.Color);
                }
                if (st.Strike)
                {
                    double sy = baseY - p.Rise - size * 0.27;
                    c.Line(px, sy, px + w, sy, thick, st.Color);
                }
            }
            if (st.Link != null)
                c.Link(px, baseY - p.Ascent, w, p.Ascent + p.Descent, st.Link);
        }
    }

    double NextX(int i) => i + 1 < Pieces.Count ? Pieces[i + 1].X : Pieces[i].X + Pieces[i].Width;

    void DrawInlineBoxes(PdfCanvas c, double lx, double baseY, int lastContent)
    {
        var boxes = new List<InlineBox>();
        foreach (var p in Pieces)
            for (var b = p.Style.Box; b != null; b = b.Parent)
                if (!boxes.Contains(b)) boxes.Add(b);
        if (boxes.Count == 0) return;

        foreach (var box in boxes.OrderBy(b => b.Depth))
        {
            int first = -1, last = -1;
            double asc = 0, desc = 0, sasc = 0, sdesc = 0;
            for (int i = 0; i < Pieces.Count; i++)
            {
                var p = Pieces[i];
                // Край блока принадлежит только своему блоку и его внешним блокам
                bool member = p.Kind == PieceKind.Spacer ? box.Contains(p.SpacerBox) : box.Contains(p.Style.Box);
                if (!member) continue;
                // конечные пробелы строки в блок не входят (кроме закрывающего края)
                if (i > lastContent && p.Kind != PieceKind.Spacer) continue;
                if (first < 0) first = i;
                last = i;
                if (p.Kind is PieceKind.Text or PieceKind.Image) { asc = Math.Max(asc, p.Ascent); desc = Math.Max(desc, p.Descent); }
                else { sasc = Math.Max(sasc, p.Ascent); sdesc = Math.Max(sdesc, p.Descent); }
            }
            if (first < 0) continue;
            if (asc == 0 && desc == 0) { asc = sasc; desc = sdesc; }
            // Для текста — область шрифта (без межстрочного запаса), как у браузеров
            var lp = Pieces[first];
            bool hasStart = lp.SpacerBox == box && lp.SpacerStart;
            var rp = Pieces[last];
            bool hasEnd = rp.SpacerBox == box && !rp.SpacerStart;
            double x1 = lx + lp.X + (hasStart ? box.MarginLeft : 0);
            double x2 = lx + rp.X + rp.Width - (hasEnd ? box.MarginRight : 0);
            if (x2 <= x1) continue;
            double bt = box.Top?.Width ?? 0, bb = box.Bottom?.Width ?? 0;
            double top = baseY - asc - box.PadTop - bt;
            double bottom = baseY + desc + box.PadBottom + bb;
            double[]? rad = null;
            if (box.Radius is { Any: true })
            {
                var r = box.Radius.Resolve(x2 - x1, bottom - top);
                // углы со стороны разрыва строки не скругляются
                rad = [hasStart ? r[0] : 0, hasEnd ? r[1] : 0, hasEnd ? r[2] : 0, hasStart ? r[3] : 0,
                       hasStart ? r[4] : 0, hasEnd ? r[5] : 0, hasEnd ? r[6] : 0, hasStart ? r[7] : 0];
            }
            if (box.Background is { } bg)
            {
                if (rad != null) c.FillRoundedRect(x1, top, x2 - x1, bottom - top, rad, bg);
                else c.FillRect(x1, top, x2 - x1, bottom - top, bg);
            }
            if (rad != null && hasStart && hasEnd && BorderDef.Same(box.Top, box.Bottom) && BorderDef.Same(box.Top, box.Left) && BorderDef.Same(box.Top, box.Right))
            {
                c.StrokeRoundedRect(x1, top, x2 - x1, bottom - top, rad, box.Top!);
                continue;
            }
            void L(BorderDef? b, double ax, double ay, double bx, double by)
            {
                if (b is { IsNone: false }) c.Line(ax, ay, bx, by, b.Width, b.Color, b.Val.Contains("dash"), b.Val.Contains("dot"));
            }
            L(box.Top, x1, top + bt / 2, x2, top + bt / 2);
            L(box.Bottom, x1, bottom - bb / 2, x2, bottom - bb / 2);
            if (hasStart) L(box.Left, x1 + (box.Left?.Width ?? 0) / 2, top, x1 + (box.Left?.Width ?? 0) / 2, bottom);
            if (hasEnd) L(box.Right, x2 - (box.Right?.Width ?? 0) / 2, top, x2 - (box.Right?.Width ?? 0) / 2, bottom);
        }
    }

    static void DrawBorder(PdfCanvas c, BorderDef b, double x1, double x2, double y)
    {
        if (b.Val == "double")
        {
            c.Line(x1, y - b.Width / 3, x2, y - b.Width / 3, b.Width / 3, b.Color);
            c.Line(x1, y + b.Width / 3, x2, y + b.Width / 3, b.Width / 3, b.Color);
        }
        else c.Line(x1, y, x2, y, b.Width, b.Color, b.Val.Contains("dash"), b.Val.Contains("dot"));
    }

    static void DrawLeader(PdfCanvas c, Piece tab, double x, double baseY)
    {
        var face = tab.Face;
        if (face == null) return;
        string leader = tab.Tab!.Leader;
        if (leader is "underscore" or "heavy")
        {
            double t = leader == "heavy" ? 1.2 : 0.6;
            c.Line(x + 1, baseY + 1.5, x + tab.Width - 1, baseY + 1.5, t, tab.Style.Color);
            return;
        }
        int ch = leader switch { "hyphen" => '-', "middleDot" => 0xB7, _ => '.' };
        if (!face.Font.TryGetGlyph(ch, out var gid)) return;
        double gw = face.Font.Advance(gid) * tab.Size;
        double step = gw * (leader == "hyphen" ? 1.3 : 1.6);
        if (step <= 0.1) return;
        // Точки выравниваются по сетке страницы, как в Word
        double start = Math.Ceiling((x + gw) / step) * step;
        double end = x + tab.Width - gw * 1.5;
        var gids = new List<ushort>(); var texts = new List<string>(); var extra = new List<double>();
        for (double px = start; px <= end; px += step)
        {
            gids.Add(gid); texts.Add(((char)ch).ToString()); extra.Add(step - gw);
        }
        if (gids.Count > 0) c.Glyphs(face, tab.Size, start, baseY, gids, texts, tab.Style.Color, extra);
    }
}

/// <summary>Ячейка размещённой строки таблицы.</summary>
public sealed class CellBox
{
    public double X, Width;
    public double MarginTop, MarginBottom, MarginLeft, MarginRight;
    public BorderDef? Top, Bottom, Left, Right;
    public Rgb? Shading;
    public string VAlign = "top";
    public RadiusSpec? Radius;
    public VMerge VMerge;
    public int GridIndex, Span;
    public List<(Frag Frag, double Y)> Items = [];
    public double ContentHeight;
    /// <summary>Для объединённой по вертикали ячейки — её содержимое рисуется поверх нескольких строк.</summary>
    public double MergedHeight;
}

public sealed class TableInfo
{
    public List<RowFrag> HeaderRows = [];
}

/// <summary>Строка таблицы (может разрываться между страницами).</summary>
public sealed class RowFrag : Frag
{
    public required TableInfo Table;
    public required List<CellBox> Cells;
    public bool IsHeader, CantSplit;

    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var cell in Cells)
            if (cell.Shading is { } shd && cell.VMerge != VMerge.Continue)
            {
                double ch = cell.VMerge == VMerge.Restart ? Math.Max(Height, cell.MergedHeight) : Height;
                if (cell.Radius is { Any: true }) c.FillRoundedRect(x + cell.X, y, cell.Width, ch, cell.Radius.Resolve(cell.Width, ch), shd);
                else c.FillRect(x + cell.X, y, cell.Width, ch, shd);
            }

        foreach (var cell in Cells)
        {
            if (cell.VMerge == VMerge.Continue) continue;
            double inner = (cell.VMerge == VMerge.Restart ? Math.Max(Height, cell.MergedHeight) : Height) - cell.MarginTop - cell.MarginBottom;
            double off = cell.VAlign switch
            {
                "center" => Math.Max(0, (inner - cell.ContentHeight) / 2),
                "bottom" => Math.Max(0, inner - cell.ContentHeight),
                _ => 0
            };
            foreach (var (f, fy) in cell.Items)
                f.Draw(c, x + cell.X + cell.MarginLeft, y + cell.MarginTop + off + fy);
        }

        foreach (var cell in Cells)
        {
            double x1 = x + cell.X, x2 = x1 + cell.Width, y1 = y, y2 = y + Height;
            if (cell.Radius is { Any: true } && BorderDef.Same(cell.Top, cell.Bottom) && BorderDef.Same(cell.Top, cell.Left) && BorderDef.Same(cell.Top, cell.Right))
            {
                c.StrokeRoundedRect(x1, y1, cell.Width, Height, cell.Radius.Resolve(cell.Width, Height), cell.Top!);
                continue;
            }
            if (cell.Top is { IsNone: false } t) DrawBorder(c, t, x1, y1, x2, y1);
            if (cell.Bottom is { IsNone: false } b) DrawBorder(c, b, x1, y2, x2, y2);
            if (cell.Left is { IsNone: false } l) DrawBorder(c, l, x1, y1, x1, y2);
            if (cell.Right is { IsNone: false } r) DrawBorder(c, r, x2, y1, x2, y2);
        }
    }

    static void DrawBorder(PdfCanvas c, BorderDef b, double x1, double y1, double x2, double y2)
    {
        if (b.Val == "double")
        {
            bool horiz = Math.Abs(y1 - y2) < 0.01;
            double d = b.Width / 3;
            if (horiz) { c.Line(x1, y1 - d, x2, y2 - d, d, b.Color); c.Line(x1, y1 + d, x2, y2 + d, d, b.Color); }
            else { c.Line(x1 - d, y1, x2 - d, y2, d, b.Color); c.Line(x1 + d, y1, x2 + d, y2, d, b.Color); }
            return;
        }
        c.Line(x1, y1, x2, y2, b.Width, b.Color, b.Val.Contains("dash"), b.Val.Contains("dot"));
    }

    public override (Frag First, Frag Remainder)? Split(double avail)
    {
        if (CantSplit) return null;
        var first = new List<CellBox>();
        var rest = new List<CellBox>();
        bool anyKept = false, anyMoved = false;
        double firstH = 0, restH = 0;
        foreach (var cell in Cells)
        {
            var a = Copy(cell); var b = Copy(cell);
            double cut = avail - cell.MarginTop - cell.MarginBottom;
            int idx = 0;
            while (idx < cell.Items.Count)
            {
                var (f, fy) = cell.Items[idx];
                if (fy + f.Height > cut + 0.01) break;
                a.Items.Add((f, fy));
                idx++;
            }
            if (cell.VMerge == VMerge.Continue) idx = cell.Items.Count;
            if (idx > 0) anyKept = true;
            if (idx < cell.Items.Count)
            {
                anyMoved = true;
                double shift = cell.Items[idx].Y - cell.Items[idx].Frag.SpaceBefore;
                for (int j = idx; j < cell.Items.Count; j++)
                    b.Items.Add((cell.Items[j].Frag, cell.Items[j].Y - shift));
                var lastB = b.Items[^1];
                b.ContentHeight = lastB.Y + lastB.Frag.Height + lastB.Frag.SpaceAfter;
            }
            else b.ContentHeight = 0;
            if (a.Items.Count > 0)
            {
                var lastA = a.Items[^1];
                a.ContentHeight = lastA.Y + lastA.Frag.Height;
            }
            else a.ContentHeight = 0;
            a.MergedHeight = b.MergedHeight = 0;
            firstH = Math.Max(firstH, a.ContentHeight + cell.MarginTop + cell.MarginBottom);
            restH = Math.Max(restH, b.ContentHeight + cell.MarginTop + cell.MarginBottom);
            first.Add(a); rest.Add(b);
        }
        if (!anyKept || !anyMoved) return null;
        var r1 = new RowFrag { Table = Table, Cells = first, Height = avail, IsHeader = IsHeader };
        var r2 = new RowFrag { Table = Table, Cells = rest, Height = restH, IsHeader = IsHeader, KeepWithNext = KeepWithNext, PageBreakAfter = PageBreakAfter, SpaceAfter = SpaceAfter };
        return (r1, r2);
    }

    static CellBox Copy(CellBox c) => new()
    {
        X = c.X, Width = c.Width, MarginTop = c.MarginTop, MarginBottom = c.MarginBottom, MarginLeft = c.MarginLeft, MarginRight = c.MarginRight,
        Top = c.Top, Bottom = c.Bottom, Left = c.Left, Right = c.Right, Shading = c.Shading, VAlign = c.VAlign, VMerge = c.VMerge, Radius = c.Radius,
        GridIndex = c.GridIndex, Span = c.Span
    };
}

/// <summary>Блок с фоном, рамкой и отступами, содержащий вложенные фрагменты (float, absolute).</summary>
public sealed class BoxFrag : Frag
{
    public double Width;
    public required BoxStyle Box;
    public List<(Frag Frag, double Y)> Items = [];

    public override void Draw(PdfCanvas c, double x, double y)
    {
        var rad = Box.Radius?.Resolve(Width, Height);
        bool round = PdfCanvas.HasRadius(rad);
        if (Box.Background is { } bg)
        {
            if (round) c.FillRoundedRect(x, y, Width, Height, rad!, bg);
            else c.FillRect(x, y, Width, Height, bg);
        }
        double cx = x + Box.PadLeft + (Box.BorderLeft?.Width ?? 0);
        double cy = y + Box.PadTop + (Box.BorderTop?.Width ?? 0);
        foreach (var (f, fy) in Items) f.Draw(c, cx, cy + fy);
        void B(BorderDef? b, double x1, double y1, double x2, double y2)
        {
            if (b is not { IsNone: false }) return;
            if (b.Val == "double")
            {
                // две линии по трети толщины у внешнего и внутреннего края
                double d = b.Width / 3;
                bool horiz = Math.Abs(y1 - y2) < 0.01;
                double sx = horiz ? 0 : d, sy = horiz ? d : 0;
                c.Line(x1 - sx, y1 - sy, x2 - sx, y2 - sy, d, b.Color);
                c.Line(x1 + sx, y1 + sy, x2 + sx, y2 + sy, d, b.Color);
                return;
            }
            c.Line(x1, y1, x2, y2, b.Width, b.Color, b.Val.Contains("dash"), b.Val.Contains("dot"));
        }
        if (round && BorderDef.Same(Box.BorderTop, Box.BorderBottom) && BorderDef.Same(Box.BorderTop, Box.BorderLeft) && BorderDef.Same(Box.BorderTop, Box.BorderRight))
        {
            c.StrokeRoundedRect(x, y, Width, Height, rad!, Box.BorderTop!);
            return;
        }
        double h2 = (Box.BorderTop?.Width ?? 0) / 2, w2 = (Box.BorderLeft?.Width ?? 0) / 2;
        B(Box.BorderTop, x, y + h2, x + Width, y + h2);
        B(Box.BorderBottom, x, y + Height - (Box.BorderBottom?.Width ?? 0) / 2, x + Width, y + Height - (Box.BorderBottom?.Width ?? 0) / 2);
        B(Box.BorderLeft, x + w2, y, x + w2, y + Height);
        B(Box.BorderRight, x + Width - (Box.BorderRight?.Width ?? 0) / 2, y, x + Width - (Box.BorderRight?.Width ?? 0) / 2, y + Height);
    }
}

/// <summary>Плавающие блоки вместе с обтекающим их текстом — размещаются на странице как единое целое.</summary>
public sealed class GroupFrag : Frag
{
    public List<(Frag Frag, double X, double Y)> Items = [];

    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var (f, fx, fy) in Items) f.Draw(c, x + fx, y + fy);
    }
}

/// <summary>Абсолютно позиционированный блок: в потоке не занимает места.</summary>
public sealed class AbsFrag : Frag
{
    public required BoxFrag Content;
    public double? Left, Top, Right, Bottom;
    public bool Fixed;
    public bool RelativeToFlow;          // относительно позиционированного предка
    public double AnchorX, AnchorWidth, TopOffset;
    public int ZIndex;

    /// <summary>Отрисовка внутри ячеек/групп (без участия разбивки на страницы).</summary>
    public override void Draw(PdfCanvas c, double x, double y)
    {
        double px = x + AnchorX + (Left ?? (Right != null ? AnchorWidth - Right.Value - Content.Width : 0));
        double py = y + TopOffset + (Top ?? 0);
        Content.Draw(c, px, py);
    }
}

/// <summary>Фрагменты, нарисованные с CSS-преобразованием (неделимы при разбивке на страницы).</summary>
public sealed class TransformFrag : Frag
{
    public List<(Frag Frag, double Y)> Items = [];
    public double BoxX, BoxWidth;
    public required BoxTransform T;

    public override void Draw(PdfCanvas c, double x, double y) =>
        TransformPainter.Paint(c, x + BoxX, y, BoxWidth, Height, T, () =>
        {
            foreach (var (f, fy) in Items) f.Draw(c, x, y + fy);
        });
}

/// <summary>
/// Рисует содержимое с преобразованием. Аффинное — одной матрицей PDF. Проективное (перспектива) —
/// кусочно: рамка делится на сетку плиток, каждая рисуется со своей аффинной матрицей
/// и обрезается по своему спроецированному четырёхугольнику.
/// </summary>
public static class TransformPainter
{
    public static void Paint(PdfCanvas c, double bx, double by, double w, double h, BoxTransform t, Action draw)
    {
        if (t.Hidden) return;
        var H = t.H;
        if (t.IsAffine)
        {
            double k = H[8];
            c.BeginTransform(bx, by, H[0] / k, H[3] / k, H[1] / k, H[4] / k, H[2] / k, H[5] / k);
            draw();
            c.EndTransform();
            return;
        }
        // Чем сильнее перспектива, тем мельче сетка
        double strength = (Math.Abs(H[6]) * w + Math.Abs(H[7]) * h) / Math.Max(1e-9, Math.Abs(H[8]));
        int n = Math.Clamp((int)Math.Ceiling(strength * 30), 4, 16);
        double dx = w / n, dy = h / n;
        // Плитки: текст контурами (иначе он повторился бы в каждой плитке), ссылки — только в текстовом слое
        c.TextAsPaths++;
        c.SuppressLinks++;
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n; j++)
            {
                double x0 = i * dx, x1 = x0 + dx, y0 = j * dy, y1 = y0 + dy;
                var p00 = t.Map(x0, y0); var p10 = t.Map(x1, y0); var p11 = t.Map(x1, y1); var p01 = t.Map(x0, y1);
                if (p00.W <= 1e-6 || p10.W <= 1e-6 || p11.W <= 1e-6 || p01.W <= 1e-6) continue;   // за наблюдателем
                // Аффинное приближение плитки по четырём углам
                double ax = ((p10.X - p00.X) + (p11.X - p01.X)) / (2 * dx), ay = ((p10.Y - p00.Y) + (p11.Y - p01.Y)) / (2 * dx);
                double cx = ((p01.X - p00.X) + (p11.X - p10.X)) / (2 * dy), cy = ((p01.Y - p00.Y) + (p11.Y - p10.Y)) / (2 * dy);
                double mx = (p00.X + p10.X + p11.X + p01.X) / 4, my = (p00.Y + p10.Y + p11.Y + p01.Y) / 4;
                double tx = (x0 + x1) / 2, ty = (y0 + y1) / 2;
                double e = mx - (ax * tx + cx * ty), f = my - (ay * tx + cy * ty);
                // Обрезка по четырёхугольнику, чуть расширенному, чтобы не было щелей между плитками
                (double, double) Grow((double X, double Y, double W) p)
                {
                    double vx = p.X - mx, vy = p.Y - my, len = Math.Sqrt(vx * vx + vy * vy);
                    double g = len > 1e-6 ? 0.35 / len : 0;
                    return (bx + p.X + vx * g, by + p.Y + vy * g);
                }
                c.BeginClipPolygon([Grow(p00), Grow(p10), Grow(p11), Grow(p01)]);
                c.BeginTransform(bx, by, ax, ay, cx, cy, e, f);
                draw();
                c.EndTransform();
                c.EndClip();
            }
        }
        c.SuppressLinks--;
        c.TextAsPaths--;

        // Один невидимый текстовый слой (поиск/копирование) с аффинным приближением всей рамки
        var q00 = t.Map(0, 0); var q10 = t.Map(w, 0); var q11 = t.Map(w, h); var q01 = t.Map(0, h);
        if (q00.W > 1e-6 && q10.W > 1e-6 && q11.W > 1e-6 && q01.W > 1e-6 && w > 0 && h > 0)
        {
            double ax = ((q10.X - q00.X) + (q11.X - q01.X)) / (2 * w), ay = ((q10.Y - q00.Y) + (q11.Y - q01.Y)) / (2 * w);
            double cx = ((q01.X - q00.X) + (q11.X - q10.X)) / (2 * h), cy = ((q01.Y - q00.Y) + (q11.Y - q10.Y)) / (2 * h);
            double mx = (q00.X + q10.X + q11.X + q01.X) / 4, my = (q00.Y + q10.Y + q11.Y + q01.Y) / 4;
            double e = mx - (ax * w / 2 + cx * h / 2), f = my - (ay * w / 2 + cy * h / 2);
            c.InvisibleTextOnly++;
            c.BeginTransform(bx, by, ax, ay, cx, cy, e, f);
            draw();
            c.EndTransform();
            c.InvisibleTextOnly--;
        }
    }
}

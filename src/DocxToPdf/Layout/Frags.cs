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
    /// <summary>Шаг сетки строк текущего раздела (docGrid), 0 — нет.</summary>
    public double GridPitch;
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

    /// <summary>Сноски, на которые ссылается фрагмент (их текст размещается внизу той же страницы).</summary>
    public virtual IEnumerable<List<Block>> Notes => [];

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
    /// <summary>Мягкие переносы: перед глифом с этим индексом допускается перенос с дефисом.</summary>
    public List<int>? Soft;
    /// <summary>Добавка после каждого глифа (распределённое выравнивание jc="distribute").</summary>
    public double GlyphExtra;
    public double Width;
    public double Ascent, Descent;
    public bool BreakBefore;
    public ImageInline? Image;
    public TabStop? Tab;
    /// <summary>Позиционная табуляция (w:ptab), из которой строится позиция табуляции.</summary>
    public TabInline? PTab;
    /// <summary>Для Spacer: чей это край (начало или конец строчного блока).</summary>
    public InlineBox? SpacerBox;
    public bool SpacerStart;
    /// <summary>Источник: индекс инлайна в абзаце и позиции символов каждого глифа.</summary>
    public int Src = -1;
    /// <summary>Поле PAGE/NUMPAGES: текст подставляется при отрисовке.</summary>
    public FieldKind? Field;
    /// <summary>Для PAGEREF — имя закладки.</summary>
    public string? FieldArg;
    /// <summary>Место примечания (выноска на правом поле).</summary>
    public CommentInfo? Comment;
    /// <summary>Направление: 0 — нейтральный, 1 — слева направо, 2 — справа налево.</summary>
    /// <summary>Уровень встраивания bidi (нечётный — справа налево).</summary>
    public byte Level;
    /// <summary>Разрыв колонки (а не страницы).</summary>
    public bool ColumnBreak;
    /// <summary>Текст сноски, на которую ссылается этот фрагмент.</summary>
    public List<Block>? Note;
    public List<int> Pos = [];
    public double X;

    public Piece CloneEmpty() => new()
    {
        Kind = Kind, Style = Style, Face = Face, Size = Size, Rise = Rise, Ascent = Ascent, Descent = Descent, Src = Src, Field = Field
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
    /// <summary>Строка заканчивается разрывом колонки.</summary>
    public bool ColumnBreakAfter;
    public List<List<Block>>? NoteBlocks;
    public override IEnumerable<List<Block>> Notes => NoteBlocks ?? [];

    public override void Draw(PdfCanvas c, double x, double y)
    {
        x += Para.OffsetX;
        y += Para.OffsetY;
        // Рамка/заливка: в группе абзацев с общей рамкой промежутки между ними тоже внутри рамки
        double boxTop = y - (IsFirst && Para.GroupPrev ? SpaceBefore : 0);
        double boxBottom = y + Height + (IsLast && Para.GroupNext ? SpaceAfter : 0);
        var bl = Para.BorderLeft; var br = Para.BorderRight;
        double bx1 = x + ParaLeft - (bl != null ? bl.Space + bl.Width : 0);
        double bx2 = x + ParaRight + (br != null ? br.Space + br.Width : 0);
        if (Para.Shading is { } shd)
            c.FillRect(bx1, boxTop, bx2 - bx1, boxBottom - boxTop, shd);
        var topB = IsFirst ? (Para.GroupPrev ? Para.BorderBetween : Para.BorderTop) : null;
        if (topB != null) DrawBorder(c, topB, bx1, bx2, y + topB.Width / 2);
        if (IsLast && !Para.GroupNext && Para.BorderBottom is { } bb)
            DrawBorder(c, bb, bx1, bx2, y + Height - bb.Width / 2);
        if (bl != null) c.Line(bx1 + bl.Width / 2, boxTop, bx1 + bl.Width / 2, boxBottom, bl.Width, bl.Color, bl.Val.Contains("dash"), bl.Val.Contains("dot"));
        if (br != null) c.Line(bx2 - br.Width / 2, boxTop, bx2 - br.Width / 2, boxBottom, br.Width, br.Color, br.Val.Contains("dash"), br.Val.Contains("dot"));

        double baseY = y + TopPad + Baseline;
        double lx = x + Left;

        // Черта изменений на левом поле у строк с исправлениями
        if (Pieces.Any(p => p.Style.RevisionMark)) c.Line(x - 8, y, x - 8, y + Height, 0.75, Rgb.Black);

        // Барная табуляция: вертикальная черта на всю высоту строки (позиция — от левого поля)
        foreach (var t in Para.Tabs)
            if (t.Align == "bar") c.Line(x + t.Pos, y, x + t.Pos, y + Height, 0.5, Rgb.Black);

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
            if (p.Field != null && c.PageNumber > 0)
            {
                // Номер страницы известен только сейчас
                int m2 = k;
                while (m2 < Pieces.Count && Pieces[m2].Field == p.Field && Pieces[m2].FieldArg == p.FieldArg && Pieces[m2].Face == p.Face) m2++;
                string? num = p.Field switch
                {
                    FieldKind.Page => c.PageNumberText ?? c.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    FieldKind.NumPages => c.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    _ => p.FieldArg != null ? c.PageRefText?.Invoke(p.FieldArg) : null,
                };
                // Неизвестная закладка — сохранённый результат как есть
                num ??= string.Concat(Pieces.Skip(k).Take(m2 - k).SelectMany(q => q.Texts));
                var fg = new List<ushort>(); var ft = new List<string>();
                double nw = 0;
                foreach (char ch in num)
                {
                    p.Face.Font.TryGetGlyph(ch, out var g); fg.Add(g); ft.Add(ch.ToString());
                    nw += p.Face.Font.Advance(g) * p.Size * p.Style.HScale;
                }
                // PAGEREF (оглавление): номер выравнивается по правому краю оценённой ширины
                double placeholderW = (m2 < Pieces.Count ? Pieces[m2].X : Pieces[m2 - 1].X + Pieces[m2 - 1].Width) - p.X;
                double dx = p.Field == FieldKind.PageRef ? Math.Max(0, placeholderW - nw) : 0;
                c.Glyphs(p.Face, p.Size, lx + p.X + dx, baseY - p.Rise, fg, ft, p.Style.Color, null, p.Style.HScale, p.Style.Outline);
                k = m2;
                continue;
            }
            var gids = new List<ushort>(); var texts = new List<string>(); var extra = new List<double>();
            bool anyExtra = false;
            int m = k;
            while (m < Pieces.Count)
            {
                var q = Pieces[m];
                if (q.Kind is not (PieceKind.Text or PieceKind.Space) || q.Face != p.Face || q.Size != p.Size || q.Rise != p.Rise || q.Style.Color != p.Style.Color || q.Gids.Count == 0
                    || q.Field != null || q.Style.HScale != p.Style.HScale || q.Style.Outline != p.Style.Outline || q.Style.ShadowColor != p.Style.ShadowColor)
                    break;
                gids.AddRange(q.Gids); texts.AddRange(q.Texts);
                // Добавка к естественной ширине глифа: разрядка, интервал между словами, кернинг (всё уже в Adv)
                double natK = q.Size * q.Style.HScale;
                for (int g = 0; g < q.Gids.Count; g++)
                {
                    double e = (g < q.Adv.Count ? q.Adv[g] - q.Face!.Font.Advance(q.Gids[g]) * natK : 0) + q.GlyphExtra;
                    if (Math.Abs(e) < 1e-6) e = 0; else anyExtra = true;
                    extra.Add(e);
                }
                // Зазор до следующего куска (выравнивание по ширине)
                double gap = m + 1 < Pieces.Count ? Pieces[m + 1].X - (q.X + q.Width + q.GlyphExtra * q.Gids.Count) : 0;
                if (Math.Abs(gap) > 0.001) { extra[^1] += gap; anyExtra = true; }
                m++;
            }
            if (p.Style.ShadowColor is { } shc)
            {
                // Тень/рельеф: смещённая копия текста под ним; невидима для поиска (не дублирует текст)
                double so = p.Size * p.Style.ShadowOffset;
                c.TextAsPaths++;
                c.Glyphs(p.Face, p.Size, lx + p.X + so, baseY - p.Rise + so, gids, texts, shc, anyExtra ? extra : null, p.Style.HScale, p.Style.Outline);
                c.TextAsPaths--;
            }
            c.Glyphs(p.Face, p.Size, lx + p.X, baseY - p.Rise, gids, texts, p.Style.Color, anyExtra ? extra : null, p.Style.HScale, p.Style.Outline);
            k = m;
        }

        // 3. Картинки, заполнители табуляции, подчёркивания, ссылки
        for (int i = 0; i < Pieces.Count; i++)
        {
            var p = Pieces[i];
            double px = lx + p.X;
            if (p.Kind == PieceKind.Image && p.Image is { } obj && (obj.Shape != null || obj.Content != null || obj.Children != null || obj.Chart != null || obj.Math != null))
            {
                // Фигура, надпись или группа: рамка неповёрнутая, поворот вокруг центра
                double sh = p.Ascent + p.Descent, sk = obj.Width > 0 ? p.Width / obj.Width : 1;
                DrawObject(c, Doc, obj, px + p.Width / 2, baseY - p.Ascent + sh / 2, sk);
            }
            else if (p.Kind == PieceKind.Image && p.Image != null && Doc.Media.TryGetValue(p.Image.PartName, out var bytes))
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
                    if (img.HasGeometry)
                    {
                        double k = img.Width > 0 ? iw / img.Width : 1;
                        c.ImageGeometry(img.PartName, bytes, px + iw / 2, iy + ih / 2, img.FrameW * k, img.FrameH * k,
                            img.CropL, img.CropT, img.CropR, img.CropB, img.Rotation, img.FlipH, img.FlipV, img.Opacity, img.Grayscale);
                    }
                    else c.Image(img.PartName, bytes, px, iy, iw, ih, img.Opacity, img.Grayscale);
                    if (clip) c.EndClip();
                }
                if (tr != null) TransformPainter.Paint(c, px, iy, iw, ih, tr, DrawImg);
                else DrawImg();
                if (img.Border is { } ib)
                {
                    // Рамка картинки по краю рамки (повёрнутой вместе с ней)
                    double k2 = img.Width > 0 ? iw / img.Width : 1;
                    var frame = new ShapeSpec { Geom = "rect", Stroke = ib.Color, StrokeWidth = ib.Width };
                    c.Shape(frame, px + iw / 2, iy + ih / 2, img.FrameW * k2, img.FrameH * k2, img.Rotation, img.FlipH, img.FlipV);
                }
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
                // «Только слова»: пробелы не подчёркиваются
                if (st.Underline && !(st.UnderlineStyle == "words" && p.Kind is PieceKind.Space or PieceKind.Tab))
                {
                    double uy = baseY - p.Rise - f.UnderlinePosition / (double)f.UnitsPerEm * size + thick / 2;
                    var ucol = st.UnderlineColor ?? st.Color;
                    string us = st.UnderlineStyle ?? "single";
                    double ut = us.Contains("thick", StringComparison.OrdinalIgnoreCase) || us.Contains("Heavy") ? thick * 2 : thick;
                    if (us.StartsWith("wav", StringComparison.OrdinalIgnoreCase))
                    {
                        c.WavyLine(px, uy + ut, w, ut, ucol);
                        if (st.DoubleUnderline) c.WavyLine(px, uy + ut * 4, w, ut, ucol);
                    }
                    else
                    {
                        bool dotted = us.StartsWith("dotted") || us == "dottedHeavy";
                        bool dashed = us.StartsWith("dash", StringComparison.OrdinalIgnoreCase) || us.StartsWith("dot");
                        c.Line(px, uy, px + w, uy, ut, ucol, dashed && !dotted, dotted);
                        if (st.DoubleUnderline) c.Line(px, uy + thick * 2, px + w, uy + thick * 2, thick, ucol);
                    }
                }
                if (st.Strike)
                {
                    double sy = baseY - p.Rise - size * 0.27;
                    if (st.DoubleStrike)
                    {
                        c.Line(px, sy - thick, px + w, sy - thick, thick, st.Color);
                        c.Line(px, sy + thick, px + w, sy + thick, thick, st.Color);
                    }
                    else c.Line(px, sy, px + w, sy, thick, st.Color);
                }
            }
            if (st.Emphasis != null && p.Kind == PieceKind.Text && p.Adv.Count == p.Texts.Count)
            {
                // Знаки ударения (w:em): точка над каждым символом (underDot — под ним)
                double r = Math.Max(0.5, st.Size * 0.07), gx = px;
                double ey = st.Emphasis == "underDot" ? baseY - p.Rise + st.Size * 0.28 : baseY - p.Rise - st.Size * 0.98;
                for (int g = 0; g < p.Texts.Count; g++)
                {
                    double adv = p.Adv[g] + p.GlyphExtra;
                    if (!string.IsNullOrWhiteSpace(p.Texts[g]))
                    {
                        double cx = gx + p.Adv[g] / 2;
                        var radii = new[] { r, r, r, r, r, r, r, r };
                        if (st.Emphasis == "circle")
                            c.StrokeRoundedRect(cx - r * 1.3, ey - r * 1.3, r * 2.6, r * 2.6, radii.Select(v => v * 1.3).ToArray(), new BorderDef { Width = Math.Max(0.3, r * 0.5), Color = st.Color });
                        else c.FillRoundedRect(cx - r, ey - r, 2 * r, 2 * r, radii, st.Color);
                    }
                    gx += adv;
                }
            }
            if (st.Link != null)
                c.Link(px, baseY - p.Ascent, w, p.Ascent + p.Descent, st.Link);
        }
    }

    /// <summary>
    /// Рисует фигуру, надпись, картинку или группу с центром (cx, cy) в масштабе k
    /// (рамка FrameW×FrameH, поворот и отражение вокруг центра).
    /// </summary>
    internal static void DrawObject(PdfCanvas c, DocxDocument doc, ImageInline obj, double cx, double cy, double k)
    {
        double fw = obj.FrameW * k, fh = obj.FrameH * k;
        if (obj.Shape != null)
        {
            c.Shape(obj.Shape, cx, cy, fw, fh, obj.Rotation, obj.FlipH, obj.FlipV);
            if (obj.Built == null) return;
        }
        bool rotated = obj.Rotation != 0 || obj.FlipH || obj.FlipV;
        if (rotated)
        {
            double a = obj.Rotation * Math.PI / 180, cos = Math.Cos(a), sin = Math.Sin(a);
            double fx = obj.FlipH ? -1 : 1, fy = obj.FlipV ? -1 : 1;
            // Поворот по часовой стрелке (ось y вниз) после отражения
            c.BeginTransform(cx, cy, cos * fx, sin * fx, -sin * fy, cos * fy, 0, 0);
        }
        if (obj.Children != null)
        {
            double gk = obj.FrameW > 0 ? fw / obj.FrameW : k;
            foreach (var (child, x, y) in obj.Children)
                DrawObject(c, doc, child, cx - fw / 2 + (x + child.FrameW / 2) * gk, cy - fh / 2 + (y + child.FrameH / 2) * gk, gk);
        }
        else if (obj.Built is Frag f)
        {
            // Текст в фигуре: по вертикали сверху, по центру или снизу (bodyPr anchor)
            double dy = obj.VAnchor switch { "ctr" => (fh - f.Height * k) / 2, "b" => fh - f.Height * k, _ => 0 };
            double ox = cx - fw / 2, oy = cy - fh / 2 + Math.Max(0, dy);
            if (Math.Abs(k - 1) > 0.001) c.BeginTransform(ox, oy, k, 0, 0, k, 0, 0);
            f.Draw(c, ox, oy);
            if (Math.Abs(k - 1) > 0.001) c.EndTransform();
        }
        else if (doc.Media.TryGetValue(obj.PartName, out var bytes))
        {
            // Картинка внутри группы
            if (obj.CropL != 0 || obj.CropT != 0 || obj.CropR != 0 || obj.CropB != 0)
                c.ImageGeometry(obj.PartName, bytes, cx, cy, fw, fh, obj.CropL, obj.CropT, obj.CropR, obj.CropB, 0, false, false, obj.Opacity, obj.Grayscale);
            else c.Image(obj.PartName, bytes, cx - fw / 2, cy - fh / 2, fw, fh, obj.Opacity, obj.Grayscale);
            if (obj.Border is { } b) c.Shape(new ShapeSpec { Stroke = b.Color, StrokeWidth = b.Width }, cx, cy, fw, fh, 0, false, false);
        }
        if (rotated) c.EndTransform();
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
    /// <summary>Точная высота строки (hRule="exact"): содержимое ячеек обрезается по строке.</summary>
    public bool ClipContent;
    /// <summary>Интервалы между ячейками: отступ ячеек сверху/снизу в строке и внешняя рамка таблицы.</summary>
    public double InsetTop, InsetBottom, OuterX, OuterWidth;
    public BorderDef? OuterTop, OuterBottom, OuterLeft, OuterRight;
    public override IEnumerable<List<Block>> Notes => Cells.SelectMany(c => c.Items.SelectMany(i => i.Frag.Notes));

    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var cell in Cells)
            if (cell.Shading is { } shd && cell.VMerge != VMerge.Continue)
            {
                double ch = (cell.VMerge == VMerge.Restart ? Math.Max(Height, cell.MergedHeight) : Height) - InsetTop - InsetBottom;
                if (cell.Radius is { Any: true }) c.FillRoundedRect(x + cell.X, y + InsetTop, cell.Width, ch, cell.Radius.Resolve(cell.Width, ch), shd);
                else c.FillRect(x + cell.X, y + InsetTop, cell.Width, ch, shd);
            }

        foreach (var cell in Cells)
        {
            if (cell.VMerge == VMerge.Continue) continue;
            double inner = (cell.VMerge == VMerge.Restart ? Math.Max(Height, cell.MergedHeight) : Height) - cell.MarginTop - cell.MarginBottom - InsetTop - InsetBottom;
            double off = cell.VAlign switch
            {
                "center" => Math.Max(0, (inner - cell.ContentHeight) / 2),
                "bottom" => Math.Max(0, inner - cell.ContentHeight),
                _ => 0
            };
            bool clip = ClipContent && cell.VMerge == VMerge.None;
            if (clip) c.BeginRoundedClip(x + cell.X, y, cell.Width, Height, new double[8]);
            foreach (var (f, fy) in cell.Items)
                f.Draw(c, x + cell.X + cell.MarginLeft, y + InsetTop + cell.MarginTop + off + fy);
            if (clip) c.EndClip();
        }

        foreach (var cell in Cells)
        {
            double x1 = x + cell.X, x2 = x1 + cell.Width, y1 = y + InsetTop, y2 = y + Height - InsetBottom;
            if (cell.Radius is { Any: true } && BorderDef.Same(cell.Top, cell.Bottom) && BorderDef.Same(cell.Top, cell.Left) && BorderDef.Same(cell.Top, cell.Right))
            {
                c.StrokeRoundedRect(x1, y1, cell.Width, y2 - y1, cell.Radius.Resolve(cell.Width, y2 - y1), cell.Top!);
                continue;
            }
            if (cell.Radius is { Any: true })
            {
                // Скруглены не все углы или стороны разные (угол области подсетки): стороны одной рамки — одним контуром с дугами
                var rad = cell.Radius.Resolve(cell.Width, y2 - y1);
                // Начало объединения по вертикали: нижние углы — у последней строки объединения
                if (cell.VMerge == VMerge.Restart) rad[2] = rad[3] = rad[6] = rad[7] = 0;
                BorderDef?[] sides = [cell.Top, cell.Right, cell.Bottom, cell.Left];
                var done = new bool[4];
                for (int i = 0; i < 4; i++)
                {
                    if (done[i] || sides[i] is not { IsNone: false } bd) continue;
                    var mask = new bool[4];
                    for (int j = 0; j < 4; j++)
                        if (sides[j] is { IsNone: false } o && BorderDef.Same(o, bd)) { mask[j] = true; done[j] = true; }
                    c.StrokeRoundedSides(x1, y1, cell.Width, y2 - y1, rad, bd, mask);
                }
                continue;
            }
            if (cell.Top is { IsNone: false } t) DrawBorder(c, t, x1, y1, x2, y1);
            if (cell.Bottom is { IsNone: false } b) DrawBorder(c, b, x1, y2, x2, y2);
            if (cell.Left is { IsNone: false } l) DrawBorder(c, l, x1, y1, x1, y2);
            if (cell.Right is { IsNone: false } r) DrawBorder(c, r, x2, y1, x2, y2);
        }

        // Внешняя рамка таблицы с интервалами между ячейками
        if (OuterWidth > 0)
        {
            double ox1 = x + OuterX, ox2 = ox1 + OuterWidth, oy2 = y + Height;
            if (OuterTop is { IsNone: false } ot) DrawBorder(c, ot, ox1, y, ox2, y);
            if (OuterBottom is { IsNone: false } ob) DrawBorder(c, ob, ox1, oy2, ox2, oy2);
            if (OuterLeft is { IsNone: false } ol) DrawBorder(c, ol, ox1, y, ox1, oy2);
            if (OuterRight is { IsNone: false } or) DrawBorder(c, or, ox2, y, ox2, oy2);
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

    /// <summary>
    /// Разрыв блока между страницами (рамка «срезается»: у первой части нет нижней рамки и отступа,
    /// у второй — верхних): содержимое делится между строками/строками таблицы, пересекающий элемент разрезается сам.
    /// </summary>
    public override (Frag First, Frag Remainder)? Split(double avail)
    {
        double topIn = Box.PadTop + (Box.BorderTop?.Width ?? 0), botIn = Box.PadBottom + (Box.BorderBottom?.Width ?? 0);
        double cut = avail - topIn;
        if (cut <= 1) return null;
        int splitIdx = -1;
        (Frag First, Frag Remainder)? splitParts = null;
        for (int i = 0; i < Items.Count; i++)
        {
            var (f, fy) = Items[i];
            if (fy >= cut - 0.01 || fy + f.Height <= cut + 0.01) continue;
            var parts = f.Split(cut - fy);
            if (parts != null && parts.Value.First.Height <= cut - fy + 0.01) { splitIdx = i; splitParts = parts; }
            else cut = fy;
            break;
        }
        if (cut <= 0.5) return null;
        var bs1 = new BoxStyle { Background = Box.Background, BorderTop = Box.BorderTop, BorderLeft = Box.BorderLeft, BorderRight = Box.BorderRight,
            PadTop = Box.PadTop, PadLeft = Box.PadLeft, PadRight = Box.PadRight };
        var bs2 = new BoxStyle { Background = Box.Background, BorderBottom = Box.BorderBottom, BorderLeft = Box.BorderLeft, BorderRight = Box.BorderRight,
            PadBottom = Box.PadBottom, PadLeft = Box.PadLeft, PadRight = Box.PadRight };
        var first = new BoxFrag { Box = bs1, Width = Width, Height = topIn + cut, SpaceBefore = SpaceBefore };
        double contentH = Height - topIn - botIn;
        var rest = new BoxFrag { Box = bs2, Width = Width, Height = Math.Max(0, contentH - cut) + botIn, SpaceAfter = SpaceAfter };
        for (int i = 0; i < Items.Count; i++)
        {
            var (f, fy) = Items[i];
            if (i == splitIdx) { first.Items.Add((splitParts!.Value.First, fy)); rest.Items.Add((splitParts.Value.Remainder, 0)); }
            else if (fy < cut && fy + f.Height <= cut + 0.01) first.Items.Add((f, fy));
            else rest.Items.Add((f, fy - cut));
        }
        if (first.Items.Count == 0) return null;
        return (first, rest);
    }

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
    /// <summary>Объекты с привязкой к странице/полям: верх объекта в группе (без MarginTop) и его высота.</summary>
    public List<(FloatBlock Fb, double Top, double Height)> AbsFloats = [];
    /// <summary>
    /// Пересборка группы, если её первый неразрываемый элемент (картинка) не помещается в остаток страницы:
    /// текст группы — во всю ширину в остаток страницы (First), плавающий блок с остальным текстом — дальше (Rest).
    /// Заменяет группу и следующие за ней TailFrags фрагментов (строки последнего абзаца вне группы).
    /// </summary>
    public Func<double, (List<Frag> First, List<Frag> Remainder, int Consumed)?>? Rebuild;
    public int TailFrags;
    public override IEnumerable<List<Block>> Notes => Items.SelectMany(i => i.Frag.Notes);

    public override void Draw(PdfCanvas c, double x, double y)
    {
        foreach (var (f, fx, fy) in Items) f.Draw(c, x + fx, y + fy);
    }

    /// <summary>
    /// Разрыв группы между страницами: разрез сдвигается вверх, пока его пересекают неразрываемые элементы
    /// (строки, картинки); пересекающие его разрываемые элементы (блок float, строки таблицы) разрезаются сами.
    /// Строки рядом с float остаются суженными и переходят на следующую страницу вместе с ним.
    /// </summary>
    public override (Frag First, Frag Remainder)? Split(double avail)
    {
        double cut = Math.Min(avail, Height);
        var splits = new Dictionary<int, (Frag First, Frag Remainder)>();
        for (int guard = 0; guard < 200; guard++)
        {
            bool changed = false;
            splits.Clear();
            for (int i = 0; i < Items.Count; i++)
            {
                var (f, _, fy) = Items[i];
                if (fy >= cut - 0.01 || fy + f.Height <= cut + 0.01) continue;
                var parts = f is GroupFrag or LineFrag ? null : f.Split(cut - fy);
                if (parts != null && parts.Value.First.Height <= cut - fy + 0.01) { splits[i] = parts.Value; continue; }
                cut = fy;   // элемент целиком переходит на следующую страницу
                changed = true;
                break;
            }
            if (!changed) break;
        }
        if (cut <= 0.5) return null;
        var first = new GroupFrag { Height = cut, SpaceBefore = SpaceBefore, PageBreakBefore = PageBreakBefore };
        var rest = new GroupFrag { Height = Math.Max(0, Height - cut), SpaceAfter = SpaceAfter, KeepWithNext = KeepWithNext };
        for (int i = 0; i < Items.Count; i++)
        {
            var (f, fx, fy) = Items[i];
            if (splits.TryGetValue(i, out var sp))
            {
                first.Items.Add((sp.First, fx, fy));
                rest.Items.Add((sp.Remainder, fx, 0));
            }
            else if (fy + f.Height <= cut + 0.01 && fy < cut) first.Items.Add((f, fx, fy));
            else rest.Items.Add((f, fx, fy - cut));
        }
        foreach (var a in AbsFloats) (a.Top < cut ? first : rest).AbsFloats.Add(a.Top < cut ? a : (a.Fb, a.Top - cut, a.Height));
        if (first.Items.Count == 0) return null;
        return (first, rest);
    }
}

/// <summary>Переход к следующей колонке/странице (после текста, вынесенного из группы перед плавающим блоком).</summary>
public sealed class BreakFrag : Frag
{
    public override void Draw(PdfCanvas c, double x, double y) { }
}

/// <summary>Метка конца позиционированного предка (нулевой высоты).</summary>
public sealed class AnchorEndFrag : Frag
{
    public required AnchorBlock Anchor;
    public double PadBottom;
    public override void Draw(PdfCanvas c, double x, double y) { }
}

/// <summary>Абсолютно позиционированный блок: в потоке не занимает места.</summary>
public sealed class AbsFrag : Frag
{
    public required BoxFrag Content;
    public double? Left, Top, Right, Bottom;
    public bool Fixed;
    public bool RelativeToFlow;          // относительно позиционированного предка
    /// <summary>Позиционированный предок (для bottom — его высота известна по метке конца).</summary>
    public AnchorBlock? Anchor;
    public double AnchorX, AnchorWidth, TopOffset;
    public int ZIndex;
    public string? HRel, VRel, AlignX, AlignY;

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
    // ---------------- 3D-сцена (transform-style: preserve-3d) ----------------

    sealed class Plane
    {
        public double Bx, By, W, H;
        public required double[] A;          // накопленная матрица 4×4 в координатах страницы
        public required Action Draw;
        public bool Extract;                 // preserve-3d: преобразованные потомки — отдельные плоскости
        public bool BackfaceHidden;
        public int Order;
    }

    sealed class Scene3D
    {
        public readonly List<Plane> Planes = [];
        public readonly Stack<double[]> Accum = new();
    }

    static double[] Mul4(double[] a, double[] b)
    {
        var r = new double[16];
        for (int i = 0; i < 4; i++)
            for (int j = 0; j < 4; j++)
                r[i * 4 + j] = a[i * 4] * b[j] + a[i * 4 + 1] * b[4 + j] + a[i * 4 + 2] * b[8 + j] + a[i * 4 + 3] * b[12 + j];
        return r;
    }

    static double[] T4(double x, double y) => [1, 0, 0, x, 0, 1, 0, y, 0, 0, 1, 0, 0, 0, 0, 1];

    /// <summary>Матрица элемента в координатах страницы (рамка в точке bx, by).</summary>
    static double[] PageMatrix(BoxTransform t, double bx, double by)
    {
        var h = t.H;
        var m = t.M ?? [h[0], h[1], 0, h[2], h[3], h[4], 0, h[5], 0, 0, 1, 0, h[6], h[7], 0, h[8]];
        return Mul4(Mul4(T4(bx, by), m), T4(-bx, -by));
    }

    static (double X, double Y, double Z, double W) Apply(double[] a, double x, double y) =>
        (a[0] * x + a[1] * y + a[3], a[4] * x + a[5] * y + a[7], a[8] * x + a[9] * y + a[11], a[12] * x + a[13] * y + a[15]);

    /// <summary>Проекция плоскости сцены в локальные координаты её рамки (3×3).</summary>
    static double[] LocalHomography(double[] a, double bx, double by)
    {
        // H = T(-bx,-by) · A[строки 0,1,3; столбцы 0,1,3] · T(bx,by)
        double[] m = [a[0], a[1], a[3], a[4], a[5], a[7], a[12], a[13], a[15]];
        double[] post = [1, 0, bx, 0, 1, by, 0, 0, 1];
        double[] pre = [1, 0, -bx, 0, 1, -by, 0, 0, 1];
        static double[] M3(double[] p, double[] q)
        {
            var r = new double[9];
            for (int i = 0; i < 3; i++)
                for (int j = 0; j < 3; j++)
                    r[i * 3 + j] = p[i * 3] * q[j] + p[i * 3 + 1] * q[3 + j] + p[i * 3 + 2] * q[6 + j];
            return r;
        }
        return M3(M3(pre, m), post);
    }

    /// <summary>Повёрнута ли плоскость обратной стороной: ориентация спроецированного контура сменилась.</summary>
    static bool IsBackFacing(BoxTransform t, double w, double h)
    {
        var p = new[] { t.Map(0, 0), t.Map(w, 0), t.Map(w, h), t.Map(0, h) };
        if (p.Any(q => q.W <= 1e-9)) return false;
        double area = 0;
        for (int i = 0; i < 4; i++)
        {
            var (x1, y1, _) = p[i]; var (x2, y2, _) = p[(i + 1) % 4];
            area += x1 * y2 - x2 * y1;
        }
        return area < 0;
    }

    static void RenderScene(PdfCanvas c, double bx, double by, double w, double h, BoxTransform t, Action draw)
    {
        var scene = new Scene3D();
        var rootA = PageMatrix(t, bx, by);
        scene.Planes.Add(new Plane { Bx = bx, By = by, W = w, H = h, A = rootA, Draw = draw, Extract = true, BackfaceHidden = t.BackfaceHidden });
        scene.Accum.Push(rootA);

        // Холостой проход: собираем плоскости потомков
        var prev = c.Scene;
        c.Scene = scene;
        c.Collecting++;
        draw();
        c.Collecting--;
        c.Scene = prev;

        // Алгоритм художника: от дальних плоскостей к ближним (z к зрителю — положительный)
        double Depth(Plane p)
        {
            var (_, _, z, wv) = Apply(p.A, p.Bx + p.W / 2, p.By + p.H / 2);
            return Math.Abs(wv) > 1e-12 ? z / wv : z;
        }
        foreach (var p in scene.Planes.OrderBy(Depth).ThenBy(p => p.Order))
        {
            var bt = new BoxTransform { H = LocalHomography(p.A, p.Bx, p.By) };
            if (p.BackfaceHidden && IsBackFacing(bt, p.W, p.H)) continue;
            var plane = p;
            Action body = plane.Extract
                ? () => { c.SkipTransformed++; plane.Draw(); c.SkipTransformed--; }
                : plane.Draw;
            PaintProjected(c, plane.Bx, plane.By, plane.W, plane.H, bt, body);
        }
    }

    public static void Paint(PdfCanvas c, double bx, double by, double w, double h, BoxTransform t, Action draw)
    {
        if (c.Collecting > 0 && c.Scene is Scene3D scene)
        {
            // Сбор плоскости: накапливаем матрицу; потомки preserve-3d разворачиваются рекурсивно
            var acc = Mul4(scene.Accum.Peek(), PageMatrix(t, bx, by));
            scene.Planes.Add(new Plane
            {
                Bx = bx, By = by, W = w, H = h, A = acc, Draw = draw, Extract = t.Preserve3D,
                BackfaceHidden = t.BackfaceHidden, Order = scene.Planes.Count
            });
            if (t.Preserve3D)
            {
                scene.Accum.Push(acc);
                draw();
                scene.Accum.Pop();
            }
            return;
        }
        if (c.SkipTransformed > 0) return;   // рисуется отдельной плоскостью сцены
        if (t.Preserve3D) { RenderScene(c, bx, by, w, h, t, draw); return; }
        PaintProjected(c, bx, by, w, h, t, draw);
    }

    /// <summary>Площадь спроецированной рамки (со знаком); ~0 — плоскость видна «с ребра».</summary>
    static double ProjectedArea(BoxTransform t, double w, double h)
    {
        var p = new[] { t.Map(0, 0), t.Map(w, 0), t.Map(w, h), t.Map(0, h) };
        if (p.Any(q => q.W <= 1e-9)) return double.NaN;
        double area = 0;
        for (int i = 0; i < 4; i++) area += p[i].X * p[(i + 1) % 4].Y - p[(i + 1) % 4].X * p[i].Y;
        return area / 2;
    }

    static void PaintProjected(PdfCanvas c, double bx, double by, double w, double h, BoxTransform t, Action draw)
    {
        if (t.Hidden) return;
        // Плоскость, повёрнутая ребром к зрителю, не рисуется (как в браузерах)
        double pa = ProjectedArea(t, w, h);
        if (!double.IsNaN(pa) && w * h > 1 && Math.Abs(pa) < 1e-3 * w * h) return;
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
        if (c.InvisibleTextOnly > 0)
        {
            // Текстовый слой внешней сцены: достаточно аффинного приближения
            DrawInvisibleText(c, bx, by, w, h, t, draw);
            return;
        }
        // Содержимое рисуется один раз в форму (прозрачностная группа: при сглаживании краёв плиток
        // слои внутри не просвечивают); текст — контурами, иначе он повторялся бы в каждой плитке
        string form = c.CaptureForm(() => { c.TextAsPaths++; draw(); c.TextAsPaths--; });
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
                // Обрезка по спроецированной плитке, расширенной на ~0,6 pt по каждой оси
                // (иначе при сглаживании между плитками просвечивают швы)
                double sx = Math.Sqrt(ax * ax + ay * ay), sy = Math.Sqrt(cx * cx + cy * cy);
                double ex = sx > 1e-9 ? 0.6 / sx : 0, ey = sy > 1e-9 ? 0.6 / sy : 0;
                var c00 = t.Map(x0 - ex, y0 - ey); var c10 = t.Map(x1 + ex, y0 - ey);
                var c11 = t.Map(x1 + ex, y1 + ey); var c01 = t.Map(x0 - ex, y1 + ey);
                if (c00.W <= 1e-6 || c10.W <= 1e-6 || c11.W <= 1e-6 || c01.W <= 1e-6) { c00 = p00; c10 = p10; c11 = p11; c01 = p01; }
                c.BeginClipPolygon([(bx + c00.X, by + c00.Y), (bx + c10.X, by + c10.Y), (bx + c11.X, by + c11.Y), (bx + c01.X, by + c01.Y)]);
                c.BeginTransform(bx, by, ax, ay, cx, cy, e, f);
                c.DoForm(form);
                c.EndTransform();
                c.EndClip();
            }
        }
        DrawInvisibleText(c, bx, by, w, h, t, draw);
    }

    /// <summary>Один невидимый текстовый слой (поиск/копирование, ссылки) с аффинным приближением всей рамки.</summary>
    static void DrawInvisibleText(PdfCanvas c, double bx, double by, double w, double h, BoxTransform t, Action draw)
    {
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

/// <summary>Содержимое ячейки с повёрнутым текстом (снизу вверх или сверху вниз).</summary>
public sealed class RotatedFrag : Frag
{
    public List<(Frag Frag, double Y)> Items = [];
    public double Length;        // длина строк (вдоль поворота) = высота области
    public double Thickness;     // суммарная высота строк до поворота
    public double CellWidth;     // ширина ячейки
    public string Direction = "btLr";

    public override void Draw(PdfCanvas c, double x, double y)
    {
        // Локальные координаты содержимого (u вдоль строки, v поперёк) → область ячейки
        double off = Math.Max(0, (CellWidth - Thickness) / 2);
        if (Direction == "btLr")
            c.BeginTransform(x, y, 0, -1, 1, 0, off, Length);          // (u,v) → (v, L − u)
        else
            c.BeginTransform(x, y, 0, 1, -1, 0, off + Thickness, 0);   // (u,v) → (T − v, u)
        foreach (var (f, fy) in Items) f.Draw(c, x, y + fy);
        c.EndTransform();
    }
}

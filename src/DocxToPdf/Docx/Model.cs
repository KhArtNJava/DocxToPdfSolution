namespace DocxToPdf.Docx;

/// <summary>Цвет RGB (компоненты 0..1).</summary>
public readonly record struct Rgb(double R, double G, double B)
{
    public static readonly Rgb Black = new(0, 0, 0);
    public static readonly Rgb White = new(1, 1, 1);

    public static Rgb? Parse(string? hex)
    {
        if (string.IsNullOrEmpty(hex) || hex.Equals("auto", StringComparison.OrdinalIgnoreCase) || hex.Length != 6)
            return null;
        try
        {
            int v = Convert.ToInt32(hex, 16);
            return new Rgb(((v >> 16) & 255) / 255.0, ((v >> 8) & 255) / 255.0, (v & 255) / 255.0);
        }
        catch { return null; }
    }

    public static Rgb? FromHighlight(string? name) => name switch
    {
        "yellow" => new Rgb(1, 1, 0),
        "green" => new Rgb(0, 1, 0),
        "cyan" => new Rgb(0, 1, 1),
        "magenta" => new Rgb(1, 0, 1),
        "blue" => new Rgb(0, 0, 1),
        "red" => new Rgb(1, 0, 0),
        "darkBlue" => new Rgb(0, 0, 0.5),
        "darkCyan" => new Rgb(0, 0.5, 0.5),
        "darkGreen" => new Rgb(0, 0.5, 0),
        "darkMagenta" => new Rgb(0.5, 0, 0.5),
        "darkRed" => new Rgb(0.5, 0, 0),
        "darkYellow" => new Rgb(0.5, 0.5, 0),
        "darkGray" => new Rgb(0.5, 0.5, 0.5),
        "lightGray" => new Rgb(0.75, 0.75, 0.75),
        "black" => new Rgb(0, 0, 0),
        "white" => new Rgb(1, 1, 1),
        _ => null
    };
}

// ------------------------------------------------------------------
//  "Сырые" свойства (как в XML, все поля nullable — для наследования стилей)
// ------------------------------------------------------------------

public sealed class RunProps
{
    public string? FontAscii, FontHAnsi, FontEastAsia;
    public string? ThemeAscii, ThemeHAnsi, ThemeEastAsia;
    public bool? Bold, Italic, Strike, DStrike, Caps, SmallCaps, Vanish;
    public string? Underline;
    public string? UnderlineColor;
    public double? Size;          // pt
    public string? Color;
    public string? Highlight;
    public string? Shading;
    public string? VertAlign;
    public string? StyleId;       // rStyle
    public double? Position;      // pt (поднятие/опускание)
    public double? Spacing;       // pt (межбуквенный интервал)
    public double? Scale;         // w:w — масштаб символов по горизонтали, %
    public bool? Outline;         // контурный текст
    public bool? Emboss, Imprint, Shadow;   // эффекты текста
    public BorderDef? Border;     // w:bdr — рамка вокруг текста
    /// <summary>Свойства для сложных письменностей (иврит, арабский…): w:szCs, w:bCs, w:iCs, w:rFonts w:cs.</summary>
    public double? SizeCs;
    public bool? BoldCs, ItalicCs;
    public string? FontCs;
    /// <summary>w:rtl / w:cs — весь фрагмент оформляется свойствами сложных письменностей.</summary>
    public bool? Rtl, Cs;
    /// <summary>Знаки ударения w:em: dot, comma, circle, underDot.</summary>
    public string? Emphasis;

    public RunProps Clone() => (RunProps)MemberwiseClone();

    /// <summary>Накладывает заданные свойства поверх текущих.</summary>
    public void Overlay(RunProps? o)
    {
        if (o == null) return;
        if (o.FontAscii != null) { FontAscii = o.FontAscii; ThemeAscii = null; }
        if (o.FontHAnsi != null) { FontHAnsi = o.FontHAnsi; ThemeHAnsi = null; }
        if (o.FontEastAsia != null) { FontEastAsia = o.FontEastAsia; ThemeEastAsia = null; }
        if (o.ThemeAscii != null) ThemeAscii = o.ThemeAscii;
        if (o.ThemeHAnsi != null) ThemeHAnsi = o.ThemeHAnsi;
        if (o.ThemeEastAsia != null) ThemeEastAsia = o.ThemeEastAsia;
        Bold = o.Bold ?? Bold;
        Italic = o.Italic ?? Italic;
        Strike = o.Strike ?? Strike;
        DStrike = o.DStrike ?? DStrike;
        Caps = o.Caps ?? Caps;
        SmallCaps = o.SmallCaps ?? SmallCaps;
        Vanish = o.Vanish ?? Vanish;
        Underline = o.Underline ?? Underline;
        UnderlineColor = o.Underline != null ? o.UnderlineColor : UnderlineColor;
        Size = o.Size ?? Size;
        Color = o.Color ?? Color;
        Highlight = o.Highlight ?? Highlight;
        Shading = o.Shading ?? Shading;
        VertAlign = o.VertAlign ?? VertAlign;
        Position = o.Position ?? Position;
        Scale = o.Scale ?? Scale;
        Outline = o.Outline ?? Outline;
        Emboss = o.Emboss ?? Emboss; Imprint = o.Imprint ?? Imprint; Shadow = o.Shadow ?? Shadow;
        Border = o.Border ?? Border;
        Spacing = o.Spacing ?? Spacing;
        SizeCs = o.SizeCs ?? SizeCs; BoldCs = o.BoldCs ?? BoldCs; ItalicCs = o.ItalicCs ?? ItalicCs;
        FontCs = o.FontCs ?? FontCs; Rtl = o.Rtl ?? Rtl; Cs = o.Cs ?? Cs;
        Emphasis = o.Emphasis ?? Emphasis;
        if (o.StyleId != null) StyleId = o.StyleId;
    }
}

public sealed record TabStop(double Pos, string Align, string Leader);

public sealed class BorderDef
{
    /// <summary>Одинаковы ли две рамки (для обводки скруглённым контуром).</summary>
    public static bool Same(BorderDef? a, BorderDef? b) =>
        a is { IsNone: false } && b is { IsNone: false } && a.Val == b.Val && Math.Abs(a.Width - b.Width) < 0.01 && a.Color == b.Color;

    public string Val = "single";
    public double Width = 0.5;    // pt
    public double Space;          // pt
    public Rgb Color = Rgb.Black;
    public bool IsNone => Val is "nil" or "none";
}

public sealed class ParaProps
{
    public string? StyleId;
    public string? Jc;
    public double? IndLeft, IndRight, IndFirstLine, IndHanging;
    public double? SpBefore, SpAfter;
    public bool? BeforeAuto, AfterAuto;
    public double? Line;          // pt (exact/atLeast) либо 240-е доли (auto)
    public string? LineRule;
    public int? NumId, Ilvl;
    public bool? PageBreakBefore, KeepNext, KeepLines, ContextualSpacing, WidowControl, SnapToGrid, Bidi, SuppressLineNumbers;
    public int? DropLines;
    public double? DropGap;
    public List<TabStop>? Tabs;   // включая val="clear"
    public string? Shading;
    public BorderDef? BorderTop, BorderBottom, BorderLeft, BorderRight, BorderBetween;
    public RunProps? MarkRunProps;

    public ParaProps Clone()
    {
        var c = (ParaProps)MemberwiseClone();
        if (Tabs != null) c.Tabs = new List<TabStop>(Tabs);
        c.MarkRunProps = MarkRunProps?.Clone();
        return c;
    }

    public void Overlay(ParaProps? o)
    {
        if (o == null) return;
        if (o.StyleId != null) StyleId = o.StyleId;
        Jc = o.Jc ?? Jc;
        IndLeft = o.IndLeft ?? IndLeft;
        IndRight = o.IndRight ?? IndRight;
        if (o.IndFirstLine != null) { IndFirstLine = o.IndFirstLine; IndHanging = null; }
        if (o.IndHanging != null) { IndHanging = o.IndHanging; IndFirstLine = null; }
        SpBefore = o.SpBefore ?? SpBefore;
        SpAfter = o.SpAfter ?? SpAfter;
        BeforeAuto = o.BeforeAuto ?? BeforeAuto;
        AfterAuto = o.AfterAuto ?? AfterAuto;
        if (o.Line != null) { Line = o.Line; LineRule = o.LineRule ?? "auto"; }
        NumId = o.NumId ?? NumId;
        Ilvl = o.Ilvl ?? Ilvl;
        PageBreakBefore = o.PageBreakBefore ?? PageBreakBefore;
        KeepNext = o.KeepNext ?? KeepNext;
        KeepLines = o.KeepLines ?? KeepLines;
        ContextualSpacing = o.ContextualSpacing ?? ContextualSpacing;
        WidowControl = o.WidowControl ?? WidowControl;
        SnapToGrid = o.SnapToGrid ?? SnapToGrid;
        SuppressLineNumbers = o.SuppressLineNumbers ?? SuppressLineNumbers;
        Bidi = o.Bidi ?? Bidi;
        if (o.Tabs != null)
        {
            Tabs ??= new List<TabStop>();
            foreach (var t in o.Tabs)
            {
                Tabs.RemoveAll(x => Math.Abs(x.Pos - t.Pos) < 0.5);
                Tabs.Add(t);
            }
        }
        Shading = o.Shading ?? Shading;
        BorderTop = o.BorderTop ?? BorderTop;
        BorderBottom = o.BorderBottom ?? BorderBottom;
        BorderLeft = o.BorderLeft ?? BorderLeft;
        BorderRight = o.BorderRight ?? BorderRight;
        BorderBetween = o.BorderBetween ?? BorderBetween;
        if (o.MarkRunProps != null)
        {
            MarkRunProps ??= new RunProps();
            MarkRunProps.Overlay(o.MarkRunProps);
        }
    }
}

// ------------------------------------------------------------------
//  Итоговая (вычисленная) модель документа, готовая к вёрстке
// ------------------------------------------------------------------

/// <summary>
/// Строчный блок (inline box): рамка, фон и горизонтальные отступы вокруг фрагмента текста.
/// Вложенные блоки образуют цепочку через Parent.
/// </summary>
public sealed class InlineBox
{
    public InlineBox? Parent;
    public Rgb? Background;
    public BorderDef? Top, Bottom, Left, Right;
    public double PadTop, PadBottom, PadLeft, PadRight;
    public double MarginLeft, MarginRight;
    public RadiusSpec? Radius;

    public double StartWidth => MarginLeft + (Left?.Width ?? 0) + PadLeft;
    public double EndWidth => MarginRight + (Right?.Width ?? 0) + PadRight;

    public int Depth { get { int d = 0; for (var b = Parent; b != null; b = b.Parent) d++; return d; } }

    public bool Contains(InlineBox? inner)
    {
        for (var b = inner; b != null; b = b.Parent) if (b == this) return true;
        return false;
    }

    /// <summary>Цепочка от внешнего блока к внутреннему.</summary>
    public static List<InlineBox> Chain(InlineBox? inner)
    {
        var list = new List<InlineBox>();
        for (var b = inner; b != null; b = b.Parent) list.Insert(0, b);
        return list;
    }
}

/// <summary>Окончательный стиль фрагмента текста.</summary>
public sealed record TextStyle
{
    public string FontAscii { get; init; } = "Times New Roman";
    public string FontHAnsi { get; init; } = "Times New Roman";
    public string FontEastAsia { get; init; } = "Times New Roman";
    public bool Bold { get; init; }
    /// <summary>Числовая насыщенность (CSS font-weight); 0 — по Bold (400/700). Для вариативных шрифтов задаёт ось wght.</summary>
    public int Weight { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool DoubleUnderline { get; init; }
    /// <summary>Стиль подчёркивания Word: single, thick, dotted, dash, wave, words …</summary>
    public string? UnderlineStyle { get; init; }
    /// <summary>Цвет подчёркивания (w:u w:color); null — цвет текста.</summary>
    public Rgb? UnderlineColor { get; init; }
    public bool DoubleStrike { get; init; }
    /// <summary>Фрагмент исправления (вставка/удаление/перемещение) — черта изменений на левом поле.</summary>
    public bool RevisionMark { get; init; }
    /// <summary>Знаки ударения (w:em): dot, comma, circle — над символами, underDot — под ними.</summary>
    public string? Emphasis { get; init; }
    public bool Strike { get; init; }
    public bool Caps { get; init; }
    /// <summary>Капитель: строчные буквы — уменьшенными прописными.</summary>
    public bool SmallCaps { get; init; }
    /// <summary>Дополнительный интервал после каждого символа, pt (letter-spacing / w:spacing).</summary>
    public double LetterSpacing { get; init; }
    /// <summary>Дополнительный интервал у пробелов, pt (word-spacing).</summary>
    public double WordSpacing { get; init; }
    /// <summary>Горизонтальный масштаб символов (w:w), 1 — 100%.</summary>
    public double HScale { get; init; } = 1;
    /// <summary>Контурный текст (w:outline).</summary>
    public bool Outline { get; init; }
    /// <summary>Тень под текстом (w:shadow, w:emboss, w:imprint): цвет тени и её смещение в долях кегля.</summary>
    public Rgb? ShadowColor { get; init; }
    public double ShadowOffset { get; init; }
    /// <summary>Рамка вокруг текста (w:bdr) — превращается в общий строчный блок соседних фрагментов.</summary>
    public BorderDef? RunBorder { get; init; }
    public double Size { get; init; } = 10;
    public Rgb Color { get; init; } = Rgb.Black;
    public Rgb? Background { get; init; }
    public int VertAlign { get; init; }   // 0 — нет, 1 — superscript, -1 — subscript
    public double Position { get; init; }
    public string? Link { get; init; }
    /// <summary>Строчный блок с рамкой/фоном/отступами (CSS border у span и т.п.); сравнивается по ссылке.</summary>
    public InlineBox? Box { get; init; }
}

public abstract class Inline { }

public sealed class TextInline(string text, TextStyle style) : Inline
{
    public string Text = text;
    public TextStyle Style = style;
}

public sealed class TabInline(TextStyle style) : Inline
{
    public TextStyle Style = style;
    /// <summary>Позиционная табуляция (w:ptab): выравнивание left/center/right.</summary>
    public string? PtabAlign;
    /// <summary>Для w:ptab: отсчёт от полей ("margin") или от отступов абзаца ("indent").</summary>
    public string PtabRelativeTo = "margin";
    public string PtabLeader = "none";
}

public enum BreakKind { Line, Page, Column }

public sealed class BreakInline(BreakKind kind, TextStyle style) : Inline
{
    public BreakKind Kind = kind;
    public TextStyle Style = style;
}

public sealed class ImageInline(string partName, double width, double height, TextStyle style) : Inline
{
    public TransformResolver? Transform;
    public RadiusSpec? Radius;           // скругление картинки (обрезка)
    public string? VAlign;               // vertical-align: middle / top / bottom …
    public string PartName = partName;
    public double Width = width;   // pt
    public double Height = height; // pt
    public TextStyle Style = style;
    /// <summary>Обрезка (a:srcRect) в долях исходной картинки: слева, сверху, справа, снизу.</summary>
    public double CropL, CropT, CropR, CropB;
    /// <summary>Поворот по часовой стрелке, градусы (a:xfrm rot).</summary>
    public double Rotation;
    public bool FlipH, FlipV;
    /// <summary>Размер неповёрнутой рамки; Width/Height — описанный прямоугольник после поворота.</summary>
    public double FrameW = width, FrameH = height;
    /// <summary>Фигура без текста (DrawingML prstGeom / VML): рисуется контурами вместо картинки.</summary>
    public ShapeSpec? Shape;
    /// <summary>Надпись в строке: содержимое и оформление рамки (как строчный блок).</summary>
    public List<Block>? Content;
    public BoxStyle? Box;
    /// <summary>Группа фигур: элементы с координатами внутри рамки группы (pt).</summary>
    public List<(ImageInline Item, double X, double Y)>? Children;
    /// <summary>Рамка картинки (pic:spPr/a:ln).</summary>
    public BorderDef? Border;
    /// <summary>Свёрстанное содержимое надписи (заполняется при вёрстке).</summary>
    public object? Built;
    /// <summary>Вертикальное положение текста в фигуре (bodyPr anchor): t, ctr, b.</summary>
    public string VAnchor = "t";
    /// <summary>Непрозрачность картинки (a:alphaModFix), 1 — непрозрачна.</summary>
    public double Opacity = 1;
    /// <summary>Картинка в оттенках серого (a:grayscl).</summary>
    public bool Grayscale;
    /// <summary>Диаграмма (c:chart).</summary>
    public ChartSpec? Chart;
    /// <summary>Формула OMML (двумерная вёрстка); Display — отдельной строкой (m:oMathPara).</summary>
    public MathNode? Math;
    public bool MathDisplay;
    /// <summary>Глубина формулы под базовой линией (заполняется при вёрстке).</summary>
    public double MathDescent;
    public bool HasGeometry => CropL != 0 || CropT != 0 || CropR != 0 || CropB != 0 || Rotation != 0 || FlipH || FlipV;
}

/// <summary>Простая фигура: прямоугольник, скруглённый прямоугольник, эллипс, линия, треугольник и т. п.</summary>
public sealed class ShapeSpec
{
    /// <summary>rect, roundRect, ellipse, line, triangle, rtTriangle, diamond, parallelogram, hexagon, octagon, rightArrow, leftArrow, upArrow, downArrow</summary>
    public string Geom = "rect";
    public Rgb? Fill;
    public Rgb? Stroke;
    public double StrokeWidth = 0.75;
    public bool Dashed, Dotted;
    /// <summary>Стрелки на концах линии (a:headEnd / a:tailEnd).</summary>
    public bool HeadArrow, TailArrow;
    /// <summary>Произвольный контур (a:custGeom): подконтуры в долях рамки 0…1.</summary>
    public List<CustomPath>? Paths;
    /// <summary>Градиентная заливка: опорные точки (0…1) и угол направления (градусы, по часовой от оси x).</summary>
    public List<(double Pos, Rgb Color)>? Gradient;
    public double GradientAngle;
    /// <summary>Радиальный градиент (a:path circle/rect/shape): центр (доли рамки), иначе — линейный.</summary>
    public bool GradientRadial;
    public double GradientCx = 0.5, GradientCy = 0.5;
    /// <summary>Первый параметр фигуры (a:avLst adj), в долях 100000; null — по умолчанию.</summary>
    public double? Adj;
}

/// <summary>Диаграмма DrawingML: тип, ряды, категории, заголовок, легенда.</summary>
public sealed class ChartSpec
{
    /// <summary>bar (горизонтальные), col (столбцы), line, area, pie, doughnut, scatter</summary>
    public string Type = "col";
    /// <summary>clustered, stacked, percentStacked, standard</summary>
    public string Grouping = "clustered";
    public string? Title;
    public List<string> Categories = [];
    public List<ChartSeries> Series = [];
    public bool Legend = true;
    public string LegendPos = "r";
    public bool Gridlines = true;
    public bool DataLabels;
    /// <summary>Объёмный вид (bar3DChart, pie3DChart …).</summary>
    public bool ThreeD;
    /// <summary>Есть ряды на вторичной оси значений (справа).</summary>
    public bool HasSecondary => Series.Any(s => s.Secondary);
}

public sealed class ChartSeries
{
    public string Name = "";
    public List<double> Values = [];
    public Rgb? Color;
    /// <summary>Тип ряда (у комбинированных диаграмм свой у каждой группы): col, bar, line, area, scatter, pie …</summary>
    public string Type = "col";
    public string Grouping = "clustered";
    /// <summary>Ряд на вторичной оси значений.</summary>
    public bool Secondary;
    /// <summary>Маркер точек: circle, square, diamond, triangle, x, star, plus, dash, dot; null — без маркеров.</summary>
    public string? Marker;
    public double MarkerSize = 5;
    /// <summary>Соединять точки линией (точечная диаграмма может быть только с маркерами).</summary>
    public bool Line = true;
    /// <summary>Цвета отдельных точек (круговая диаграмма, c:dPt).</summary>
    public Dictionary<int, Rgb> PointColors = [];
}

/// <summary>Подконтур a:path: команды M/L/C/Z с точками в долях рамки.</summary>
public sealed class CustomPath
{
    public List<(char Cmd, double[] Pts)> Commands = [];
    public bool Fill = true, Stroke = true;
}

/// <summary>Параметры чтения DOCX: показ исправлений и примечаний.</summary>
public sealed class DocxOptions
{
    /// <summary>auto (как в настройках документа; по умолчанию — с разметкой), final — итоговый вид, markup — с разметкой.</summary>
    public string Revisions = "auto";
    /// <summary>auto (выноски, если позволяет правое поле, иначе — список в конце), balloons, end, none.</summary>
    public string Comments = "auto";
}

/// <summary>Примечание (w:comment): номер, автор, текст; цвет автора.</summary>
public sealed class CommentInfo
{
    public int Number;
    public string Author = "", Initials = "";
    public List<Block> Blocks = [];
    public Rgb Color;
}

/// <summary>Место примечания в тексте (конец диапазона): выноска на правом поле.</summary>
public sealed class CommentInline(CommentInfo info, TextStyle style) : Inline
{
    public CommentInfo Info = info;
    public TextStyle Style = style;
}

// ---------------- Формулы (OMML) ----------------

public abstract class MathNode { }
/// <summary>Последовательность элементов формулы.</summary>
public sealed class MRow(List<MathNode> items) : MathNode { public List<MathNode> Items = items; }
/// <summary>Текст формулы: Plain — прямой шрифт (m:sty p, m:nor), иначе буквы курсивом.</summary>
public sealed class MText(string text, bool plain, bool bold) : MathNode { public string Text = text; public bool Plain = plain, Bold = bold; }
public sealed class MFrac(MathNode num, MathNode den, string type) : MathNode { public MathNode Num = num, Den = den; public string Type = type; }
public sealed class MScript(MathNode b, MathNode? sub, MathNode? sup, bool pre) : MathNode { public MathNode Base = b; public MathNode? Sub = sub, Sup = sup; public bool Pre = pre; }
public sealed class MRad(MathNode? deg, MathNode e) : MathNode { public MathNode? Deg = deg; public MathNode E = e; }
public sealed class MNary(string chr, MathNode? sub, MathNode? sup, MathNode e, bool undOvr) : MathNode { public string Chr = chr; public MathNode? Sub = sub, Sup = sup; public MathNode E = e; public bool UndOvr = undOvr; }
public sealed class MDelim(string beg, string end, string sep, List<MathNode> items) : MathNode { public string Beg = beg, End = end, Sep = sep; public List<MathNode> Items = items; }
public sealed class MMatrix(List<List<MathNode>> rows) : MathNode { public List<List<MathNode>> Rows = rows; }
public sealed class MFunc(MathNode name, MathNode e) : MathNode { public MathNode Name = name, E = e; }
public sealed class MAcc(string chr, MathNode e) : MathNode { public string Chr = chr; public MathNode E = e; }
public sealed class MBar(bool top, MathNode e) : MathNode { public bool Top = top; public MathNode E = e; }
public sealed class MGroupChr(string chr, bool top, MathNode e) : MathNode { public string Chr = chr; public bool Top = top; public MathNode E = e; }
public sealed class MLim(MathNode e, MathNode lim, bool upper) : MathNode { public MathNode E = e, Lim = lim; public bool Upper = upper; }
public sealed class MEqArr(List<MathNode> rows) : MathNode { public List<MathNode> Rows = rows; }
public sealed class MBox(MathNode e, bool border) : MathNode { public MathNode E = e; public bool Border = border; }

public enum FieldKind { Page, NumPages, PageRef }

/// <summary>Ссылка на сноску: номер в тексте и текст сноски внизу страницы.</summary>
public sealed class FootnoteInline(string mark, TextStyle style, List<Block> blocks) : Inline
{
    public string Mark = mark;
    public TextStyle Style = style;
    public List<Block> Blocks = blocks;
}

public sealed class FieldInline(FieldKind kind, TextStyle style) : Inline
{
    public FieldKind Kind = kind;
    public TextStyle Style = style;
    /// <summary>Для PAGEREF — имя закладки.</summary>
    public string? Arg;
    /// <summary>Сохранённый в документе результат (для оценки ширины).</summary>
    public string? Cached;
}

public abstract class Block { }

public sealed class Paragraph : Block
{
    public string StyleId = "";
    public string Align = "left";        // left, center, right, both
    public double IndLeft, IndRight, IndFirst;  // IndFirst < 0 — выступ
    public double SpaceBefore, SpaceAfter;
    public double LineValue = 240;
    public string LineRule = "auto";
    public bool PageBreakBefore, KeepNext, KeepLines, ContextualSpacing, WidowControl = true;
    public List<TabStop> Tabs = [];
    public Rgb? Shading;
    public BorderDef? BorderTop, BorderBottom, BorderLeft, BorderRight, BorderBetween;
    /// <summary>Закладки, начинающиеся в абзаце (цели внутренних ссылок).</summary>
    public List<string>? Bookmarks;
    /// <summary>Выравнивание номера списка (lvlJc): left/center/right.</summary>
    public string? LabelAlign;
    /// <summary>Привязка строк к сетке документа (docGrid).</summary>
    public bool SnapToGrid = true;
    /// <summary>w:suppressLineNumbers — строки абзаца не нумеруются.</summary>
    public bool SuppressLineNumbers;
    /// <summary>Буквица (w:framePr dropCap): высота в строках, отступ до текста.</summary>
    public int DropLines;
    public double DropGap;
    /// <summary>Абзац справа налево (w:bidi).</summary>
    public bool Rtl;
    /// <summary>Абзац входит в группу с общей рамкой (соседи с одинаковыми границами).</summary>
    public bool GroupPrev, GroupNext;
    public TextStyle MarkStyle = new();
    public List<Inline> Inlines = [];
    /// <summary>Свойства раздела, если абзац завершает раздел.</summary>
    public SectionProps? Section;
    /// <summary>CSS clear: left/right/both.</summary>
    public string? Clear;
    /// <summary>Сдвиг при position: relative.</summary>
    public double OffsetX, OffsetY;
    /// <summary>CSS ::first-line — преобразование стиля текста первой строки.</summary>
    public Func<TextStyle, TextStyle>? FirstLine;

    public Paragraph ShallowCopy() => (Paragraph)MemberwiseClone();
}

public sealed class CellBorders
{
    public BorderDef? Top, Bottom, Left, Right, InsideH, InsideV;

    public void Overlay(CellBorders? o)
    {
        if (o == null) return;
        Top = o.Top ?? Top; Bottom = o.Bottom ?? Bottom;
        Left = o.Left ?? Left; Right = o.Right ?? Right;
        InsideH = o.InsideH ?? InsideH; InsideV = o.InsideV ?? InsideV;
    }
}

public sealed class Margins
{
    public double? Top, Bottom, Left, Right;
    public void Overlay(Margins? o)
    {
        if (o == null) return;
        Top = o.Top ?? Top; Bottom = o.Bottom ?? Bottom;
        Left = o.Left ?? Left; Right = o.Right ?? Right;
    }
}

public enum VMerge { None, Restart, Continue }

public sealed class TableCell
{
    public RadiusSpec? Radius;
    public int GridSpan = 1;
    public VMerge VMerge;
    public CellBorders? Borders;
    public Rgb? Shading;
    public Margins? Margins;
    public string VAlign = "top";
    /// <summary>Направление текста в ячейке: null — обычное, "btLr" — снизу вверх, "tbRl" — сверху вниз.</summary>
    public string? TextDirection;
    public double? Width;
    public List<Block> Blocks = [];
}

public sealed class TableRow
{
    public double Height;
    public string HeightRule = "atLeast";
    /// <summary>Роль строки при распределении высоты (align-content): 0 — содержимое, 1 — промежуток, 2 — служебная сверху, 3 — служебная снизу.</summary>
    public int Role;
    /// <summary>Высота строки не задана явно (растягивается при align-content: stretch).</summary>
    public bool AutoHeight = true;
    public bool IsHeader, CantSplit;
    public List<TableCell> Cells = [];
}

public sealed class Table : Block
{
    public List<double> Grid = [];
    public CellBorders Borders = new();
    public Margins CellMargins = new() { Left = 5.4, Right = 5.4, Top = 0, Bottom = 0 };
    public double Indent;
    public string Align = "left";
    public double? Width;
    /// <summary>Высота области содержимого (CSS height контейнера сетки) и распределение лишнего места (align-content).</summary>
    public double MinContentHeight;
    public string AlignContent = "normal";
    /// <summary>Таблица справа налево (w:bidiVisual): первая колонка — справа.</summary>
    public bool Rtl;
    /// <summary>Интервал между ячейками (w:tblCellSpacing), pt: отступ ячейки от соседей и края таблицы.</summary>
    public double CellSpacing;
    /// <summary>Ширина в долях ширины текста (tblW type="pct"), 1 = 100 %</summary>
    public double? WidthPct;
    public List<TableRow> Rows = [];
    public bool PageBreakBefore;
    public double SpaceBefore, SpaceAfter;
    public string? Clear;
}

/// <summary>
/// Радиусы скругления углов (CSS border-radius), в т.ч. эллиптические.
/// Порядок: горизонтальные TL, TR, BR, BL, затем вертикальные TL, TR, BR, BL.
/// Проценты хранятся долями и вычисляются при отрисовке: горизонтальные — от ширины, вертикальные — от высоты.
/// </summary>
public sealed class RadiusSpec
{
    public double[] V = new double[8];
    public bool[] Pct = new bool[8];

    public bool Any => V.Any(v => v > 0);

    public double[] Resolve(double w, double h)
    {
        var r = new double[8];
        for (int i = 0; i < 8; i++) r[i] = Pct[i] ? V[i] * (i < 4 ? w : h) : V[i];
        return r;
    }
}

/// <summary>Оформление блока-контейнера (фон, рамка, внутренние отступы).</summary>
public sealed class BoxStyle
{
    public Rgb? Background;
    public BorderDef? BorderTop, BorderBottom, BorderLeft, BorderRight;
    public double PadTop, PadBottom, PadLeft, PadRight;
    public double MinHeight;             // CSS height (рамка целиком)
    public RadiusSpec? Radius;           // border-radius
    public double HorizontalExtra => PadLeft + PadRight + (BorderLeft?.Width ?? 0) + (BorderRight?.Width ?? 0);
}

/// <summary>Плавающий блок (CSS float) — текст обтекает его.</summary>
public sealed class FloatBlock : Block
{
    /// <summary>Буквица: высота в строках следующего абзаца (0 — обычный плавающий блок).</summary>
    public int DropLines;
    public TransformResolver? Transform;
    public string Side = "left";
    public double Width;                 // ширина рамки блока (без внешних отступов)
    public double MarginTop, MarginBottom, MarginLeft, MarginRight;
    public BoxStyle Box = new();
    public List<Block> Blocks = [];
    /// <summary>Для рамок w:framePr: подпись свойств рамки (соседние абзацы с той же рамкой объединяются).</summary>
    public string? FrameKey;
    /// <summary>Обтекание по контуру (wp:wrapPolygon): точки в долях рамки; расстояние от контура до текста, pt.</summary>
    public List<(double X, double Y)>? WrapPolygon;
    public double WrapGap;
    /// <summary>
    /// Обтекаемый объект, привязанный по вертикали к странице/полям: положение (VRel, выравнивание, смещение) уточняется
    /// после разбивки на страницы (MarginTop — сдвиг от абзаца привязки); если сдвинуть нельзя — выводится AbsFallback без обтекания.
    /// </summary>
    public string? AbsVRel, AbsVAlign;
    public double AbsVOffset;
    public PositionedBlock? AbsFallback;
    public bool UseFallback;
}

/// <summary>Абсолютно/фиксированно позиционированный блок (вне потока).</summary>
public sealed class PositionedBlock : Block
{
    /// <summary>
    /// Относительно чего позиционировать (DOCX-якоря): "page", "margin", "flow" (абзац); null — правила CSS.
    /// </summary>
    public string? HRel, VRel;
    /// <summary>Выравнивание внутри области: left/center/right и top/center/bottom.</summary>
    public string? AlignX, AlignY;
    public TransformResolver? Transform;
    public string Mode = "absolute";     // absolute | fixed
    public double? Left, Top, Right, Bottom;
    public double Width;
    public int ZIndex;
    public BoxStyle Box = new();
    public List<Block> Blocks = [];
}

/// <summary>
/// Преобразование рамки элемента (CSS transform, включая 3D и перспективу) как проективное
/// отображение плоскости: [X, Y, W] = H · [x, y, 1], результат (X/W, Y/W).
/// Координаты локальные — от левого верхнего угла рамки, ось y направлена вниз.
/// Без перспективы H аффинна (H[6] = H[7] = 0) и выражается одной матрицей PDF.
/// </summary>
public sealed class BoxTransform
{
    public required double[] H;          // 3×3 по строкам
    public bool Hidden;                  // backface-visibility: hidden и элемент повёрнут обратной стороной
    /// <summary>Полная матрица 4×4 (по строкам) в локальных координатах рамки — для 3D-сцены.</summary>
    public double[]? M;
    /// <summary>transform-style: preserve-3d — потомки находятся в общей 3D-сцене с элементом.</summary>
    public bool Preserve3D;
    /// <summary>backface-visibility: hidden (обратная сторона определяется уже в сцене).</summary>
    public bool BackfaceHidden;

    public bool IsAffine => Math.Abs(H[6]) < 1e-12 && Math.Abs(H[7]) < 1e-12;

    public (double X, double Y, double W) Map(double x, double y)
    {
        double w = H[6] * x + H[7] * y + H[8];
        return ((H[0] * x + H[1] * y + H[2]) / w, (H[3] * x + H[4] * y + H[5]) / w, w);
    }
}

/// <summary>Вычисляет преобразование по размеру рамки элемента (ширина, высота) — для % в translate и transform-origin.</summary>
public delegate BoxTransform TransformResolver(double width, double height);

/// <summary>Блоки, отрисовываемые с преобразованием (на раскладку не влияет).</summary>
public sealed class TransformBlock : Block
{
    public List<Block> Blocks = [];
    public double BoxX, BoxWidth;        // рамка элемента относительно колонки
    public required TransformResolver Resolve;
}

/// <summary>Строки flex-контейнера с переносом и заданной высотой: распределение по высоте (align-content).</summary>
public sealed class FlexLinesBlock : Block
{
    public List<Table> Lines = [];
    public double MinHeight;
    public string AlignContent = "normal";
}

/// <summary>Конец позиционированного (relative) предка: по нему определяется высота предка для bottom у absolute-потомков.</summary>
public sealed class AnchorEndBlock : Block
{
    public required AnchorBlock Anchor;
    public double PadBottom;
}

/// <summary>Точка привязки для absolute-потомков позиционированного предка.</summary>
public sealed class AnchorBlock : Block
{
    public double X;                     // левый край предка относительно колонки
    public double Width;                 // ширина предка
    public double TopOffset;             // смещение верха предка относительно точки привязки
    public double? Height;               // высота предка (для top/bottom в %), если известна
    public List<PositionedBlock> Children = [];
}

public sealed class HeaderFooter
{
    public List<Block> Blocks = [];
}

public sealed class SectionProps
{
    /// <summary>Нумерация строк (w:lnNumType): шаг вывода; 0 — нумерации нет.</summary>
    public int LineNumberCountBy;
    public int LineNumberStart = 1;
    /// <summary>Расстояние от номера до текста, pt; null — авто (0,25").</summary>
    public double? LineNumberDistance;
    /// <summary>newPage, newSection, continuous</summary>
    public string LineNumberRestart = "newPage";
    /// <summary>Вертикальное выравнивание текста на странице (w:vAlign): top, center, bottom, both.</summary>
    public string VAlign = "top";
    public double PageWidth = 595.3, PageHeight = 841.9;
    public double MarginTop = 56.7, MarginBottom = 56.7, MarginLeft = 85.05, MarginRight = 42.55;
    public double HeaderDistance = 35.4, FooterDistance = 35.4;
    public string Type = "nextPage";
    public bool TitlePage;
    public HeaderFooter? HeaderDefault, HeaderFirst, HeaderEven;
    public HeaderFooter? FooterDefault, FooterFirst, FooterEven;
    public int? PageNumberStart;
    /// <summary>Формат номеров страниц (w:pgNumType w:fmt): decimal, upperRoman, lowerRoman, upperLetter …</summary>
    public string PageNumberFormat = "decimal";
    /// <summary>Рамка страницы (w:pgBorders); Space — отступ от края страницы или текста.</summary>
    public BorderDef? PageBorderTop, PageBorderBottom, PageBorderLeft, PageBorderRight;
    public bool PageBorderFromText;
    public string PageBorderDisplay = "allPages";
    /// <summary>Сетка строк (w:docGrid type="lines"): шаг в pt, 0 — нет.</summary>
    public double LinePitch;
    /// <summary>Колонки текста.</summary>
    public int Columns = 1;
    public double ColumnSpace = 36;
    public List<(double Width, double Space)>? ColumnWidths;
    public bool ColumnSeparator;
    /// <summary>position: fixed — выводится на каждой странице.</summary>
    public List<PositionedBlock> FixedBlocks = [];
}

public sealed class Section
{
    public SectionProps Props = new();
    public List<Block> Blocks = [];
}

public sealed class DocxDocument
{
    public List<Section> Sections = [];
    public Dictionary<string, byte[]> Media = new(StringComparer.OrdinalIgnoreCase);
    public bool EvenAndOddHeaders;
    /// <summary>Зеркальные поля (w:mirrorMargins): на чётных страницах левое и правое поля меняются местами.</summary>
    public bool MirrorMargins;
    /// <summary>Примечания выводятся выносками на правом поле.</summary>
    public bool CommentBalloons;
    /// <summary>Цвет фона страниц (w:background).</summary>
    public Rgb? Background;
    public double DefaultTabStop = 35.4;
    public string? Title, Author, Subject;
}

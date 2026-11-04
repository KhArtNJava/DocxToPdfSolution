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
    public double? Size;          // pt
    public string? Color;
    public string? Highlight;
    public string? Shading;
    public string? VertAlign;
    public string? StyleId;       // rStyle
    public double? Position;      // pt (поднятие/опускание)
    public double? Spacing;       // pt (межбуквенный интервал)

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
        Size = o.Size ?? Size;
        Color = o.Color ?? Color;
        Highlight = o.Highlight ?? Highlight;
        Shading = o.Shading ?? Shading;
        VertAlign = o.VertAlign ?? VertAlign;
        Position = o.Position ?? Position;
        Spacing = o.Spacing ?? Spacing;
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
    public bool? PageBreakBefore, KeepNext, KeepLines, ContextualSpacing, WidowControl;
    public List<TabStop>? Tabs;   // включая val="clear"
    public string? Shading;
    public BorderDef? BorderTop, BorderBottom;
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
    public double[]? Radius;

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
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool DoubleUnderline { get; init; }
    public bool Strike { get; init; }
    public bool Caps { get; init; }
    /// <summary>Капитель: строчные буквы — уменьшенными прописными.</summary>
    public bool SmallCaps { get; init; }
    /// <summary>Дополнительный интервал после каждого символа, pt (letter-spacing / w:spacing).</summary>
    public double LetterSpacing { get; init; }
    /// <summary>Дополнительный интервал у пробелов, pt (word-spacing).</summary>
    public double WordSpacing { get; init; }
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
    public double[]? Radius;             // скругление картинки (обрезка)
    public string? VAlign;               // vertical-align: middle / top / bottom …
    public string PartName = partName;
    public double Width = width;   // pt
    public double Height = height; // pt
    public TextStyle Style = style;
}

public enum FieldKind { Page, NumPages }

public sealed class FieldInline(FieldKind kind, TextStyle style) : Inline
{
    public FieldKind Kind = kind;
    public TextStyle Style = style;
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
    public BorderDef? BorderTop, BorderBottom;
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
    public double[]? Radius;
    public int GridSpan = 1;
    public VMerge VMerge;
    public CellBorders? Borders;
    public Rgb? Shading;
    public Margins? Margins;
    public string VAlign = "top";
    public double? Width;
    public List<Block> Blocks = [];
}

public sealed class TableRow
{
    public double Height;
    public string HeightRule = "atLeast";
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
    public List<TableRow> Rows = [];
    public bool PageBreakBefore;
    public double SpaceBefore, SpaceAfter;
    public string? Clear;
}

/// <summary>Оформление блока-контейнера (фон, рамка, внутренние отступы).</summary>
public sealed class BoxStyle
{
    public Rgb? Background;
    public BorderDef? BorderTop, BorderBottom, BorderLeft, BorderRight;
    public double PadTop, PadBottom, PadLeft, PadRight;
    public double MinHeight;             // CSS height (рамка целиком)
    public double[]? Radius;             // border-radius: TL, TR, BR, BL (pt)
    public double HorizontalExtra => PadLeft + PadRight + (BorderLeft?.Width ?? 0) + (BorderRight?.Width ?? 0);
}

/// <summary>Плавающий блок (CSS float) — текст обтекает его.</summary>
public sealed class FloatBlock : Block
{
    public TransformResolver? Transform;
    public string Side = "left";
    public double Width;                 // ширина рамки блока (без внешних отступов)
    public double MarginTop, MarginBottom, MarginLeft, MarginRight;
    public BoxStyle Box = new();
    public List<Block> Blocks = [];
}

/// <summary>Абсолютно/фиксированно позиционированный блок (вне потока).</summary>
public sealed class PositionedBlock : Block
{
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
    public double PageWidth = 595.3, PageHeight = 841.9;
    public double MarginTop = 56.7, MarginBottom = 56.7, MarginLeft = 85.05, MarginRight = 42.55;
    public double HeaderDistance = 35.4, FooterDistance = 35.4;
    public string Type = "nextPage";
    public bool TitlePage;
    public HeaderFooter? HeaderDefault, HeaderFirst, HeaderEven;
    public HeaderFooter? FooterDefault, FooterFirst, FooterEven;
    public int? PageNumberStart;
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
    public double DefaultTabStop = 35.4;
    public string? Title, Author, Subject;
}

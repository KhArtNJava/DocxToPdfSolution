using System.Globalization;
using DocxToPdf.Docx;
using DocxToPdf.Images;

namespace DocxToPdf.Html;

/// <summary>
/// Фреймы: &lt;iframe&gt; (src, srcdoc) и &lt;frameset&gt;/&lt;frame&gt;. Документ фрейма разбирается отдельным
/// HtmlReader со своими стилями (общие с родителем только ресурсы веб-архива и картинки) и выводится
/// в рамке заданного размера с обрезкой по ней. В MHTML документы фреймов — части «frame-…» (cid:frame-…@mhtml.blink).
/// </summary>
public sealed partial class HtmlReader
{
    string _mediaPrefix = "html-image-";      // префикс имён картинок (у вложенных документов — свой)
    int _frameCounter;
    int _frameDepth;
    double _viewportHeight = 700;             // высота области просмотра (страницы или фрейма), pt
    readonly HashSet<string> _frameChain = new(StringComparer.OrdinalIgnoreCase);   // адреса документов-предков (защита от рекурсии)

    const int MaxFrameDepth = 8;

    /// <summary>Читатель вложенного документа: те же шрифты, настройки, архив и общий набор картинок.</summary>
    HtmlReader ChildReader(string baseDir, string? baseUrl, string? key)
    {
        var r = new HtmlReader(_fonts, _opt, baseDir)
        {
            _mht = _mht, _baseUrl = baseUrl, _doc = _doc, _frameDepth = _frameDepth + 1,
            _mediaPrefix = _mediaPrefix + "f" + (++_frameCounter).ToString(CultureInfo.InvariantCulture) + "-",
        };
        foreach (var k in _frameChain) r._frameChain.Add(k);
        if (key != null) r._frameChain.Add(key);
        r._ua.Parse("body { margin: 8px }");
        return r;
    }

    /// <summary>Документ фрейма: srcdoc, часть веб-архива, файл или адрес. Картинка вместо документа — как картинка.</summary>
    (string Html, string BaseDir, string? BaseUrl, string? Key)? LoadFrame(Node n)
    {
        if (n.Tag == "iframe" && n.Attr("srcdoc") is { } srcdoc) return (srcdoc, _baseDir, _baseUrl, null);
        string? src = n.Attr("src")?.Trim();
        if (string.IsNullOrEmpty(src) || src.StartsWith("about:", StringComparison.OrdinalIgnoreCase)
            || src.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase)) return null;
        if (_frameChain.Contains(src)) return null;
        if (_mht != null && !src.StartsWith("data:", StringComparison.OrdinalIgnoreCase) && _mht.Find(src, _baseUrl) is { } part)
        {
            string key = part.ContentId ?? part.Location ?? src;
            if (_frameChain.Contains(key) || ReferenceEquals(part, _mht.Main)) return null;
            if (part.ContentType.StartsWith("image/", StringComparison.Ordinal)) return (ImageDocument(src), _baseDir, _baseUrl, key);
            return (DecodePart(part), _baseDir, part.Location ?? _baseUrl, key);
        }
        string dir = _baseDir, fileKey = src;
        if (!src.Contains(':') || src.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                string path = src.StartsWith("file:", StringComparison.OrdinalIgnoreCase) ? new Uri(src).LocalPath : Uri.UnescapeDataString(src.Split('?', '#')[0]);
                if (!Path.IsPathRooted(path)) path = Path.Combine(_baseDir, path.Replace('/', Path.DirectorySeparatorChar));
                fileKey = Path.GetFullPath(path);
                dir = Path.GetDirectoryName(fileKey) ?? _baseDir;
            }
            catch { /* остаётся каталог родителя */ }
        }
        if (_frameChain.Contains(fileKey)) return null;   // документ-предок (фрейм сам в себе)
        var bytes = LoadResource(src);
        if (bytes == null) return null;
        if (ImageDecoder.Decode(bytes) != null) return (ImageDocument(src), _baseDir, _baseUrl, fileKey);
        return (DecodeBytes(bytes), dir, _baseUrl, fileKey);
    }

    static string ImageDocument(string src) =>
        "<html><body><img src=\"" + System.Net.WebUtility.HtmlEncode(src) + "\"></body></html>";

    /// <summary>Содержимое фрейма шириной/высотой области просмотра; фон — от html/body документа фрейма.</summary>
    List<Block> FrameBlocks(Node n, double width, double height, out Rgb? background)
    {
        background = null;
        if (_frameDepth >= MaxFrameDepth || LoadFrame(n) is not { } fr) return [];
        var r = ChildReader(fr.BaseDir, fr.BaseUrl, fr.Key);
        try { return r.ConvertFrame(fr.Html, width, height, out background); }
        catch { return []; }
    }

    List<Block> ConvertFrame(string html, double width, double height, out Rgb? background)
    {
        var root = Prepare(html);
        _contentWidth = Math.Max(5, width);
        _viewportHeight = Math.Max(5, height);
        ComputeStyles(root, new CssStyle { Display = "block" });
        var htmlEl = root.Descendants().FirstOrDefault(d => d.Tag == "html");
        var body = root.Descendants().FirstOrDefault(d => d.Tag == "body");
        // Фон html (или body) переходит на холст фрейма — рисуется ячейкой на всю область, а не блоком элемента
        if (htmlEl?.Style?.Background is { } hb) { background = hb; htmlEl.Style.Background = null; }
        else if (body?.Style?.Background is { } bb) { background = bb; body.Style.Background = null; }
        else background = null;

        var blocks = new List<Block>();
        var flow = new Flow(blocks);
        // Фрейм — содержащий блок для absolute- и fixed-элементов своего документа
        var ab = new AnchorBlock { X = 0, Width = _contentWidth, TopOffset = 0, Height = _viewportHeight };
        blocks.Add(ab);
        var ctx = new BlockCtx { IndL = 0, IndR = 0, Owner = root.Style!, Width = _contentWidth, Anchor = ab };
        WalkChildren(root, flow, ctx);
        Flush(flow);
        foreach (var pb in _fixedBlocks) { pb.Mode = "absolute"; ab.Children.Add(pb); }
        _fixedBlocks.Clear();
        if (ab.Children.Count == 0) blocks.Remove(ab);
        else blocks.Add(new AnchorEndBlock { Anchor = ab });
        return blocks;
    }

    /// <summary>Высота области содержимого &lt;iframe&gt; (без рамки и отступов), pt.</summary>
    static double FrameContentHeight(CssStyle s)
    {
        double bh = BoxHeight(s) ?? 150 * 0.75 + s.PaddingTop + s.PaddingBottom
            + (s.BorderTop?.ToBorderDef()?.Width ?? 0) + (s.BorderBottom?.ToBorderDef()?.Width ?? 0);
        return Math.Max(1, bh - s.PaddingTop - s.PaddingBottom - (s.BorderTop?.ToBorderDef()?.Width ?? 0) - (s.BorderBottom?.ToBorderDef()?.Width ?? 0));
    }

    /// <summary>Таблица из одной ячейки точной высоты с содержимым фрейма (без оформления — оно снаружи).</summary>
    Table FrameTable(Node n, CssStyle s, double width)
    {
        double h = FrameContentHeight(s);
        var cell = new TableCell { Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 }, Borders = NoBorders() };
        cell.Blocks = FrameBlocks(n, width, h, out var bg);
        cell.Shading = bg;
        if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
        var t = new Table { Grid = [width], Borders = new CellBorders(), CellMargins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 } };
        var row = new TableRow { Height = h, HeightRule = "exact", AutoHeight = false, CantSplit = true };
        row.Cells.Add(cell);
        t.Rows.Add(row);
        return t;
    }

    static CellBorders NoBorders() => new()
    {
        Top = new BorderDef { Val = "nil" }, Bottom = new BorderDef { Val = "nil" },
        Left = new BorderDef { Val = "nil" }, Right = new BorderDef { Val = "nil" },
    };

    // =================================================================
    //  frameset
    // =================================================================

    /// <summary>&lt;frameset&gt; на месте body: сетка фреймов на всю область страницы.</summary>
    void AddFrameset(Node fs, Flow f, BlockCtx b)
    {
        Flush(f);
        double w = Math.Max(20, _contentWidth - b.IndL - b.IndR);
        double h = Math.Max(20, _viewportHeight - f.PendingMargin - 1);
        var t = FramesetTable(fs, w, h);
        t.Indent = b.IndL;
        t.SpaceBefore = f.PendingMargin; t.PageBreakBefore = f.PendingBreak;
        f.PendingMargin = 0; f.PendingBreak = false;
        f.Target.Add(t);
    }

    Table FramesetTable(Node fs, double w, double h)
    {
        var cols = FrameTracks(fs.Attr("cols"), w);
        var rows = FrameTracks(fs.Attr("rows"), h);
        bool border = fs.Attr("frameborder")?.Trim().ToLowerInvariant() is not ("0" or "no") && fs.Attr("border")?.Trim() != "0";
        var line = border ? new BorderDef { Val = "single", Width = 0.75, Color = new Rgb(0.6, 0.6, 0.6) } : null;
        var items = fs.Children.Where(c => c.Tag is "frame" or "frameset").ToList();
        var t = new Table { Grid = cols, Borders = new CellBorders(), CellMargins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 } };
        int k = 0;
        for (int ri = 0; ri < rows.Count; ri++)
        {
            var row = new TableRow { Height = rows[ri], HeightRule = "exact", AutoHeight = false, CantSplit = true };
            for (int ci = 0; ci < cols.Count; ci++)
            {
                var item = k < items.Count ? items[k++] : null;
                bool frameBorder = line != null && item?.Attr("frameborder")?.Trim().ToLowerInvariant() is not ("0" or "no");
                var cell = new TableCell { Margins = new Margins { Top = 0, Bottom = 0, Left = 0, Right = 0 }, Borders = NoBorders() };
                if (frameBorder && item?.Tag == "frame")
                    cell.Borders = new CellBorders { Top = line, Bottom = line, Left = line, Right = line };
                double bw = cell.Borders.Left!.IsNone ? 0 : 0.75;
                double cw = cols[ci] - 2 * bw, chh = rows[ri] - 2 * bw;
                if (item?.Tag == "frameset") cell.Blocks.Add(FramesetTable(item, cols[ci], rows[ri]));
                else if (item != null)
                {
                    cell.Blocks = FrameBlocks(item, cw, chh, out var bg);
                    cell.Shading = bg;
                }
                if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
                row.Cells.Add(cell);
            }
            t.Rows.Add(row);
        }
        return t;
    }

    /// <summary>Размеры дорожек rows/cols: пиксели, проценты и доли «*»; без атрибута — одна дорожка на всё.</summary>
    static List<double> FrameTracks(string? spec, double total)
    {
        var parts = (spec ?? "").Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0) return [total];
        var size = new double[parts.Length];
        var rel = new double[parts.Length];
        double fixedSum = 0, relSum = 0;
        for (int i = 0; i < parts.Length; i++)
        {
            string p = parts[i];
            if (p.EndsWith('*'))
            {
                rel[i] = double.TryParse(p[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var r) && r > 0 ? r : 1;
                relSum += rel[i];
            }
            else if (p.EndsWith('%'))
            {
                size[i] = double.TryParse(p[..^1], NumberStyles.Float, CultureInfo.InvariantCulture, out var pc) ? Math.Max(0, pc) * total / 100 : 0;
                fixedSum += size[i];
            }
            else
            {
                size[i] = double.TryParse(p.TrimEnd('p', 'x'), NumberStyles.Float, CultureInfo.InvariantCulture, out var px) ? Math.Max(0, px) * 0.75 : 0;
                fixedSum += size[i];
            }
        }
        double rest = total - fixedSum;
        if (relSum > 0)
        {
            for (int i = 0; i < parts.Length; i++) if (rel[i] > 0) size[i] = Math.Max(0, rest) * rel[i] / relSum;
        }
        double sum = size.Sum();
        if (sum > 0 && Math.Abs(sum - total) > 0.01)
        {
            // Не хватило или осталось место (нет «*») — пропорционально
            for (int i = 0; i < size.Length; i++) size[i] *= total / sum;
        }
        else if (sum <= 0) for (int i = 0; i < size.Length; i++) size[i] = total / size.Length;
        return [.. size.Select(x => Math.Max(1, x))];
    }
}

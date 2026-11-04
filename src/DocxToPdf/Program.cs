using System.Diagnostics;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Html;
using DocxToPdf.Layout;
using DocxToPdf.Pdf;

namespace DocxToPdf;

public static class Converter
{
    /// <summary>Конвертирует DOCX- или HTML-файл в PDF (формат определяется по содержимому).</summary>
    public static int Convert(string input, string output, FontManager fonts, HtmlOptions? htmlOptions = null)
    {
        var doc = IsZip(input) ? DocxReader.Read(input) : HtmlReader.Read(input, fonts, htmlOptions);
        var ctx = new LayoutContext(doc, fonts);
        var pages = new Paginator(ctx).Run();
        var pdf = new PdfDocument { Title = doc.Title, Author = doc.Author, Subject = doc.Subject };
        Paginator.Render(pages, pdf);
        pdf.Save(output);
        return pages.Count;
    }

    static bool IsZip(string path)
    {
        using var fs = File.OpenRead(path);
        return fs.ReadByte() == 'P' && fs.ReadByte() == 'K';
    }
}

public static class Program
{
    public static int Main(string[] args)
    {
        var inputs = new List<string>();
        var fontDirs = new List<string>();
        string? output = null;
        var html = new HtmlOptions();
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-o" or "--output" when i + 1 < args.Length: output = args[++i]; break;
                case "--fonts" when i + 1 < args.Length: fontDirs.Add(args[++i]); break;
                case "--page-size" when i + 1 < args.Length:
                    (html.PageWidth, html.PageHeight) = args[++i].ToLowerInvariant() switch
                    {
                        "a3" => (841.9, 1190.6), "a5" => (419.5, 595.3), "letter" => (612.0, 792.0), "legal" => (612.0, 1008.0), _ => (595.3, 841.9)
                    };
                    break;
                case "--landscape": html.Landscape = true; break;
                case "--margin" when i + 1 < args.Length:
                    if (double.TryParse(args[++i], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var mm))
                        html.MarginTop = html.MarginBottom = html.MarginLeft = html.MarginRight = mm * 72 / 25.4;
                    break;
                case "--page-numbers": html.PageNumbers = true; break;
                case "--no-remote": html.AllowRemote = false; break;
                case "-h" or "--help": PrintUsage(); return 0;
                default: inputs.Add(args[i]); break;
            }
        }
        if (inputs.Count == 0) { PrintUsage(); return 1; }
        if (output != null && inputs.Count > 1 && !Directory.Exists(output))
        {
            Console.Error.WriteLine("При нескольких входных файлах -o должен указывать на существующий каталог.");
            return 1;
        }

        FontManager fonts;
        try { fonts = new FontManager(fontDirs); }
        catch (Exception e) { Console.Error.WriteLine("Ошибка: " + e.Message); return 2; }

        int failed = 0;
        foreach (var input in inputs)
        {
            string target = output == null
                ? Path.ChangeExtension(input, ".pdf")
                : Directory.Exists(output) ? Path.Combine(output, Path.GetFileNameWithoutExtension(input) + ".pdf") : output;
            var sw = Stopwatch.StartNew();
            try
            {
                int pages = Converter.Convert(input, target, fonts, html);
                Console.WriteLine($"{input} -> {target}: {pages} стр., {sw.ElapsedMilliseconds} мс");
            }
            catch (Exception e)
            {
                failed++;
                Console.Error.WriteLine($"{input}: ошибка: {e.Message}");
#if DEBUG
                Console.Error.WriteLine(e);
#endif
            }
        }
        return failed == 0 ? 0 : 3;
    }

    static void PrintUsage()
    {
        Console.WriteLine("""
            docx2pdf — конвертер DOCX и HTML в PDF без сторонних библиотек

            Использование:
              docx2pdf <файл.docx|файл.html> [ещё ...] [-o <выход.pdf | каталог>] [опции]

            Опции:
              -o, --output     Выходной PDF (или каталог для нескольких файлов).
                               По умолчанию — рядом с исходным файлом.
              --fonts <dir>    Дополнительный каталог со шрифтами TTF/TTC (можно несколько раз).

            Только для HTML (CSS @page имеет приоритет):
              --page-size <A3|A4|A5|Letter|Legal>   Размер страницы (по умолчанию A4).
              --landscape      Альбомная ориентация.
              --margin <мм>    Поля страницы в миллиметрах (по умолчанию 20).
              --page-numbers   Номера страниц внизу ("1 / 5").
              --no-remote      Не загружать картинки и стили по http(s).
            """);
    }
}

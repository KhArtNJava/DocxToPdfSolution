using System.Diagnostics;
using DocxToPdf.Docx;
using DocxToPdf.Fonts;
using DocxToPdf.Layout;
using DocxToPdf.Pdf;

namespace DocxToPdf;

public static class Converter
{
    /// <summary>Конвертирует DOCX-файл в PDF.</summary>
    public static int Convert(string input, string output, FontManager fonts)
    {
        var doc = DocxReader.Read(input);
        var ctx = new LayoutContext(doc, fonts);
        var pages = new Paginator(ctx).Run();
        var pdf = new PdfDocument { Title = doc.Title, Author = doc.Author, Subject = doc.Subject };
        Paginator.Render(pages, pdf);
        pdf.Save(output);
        return pages.Count;
    }
}

public static class Program
{
    public static int Main(string[] args)
    {
        var inputs = new List<string>();
        var fontDirs = new List<string>();
        string? output = null;
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-o" or "--output" when i + 1 < args.Length: output = args[++i]; break;
                case "--fonts" when i + 1 < args.Length: fontDirs.Add(args[++i]); break;
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
                int pages = Converter.Convert(input, target, fonts);
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
            docx2pdf — конвертер DOCX в PDF без сторонних библиотек

            Использование:
              docx2pdf <файл.docx> [ещё.docx ...] [-o <выход.pdf | каталог>] [--fonts <каталог>]

            Опции:
              -o, --output   Выходной PDF (или каталог для нескольких файлов).
                             По умолчанию — рядом с исходным файлом.
              --fonts        Дополнительный каталог со шрифтами TTF/TTC (можно несколько раз).
            """);
    }
}

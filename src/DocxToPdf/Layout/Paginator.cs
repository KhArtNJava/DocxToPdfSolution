using System.Globalization;
using DocxToPdf.Docx;
using DocxToPdf.Pdf;

namespace DocxToPdf.Layout;

/// <summary>Свёрстанная страница.</summary>
public sealed class Page
{
    public required SectionProps Section;
    public double Width, Height;
    public double BodyLeft;
    public bool FirstOfSection;
    public int Number;   // отображаемый номер страницы
    public List<(Frag Frag, double X, double Y)> Items = [];
    /// <summary>Позиционированные блоки: под потоком (z-index &lt; 0) и над ним.</summary>
    public List<(Frag Frag, double X, double Y, int Z)> Underlays = [], Overlays = [];
}

/// <summary>Разбивает поток фрагментов на страницы и добавляет колонтитулы.</summary>
public sealed class Paginator(LayoutContext ctx)
{
    readonly LayoutContext _ctx = ctx;
    readonly List<Page> _pages = [];

    Page? _page;
    double _y, _top, _bottom;
    bool _natural;          // страница начата из-за переполнения (а не явного разрыва)
    TableInfo? _lastTable;
    readonly List<AbsFrag> _fixed = [];
    int _sectionFirstPage;
    int _pageNumber;

    public List<Page> Run()
    {
        var doc = _ctx.Doc;
        for (int si = 0; si < doc.Sections.Count; si++)
        {
            var section = doc.Sections[si];
            var sp = section.Props;
            var (top, bottom) = BodyArea(sp);
            double width = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
            _ctx.MaxImageHeight = Math.Max(50, bottom - top - 2);
            var layout = new BlockLayout(_ctx);
            var frags = layout.LayoutBlocks(section.Blocks, width);
            foreach (var fb in sp.FixedBlocks)
                _fixed.Add(new AbsFrag
                {
                    Content = layout.BuildBoxTransformed(fb), Left = fb.Left, Top = fb.Top, Right = fb.Right,
                    Bottom = fb.Bottom, Fixed = true, ZIndex = fb.ZIndex
                });

            bool continuous = sp.Type == "continuous" && _page != null;
            if (!continuous)
            {
                if (sp.PageNumberStart != null) _pageNumber = sp.PageNumberStart.Value - 1;
                NewPage(sp, natural: false, firstOfSection: true);
                if (sp.Type is "evenPage" or "oddPage" && _pages.Count > 1)
                {
                    bool wantEven = sp.Type == "evenPage";
                    if ((_pageNumber % 2 == 0) != wantEven) NewPage(sp, natural: false, firstOfSection: true);
                }
            }
            else
            {
                if (sp.PageNumberStart != null) _pageNumber = sp.PageNumberStart.Value;
                _page!.Section = sp;
            }
            _top = top; _bottom = bottom;
            if (!continuous) _sectionFirstPage = _pages.Count - 1;
            if (_page!.Items.Count == 0) _y = Math.Max(_y, top);
            Place(frags, sp);
        }
        if (_pages.Count == 0) NewPage(doc.Sections.LastOrDefault()?.Props ?? new SectionProps(), false, true);

        // position: fixed — на каждой странице (относительно страницы целиком)
        foreach (var page in _pages)
            foreach (var a in _fixed)
            {
                double w = a.Content.Width, h = a.Content.Height;
                // Для печати fixed отсчитывается от области страницы (внутри полей)
                var ps = page.Section;
                double x = a.Left != null ? ps.MarginLeft + a.Left.Value : a.Right != null ? page.Width - ps.MarginRight - a.Right.Value - w : ps.MarginLeft;
                double y = a.Top != null ? ps.MarginTop + a.Top.Value : a.Bottom != null ? page.Height - ps.MarginBottom - a.Bottom.Value - h : ps.MarginTop;
                (a.ZIndex < 0 ? page.Underlays : page.Overlays).Add((a.Content, x, y, a.ZIndex));
            }

        AddHeadersFooters();
        return _pages;
    }

    (double Top, double Bottom) BodyArea(SectionProps sp)
    {
        double top = sp.MarginTop, bottom = sp.PageHeight - sp.MarginBottom;
        double width = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
        var saved = _ctx.FieldText;
        _ctx.FieldText = _ => "1";
        foreach (var h in new[] { sp.HeaderDefault, sp.HeaderFirst, sp.HeaderEven })
            if (h != null) top = Math.Max(top, sp.HeaderDistance + Measure(h, width));
        foreach (var f in new[] { sp.FooterDefault, sp.FooterFirst, sp.FooterEven })
            if (f != null) bottom = Math.Min(bottom, sp.PageHeight - sp.FooterDistance - Measure(f, width));
        _ctx.FieldText = saved;
        if (bottom - top < 72) { top = sp.MarginTop; bottom = sp.PageHeight - sp.MarginBottom; }
        return (top, bottom);
    }

    double Measure(HeaderFooter hf, double width)
    {
        var frags = new BlockLayout(_ctx).LayoutBlocks(hf.Blocks, width);
        double h = 0;
        for (int i = 0; i < frags.Count; i++)
            h += frags[i].SpaceBefore + frags[i].Height + (i < frags.Count - 1 ? frags[i].SpaceAfter : 0);
        return h;
    }

    void NewPage(SectionProps sp, bool natural, bool firstOfSection)
    {
        _pageNumber++;
        _page = new Page
        {
            Section = sp, Width = sp.PageWidth, Height = sp.PageHeight, BodyLeft = sp.MarginLeft,
            FirstOfSection = firstOfSection, Number = _pageNumber
        };
        _pages.Add(_page);
        var (top, bottom) = BodyArea(sp);
        _top = top; _bottom = bottom;
        _y = top;
        _natural = natural;
    }

    bool PageEmpty => _page!.Items.Count == 0;

    void Put(Frag f, double spaceBefore)
    {
        _y += spaceBefore;
        _page!.Items.Add((f, _page.Section.MarginLeft, _y));
        _y += f.Height + f.SpaceAfter;
        if (f is RowFrag rf) _lastTable = rf.Table;
        else _lastTable = null;
    }

    double ChainHeight(List<Frag> frags, int i, double firstSpace)
    {
        double h = firstSpace + frags[i].Height;
        int count = 0;
        while (frags[i].KeepWithNext && i + 1 < frags.Count && count++ < 40)
        {
            h += frags[i].SpaceAfter + frags[i + 1].SpaceBefore + frags[i + 1].Height;
            i++;
            if (frags[i].PageBreakBefore) break;
        }
        return h;
    }

    void PlaceAbs(AbsFrag a, SectionProps sp)
    {
        double ox, oy, cbW, cbH;
        if (a.RelativeToFlow)
        {
            ox = sp.MarginLeft + a.AnchorX; oy = _y + a.TopOffset; cbW = a.AnchorWidth; cbH = _bottom - oy;
        }
        else
        {
            ox = sp.MarginLeft; oy = _top; cbW = sp.PageWidth - sp.MarginLeft - sp.MarginRight; cbH = _bottom - _top;
        }
        double w = a.Content.Width, h = a.Content.Height;
        double x = a.Left != null ? ox + a.Left.Value : a.Right != null ? ox + cbW - a.Right.Value - w : ox;
        double y = a.Top != null ? oy + a.Top.Value : a.Bottom != null ? oy + cbH - a.Bottom.Value - h : _y;
        // Относительно начального блока (без позиционированного предка) — на первой странице раздела
        var target = !a.RelativeToFlow && (a.Top != null || a.Bottom != null) ? _pages[_sectionFirstPage] : _page!;
        var list = a.ZIndex < 0 ? target.Underlays : target.Overlays;
        list.Add((a.Content, x, y, a.ZIndex));
    }

    void Place(List<Frag> frags, SectionProps sp)
    {
        for (int i = 0; i < frags.Count; i++)
        {
            Frag f = frags[i];
            if (f is AbsFrag abs)
            {
                if (abs.Fixed) _fixed.Add(abs);
                else PlaceAbs(abs, sp);
                continue;
            }
            if (f.PageBreakBefore && !PageEmpty) NewPage(sp, natural: false, firstOfSection: false);

            double sb = PageEmpty && _natural ? 0 : f.SpaceBefore;
            double areaH = _bottom - _top;

            // Не отрывать от следующего (keepNext, висячие строки)
            if (f.KeepWithNext && !PageEmpty)
            {
                double ch = ChainHeight(frags, i, sb);
                if (_y + ch > _bottom + 0.01 && ch <= areaH)
                {
                    NewPage(sp, natural: true, firstOfSection: false);
                    sb = 0;
                }
            }

            while (true)
            {
                RepeatHeaderRows(f);
                if (PageEmpty && _natural) sb = 0;
                double need = sb + f.Height;
                if (_y + need <= _bottom + 0.01) break;

                // Пытаемся разрезать (строки таблицы)
                double room = _bottom - _y - sb;
                if (room > 12)
                {
                    var parts = f.Split(room);
                    if (parts != null)
                    {
                        Put(parts.Value.First, sb);
                        NewPage(sp, natural: true, firstOfSection: false);
                        f = parts.Value.Remainder;
                        sb = 0;
                        continue;
                    }
                }
                if (PageEmpty) break; // не помещается даже на пустую страницу — выводим как есть
                NewPage(sp, natural: true, firstOfSection: false);
                sb = 0;
            }
            Put(f, sb);
            if (f.PageBreakAfter) NewPage(sp, natural: false, firstOfSection: false);
        }
    }

    void RepeatHeaderRows(Frag f)
    {
        if (!PageEmpty || f is not RowFrag rf || rf.IsHeader || rf.Table.HeaderRows.Count == 0) return;
        if (_lastTable != rf.Table && !_pages.Any(p => p != _page && p.Items.Any(it => it.Frag is RowFrag r && r.Table == rf.Table)))
            return;
        foreach (var h in rf.Table.HeaderRows) Put(h, 0);
    }

    void AddHeadersFooters()
    {
        int total = _pages.Count;
        bool evenOdd = _ctx.Doc.EvenAndOddHeaders;
        foreach (var page in _pages)
        {
            var sp = page.Section;
            HeaderFooter? header, footer;
            if (sp.TitlePage && page.FirstOfSection) { header = sp.HeaderFirst; footer = sp.FooterFirst; }
            else if (evenOdd && page.Number % 2 == 0) { header = sp.HeaderEven; footer = sp.FooterEven; }
            else { header = sp.HeaderDefault; footer = sp.FooterDefault; }

            double width = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
            _ctx.FieldText = k => k == FieldKind.Page
                ? page.Number.ToString(CultureInfo.InvariantCulture)
                : total.ToString(CultureInfo.InvariantCulture);

            if (header != null)
            {
                var frags = new BlockLayout(_ctx).LayoutBlocks(header.Blocks, width);
                double y = sp.HeaderDistance;
                for (int i = 0; i < frags.Count; i++)
                {
                    y += frags[i].SpaceBefore;
                    page.Items.Add((frags[i], sp.MarginLeft, y));
                    y += frags[i].Height + frags[i].SpaceAfter;
                }
            }
            if (footer != null)
            {
                var frags = new BlockLayout(_ctx).LayoutBlocks(footer.Blocks, width);
                double h = 0;
                for (int i = 0; i < frags.Count; i++)
                    h += frags[i].SpaceBefore + frags[i].Height + (i < frags.Count - 1 ? frags[i].SpaceAfter : 0);
                double y = sp.PageHeight - sp.FooterDistance - h;
                foreach (var f in frags)
                {
                    y += f.SpaceBefore;
                    page.Items.Add((f, sp.MarginLeft, y));
                    y += f.Height + f.SpaceAfter;
                }
            }
            _ctx.FieldText = null;
        }
    }

    /// <summary>Рисует все страницы в PDF.</summary>
    public static void Render(List<Page> pages, PdfDocument pdf)
    {
        foreach (var page in pages)
        {
            var canvas = pdf.AddPage(page.Width, page.Height);
            foreach (var (frag, x, y, _) in page.Underlays.OrderBy(u => u.Z)) frag.Draw(canvas, x, y);
            foreach (var (frag, x, y) in page.Items) frag.Draw(canvas, x, y);
            foreach (var (frag, x, y, _) in page.Overlays.OrderBy(u => u.Z)) frag.Draw(canvas, x, y);
        }
    }
}

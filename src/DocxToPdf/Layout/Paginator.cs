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
    /// <summary>Разделительные линии колонок.</summary>
    public List<(double X, double Y1, double Y2)> ColumnLines = [];
    public Rgb? Background;
    /// <summary>Левое поле этой страницы (с учётом зеркальных полей) и текст номера страницы.</summary>
    public double MarginLeft;
    public string NumberText = "1";
}

/// <summary>Разбивает поток фрагментов на страницы (и колонки), размещает сноски и колонтитулы.</summary>
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

    // Колонки текущего раздела
    double[] _colX = [0];
    double[] _colW = [500];
    int _col;
    double _colTop, _colMaxY, _balanceLimit;
    SectionProps? _sp;

    // Сноски текущей страницы
    readonly List<Frag> _notes = [];
    double _footH;
    const double NoteSeparator = 10;

    double Bottom => (_col < _colX.Length - 1 && _balanceLimit > 0 ? Math.Min(_bottom, _balanceLimit) : _bottom) - _footH;

    public List<Page> Run()
    {
        var doc = _ctx.Doc;
        for (int si = 0; si < doc.Sections.Count; si++)
        {
            var section = doc.Sections[si];
            var sp = section.Props;
            _sp = sp;
            var (top, bottom) = BodyArea(sp);
            double width = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
            SetupColumns(sp, width);
            _ctx.MaxImageHeight = Math.Max(50, bottom - top - 2);
            _ctx.GridPitch = sp.LinePitch;
            var layout = new BlockLayout(_ctx);
            var frags = layout.LayoutBlocks(section.Blocks, _colW[0]);
            _ctx.GridPitch = 0;
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
                _top = top; _bottom = bottom;
            }
            if (!continuous) _sectionFirstPage = _pages.Count - 1;
            if (_page!.Items.Count == 0) _y = Math.Max(_y, top);
            _col = 0; _colTop = _y; _colMaxY = _y;

            // Балансировка колонок перед непрерывным разделом (как в Word)
            _balanceLimit = 0;
            bool nextContinuous = si + 1 < doc.Sections.Count && doc.Sections[si + 1].Props.Type == "continuous";
            if (_colX.Length > 1 && nextContinuous)
            {
                double total = frags.Sum(f => f.SpaceBefore + f.Height + f.SpaceAfter);
                _balanceLimit = _colTop + total / _colX.Length + 2;
            }

            Place(frags, sp);
            FinishColumns();
            _y = Math.Max(_y, _colMaxY);
        }
        if (_pages.Count == 0) NewPage(doc.Sections.LastOrDefault()?.Props ?? new SectionProps(), false, true);
        FlushNotes();

        // position: fixed — на каждой странице (относительно страницы целиком)
        foreach (var page in _pages)
            foreach (var a in _fixed)
            {
                double w = a.Content.Width, h = a.Content.Height;
                // Для печати fixed отсчитывается от области страницы (внутри полей)
                var ps = page.Section;
                double mr = page.Width - page.MarginLeft - (ps.PageWidth - ps.MarginLeft - ps.MarginRight);
                double x = a.Left != null ? page.MarginLeft + a.Left.Value : a.Right != null ? page.Width - mr - a.Right.Value - w : page.MarginLeft;
                double y = a.Top != null ? ps.MarginTop + a.Top.Value : a.Bottom != null ? page.Height - ps.MarginBottom - a.Bottom.Value - h : ps.MarginTop;
                (a.ZIndex < 0 ? page.Underlays : page.Overlays).Add((a.Content, x, y, a.ZIndex));
            }

        AddHeadersFooters();
        return _pages;
    }

    void SetupColumns(SectionProps sp, double width)
    {
        int n = Math.Max(1, sp.Columns);
        _colX = new double[n]; _colW = new double[n];
        if (sp.ColumnWidths is { Count: > 0 } cw && cw.Sum(c => c.Width) > 10)
        {
            double x = 0;
            for (int i = 0; i < n; i++) { _colX[i] = x; _colW[i] = cw[i].Width; x += cw[i].Width + cw[i].Space; }
        }
        else
        {
            double w = n == 1 ? width : (width - sp.ColumnSpace * (n - 1)) / n;
            for (int i = 0; i < n; i++) { _colX[i] = i * (w + sp.ColumnSpace); _colW[i] = w; }
        }
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
        double pitch = _ctx.GridPitch;
        _ctx.GridPitch = 0;
        var frags = new BlockLayout(_ctx).LayoutBlocks(hf.Blocks, width).Where(f => f is not AbsFrag).ToList();
        _ctx.GridPitch = pitch;
        double h = 0;
        for (int i = 0; i < frags.Count; i++)
            h += frags[i].SpaceBefore + frags[i].Height + (i < frags.Count - 1 ? frags[i].SpaceAfter : 0);
        return h;
    }

    void NewPage(SectionProps sp, bool natural, bool firstOfSection)
    {
        FinishColumns();
        FlushNotes();
        _pageNumber++;
        _page = new Page
        {
            Section = sp, Width = sp.PageWidth, Height = sp.PageHeight, BodyLeft = sp.MarginLeft,
            FirstOfSection = firstOfSection, Number = _pageNumber, Background = _ctx.Doc.Background,
            // Зеркальные поля: на чётных страницах внутреннее (широкое) поле справа
            MarginLeft = _ctx.Doc.MirrorMargins && _pageNumber % 2 == 0 ? sp.MarginRight : sp.MarginLeft,
            NumberText = FormatPageNumber(_pageNumber, sp.PageNumberFormat),
        };
        _pages.Add(_page);
        var (top, bottom) = BodyArea(sp);
        _top = top; _bottom = bottom;
        _y = top;
        _natural = natural;
        _col = 0; _colTop = top; _colMaxY = top;
    }

    /// <summary>Следующая колонка, а после последней — новая страница.</summary>
    void NextColumn(SectionProps sp, bool natural)
    {
        if (_col < _colX.Length - 1)
        {
            _colMaxY = Math.Max(_colMaxY, _y);
            _col++;
            _y = _colTop;
            _natural = natural;
            return;
        }
        NewPage(sp, natural, firstOfSection: false);
    }

    /// <summary>Завершает колонки на текущей странице: разделительные линии.</summary>
    void FinishColumns()
    {
        if (_page == null || _colX.Length < 2) return;
        _colMaxY = Math.Max(_colMaxY, _y);
        if (_sp?.ColumnSeparator == true && _colMaxY > _colTop + 1)
            for (int i = 1; i <= Math.Min(_col, _colX.Length - 1); i++)
            {
                double x = _page.MarginLeft + (_colX[i - 1] + _colW[i - 1] + _colX[i]) / 2;
                _page.ColumnLines.Add((x, _colTop, _colMaxY));
            }
    }

    /// <summary>Переносит сноски текущей страницы на страницу (внизу области текста).</summary>
    void FlushNotes()
    {
        if (_page == null || _notes.Count == 0) { _notes.Clear(); _footH = 0; return; }
        var sp = _page.Section;
        double y = _bottom - _footH + NoteSeparator;
        // Разделитель сносок: короткая линия
        var sep = new SeparatorFrag { Width = 144, Height = 0 };
        _page.Items.Add((sep, _page.MarginLeft, y - NoteSeparator / 2));
        foreach (var f in _notes)
        {
            y += f.SpaceBefore;
            _page.Items.Add((f, _page.MarginLeft, y));
            y += f.Height + f.SpaceAfter;
        }
        _notes.Clear();
        _footH = 0;
    }

    bool PageEmpty => _page!.Items.Count == 0 || (_col > 0 && Math.Abs(_y - _colTop) < 0.01);

    void Put(Frag f, double spaceBefore)
    {
        _y += spaceBefore;
        _page!.Items.Add((f, _page.MarginLeft + _colX[_col], _y));
        CollectBookmarks(f, _page.MarginLeft + _colX[_col], _y);
        _y += f.Height + f.SpaceAfter;
        _colMaxY = Math.Max(_colMaxY, _y - f.SpaceAfter);
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

    /// <summary>Положение абсолютного блока на странице.</summary>
    (double X, double Y) AbsPosition(AbsFrag a, SectionProps sp, double flowY, double top, double bottom, double ml)
    {
        double w = a.Content.Width, h = a.Content.Height;
        double contentW = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
        if (a.HRel != null || a.VRel != null)
        {
            // Якоря DOCX: относительно страницы, полей или абзаца, с выравниванием
            var (cx, cw) = a.HRel switch
            {
                "page" => (0.0, sp.PageWidth),
                "flow" => (ml + a.AnchorX, a.AnchorWidth > 0 ? a.AnchorWidth : contentW),
                _ => (ml, contentW)
            };
            var (cy, ch) = a.VRel switch
            {
                "page" => (0.0, sp.PageHeight),
                "margin" => (sp.MarginTop, sp.PageHeight - sp.MarginTop - sp.MarginBottom),
                _ => (flowY + a.TopOffset, bottom - flowY)
            };
            double x = a.AlignX switch
            {
                "center" => cx + (cw - w) / 2,
                "right" or "outside" => cx + cw - w,
                "left" or "inside" => cx,
                _ => cx + (a.Left ?? 0)
            };
            double y = a.AlignY switch
            {
                "center" => cy + (ch - h) / 2,
                "bottom" or "outside" => cy + ch - h,
                "top" or "inside" => cy,
                _ => cy + (a.Top ?? 0)
            };
            return (x, y);
        }
        double ox, oy, cbW, cbH;
        if (a.RelativeToFlow) { ox = ml + a.AnchorX; oy = flowY + a.TopOffset; cbW = a.AnchorWidth; cbH = bottom - oy; }
        else { ox = ml; oy = top; cbW = contentW; cbH = bottom - top; }
        double x2 = a.Left != null ? ox + a.Left.Value : a.Right != null ? ox + cbW - a.Right.Value - w : ox;
        double y2 = a.Top != null ? oy + a.Top.Value : a.Bottom != null ? oy + cbH - a.Bottom.Value - h : flowY;
        return (x2, y2);
    }

    void PlaceAbs(AbsFrag a, SectionProps sp)
    {
        var (x, y) = AbsPosition(a, sp, _y, _top, _bottom, _page!.MarginLeft);
        // CSS: относительно начального блока (без позиционированного предка) — на первой странице раздела
        var target = a.HRel == null && !a.RelativeToFlow && (a.Top != null || a.Bottom != null) ? _pages[_sectionFirstPage] : _page!;
        var list = a.ZIndex < 0 ? target.Underlays : target.Overlays;
        list.Add((a.Content, x, y, a.ZIndex));
    }

    /// <summary>Высота сносок фрагмента (с разделителем, если это первые сноски страницы).</summary>
    (List<Frag> Frags, double Height) LayoutNotes(Frag f)
    {
        var list = new List<Frag>();
        double h = 0;
        double width = _page!.Section.PageWidth - _page.Section.MarginLeft - _page.Section.MarginRight;
        foreach (var blocks in f.Notes)
        {
            var frags = new BlockLayout(_ctx).LayoutBlocks(blocks, width);
            foreach (var nf in frags) { h += nf.SpaceBefore + nf.Height + nf.SpaceAfter; list.Add(nf); }
        }
        if (list.Count > 0 && _notes.Count == 0) h += NoteSeparator;
        return (list, h);
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
            if (f.PageBreakBefore && !(_page!.Items.Count == 0)) NewPage(sp, natural: false, firstOfSection: false);

            double sb = PageEmpty && _natural ? 0 : f.SpaceBefore;
            double areaH = _bottom - _top;

            // Не отрывать от следующего (keepNext, висячие строки)
            if (f.KeepWithNext && !PageEmpty)
            {
                double ch = ChainHeight(frags, i, sb);
                if (_y + ch > Bottom + 0.01 && ch <= areaH)
                {
                    NextColumn(sp, natural: true);
                    sb = 0;
                }
            }

            var (noteFrags, noteH) = LayoutNotes(f);
            while (true)
            {
                RepeatHeaderRows(f);
                if (PageEmpty && _natural) sb = 0;
                if (noteFrags.Count > 0) (noteFrags, noteH) = LayoutNotes(f);   // разделитель зависит от страницы
                double need = sb + f.Height;
                if (_y + need <= Bottom - noteH + 0.01) break;

                // Пытаемся разрезать (строки таблицы)
                double room = Bottom - _y - sb;
                if (room > 12 && noteFrags.Count == 0)
                {
                    var parts = f.Split(room);
                    if (parts != null)
                    {
                        Put(parts.Value.First, sb);
                        NextColumn(sp, natural: true);
                        f = parts.Value.Remainder;
                        sb = 0;
                        continue;
                    }
                }
                if (PageEmpty) break; // не помещается даже на пустую страницу — выводим как есть
                NextColumn(sp, natural: true);
                sb = 0;
            }
            Put(f, sb);
            if (noteFrags.Count > 0)
            {
                _notes.AddRange(noteFrags);
                _footH += noteH;
            }
            if (f.PageBreakAfter) NewPage(sp, natural: false, firstOfSection: false);
            else if (f is LineFrag { ColumnBreakAfter: true }) NextColumn(sp, natural: false);
        }
    }

    /// <summary>Цели внутренних ссылок: закладка → (страница, x, y).</summary>
    public readonly Dictionary<string, (int Page, double X, double Y)> Destinations = [];

    void CollectBookmarks(Frag f, double x, double y)
    {
        switch (f)
        {
            case LineFrag lf when lf.IsFirst && lf.Para.Bookmarks != null:
                foreach (var b in lf.Para.Bookmarks) Destinations.TryAdd(b, (_pages.Count - 1, x, y));
                break;
            case RowFrag rf:
                foreach (var cell in rf.Cells)
                    foreach (var (cf, cy) in cell.Items) CollectBookmarks(cf, x + cell.X + cell.MarginLeft, y + cell.MarginTop + cy);
                break;
            case GroupFrag gf:
                foreach (var (gfr, gx, gy) in gf.Items) CollectBookmarks(gfr, x + gx, y + gy);
                break;
            case TransformFrag tf:
                foreach (var (tfr, ty) in tf.Items) CollectBookmarks(tfr, x, y + ty);
                break;
        }
    }

    /// <summary>Номер страницы в формате раздела (w:pgNumType w:fmt).</summary>
    public static string FormatPageNumber(int n, string fmt)
    {
        static string Roman(int v)
        {
            if (v <= 0 || v >= 4000) return v.ToString(CultureInfo.InvariantCulture);
            int[] vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
            string[] syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < vals.Length; i++) while (v >= vals[i]) { sb.Append(syms[i]); v -= vals[i]; }
            return sb.ToString();
        }
        static string Letters(int v, string abc)
        {
            if (v <= 0) return v.ToString(CultureInfo.InvariantCulture);
            char c = abc[(v - 1) % abc.Length];
            return new string(c, (v - 1) / abc.Length + 1);   // как в Word: A … Z, AA, BB …
        }
        string d = n.ToString(CultureInfo.InvariantCulture);
        return fmt switch
        {
            "upperRoman" => Roman(n),
            "lowerRoman" => Roman(n).ToLowerInvariant(),
            "upperLetter" => Letters(n, "ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
            "lowerLetter" => Letters(n, "abcdefghijklmnopqrstuvwxyz"),
            "russianUpper" => Letters(n, "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ"),
            "russianLower" => Letters(n, "абвгдежзиклмнопрстуфхцчшщэюя"),
            "numberInDash" => "- " + d + " -",
            "decimalZero" => n < 10 ? "0" + d : d,
            _ => d
        };
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
            _ctx.FieldText = k => k == FieldKind.Page ? page.NumberText : total.ToString(CultureInfo.InvariantCulture);
            var (top, bottom) = BodyArea(sp);

            void PlaceBlocks(List<Frag> frags, double y)
            {
                foreach (var f in frags)
                {
                    if (f is AbsFrag a)
                    {
                        // Объекты колонтитула (водяные знаки, логотипы) позиционируются относительно страницы
                        var (ax, ay) = AbsPosition(a, sp, y, top, bottom, page.MarginLeft);
                        (a.ZIndex < 0 || a.HRel != null ? page.Underlays : page.Overlays).Add((a.Content, ax, ay, a.ZIndex));
                        continue;
                    }
                    y += f.SpaceBefore;
                    page.Items.Add((f, page.MarginLeft, y));
                    y += f.Height + f.SpaceAfter;
                }
            }
            if (header != null)
                PlaceBlocks(new BlockLayout(_ctx).LayoutBlocks(header.Blocks, width), sp.HeaderDistance);
            if (footer != null)
            {
                var frags = new BlockLayout(_ctx).LayoutBlocks(footer.Blocks, width);
                var flow = frags.Where(f => f is not AbsFrag).ToList();
                double h = 0;
                for (int i = 0; i < flow.Count; i++)
                    h += flow[i].SpaceBefore + flow[i].Height + (i < flow.Count - 1 ? flow[i].SpaceAfter : 0);
                PlaceBlocks(frags, sp.PageHeight - sp.FooterDistance - h);
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
            canvas.PageNumber = page.Number;
            canvas.PageNumberText = page.NumberText;
            canvas.PageCount = pages.Count;
            if (page.Background is { } bg) canvas.FillRect(0, 0, page.Width, page.Height, bg);
            foreach (var (frag, x, y, _) in page.Underlays.OrderBy(u => u.Z)) frag.Draw(canvas, x, y);
            foreach (var (frag, x, y) in page.Items) frag.Draw(canvas, x, y);
            foreach (var (x, y1, y2) in page.ColumnLines) canvas.Line(x, y1, x, y2, 0.5, Rgb.Black);
            foreach (var (frag, x, y, _) in page.Overlays.OrderBy(u => u.Z)) frag.Draw(canvas, x, y);
            DrawPageBorders(canvas, page);
        }
    }

    static void DrawPageBorders(PdfCanvas c, Page page)
    {
        var sp = page.Section;
        if (sp.PageBorderTop == null && sp.PageBorderBottom == null && sp.PageBorderLeft == null && sp.PageBorderRight == null) return;
        if (sp.PageBorderDisplay == "firstPage" && !page.FirstOfSection) return;
        if (sp.PageBorderDisplay == "notFirstPage" && page.FirstOfSection) return;
        double Sp(BorderDef? b) => b?.Space ?? 0;
        double x1, y1, x2, y2;
        if (sp.PageBorderFromText)
        {
            double contentW = sp.PageWidth - sp.MarginLeft - sp.MarginRight;
            x1 = page.MarginLeft - Sp(sp.PageBorderLeft); x2 = page.MarginLeft + contentW + Sp(sp.PageBorderRight);
            y1 = sp.MarginTop - Sp(sp.PageBorderTop); y2 = page.Height - sp.MarginBottom + Sp(sp.PageBorderBottom);
        }
        else
        {
            x1 = Sp(sp.PageBorderLeft); x2 = page.Width - Sp(sp.PageBorderRight);
            y1 = Sp(sp.PageBorderTop); y2 = page.Height - Sp(sp.PageBorderBottom);
        }
        void L(BorderDef? b, double ax, double ay, double bx, double by)
        {
            if (b is { IsNone: false }) c.Line(ax, ay, bx, by, b.Width, b.Color, b.Val.Contains("dash"), b.Val.Contains("dot"));
        }
        L(sp.PageBorderTop, x1, y1, x2, y1);
        L(sp.PageBorderBottom, x1, y2, x2, y2);
        L(sp.PageBorderLeft, x1, y1, x1, y2);
        L(sp.PageBorderRight, x2, y1, x2, y2);
    }
}

/// <summary>Короткая линия-разделитель сносок.</summary>
public sealed class SeparatorFrag : Frag
{
    public double Width;
    public override void Draw(PdfCanvas c, double x, double y) => c.Line(x, y, x + Width, y, 0.5, Rgb.Black);
}

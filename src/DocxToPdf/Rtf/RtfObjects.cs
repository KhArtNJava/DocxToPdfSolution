using System.Globalization;
using System.Text;
using DocxToPdf.Docx;
using DocxToPdf.Images;
using DocxToPdf.Pdf;

namespace DocxToPdf.Rtf;

public sealed partial class RtfReader
{
    // ================================================================
    //  Таблицы
    // ================================================================

    sealed class CellDef
    {
        public int Right;
        public CellBorders Borders = new();
        public int Shade = -1, ShadePct, ShadeFore;
        public string VAlign = "top";
        public VMerge VMerge;
        public bool HFirst, HMerge;
        public double? PadL, PadR, PadT, PadB;
        public string? TextDir;
    }

    sealed class RowDef
    {
        public int Left, Gap = 108;
        public List<CellDef> Cells = [];
        public int Height;
        public bool Header, Keep, Rtl;
        public string Align = "left";
        public CellBorders Borders = new();
        public double? PadL, PadR, PadT, PadB;
        public CellDef Pending = new();
        // Плавающая (позиционированная) таблица: \tpos*, \tdfrmtxt*
        public bool Floating;
        public string? PosXAlign;
        public int PosX, PosY, TxtL, TxtR, TxtT, TxtB;
        public bool PosYPage;
        public RowDef Clone()
        {
            var r = (RowDef)MemberwiseClone();
            r.Cells = [.. Cells];
            return r;
        }
    }

    sealed class TableBuilder
    {
        public readonly List<(RowDef Def, List<List<Block>> Cells)> Rows = [];
        public List<List<Block>> Cells = [];
        public List<Block> CurCell = [];
    }

    readonly Dictionary<int, RowDef> _rowDefs = [];
    int _rowDefLevel = 1;
    RowDef RD => _rowDefs.TryGetValue(_rowDefLevel, out var r) ? r : _rowDefs[_rowDefLevel] = new RowDef();

    void TableWord(Word w, State st)
    {
        int v = w.Param ?? 1;
        switch (w.Name)
        {
            case "trowd": _rowDefs[_rowDefLevel] = new RowDef(); break;
            case "tposx": RD.Floating = true; RD.PosX = v; break;
            case "tposxl" or "tposxi": RD.Floating = true; RD.PosXAlign = "left"; break;
            case "tposxc": RD.Floating = true; RD.PosXAlign = "center"; break;
            case "tposxr" or "tposxo": RD.Floating = true; RD.PosXAlign = "right"; break;
            case "tposy": RD.Floating = true; RD.PosY = v; break;
            case "tposnegy": RD.Floating = true; RD.PosY = v; break;
            case "tpvpg" or "tpvmrg": RD.PosYPage = true; break;
            case "tdfrmtxtLeft": RD.TxtL = v; break;
            case "tdfrmtxtRight": RD.TxtR = v; break;
            case "tdfrmtxtTop": RD.TxtT = v; break;
            case "tdfrmtxtBottom": RD.TxtB = v; break;
            case "trgaph": RD.Gap = v; break;
            case "trleft": RD.Left = v; break;
            case "trrh": RD.Height = v; break;
            case "trhdr": RD.Header = true; break;
            case "trkeep": RD.Keep = true; break;
            case "trqc": RD.Align = "center"; break;
            case "trqr": RD.Align = "right"; break;
            case "trql": RD.Align = "left"; break;
            case "rtlrow": RD.Rtl = true; break;
            case "trpaddl": RD.PadL = Tw(v); break;
            case "trpaddr": RD.PadR = Tw(v); break;
            case "trpaddt": RD.PadT = Tw(v); break;
            case "trpaddb": RD.PadB = Tw(v); break;
            case "trbrdrt": _curBorder = RD.Borders.Top = new BorderDef(); break;
            case "trbrdrb": _curBorder = RD.Borders.Bottom = new BorderDef(); break;
            case "trbrdrl": _curBorder = RD.Borders.Left = new BorderDef(); break;
            case "trbrdrr": _curBorder = RD.Borders.Right = new BorderDef(); break;
            case "trbrdrh": _curBorder = RD.Borders.InsideH = new BorderDef(); break;
            case "trbrdrv": _curBorder = RD.Borders.InsideV = new BorderDef(); break;
            case "clbrdrt": _curBorder = RD.Pending.Borders.Top = new BorderDef(); break;
            case "clbrdrb": _curBorder = RD.Pending.Borders.Bottom = new BorderDef(); break;
            case "clbrdrl": _curBorder = RD.Pending.Borders.Left = new BorderDef(); break;
            case "clbrdrr": _curBorder = RD.Pending.Borders.Right = new BorderDef(); break;
            case "clcbpat": RD.Pending.Shade = v; break;
            case "clshdng": RD.Pending.ShadePct = v; break;
            case "clcfpat": RD.Pending.ShadeFore = v; break;
            case "clvertalt": RD.Pending.VAlign = "top"; break;
            case "clvertalc": RD.Pending.VAlign = "center"; break;
            case "clvertalb": RD.Pending.VAlign = "bottom"; break;
            case "clvmgf": RD.Pending.VMerge = VMerge.Restart; break;
            case "clvmrg": RD.Pending.VMerge = VMerge.Continue; break;
            case "clmgf": RD.Pending.HFirst = true; break;
            case "clmrg": RD.Pending.HMerge = true; break;
            // Word меняет местами \clpadl и \clpadt (известная особенность формата)
            case "clpadl": RD.Pending.PadT = Tw(v); break;
            case "clpadt": RD.Pending.PadL = Tw(v); break;
            case "clpadr": RD.Pending.PadR = Tw(v); break;
            case "clpadb": RD.Pending.PadB = Tw(v); break;
            case "cltxbtlr": RD.Pending.TextDir = "btLr"; break;
            case "cltxtbrl" or "cltxtbrlv": RD.Pending.TextDir = "tbRl"; break;
            case "cellx":
                RD.Pending.Right = v;
                RD.Cells.Add(RD.Pending);
                RD.Pending = new CellDef();
                break;
        }
    }

    void EndCell(int level, State st)
    {
        // Последний абзац ячейки заканчивается \cell (без \par)
        if (_sink.Inlines.Count > 0 || _pendingLabel != null || _sink.Pending.Count > 0)
        {
            var p = st.P;
            bool wasIn = p.InTable; int wasItap = p.Itap;
            p.InTable = true; p.Itap = level;
            FlushParagraph(st);
            p.InTable = wasIn; p.Itap = wasItap;
        }
        while (_sink.Tables.Count < level) _sink.Tables.Add(new TableBuilder());
        CloseTables(level, st);
        var tb = _sink.Tables[level - 1];
        if (tb.CurCell.Count == 0) tb.CurCell.Add(MakeParagraph(st.P, st.C));
        tb.Cells.Add(tb.CurCell);
        tb.CurCell = [];
    }

    void EndRow(int level, State st)
    {
        while (_sink.Tables.Count < level) _sink.Tables.Add(new TableBuilder());
        CloseTables(level, st);
        var tb = _sink.Tables[level - 1];
        if (tb.CurCell.Count > 0) { tb.Cells.Add(tb.CurCell); tb.CurCell = []; }
        var def = _rowDefs.TryGetValue(level, out var d) ? d.Clone() : new RowDef();
        tb.Rows.Add((def, tb.Cells));
        tb.Cells = [];
    }

    /// <summary>Закрывает таблицы глубже уровня level: готовая таблица — блок в объемлющей ячейке (или в тексте).</summary>
    void CloseTables(int level, State st)
    {
        while (_sink.Tables.Count > level)
        {
            int idx = _sink.Tables.Count - 1;
            var tb = _sink.Tables[idx];
            _sink.Tables.RemoveAt(idx);
            if (tb.Cells.Count > 0 || tb.CurCell.Count > 0)
            {
                if (tb.CurCell.Count > 0) tb.Cells.Add(tb.CurCell);
                var def = _rowDefs.TryGetValue(idx + 1, out var d) ? d.Clone() : new RowDef();
                tb.Rows.Add((def, tb.Cells));
            }
            if (tb.Rows.Count == 0) continue;
            var table = BuildTable(tb);
            var parent = idx == 0 ? _sink.Blocks : _sink.Tables[idx - 1].CurCell;
            var fd = tb.Rows[0].Def;
            if (fd.Floating && idx == 0)
            {
                // Плавающая таблица: текст обтекает её с другой стороны
                double tw = table.Grid.Sum();
                double bodyW = _sect.PageWidth - _sect.MarginLeft - _sect.MarginRight;
                string side = fd.PosXAlign == "right" || fd.PosXAlign == null && Tw(fd.PosX) + tw / 2 > bodyW / 2 ? "right" : "left";
                table.Indent = 0; table.Align = "left";
                parent.Add(new FloatBlock
                {
                    Side = side, Width = tw, Blocks = [table],
                    MarginTop = fd.PosYPage ? 0 : Math.Max(0, Tw(fd.PosY)), MarginBottom = Tw(fd.TxtB),
                    MarginLeft = side == "right" ? Tw(fd.TxtL) : fd.PosXAlign == null ? Math.Max(0, Tw(fd.PosX)) : fd.PosXAlign == "center" ? Math.Max(0, (bodyW - tw) / 2) : 0,
                    MarginRight = side == "left" ? Tw(fd.TxtR) : 0,
                });
            }
            else parent.Add(table);
        }
    }

    Table BuildTable(TableBuilder tb)
    {
        // Сетка: все границы ячеек всех строк
        var bounds = new SortedSet<int>();
        foreach (var (def, _) in tb.Rows)
        {
            bounds.Add(def.Left);
            foreach (var c in def.Cells) bounds.Add(c.Right);
        }
        var bl = new List<int>();
        foreach (var b in bounds) if (bl.Count == 0 || b - bl[^1] > 10) bl.Add(b);
        if (bl.Count < 2) bl = [0, 9000];
        int Nearest(int x) { int best = 0; for (int i = 1; i < bl.Count; i++) if (Math.Abs(bl[i] - x) < Math.Abs(bl[best] - x)) best = i; return best; }
        var first = tb.Rows[0].Def;
        var table = new Table
        {
            Grid = Enumerable.Range(0, bl.Count - 1).Select(i => Tw(bl[i + 1] - bl[i])).ToList(),
            Indent = Tw(bl[0]),
            Align = first.Align,
            Rtl = first.Rtl,
            Borders = first.Borders,
            CellMargins = new Margins { Left = first.PadL ?? Tw(first.Gap), Right = first.PadR ?? Tw(first.Gap), Top = first.PadT ?? 0, Bottom = first.PadB ?? 0 },
        };
        foreach (var (def, cells) in tb.Rows)
        {
            var row = new TableRow
            {
                Height = Tw(Math.Abs(def.Height)), HeightRule = def.Height < 0 ? "exact" : "atLeast",
                IsHeader = def.Header, CantSplit = def.Keep,
            };
            int start = Nearest(def.Left);
            if (start > 0) row.Cells.Add(new TableCell { GridSpan = start, Borders = Nil() });
            for (int j = 0; j < def.Cells.Count; j++)
            {
                var cd = def.Cells[j];
                int end = Math.Max(start + 1, Nearest(cd.Right));
                var content = j < cells.Count ? cells[j] : [new Paragraph()];
                if (cd.HMerge && row.Cells.Count > 0)
                {
                    // Старое горизонтальное объединение: продолжение предыдущей ячейки
                    row.Cells[^1].GridSpan += end - start;
                    start = end;
                    continue;
                }
                var tc = new TableCell
                {
                    GridSpan = end - start,
                    VMerge = cd.VMerge,
                    VAlign = cd.VAlign,
                    Borders = cd.Borders,
                    TextDirection = cd.TextDir,
                    Blocks = content.Count > 0 ? content : [new Paragraph()],
                };
                if (cd.Shade > 0) tc.Shading = ColorAt(cd.Shade);
                else if (cd.ShadePct > 0)
                {
                    var fg = ColorAt(cd.ShadeFore) ?? Rgb.Black;
                    double a = cd.ShadePct / 10000.0;
                    tc.Shading = new Rgb(1 - (1 - fg.R) * a, 1 - (1 - fg.G) * a, 1 - (1 - fg.B) * a);
                }
                if (cd.PadL != null || cd.PadR != null || cd.PadT != null || cd.PadB != null)
                    tc.Margins = new Margins { Left = cd.PadL, Right = cd.PadR, Top = cd.PadT, Bottom = cd.PadB };
                row.Cells.Add(tc);
                start = end;
            }
            // Ячейки текста без определения (повреждённые файлы): добавляем по одной колонке
            if (row.Cells.Count == 0) row.Cells.Add(new TableCell { GridSpan = table.Grid.Count, Blocks = cells.FirstOrDefault() ?? [new Paragraph()] });
            int used = row.Cells.Sum(c => c.GridSpan);
            if (used < table.Grid.Count) row.Cells.Add(new TableCell { GridSpan = table.Grid.Count - used, Borders = Nil() });
            table.Rows.Add(row);
        }
        return table;
    }

    static CellBorders Nil() => new() { Top = new BorderDef { Val = "nil" }, Bottom = new BorderDef { Val = "nil" }, Left = new BorderDef { Val = "nil" }, Right = new BorderDef { Val = "nil" } };

    // ================================================================
    //  Таблица списков (если нет готовой метки \listtext)
    // ================================================================

    sealed class ListLevel
    {
        public int Nfc, Start = 1, Follow, Jc;
        public string Template = "";                  // символы < 9 — номера уровней
        public List<Node> Words = [];                 // оформление метки (\f, \fs, \b …)
    }

    readonly Dictionary<int, List<ListLevel>> _lists = [];
    readonly Dictionary<int, (int ListId, Dictionary<int, int> Starts)> _overrides = [];
    readonly Dictionary<int, int[]> _counters = [];
    readonly Dictionary<int, bool[]> _started = [];

    void ReadListTable(Group g)
    {
        foreach (var lg in g.Items.OfType<Group>())
        {
            if (lg.Items.FirstOrDefault() is not Word { Name: "list" }) continue;
            int id = lg.Items.OfType<Word>().LastOrDefault(w => w.Name == "listid")?.Param ?? 0;
            var levels = new List<ListLevel>();
            foreach (var lev in lg.Items.OfType<Group>().Where(x => x.Items.FirstOrDefault() is Word { Name: "listlevel" }))
                levels.Add(ReadLevel(lev));
            _lists[id] = levels;
        }
    }

    ListLevel ReadLevel(Group lev)
    {
        var l = new ListLevel();
        foreach (var it in lev.Items)
        {
            if (it is Word w)
            {
                switch (w.Name)
                {
                    case "levelnfc" or "levelnfcn": l.Nfc = w.Param ?? 0; break;
                    case "levelstartat": l.Start = w.Param ?? 1; break;
                    case "levelfollow": l.Follow = w.Param ?? 0; break;
                    case "leveljc" or "leveljcn": l.Jc = w.Param ?? 0; break;
                    case "listlevel" or "levelnumbers" or "leveltext" or "levelindent" or "levelspace" or "levellegal" or "levelnorestart" or "levelold" or "levelprev" or "levelprevspace" or "lvltentative" or "leveltemplateid" or "li" or "lin" or "fi" or "jclisttab" or "tx": break;
                    default: l.Words.Add(w); break;
                }
            }
            else if (it is Group sg && sg.Items.FirstOrDefault() is Word { Name: "leveltext" })
            {
                // Первый символ — длина шаблона
                var sb = new StringBuilder();
                for (int i = 1; i < sg.Items.Count; i++)
                {
                    var x = sg.Items[i];
                    if (x is Bytes b) foreach (var ch in b.Data) sb.Append(b.Hex && ch < 9 ? (char)ch : (char)Enc(_docCodepage).GetString([ch])[0]);
                    else if (x is Word { Name: "u" } u) { int code = u.Param ?? 0; if (code < 0) code += 65536; sb.Append((char)code); i = SkipFallback(sg, i, 1); }
                }
                string t = sb.ToString();
                if (t.Length > 0) { int len = t[0]; t = t[1..]; if (len < t.Length) t = t[..len]; }
                l.Template = t.TrimEnd(';');
            }
        }
        return l;
    }

    void ReadOverrides(Group g)
    {
        foreach (var og in g.Items.OfType<Group>())
        {
            int listId = 0, ls = 0;
            var starts = new Dictionary<int, int>();
            foreach (var w in og.Items.OfType<Word>())
            {
                if (w.Name == "listid") listId = w.Param ?? 0;
                else if (w.Name == "ls") ls = w.Param ?? 0;
            }
            int lvl = 0;
            foreach (var lo in og.Items.OfType<Group>().Where(x => x.Items.FirstOrDefault() is Word { Name: "lfolevel" }))
            {
                if (lo.Items.OfType<Word>().FirstOrDefault(w => w.Name == "listoverridestartat") != null)
                {
                    var start = lo.Items.OfType<Word>().FirstOrDefault(w => w.Name == "levelstartat")
                        ?? lo.Items.OfType<Group>().SelectMany(x => x.Items.OfType<Word>()).FirstOrDefault(w => w.Name == "levelstartat");
                    if (start != null) starts[lvl] = start.Param ?? 1;
                }
                lvl++;
            }
            if (ls > 0) _overrides[ls] = (listId, starts);
        }
    }

    static string ListNum(int n, int nfc) => nfc switch
    {
        1 => Roman(n), 2 => Roman(n).ToLowerInvariant(),
        3 => ((char)('A' + (n - 1) % 26)).ToString(), 4 => ((char)('a' + (n - 1) % 26)).ToString(),
        22 => n < 10 ? "0" + n : n.ToString(CultureInfo.InvariantCulture),
        255 => "",
        _ => n.ToString(CultureInfo.InvariantCulture),
    };

    List<Inline>? ListLabel(ParaProps p, CharProps c)
    {
        if (!_overrides.TryGetValue(p.Ls, out var ov) || !_lists.TryGetValue(ov.ListId, out var levels) || levels.Count == 0) return null;
        int lvl = Math.Clamp(p.Ilvl, 0, levels.Count - 1);
        var level = levels[lvl];
        if (!_counters.TryGetValue(p.Ls, out var cnt)) { _counters[p.Ls] = cnt = new int[10]; _started[p.Ls] = new bool[10]; }
        var started = _started[p.Ls];
        if (!started[lvl]) { cnt[lvl] = ov.Starts.TryGetValue(lvl, out var s0) ? s0 : level.Start; started[lvl] = true; }
        else cnt[lvl]++;
        for (int k = lvl + 1; k < 10; k++) started[k] = false;
        var sb = new StringBuilder();
        foreach (char ch in level.Template)
        {
            if (ch < 9)
            {
                int l2 = ch;
                var lv2 = l2 < levels.Count ? levels[l2] : level;
                int n = l2 <= lvl && started[l2] ? cnt[l2] : lv2.Start;
                sb.Append(ListNum(n, level.Nfc == 23 ? 0 : (l2 == lvl ? level.Nfc : lv2.Nfc)));
            }
            else sb.Append(ch);
        }
        var s = new State { C = c.Clone() };
        s.C.Underline = null;
        foreach (var w in level.Words.OfType<Word>()) ControlWord(w, s);
        var ts = Style(s.C) with { Link = null };
        var res = new List<Inline>();
        if (sb.Length > 0) res.Add(new TextInline(sb.ToString(), ts));
        if (level.Follow == 0) res.Add(new TabInline(ts));
        else if (level.Follow == 1) res.Add(new TextInline(" ", ts));
        return res;
    }

    // ================================================================
    //  Картинки
    // ================================================================

    int _imgCount;

    /// <summary>{\pict …}: PNG, JPEG, EMF, WMF, DIB; размер — \picwgoal/\pichgoal и масштаб, обрезка \piccrop*.</summary>
    ImageInline? Picture(Group g, State st, bool emit = true)
    {
        string kind = "";
        int picw = 0, pich = 0, wgoal = 0, hgoal = 0, sx = 100, sy = 100, cl = 0, cr = 0, ct = 0, cb = 0;
        var hex = new StringBuilder();
        byte[]? bin = null;
        foreach (var it in g.Items)
        {
            switch (it)
            {
                case Word w:
                    switch (w.Name)
                    {
                        case "pngblip": kind = "png"; break;
                        case "jpegblip": kind = "jpeg"; break;
                        case "emfblip": kind = "emf"; break;
                        case "wmetafile": kind = "wmf"; break;
                        case "dibitmap": kind = "dib"; break;
                        case "wbitmap": kind = "ddb"; break;
                        case "macpict": kind = "pict"; break;
                        case "picw": picw = w.Param ?? 0; break;
                        case "pich": pich = w.Param ?? 0; break;
                        case "picwgoal": wgoal = w.Param ?? 0; break;
                        case "pichgoal": hgoal = w.Param ?? 0; break;
                        case "picscalex": sx = w.Param ?? 100; break;
                        case "picscaley": sy = w.Param ?? 100; break;
                        case "piccropl": cl = w.Param ?? 0; break;
                        case "piccropr": cr = w.Param ?? 0; break;
                        case "piccropt": ct = w.Param ?? 0; break;
                        case "piccropb": cb = w.Param ?? 0; break;
                    }
                    break;
                case Bytes b when !b.Hex: hex.Append(Encoding.ASCII.GetString(b.Data)); break;
                case Bin bb: bin = bb.Data; break;
            }
        }
        byte[] data = bin ?? HexBytes(hex);
        if (data.Length == 0) return null;
        if (kind == "dib") data = DibToBmp(data);
        else if (kind is "ddb" or "pict") return null;
        bool meta = kind is "emf" or "wmf" || Metafile.IsMetafile(data);
        PdfImageData? info = meta ? null : ImageDecoder.Decode(data);
        if (!meta && info == null) return null;
        // Размер: желаемый (твипы) или по изображению (96 dpi)
        double w0 = wgoal > 0 ? Tw(wgoal) : info != null ? info.Width * 0.75 : picw * 72.0 / 2540;
        double h0 = hgoal > 0 ? Tw(hgoal) : info != null ? info.Height * 0.75 : pich * 72.0 / 2540;
        if (w0 <= 0 || h0 <= 0) { w0 = 100; h0 = 75; }
        string name = "rtf-img-" + (++_imgCount);
        _doc.Media[name] = data;
        double cropW = Tw(cl + cr), cropH = Tw(ct + cb);
        double dw = Math.Max(1, (w0 - cropW) * sx / 100), dh = Math.Max(1, (h0 - cropH) * sy / 100);
        var img = new ImageInline(name, dw, dh, Style(st.C) with { Link = _link }) { FrameW = dw, FrameH = dh };
        if (cl != 0 || cr != 0 || ct != 0 || cb != 0)
        {
            img.CropL = Tw(cl) / w0; img.CropR = Tw(cr) / w0; img.CropT = Tw(ct) / h0; img.CropB = Tw(cb) / h0;
        }
        if (emit && !st.C.Hidden) _sink.Inlines.Add(img);
        return img;
    }

    static byte[] HexBytes(StringBuilder s)
    {
        var res = new List<byte>(s.Length / 2);
        int hi = -1;
        for (int i = 0; i < s.Length; i++)
        {
            int v = Hex((byte)s[i]);
            if (v < 0) continue;
            if (hi < 0) hi = v; else { res.Add((byte)(hi << 4 | v)); hi = -1; }
        }
        return res.ToArray();
    }

    /// <summary>DIB (BITMAPINFOHEADER + пиксели) → файл BMP.</summary>
    static byte[] DibToBmp(byte[] dib)
    {
        if (dib.Length < 40) return dib;
        int hdr = BitConverter.ToInt32(dib, 0);
        int bits = BitConverter.ToUInt16(dib, 14);
        int used = BitConverter.ToInt32(dib, 32);
        int compression = BitConverter.ToInt32(dib, 16);
        int colors = used > 0 ? used : bits <= 8 ? 1 << bits : 0;
        int masks = compression == 3 && hdr == 40 ? 12 : 0;
        int off = 14 + hdr + masks + colors * 4;
        var res = new byte[14 + dib.Length];
        res[0] = (byte)'B'; res[1] = (byte)'M';
        BitConverter.GetBytes(res.Length).CopyTo(res, 2);
        BitConverter.GetBytes(off).CopyTo(res, 10);
        dib.CopyTo(res, 14);
        return res;
    }

    /// <summary>{\object …}: внедрённый объект — его результат (картинка).</summary>
    void ObjectGroup(Group g, State st)
    {
        var result = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "result" });
        if (result != null) Walk(StripFirst(result, false), st.Clone());
    }

    // ================================================================
    //  Фигуры (\shp)
    // ================================================================

    void Shape(Group g, State st)
    {
        var inst = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.Count > 1 && x.Items[0] is Sym { C: '*' } && x.Items[1] is Word { Name: "shpinst" });
        var rslt = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.Count > 0 && x.Items[0] is Word { Name: "shprslt" });
        if (inst == null) { if (rslt != null) Walk(StripFirst(rslt, false), st.Clone()); return; }

        int left = 0, top = 0, right = 0, bottom = 0, wr = 3, wrk = 0;
        string bx = "column", by = "para";
        bool behind = false, isGroup = false;
        var props = new Dictionary<string, string>();
        Group? pib = null, txt = null;
        foreach (var it in inst.Items)
        {
            if (it is Word w)
            {
                int v = w.Param ?? 0;
                switch (w.Name)
                {
                    case "shpleft": left = v; break;
                    case "shptop": top = v; break;
                    case "shpright": right = v; break;
                    case "shpbottom": bottom = v; break;
                    case "shpbxpage": bx = "page"; break;
                    case "shpbxmargin": bx = "margin"; break;
                    case "shpbxcolumn": bx = "column"; break;
                    case "shpbxignore": bx = "ignore"; break;
                    case "shpbypage": by = "page"; break;
                    case "shpbymargin": by = "margin"; break;
                    case "shpbypara": by = "para"; break;
                    case "shpbyignore": by = "ignore"; break;
                    case "shpwr": wr = v; break;
                    case "shpwrk": wrk = v; break;
                    case "shpfblwtxt": behind = v != 0; break;
                }
            }
            else if (it is Group sg)
            {
                var first = sg.Items.FirstOrDefault();
                if (first is Word { Name: "sp" })
                {
                    var sn = sg.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "sn" });
                    var sv = sg.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "sv" });
                    if (sn == null) continue;
                    string key = PlainText(StripFirst(sn, false)).Trim();
                    if (sv != null && sv.Items.OfType<Group>().FirstOrDefault(x => x.Items.FirstOrDefault() is Word { Name: "pict" }) is { } pg) { if (key == "pib") pib = pg; }
                    else props[key] = sv != null ? PlainText(StripFirst(sv, false)).Trim() : "";
                }
                else if (sg.Items.Count > 1 && sg.Items[0] is Sym { C: '*' } && sg.Items[1] is Word { Name: "shptxt" }) txt = sg;
                else if (first is Word { Name: "shptxt" }) txt = sg;
                else if (first is Word { Name: "shpgrp" }) isGroup = true;
            }
        }
        if (isGroup || g.Items.OfType<Group>().Any(x => x.Items.FirstOrDefault() is Word { Name: "shpgrp" }))
        {
            if (rslt != null) Walk(StripFirst(rslt, false), st.Clone());
            return;
        }
        double w0 = Math.Max(1, Tw(right - left)), h0 = Math.Max(1, Tw(bottom - top));
        int P(string k, int def) => props.TryGetValue(k, out var s) && int.TryParse(s, out var x) ? x : def;
        bool Bool(string k, bool def) => props.TryGetValue(k, out var s) ? s is "1" or "true" : def;
        static Rgb Bgr(int v) => new((v & 0xFF) / 255.0, (v >> 8 & 0xFF) / 255.0, (v >> 16 & 0xFF) / 255.0);
        int type = P("shapeType", 1);
        var style = Style(st.C);

        ImageInline? item = null;
        if (pib != null) item = Picture(pib, st, emit: false);
        else if (type == 75 && props.ContainsKey("pibName")) item = null;
        if (item != null) { item.Width = item.FrameW = w0; item.Height = item.FrameH = h0; }
        else if (props.TryGetValue("gtextUNICODE", out var wordArt) && wordArt.Length > 0)
        {
            // WordArt / водяной знак: текст по размеру фигуры
            var fill = Bool("fFilled", true) ? Bgr(P("fillColor", 0xC0C0C0)) : new Rgb(0.75, 0.75, 0.75);
            string font = props.GetValueOrDefault("gtextFont") ?? "Calibri";
            double fs = Math.Min(h0 * 0.85, w0 / Math.Max(1, wordArt.Length * 0.62));
            var ts = new TextStyle { FontAscii = font, FontHAnsi = font, FontEastAsia = font, Size = fs, Color = fill, Bold = true };
            var p = new Paragraph { Align = "center", MarkStyle = ts, SnapToGrid = false };
            p.Inlines.Add(new TextInline(wordArt, ts));
            item = new ImageInline("", w0, h0, style) { Content = [p], Box = new BoxStyle { MinHeight = h0, PadTop = Math.Max(0, (h0 - fs * 1.2) / 2) }, FrameW = w0, FrameH = h0 };
        }
        else
        {
            var shape = new ShapeSpec
            {
                Geom = type switch
                {
                    1 or 202 => "rect", 2 => "roundRect", 3 => "ellipse", 4 => "diamond", 5 => "triangle", 6 => "rtTriangle",
                    7 => "parallelogram", 9 => "hexagon", 10 => "octagon", 12 => "star5", 13 => "rightArrow", 66 => "leftArrow",
                    67 => "downArrow", 68 => "upArrow", 74 => "heart", 56 => "homePlate", 187 => "star4", 58 => "star8",
                    20 or 32 => "line", 11 => "plus", 116 => "flowChartTerminator", 110 => "flowChartDecision", 120 => "flowChartConnector",
                    _ => "rect",
                },
                Fill = Bool("fFilled", true) ? Bgr(P("fillColor", 0xFFFFFF)) : null,
                Stroke = Bool("fLine", true) ? Bgr(P("lineColor", 0)) : null,
                StrokeWidth = P("lineWidth", 9525) / 12700.0,
                Dashed = P("lineDashing", 0) is >= 1 and not 2,
                Dotted = P("lineDashing", 0) == 2,
                TailArrow = P("lineEndArrowhead", 0) > 0, HeadArrow = P("lineStartArrowhead", 0) > 0,
            };
            if (shape.Geom == "line") shape.Fill = null;
            item = new ImageInline("", w0, h0, style) { Shape = shape, FrameW = w0, FrameH = h0 };
            if (Bool("fFlipH", false)) item.FlipH = true;
            if (Bool("fFlipV", false)) item.FlipV = true;
            if (txt != null)
            {
                // Текст надписи
                var saved = _sink;
                var savedLabel = _pendingLabel; _pendingLabel = null;
                _sink = new Sink();
                var s2 = new State { C = _defChar.Clone(), P = _defPara.Clone(), Uc = st.Uc };
                Walk(StripFirst(txt, txt.Items[0] is Sym), s2);
                FlushParagraph(s2, force: false);
                CloseTables(0, s2);
                var blocks = _sink.Blocks;
                _sink = saved; _pendingLabel = savedLabel;
                if (blocks.Count > 0)
                {
                    item.Content = blocks;
                    item.Box = new BoxStyle
                    {
                        PadLeft = P("dxTextLeft", 91440) / 12700.0, PadRight = P("dxTextRight", 91440) / 12700.0,
                        PadTop = P("dyTextTop", 45720) / 12700.0, PadBottom = P("dyTextBottom", 45720) / 12700.0,
                    };
                    item.VAnchor = P("anchorText", 0) switch { 1 or 4 => "ctr", 2 or 5 => "b", _ => "t" };
                    if (type == 202 && !Bool("fFilled", true) && !Bool("fLine", true)) item.Shape = null;
                }
            }
        }
        if (item == null) { if (rslt != null) Walk(StripFirst(rslt, false), st.Clone()); return; }
        if (props.TryGetValue("rotation", out var rotS) && int.TryParse(rotS, out var rot) && rot != 0) item.Rotation = ((rot / 65536.0) % 360 + 360) % 360;

        // Положение: выравнивание (posh/posv) или смещение от страницы, полей, колонки, абзаца
        int posh = P("posh", 0), posv = P("posv", 0), posrelh = P("posrelh", bx switch { "page" => 1, "margin" => 0, _ => 2 }), posrelv = P("posrelv", by switch { "page" => 1, "margin" => 0, _ => 2 });
        if (bx != "ignore") posrelh = bx switch { "page" => 1, "margin" => 0, _ => 2 };
        if (by != "ignore") posrelv = by switch { "page" => 1, "margin" => 0, _ => 2 };
        string? alignX = posh switch { 1 or 4 => "left", 2 => "center", 3 or 5 => "right", _ => null };
        string? alignY = posv switch { 1 => "top", 2 => "center", 3 => "bottom", _ => null };
        double x = Tw(left), y = Tw(top);
        string hRel = posrelh switch { 1 => "page", 0 => "margin", _ => "margin" };
        string vRel = posrelv switch { 1 => "page", 0 => "margin", _ => "flow" };
        double dl = P("dxWrapDistLeft", 114300) / 12700.0, dr = P("dxWrapDistRight", 114300) / 12700.0, db = P("dyWrapDistBottom", 0) / 12700.0;

        var para = new Paragraph { MarkStyle = style with { Size = 1 }, SnapToGrid = false };
        para.Inlines.Add(item);
        var content = new List<Block> { para };
        double bodyW = _sect.PageWidth - _sect.MarginLeft - _sect.MarginRight;
        if (wr == 1 && vRel == "flow")
        {
            // Сверху и снизу: отдельной строкой
            para.Align = alignX switch { "center" => "center", "right" => "right", _ => "left" };
            if (alignX == null) para.IndLeft = Math.Max(0, hRel == "page" ? x - _sect.MarginLeft : x);
            para.SpaceBefore = Math.Max(0, y); para.SpaceAfter = db;
            _sink.Pending.Add(para);
            return;
        }
        if (wr is 2 or 4 or 5 && !behind && vRel == "flow" && y < _sect.PageHeight * 0.4)
        {
            double xRel = hRel == "page" ? x - _sect.MarginLeft : x;
            string side = wrk switch { 1 => "right", 2 => "left", _ => alignX == "right" ? "right" : alignX == "left" ? "left" : xRel + w0 / 2 > bodyW / 2 ? "right" : "left" };
            _sink.Pending.Add(new FloatBlock
            {
                Side = side, Width = item.Width, Blocks = content,
                MarginTop = alignY == null ? Math.Max(0, y) : 0, MarginBottom = db,
                MarginLeft = side == "left" ? (alignX == null ? Math.Max(0, xRel) : alignX == "center" ? Math.Max(0, (bodyW - w0) / 2) : 0) : dl,
                MarginRight = side == "left" ? dr : 0,
            });
            return;
        }
        _sink.Pending.Add(new PositionedBlock
        {
            Width = item.Width, Blocks = content, ZIndex = behind ? -1 : 1,
            HRel = hRel, VRel = vRel, AlignX = alignX, AlignY = alignY,
            Left = alignX == null ? x : null, Top = alignY == null ? y : null,
        });
    }
}

using System.Text;
using System.Xml.Linq;
using DocxToPdf.Docx;

namespace DocxToPdf.Rtf;

/// <summary>RTF: формулы Word (\mmath), примечания (\annotation).</summary>
public sealed partial class RtfReader
{
    // ================================================================
    //  Формулы (\mmath)
    // ================================================================

    static readonly XNamespace MNs = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>
    /// {\mmath …}: формула Office Math. Управляющие слова \mXxx соответствуют элементам OMML m:xxx
    /// (\moMathPara → m:oMathPara, \mf → m:f, \mr → m:r …), текст свойств (\mchr, \mbegChr, \mtype …) — атрибут m:val.
    /// Дерево OMML разбирается так же, как в DOCX.
    /// </summary>
    void RtfMath(Group g, State st, bool star)
    {
        if (st.C.Hidden) return;
        var style = Style(st.C) with { Link = _link };
        foreach (var sub in StripFirst(g, star).Items.OfType<Group>())
        {
            if (MathXml(sub) is not { } x) continue;
            if (x.Name.LocalName is not ("oMathPara" or "oMath")) x = new XElement(MNs + "oMath", x);
            DocxReader.MathToInlines(x, style, _sink.Inlines);
        }
    }

    /// <summary>Имя элемента OMML для группы формулы ({\mf …} → "f"), иначе null.</summary>
    static (string Local, Word Name)? MathName(Group g)
    {
        int k = g.Items.Count > 0 && g.Items[0] is Sym { C: '*' } ? 1 : 0;
        if (k >= g.Items.Count || g.Items[k] is not Word w || w.Name.Length < 2 || w.Name[0] != 'm') return null;
        if (w.Name is "mmath" or "mmathPict" or "mmathPr") return null;
        return (char.ToLowerInvariant(w.Name[1]) + w.Name[2..], w);
    }

    XElement? MathXml(Group g)
    {
        if (MathName(g) is not { } nm) return null;
        var el = new XElement(MNs + nm.Local);
        var text = new Group();
        var children = new List<XElement>();
        bool afterName = false;
        foreach (var it in g.Items)
        {
            if (!afterName) { if (ReferenceEquals(it, nm.Name)) afterName = true; continue; }
            if (it is Group sg && MathName(sg) != null) { if (MathXml(sg) is { } c) children.Add(c); }
            else if (it is Group or Bytes or Sym || it is Word { Name: "u" }) text.Items.Add(it);
        }
        string t = PlainText(text);
        if (nm.Local == "r")
        {
            // Фрагмент формулы: свойства (\msty, \mscr, \mnor) — в m:rPr, текст — m:t
            // Свойства: {\mrPr{\msty p}} или сразу {\msty p}
            var props = children.SelectMany(c => c.Name.LocalName == "rPr" ? c.Elements() : [c]).ToList();
            if (props.Count > 0) el.Add(new XElement(MNs + "rPr", props));
            if (t.Length > 0) el.Add(new XElement(MNs + "t", t));
            return el;
        }
        if (children.Count > 0) el.Add(children);
        else if (nm.Name.Param is int p) el.SetAttributeValue(MNs + "val", p);
        else if (t.Trim().Length > 0) el.SetAttributeValue(MNs + "val", t.Trim());
        return el;
    }

    // ================================================================
    //  Примечания (\annotation)
    // ================================================================

    // {\*\atrfstart N}текст{\*\atrfend N}{\*\atnid инициалы}{\*\atnauthor автор}\chatn{\*\annotation{\*\atnref N}…текст…}
    string? _commentMode;
    bool _commentModeSet;
    readonly Dictionary<string, CommentInfo> _atnById = [];
    readonly List<CommentInfo> _atnQueue = [];
    int _atnNext;
    readonly HashSet<string> _activeAtn = [];
    readonly List<CommentInfo> _atnList = [];
    CommentInfo? _lastAtn;

    /// <summary>Предварительный проход: автор и инициалы каждого примечания (нужны раньше, чем встретится само примечание).</summary>
    void ScanAnnotations(Group g)
    {
        string author = "", initials = "";
        foreach (var it in g.Items)
        {
            if (it is not Group sg) continue;
            int k = sg.Items.Count > 0 && sg.Items[0] is Sym { C: '*' } ? 1 : 0;
            string? dest = k < sg.Items.Count && sg.Items[k] is Word w ? w.Name : null;
            switch (dest)
            {
                case "atnauthor": author = PlainText(StripFirst(sg, k == 1)).Trim(); break;
                case "atnid": initials = PlainText(StripFirst(sg, k == 1)).Trim(); break;
                case "annotation":
                {
                    var info = new CommentInfo { Author = author, Initials = initials, Color = AtnColor(author) };
                    var refG = sg.Items.OfType<Group>().FirstOrDefault(x => x.Items.Skip(x.Items.Count > 0 && x.Items[0] is Sym ? 1 : 0).FirstOrDefault() is Word { Name: "atnref" });
                    if (refG != null) _atnById[PlainText(StripFirst(refG, refG.Items[0] is Sym)).Trim()] = info;
                    _atnQueue.Add(info);
                    author = initials = "";
                    break;
                }
                default: ScanAnnotations(sg); break;
            }
        }
    }

    Rgb AtnColor(string author)
    {
        if (!_authorIdx.TryGetValue(author, out int ai)) _authorIdx[author] = ai = _authorIdx.Count;
        return AuthorColors[ai % AuthorColors.Length];
    }

    /// <summary>Режим вывода примечаний: выноски (если правое поле достаточно широкое), список в конце или нет.</summary>
    string? CommentMode()
    {
        if (_commentModeSet) return _commentMode;
        _commentModeSet = true;
        _commentMode = _options.Comments switch
        {
            "none" => null, "balloons" => "balloons", "end" => "end",
            _ => _sect.MarginRight >= 70 ? "balloons" : "end",
        };
        return _commentMode;
    }

    /// <summary>Подсветка диапазона примечания — бледным цветом автора.</summary>
    Rgb? AnnotationBackground()
    {
        if (_activeAtn.Count == 0 || _commentMode == null) return null;
        var c = _atnById.TryGetValue(_activeAtn.First(), out var info) ? info.Color : AuthorColors[0];
        return new Rgb(1 - (1 - c.R) * 0.2, 1 - (1 - c.G) * 0.2, 1 - (1 - c.B) * 0.2);
    }

    /// <summary>\chatn — место примечания: выноска на поле или номер [n].</summary>
    void AnnotationMark(State st)
    {
        if (CommentMode() == null || _atnNext >= _atnQueue.Count) return;
        var info = _atnQueue[_atnNext++];
        _lastAtn = info;
        if (info.Number == 0) { info.Number = _atnList.Count + 1; _atnList.Add(info); }
        var style = Style(st.C);
        if (_commentMode == "balloons") _sink.Inlines.Add(new CommentInline(info, style));
        else _sink.Inlines.Add(new TextInline($"[{info.Number}]", style with { VertAlign = 1, Color = info.Color, Underline = false, Strike = false, Background = null }));
    }

    /// <summary>{\*\annotation …}: текст примечания.</summary>
    void Annotation(Group g, State st, bool star)
    {
        if (CommentMode() == null) return;
        var refG = g.Items.OfType<Group>().FirstOrDefault(x => x.Items.Skip(x.Items.Count > 0 && x.Items[0] is Sym ? 1 : 0).FirstOrDefault() is Word { Name: "atnref" });
        CommentInfo? info = refG != null && _atnById.TryGetValue(PlainText(StripFirst(refG, refG.Items[0] is Sym)).Trim(), out var byId) ? byId : _lastAtn;
        if (info == null) return;
        if (info.Number == 0)
        {
            // Примечание без \chatn — метка здесь
            if (_atnNext < _atnQueue.Count && ReferenceEquals(_atnQueue[_atnNext], info)) _atnNext++;
            _lastAtn = info;
            info.Number = _atnList.Count + 1; _atnList.Add(info);
            var style = Style(st.C);
            if (_commentMode == "balloons") _sink.Inlines.Add(new CommentInline(info, style));
            else _sink.Inlines.Add(new TextInline($"[{info.Number}]", style with { VertAlign = 1, Color = info.Color, Underline = false, Strike = false, Background = null }));
        }
        if (info.Blocks.Count > 0) return;
        var saved = _sink;
        var savedLabel = _pendingLabel; _pendingLabel = null;
        var savedActive = _activeAtn.ToList(); _activeAtn.Clear();
        var savedMode = _commentMode; _commentMode = null;
        _sink = new Sink();
        var s2 = new State { C = _defChar.Clone(), P = _defPara.Clone(), Uc = st.Uc };
        Walk(StripFirst(g, star), s2);
        FlushParagraph(s2, force: false);
        CloseTables(0, s2);
        info.Blocks.AddRange(_sink.Blocks);
        _sink = saved; _pendingLabel = savedLabel; _commentMode = savedMode;
        foreach (var a in savedActive) _activeAtn.Add(a);
    }

    /// <summary>Режим «список в конце»: заголовок и тексты примечаний с номерами.</summary>
    void AppendAnnotationList()
    {
        _doc.CommentBalloons = _commentMode == "balloons" && _atnList.Count > 0;
        if (_commentMode != "end" || _atnList.Count == 0) return;
        var hs = new TextStyle { FontAscii = "Times New Roman", FontHAnsi = "Times New Roman", FontEastAsia = "Times New Roman", Size = 12, Bold = true };
        var head = new Paragraph { MarkStyle = hs, SpaceBefore = 18, SpaceAfter = 6, KeepNext = true };
        head.Inlines.Add(new TextInline("Примечания", hs));
        _sink.Blocks.Add(head);
        foreach (var c in _atnList)
        {
            var ms = hs with { Size = 9, Color = c.Color };
            var p = new Paragraph { MarkStyle = ms, SpaceAfter = 2 };
            p.Inlines.Add(new TextInline($"[{c.Number}] {c.Author}:", ms));
            _sink.Blocks.Add(p);
            _sink.Blocks.AddRange(c.Blocks);
        }
    }
}

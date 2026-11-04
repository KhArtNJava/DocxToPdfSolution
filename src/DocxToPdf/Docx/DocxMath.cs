using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>Формулы Office Math (OMML) → линейная запись: дроби, индексы, корни, скобки, суммы, матрицы.</summary>
public sealed partial class DocxReader
{
    static string MVal(XElement? e, string name) => (string?)e?.Element(M + name)?.Attribute(M + "val") ?? "";

    /// <summary>«Простое» выражение — без скобок в линейной записи (одна буква/число).</summary>
    static bool SimpleMath(XElement? e)
    {
        if (e == null) return true;
        var runs = e.Elements().ToList();
        if (runs.Count != 1 || runs[0].Name != M + "r") return false;
        string t = string.Concat(runs[0].Elements(M + "t").Select(x => x.Value));
        return t.Length <= 3 && t.All(char.IsLetterOrDigit);
    }

    void MathToInlines(XElement e, TextStyle st, List<Inline> output)
    {
        void Text(string s, TextStyle? style = null) { if (s.Length > 0) output.Add(new TextInline(s, style ?? st)); }
        void Arg(XElement? x, TextStyle style) { if (x != null) MathToInlines(x, style, output); }
        TextStyle Sup(TextStyle s) => s with { VertAlign = s.VertAlign == -1 ? 0 : 1, Size = s.VertAlign != 0 ? s.Size * 0.8 : s.Size };
        TextStyle Sub(TextStyle s) => s with { VertAlign = s.VertAlign == 1 ? 0 : -1, Size = s.VertAlign != 0 ? s.Size * 0.8 : s.Size };
        void Wrapped(XElement? x) { if (SimpleMath(x)) Arg(x, st); else { Text("("); Arg(x, st); Text(")"); } }

        foreach (var c in e.Elements())
        {
            var n = c.Name;
            if (n.Namespace == W) { if (n == W + "r") foreach (var t in c.Elements(W + "t")) Text(t.Value); continue; }
            if (n.Namespace != M) continue;
            switch (n.LocalName)
            {
                case "r":
                {
                    string t = string.Concat(c.Elements(M + "t").Select(x => x.Value));
                    // Переменные — курсивом, если не задан прямой стиль (m:sty p)
                    bool plain = MVal(c.Element(M + "rPr"), "sty") == "p" || c.Element(M + "rPr")?.Element(M + "nor") != null;
                    var rs = st;
                    if (!plain && t.Any(char.IsLetter) && t.Length <= 2) rs = st with { Italic = true };
                    Text(t, rs);
                    break;
                }
                case "f":
                    Wrapped(c.Element(M + "num"));
                    Text(MVal(c.Element(M + "fPr"), "type") == "lin" ? "/" : "/");
                    Wrapped(c.Element(M + "den"));
                    break;
                case "sSup": Arg(c.Element(M + "e"), st); Arg(c.Element(M + "sup"), Sup(st)); break;
                case "sSub": Arg(c.Element(M + "e"), st); Arg(c.Element(M + "sub"), Sub(st)); break;
                case "sSubSup": Arg(c.Element(M + "e"), st); Arg(c.Element(M + "sub"), Sub(st)); Arg(c.Element(M + "sup"), Sup(st)); break;
                case "sPre": Arg(c.Element(M + "sub"), Sub(st)); Arg(c.Element(M + "sup"), Sup(st)); Arg(c.Element(M + "e"), st); break;
                case "rad":
                {
                    var deg = c.Element(M + "deg");
                    bool hide = MVal(c.Element(M + "radPr"), "degHide") is "1" or "on" or "true";
                    if (!hide && deg != null && deg.HasElements) Arg(deg, Sup(st));
                    Text("√");
                    Wrapped(c.Element(M + "e"));
                    break;
                }
                case "d":
                {
                    var pr = c.Element(M + "dPr");
                    string beg = pr?.Element(M + "begChr") != null ? MVal(pr, "begChr") : "(";
                    string end = pr?.Element(M + "endChr") != null ? MVal(pr, "endChr") : ")";
                    string sep = pr?.Element(M + "sepChr") != null ? MVal(pr, "sepChr") : "|";
                    Text(beg);
                    bool first = true;
                    foreach (var arg in c.Elements(M + "e")) { if (!first) Text(sep); Arg(arg, st); first = false; }
                    Text(end);
                    break;
                }
                case "nary":
                {
                    var pr = c.Element(M + "naryPr");
                    string chr = pr?.Element(M + "chr") != null ? MVal(pr, "chr") : "∫";
                    Text(chr, st with { Size = st.Size * 1.2 });
                    Arg(c.Element(M + "sub"), Sub(st));
                    Arg(c.Element(M + "sup"), Sup(st));
                    Text(" ");
                    Arg(c.Element(M + "e"), st);
                    break;
                }
                case "func": Arg(c.Element(M + "fName"), st); Text(" "); Wrapped(c.Element(M + "e")); break;
                case "limLow": Arg(c.Element(M + "e"), st); Arg(c.Element(M + "lim"), Sub(st)); break;
                case "limUpp": Arg(c.Element(M + "e"), st); Arg(c.Element(M + "lim"), Sup(st)); break;
                case "acc":
                {
                    var pr = c.Element(M + "accPr");
                    string chr = pr?.Element(M + "chr") != null ? MVal(pr, "chr") : "̂";
                    Arg(c.Element(M + "e"), st);
                    Text(chr);
                    break;
                }
                case "bar": Arg(c.Element(M + "e"), st); Text("̅"); break;
                case "eqArr":
                {
                    bool first = true;
                    foreach (var row in c.Elements(M + "e"))
                    {
                        if (!first) output.Add(new BreakInline(BreakKind.Line, st));
                        Arg(row, st);
                        first = false;
                    }
                    break;
                }
                case "m":
                {
                    Text("[");
                    bool firstRow = true;
                    foreach (var mr in c.Elements(M + "mr"))
                    {
                        if (!firstRow) Text("; ");
                        bool firstCell = true;
                        foreach (var me in mr.Elements(M + "e")) { if (!firstCell) Text("  "); Arg(me, st); firstCell = false; }
                        firstRow = false;
                    }
                    Text("]");
                    break;
                }
                case "rPr" or "ctrlPr" or "oMathParaPr" or "fPr" or "dPr" or "naryPr" or "radPr" or "accPr" or "sSupPr" or "sSubPr" or "funcPr":
                    break;
                default:
                    MathToInlines(c, st, output);   // oMath, e, groupChr, box, borderBox, phant …
                    break;
            }
        }
    }
}

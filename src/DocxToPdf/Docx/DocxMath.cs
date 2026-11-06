using System.Xml.Linq;

namespace DocxToPdf.Docx;

/// <summary>Формулы Office Math (OMML) → дерево для двумерной вёрстки (Layout/MathLayout.cs).</summary>
public sealed partial class DocxReader
{
    static string MVal(XElement? e, string name) => (string?)e?.Element(M + name)?.Attribute(M + "val") ?? "";
    static bool MHas(XElement? e, string name) => e?.Element(M + name) != null;
    static bool MOn(XElement? e, string name) => e?.Element(M + name) is { } x && (string?)x.Attribute(M + "val") is null or "1" or "on" or "true";

    /// <summary>Содержимое элемента (m:e, m:num …) как строка формулы.</summary>
    internal static MathNode MathRow(XElement? e)
    {
        var items = new List<MathNode>();
        if (e != null)
            foreach (var c in e.Elements())
                if (MathElement(c) is { } node) items.Add(node);
        return items.Count == 1 ? items[0] : new MRow(items);
    }

    static MathNode? MathElement(XElement c)
    {
        var n = c.Name;
        if (n.Namespace == W)
        {
            // Обычные фрагменты внутри формулы
            if (n == W + "r") { var t = string.Concat(c.Elements(W + "t").Select(x => x.Value)); return t.Length > 0 ? new MText(t, true, false) : null; }
            return null;
        }
        if (n.Namespace != M) return null;
        var pr = c.Elements().FirstOrDefault(x => x.Name.LocalName.EndsWith("Pr"));
        switch (n.LocalName)
        {
            case "r":
            {
                string t = string.Concat(c.Elements(M + "t").Select(x => x.Value));
                if (t.Length == 0) return null;
                var rp = c.Element(M + "rPr");
                string sty = MVal(rp, "sty");
                bool plain = sty is "p" or "b" || MHas(rp, "nor");
                return new MText(t, plain, sty is "b" or "bi");
            }
            case "f":
            {
                string type = MVal(pr, "type");
                return new MFrac(MathRow(c.Element(M + "num")), MathRow(c.Element(M + "den")), type.Length > 0 ? type : "bar");
            }
            case "sSup": return new MScript(MathRow(c.Element(M + "e")), null, MathRow(c.Element(M + "sup")), false);
            case "sSub": return new MScript(MathRow(c.Element(M + "e")), MathRow(c.Element(M + "sub")), null, false);
            case "sSubSup": return new MScript(MathRow(c.Element(M + "e")), MathRow(c.Element(M + "sub")), MathRow(c.Element(M + "sup")), false);
            case "sPre": return new MScript(MathRow(c.Element(M + "e")), MathRow(c.Element(M + "sub")), MathRow(c.Element(M + "sup")), true);
            case "rad":
            {
                bool hide = MOn(pr, "degHide");
                var deg = c.Element(M + "deg");
                return new MRad(hide || deg == null || !deg.HasElements ? null : MathRow(deg), MathRow(c.Element(M + "e")));
            }
            case "nary":
            {
                string chr = MHas(pr, "chr") ? MVal(pr, "chr") : "∫";
                bool integral = chr is "∫" or "∬" or "∭" or "∮" or "∯" or "∰";
                string loc = MVal(pr, "limLoc");
                bool undOvr = loc.Length > 0 ? loc == "undOvr" : !integral;
                var sub = MOn(pr, "subHide") ? null : c.Element(M + "sub");
                var sup = MOn(pr, "supHide") ? null : c.Element(M + "sup");
                return new MNary(chr, sub is { HasElements: true } ? MathRow(sub) : null, sup is { HasElements: true } ? MathRow(sup) : null,
                    MathRow(c.Element(M + "e")), undOvr);
            }
            case "d":
            {
                string beg = MHas(pr, "begChr") ? MVal(pr, "begChr") : "(";
                string end = MHas(pr, "endChr") ? MVal(pr, "endChr") : ")";
                string sep = MHas(pr, "sepChr") ? MVal(pr, "sepChr") : "|";
                return new MDelim(beg, end, sep, c.Elements(M + "e").Select(MathRow).ToList());
            }
            case "m":
                return new MMatrix(c.Elements(M + "mr").Select(r => r.Elements(M + "e").Select(MathRow).ToList()).ToList());
            case "func": return new MFunc(MathRow(c.Element(M + "fName")), MathRow(c.Element(M + "e")));
            case "acc": return new MAcc(MHas(pr, "chr") ? MVal(pr, "chr") : "̂", MathRow(c.Element(M + "e")));
            case "bar": return new MBar(MVal(pr, "pos") == "top", MathRow(c.Element(M + "e")));
            case "groupChr":
            {
                string chr = MHas(pr, "chr") ? MVal(pr, "chr") : "⏟";
                bool top = MVal(pr, "pos") == "top";
                return new MGroupChr(chr, top, MathRow(c.Element(M + "e")));
            }
            case "limLow": return new MLim(MathRow(c.Element(M + "e")), MathRow(c.Element(M + "lim")), false);
            case "limUpp": return new MLim(MathRow(c.Element(M + "e")), MathRow(c.Element(M + "lim")), true);
            case "eqArr": return new MEqArr(c.Elements(M + "e").Select(MathRow).ToList());
            case "box": return new MBox(MathRow(c.Element(M + "e")), false);
            case "borderBox": return new MBox(MathRow(c.Element(M + "e")), true);
            case "phant": return MOn(pr, "show") || pr == null ? MathRow(c.Element(M + "e")) : null;
            case "oMath": return MathRow(c);
            default:
                if (n.LocalName.EndsWith("Pr")) return null;
                return c.HasElements ? MathRow(c) : null;
        }
    }

    /// <summary>Формула (m:oMath / m:oMathPara) → строчные объекты двумерной вёрстки.</summary>
    internal static void MathToInlines(XElement e, TextStyle st, List<Inline> output)
    {
        if (e.Name == M + "oMathPara")
        {
            // Каждая формула блока — отдельной строкой
            bool first = true;
            foreach (var om in e.Elements(M + "oMath"))
            {
                if (!first) output.Add(new BreakInline(BreakKind.Line, st));
                first = false;
                output.Add(new ImageInline("", 0, 0, st) { Math = MathRow(om), MathDisplay = true });
            }
            return;
        }
        output.Add(new ImageInline("", 0, 0, st) { Math = MathRow(e), MathDisplay = false });
    }
}

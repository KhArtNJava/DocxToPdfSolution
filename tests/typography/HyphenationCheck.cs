using DocxToPdf.Layout;
var lang = args[0];
var h = Hyphenator.For(lang)!;
foreach (var w in File.ReadAllLines(args[1]))
{
    var pts = h.Points(w);
    var sb = new System.Text.StringBuilder(); int last = 0;
    foreach (var p in pts) { sb.Append(w[last..p]).Append('-'); last = p; }
    sb.Append(w[last..]);
    Console.WriteLine(sb);
}

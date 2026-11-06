using System.Globalization;
using System.Text.RegularExpressions;
using DocxToPdf.Fonts;
string dir = args[0];
int fails = 0;
double[] Nums(string s) => Regex.Matches(s, @"-?\d+(\.\d+)?").Select(m => double.Parse(m.Value, CultureInfo.InvariantCulture)).ToArray();
foreach (var (vf, rf) in new[] { ("VarTestTT.ttf", "ref_VarTestTT_{0}.ttf"), ("VarTestTT_noHVAR.ttf", "ref_VarTestTT_{0}.ttf"), ("VarTestCFF2.otf", "ref_VarTestCFF2_{0}.otf") })
foreach (var w in new[] { 400, 550, 700 })
{
    var f = TrueTypeFont.Load(Path.Combine(dir, vf), 0, new Dictionary<string, double> { ["wght"] = w });
    var r = TrueTypeFont.Load(Path.Combine(dir, string.Format(rf, w)), 0);
    int bad = 0, advBad = 0; double maxd = 0;
    for (ushort g = 0; g < f.NumGlyphs; g++)
    {
        if (Math.Abs(f.Advance(g) - r.Advance(g)) * f.UnitsPerEm > 1.01) { advBad++; if (advBad < 3) Console.WriteLine($"  adv g{g}: {f.Advance(g) * f.UnitsPerEm} vs {r.Advance(g) * r.UnitsPerEm}"); }
        var a = f.GlyphPath(g); var b = r.GlyphPath(g);
        var na = Nums(a); var nb = Nums(b);
        bool ok = na.Length == nb.Length && Regex.Replace(a, @"[-\d.]+ ", "") == Regex.Replace(b, @"[-\d.]+ ", "");
        if (ok) for (int i = 0; i < na.Length; i++) { maxd = Math.Max(maxd, Math.Abs(na[i] - nb[i])); if (Math.Abs(na[i] - nb[i]) > 2.01) ok = false; }
        if (!ok) { bad++; if (bad < 3) Console.WriteLine($"  g{g}:\n    {a[..Math.Min(200, a.Length)]}\n    {b[..Math.Min(200, b.Length)]}"); }
    }
    // Подмножество из всех глифов
    var sub = f.Subset(Enumerable.Range(0, f.NumGlyphs).Select(i => (ushort)i));
    File.WriteAllBytes(Path.Combine(dir, $"out_{Path.GetFileNameWithoutExtension(vf)}_{w}" + (f.IsCff ? ".cff" : ".ttf")), sub);
    Console.WriteLine($"{vf} wght={w}: glyphs={f.NumGlyphs} pathBad={bad} advBad={advBad} maxDiff={maxd:0.##} bold={f.IsBold} ps={f.PostScriptName}");
    fails += bad + advBad;
}
foreach (var vf in new[] { "VarTestTT.ttf", "VarTestCFF2.otf" })
    foreach (var fi in TrueTypeFont.ReadInfo(Path.Combine(dir, vf)))
        Console.WriteLine($"  face: {string.Join("/", fi.Families)} bold={fi.Bold} it={fi.Italic} ps={fi.PostScriptName} coords={(fi.Coords == null ? "-" : string.Join(",", fi.Coords.Select(k => k.Key + "=" + k.Value)))}");
Console.WriteLine("FAILS " + fails);

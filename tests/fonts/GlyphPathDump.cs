using DocxToPdf.Fonts;
var f = TrueTypeFont.Load(args[0], 0);
using var w = new StreamWriter(args[1]);
for (ushort g = 0; g < f.NumGlyphs; g++) w.WriteLine(f.GlyphPath(g));

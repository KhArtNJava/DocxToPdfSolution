# Шрифты OpenType CFF (.otf)

- `otf-fonts.html` — Inter (обычный, жирный, курсив), Loma (тайский), Unifont (CJK, символы); сравнение с Chromium.
- `otf-3d-paths.html` — текст в 3D-преобразовании выводится контурами глифов (интерпретатор Type 2 charstrings).
- `otf.docx` — DOCX со шрифтами Inter и Loma; сравнение с LibreOffice.
- `GlyphPathDump.cs` — выгрузка контуров всех глифов шрифта (TrueTypeFont.GlyphPath) для сравнения с fontTools.

Проверка встроенных подмножеств: `fontTools.cffLib.CFFFontSet` разбирает FontFile3, ROS = Adobe-Identity-0,
контуры всех CID совпадают с исходным шрифтом; контуры GlyphPath совпали с fontTools на всех глифах
Inter-Regular, InterDisplay-BlackItalic и Loma.

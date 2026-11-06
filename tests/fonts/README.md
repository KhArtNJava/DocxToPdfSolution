# Шрифты OpenType CFF (.otf)

- `otf-fonts.html` — Inter (обычный, жирный, курсив), Loma (тайский), Unifont (CJK, символы); сравнение с Chromium.
- `otf-3d-paths.html` — текст в 3D-преобразовании выводится контурами глифов (интерпретатор Type 2 charstrings).
- `otf.docx` — DOCX со шрифтами Inter и Loma; сравнение с LibreOffice.
- `GlyphPathDump.cs` — выгрузка контуров всех глифов шрифта (TrueTypeFont.GlyphPath) для сравнения с fontTools.

Проверка встроенных подмножеств: `fontTools.cffLib.CFFFontSet` разбирает FontFile3, ROS = Adobe-Identity-0,
контуры всех CID совпадают с исходным шрифтом; контуры GlyphPath совпали с fontTools на всех глифах
Inter-Regular, InterDisplay-BlackItalic и Loma.

## Вариативные шрифты (`variable/`)

- `VarTestTT.ttf` (glyf + gvar + HVAR + avar) и `VarTestCFF2.otf` (CFF2 + HVAR) — тестовые шрифты с осью wght 400…700
  (на основе Inter, OFL; мастер Bold — растянутый по горизонтали, чтобы разница была заметна).
- `variable-fonts.html` — font-weight 400/550/700, промежуточные 450/600/650, шорткат `font`; сравнение с Chromium.
- `variable-fonts.docx` — обычный/полужирный/курсив; сравнение с LibreOffice.
- `VariationCheck.cs` — сверка контуров и ширин всех глифов с экземплярами fontTools
  (`fonttools varLib.instancer VarTestTT.ttf wght=700 -o ref_VarTestTT_700.ttf`, то же для 400/550 и CFF2),
  а также вариант без HVAR (ширины по фантомным точкам gvar). Результат: все 297 глифов × 3 насыщенности × 3 шрифта
  совпали (расхождение ≤ 1 единицы — округление в fontTools); встроенные подмножества разобраны fontTools, контуры совпали.

# Кернинг и переносы

- `kerning-hyphens.html` — кернинг 40 pt (AVATAR, Wave, To., Yo, TAV) с `font-kerning: none` для сравнения,
  колонки по ширине с `hyphens: auto` (русский и английский через `lang`), `hyphens: manual` (по умолчанию), мягкие
  переносы `&shy;`, `hyphens: none`, очень узкая колонка. Сравнение с Chromium: ширины слов совпадают
  (кернинг), переносы русского текста — те же.
- `hyphens-kerning.docx` (+ `.odt`, `.rtf` — тот же документ, сохранённый LibreOffice) — `w:kern`, `w:autoHyphenation`,
  `w:hyphenationZone`, `w:suppressAutoHyphens` в третьей колонке, английский абзац (`w:lang`). Все три формата
  переносятся одинаково; положения слов строки с кернингом совпадают с LibreOffice с точностью до 0,05 pt.
- `hyphen-limit-1.docx` — `w:consecutiveHyphenLimit=1` (двух строк с переносом подряд нет);
  `hyphen-no-caps.docx` — `w:doNotHyphenateCaps`.
- `HyphenationCheck.cs` — переносы списка слов (`Hyphenator.For(lang).Points`); сверено с pyphen на тех же словарях:
  1278 русских слов — без расхождений (английский отличается только тем, что pyphen не учитывает RIGHTHYPHENMIN 3 словаря).

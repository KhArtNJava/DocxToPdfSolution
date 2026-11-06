# Веб-архивы MHTML

- `chromium-snapshot.mhtml` — архив Chromium (CDP `Page.captureSnapshot`): страница в windows-1251 (quoted-printable),
  внешняя таблица стилей, встроенный `<style>` как часть `cid:css-…@mhtml.blink`, JPEG/PNG (имя с пробелом, `%20`)/GIF,
  корневая ссылка `/img/…`, картинка из `::before { content: url(…) }`. Результат совпадает попиксельно с конвертацией
  той же страницы из исходных файлов и визуально — с печатью Chromium.
- `word-single-file.mht` — в стиле Word «Веб-страница в одном файле»: `file:///C:/…/test_files/…`, CSS внутри `<!-- -->`,
  VML-картинка в условном комментарии и запасной `<img>` в `<![if !vml]>`. Совпадает попиксельно с распакованной
  страницей (test.htm + test_files).
- `email-cid.mhtml` — письмо: multipart/related → multipart/alternative (text/plain + text/html в base64), картинки по
  `cid:` (в т.ч. `cid:%3C…%3E`), переводы строк LF, нет завершающей границы, отсутствующая картинка → alt-текст.
- `single-part.mht` — MHTML без multipart (одна часть text/html, мягкие переносы quoted-printable).
- `cell-margins.html` — нижнее поле последнего абзаца остаётся внутри ячейки таблицы (как в Chromium).
- `frames.mhtml` — архив Chromium со страницей с фреймами: `<iframe>` с размерами из атрибутов и рамкой по умолчанию
  (документ со своими стилями, картинкой из архива, absolute-элементом у низа фрейма и вложенным фреймом второго уровня,
  содержимое обрезается по фрейму), плавающий фрейм со `srcdoc` и рамкой из CSS, фрейм по центру без рамки
  (`frameborder="0"`, размер по умолчанию 300×150). Сравнение с печатью Chromium — `docs/compare/compare-frames.png`.
- `frameset.mhtml` — `<frameset cols="25%,*">` с вложенным `<frameset rows="120,*">`: меню, два документа, `<noframes>`.

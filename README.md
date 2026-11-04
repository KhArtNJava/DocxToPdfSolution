# docx2pdf — конвертер DOCX → PDF на чистом C# (.NET 10)

Без сторонних библиотек: используется только BCL (`System.IO.Compression`, `System.Xml.Linq`).
Весь код (чтение DOCX, разбор TrueType, вёрстка, генерация PDF) написан с нуля.

## Сборка и запуск

```bash
dotnet build -c Release
dotnet bin/Release/net10.0/docx2pdf.dll document.docx -o document.pdf
dotnet bin/Release/net10.0/docx2pdf.dll *.docx -o out_dir/ --fonts ~/my-fonts
```

## Устройство

| Файл | Назначение |
|---|---|
| `Docx/DocxReader.cs` | ZIP-пакет, связи (rels), тема, стили с наследованием, нумерация, настройки, разделы, колонтитулы, таблицы, рисунки |
| `Docx/Model.cs` | Модель документа (свойства в пунктах, разрешённые стили) |
| `Fonts/TrueTypeFont.cs` | Парсер TTF/TTC (cmap 4/12, hmtx, OS/2, post, name) и сабсеттинг glyf/loca |
| `Fonts/FontManager.cs` | Поиск системных шрифтов (Windows/Linux/macOS), замены (Calibri→Carlito и т.п.), запасные шрифты для отсутствующих символов, Symbol/Wingdings |
| `Images/ImageDecoder.cs` | JPEG (как есть, DCTDecode), PNG (все типы, Adam7, прозрачность), GIF (LZW), BMP |
| `Layout/BlockLayout.cs` | Разбиение на строки, табуляция, выравнивание, межстрочный интервал, таблицы |
| `Layout/Paginator.cs` | Разбивка на страницы, keepNext, висячие строки, разрыв строк таблиц, повтор шапки, колонтитулы |
| `Pdf/PdfDocument.cs` | Запись PDF 1.7: Type0/CIDFontType2 + Identity-H, ToUnicode, FlateDecode, изображения с SMask, ссылки |

## Поддерживается

- Шрифты: встраивание подмножеств TrueType, кириллица и любой Unicode, синтез жирного/курсива, если нет файла начертания
- Текст: жирный, курсив, подчёркивание (в т.ч. двойное), зачёркивание, цвет, выделение, верхний/нижний индекс, ПРОПИСНЫЕ, скрытый текст
- Абзацы: выравнивание (в т.ч. по ширине), отступы, выступ, интервалы до/после, межстрочный (auto/exact/atLeast), contextualSpacing, заливка, верхняя/нижняя граница, разрывы строк и страниц, pageBreakBefore, keepNext/keepLines, контроль висячих строк
- Табуляция: left/center/right/decimal, заполнители (точки, дефисы, подчёркивание) — оглавления
- Списки: маркированные и нумерованные, многоуровневые (`%1.%2.`), decimal/буквы/римские/русские буквы, startOverride
- Таблицы: сетка, gridSpan, vMerge, границы таблицы/ячеек/стиля, заливка, поля ячеек, vAlign, высота строк, вложенные таблицы, разрыв строки между страницами, повтор строк-заголовков
- Рисунки: inline и anchor (как inline), VML imagedata; масштабирование под ширину колонки
- Разделы: размер и ориентация страницы, поля, nextPage/continuous/evenPage/oddPage, нумерация с заданного номера
- Колонтитулы: default/first/even, поля PAGE и NUMPAGES
- Гиперссылки (кликабельные), метаданные (заголовок, автор)

## Ограничения

- Плавающие объекты (anchor) выводятся в потоке, обтекание не поддерживается; надписи (text box) и фигуры не рисуются
- EMF/WMF/TIFF выводятся серой заглушкой
- Нет колонок, сносок, примечаний, RTL и сложного шейпинга (арабский, индийские письменности), кернинга и автопереносов
- Только TrueType-контуры (OTF с CFF пропускаются)
- Условное форматирование стилей таблиц (firstRow, banded rows) не применяется

Пример: `src/DocxToPdf/samples/sample.docx` → `src/DocxToPdf/samples/sample.pdf`.

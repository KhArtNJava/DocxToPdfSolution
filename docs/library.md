# Библиотека DocxToPdf

`DocxToPdf.dll` — конвертер DOCX, ODT, RTF, HTML и MHTML в PDF для .NET 10. Это тот же код, что у программы
`docx2pdf`, собранный как библиотека. Сторонних зависимостей нет, только стандартная библиотека .NET.

## Состав архива

| Файл | Назначение |
|---|---|
| `DocxToPdf.dll` | Сборка библиотеки (AnyCPU, .NET 10) |
| `DocxToPdf.xml` | XML-документация: подсказки IntelliSense в Visual Studio / Rider / VS Code |
| `DocxToPdf.pdb` | Отладочные символы |
| `DocxToPdf.deps.json` | Описание зависимостей (только .NET) |
| `licenses/` | Лицензии встроенных словарей переносов (LibreOffice: LGPL/MPL/GPL и др.) |

## Требования

- .NET 10 (рантайм `Microsoft.NETCore.App` 10.0 или новее); Windows, Linux или macOS.
- Шрифты TrueType/OpenType в системе. Библиотека ищет их в системных каталогах:
  `C:\Windows\Fonts`, `/usr/share/fonts`, `~/.fonts`, `/Library/Fonts` и др.
  Дополнительные каталоги передаются в `FontManager`.
  Если нужных гарнитур нет, подбираются замены: Times New Roman → Liberation Serif и т.п.

## Подключение

Скопируйте `DocxToPdf.dll` и `DocxToPdf.xml` в каталог проекта, например в `lib/`, и добавьте ссылку в `.csproj`:

```xml
<ItemGroup>
  <Reference Include="DocxToPdf">
    <HintPath>lib/DocxToPdf.dll</HintPath>
  </Reference>
</ItemGroup>
```

## Быстрый старт

```csharp
using DocxToPdf;          // Converter
using DocxToPdf.Fonts;    // FontManager
using DocxToPdf.Docx;     // DocxOptions
using DocxToPdf.Html;     // HtmlOptions

// Поиск шрифтов — дорогая операция: один FontManager на всё приложение
var fonts = new FontManager();

int pages = Converter.Convert("report.docx", "report.pdf", fonts);
Console.WriteLine($"Готово: {pages} стр.");
```

Формат определяется по содержимому файла, а не по расширению:
- `{\rtf` — RTF;
- ZIP с `word/document.xml` — DOCX (`.docm`, `.dotx`);
- ZIP с `content.xml` — OpenDocument Text (`.odt`, `.ott`);
- XML `office:document` — плоский ODT (`.fodt`);
- MIME-заголовки — веб-архив MHTML (`.mht`, `.mhtml`);
- всё остальное — HTML.

## Настройки

### DOCX, ODT, RTF — `DocxOptions` (`DocxToPdf.Docx`)

```csharp
var docx = new DocxOptions
{
    Revisions = "final",   // auto | final | markup — исправления: как в документе, итоговый вид, с разметкой
    Comments  = "end",     // auto | balloons | end | none — примечания: выноски на поле, список в конце, не выводить
};
Converter.Convert("contract.docx", "contract.pdf", fonts, docxOptions: docx);
```

### HTML, MHTML — `HtmlOptions` (`DocxToPdf.Html`)

Размеры задаются в пунктах (1 pt = 1/72 дюйма). Если в документе есть правила CSS `@page`, они важнее этих настроек.

```csharp
var html = new HtmlOptions
{
    PageWidth = 595.3, PageHeight = 841.9,   // A4 (по умолчанию); Letter — 612 × 792
    MarginTop = 56.7, MarginBottom = 56.7,   // 20 мм
    MarginLeft = 56.7, MarginRight = 56.7,
    Landscape = false,
    PageNumbers = true,                      // «1 / 5» внизу страницы
    AllowRemote = false,                     // не загружать картинки и стили по http(s)
};
Converter.Convert("page.html", "page.pdf", fonts, htmlOptions: html);
```

### Шрифты — `FontManager` (`DocxToPdf.Fonts`)

```csharp
// Свои каталоги шрифтов (TTF/OTF/TTC) в дополнение к системным
var fonts = new FontManager(["/opt/app/fonts", @"D:\corporate-fonts"]);
```

Если не найдено ни одного шрифта, конструктор бросает `InvalidOperationException`. Это бывает, например,
в минимальном Docker-образе. Тогда установите пакет шрифтов (`fonts-liberation`, `fonts-dejavu`) или передайте
каталог со шрифтами.

## Пакетная конвертация и потоки

Один `FontManager` используется для всех файлов:

```csharp
var fonts = new FontManager();
foreach (var file in Directory.GetFiles("in", "*.docx"))
{
    string pdf = Path.Combine("out", Path.GetFileNameWithoutExtension(file) + ".pdf");
    try { Converter.Convert(file, pdf, fonts); }
    catch (Exception e) { Console.Error.WriteLine($"{file}: {e.Message}"); }
}
```

`Converter.Convert` можно вызывать из разных потоков, например из веб-сервера. Шрифты, словари переносов и
разобранные глифы общие для всех вызовов, поэтому внутри процесса конвертации идут строго по очереди.
Чтобы конвертировать параллельно, запустите несколько процессов, например программу `docx2pdf`.

## Ошибки

| Исключение | Когда |
|---|---|
| `NotSupportedException` | Документ OpenDocument другого типа: таблица `.ods`, презентация `.odp` и т.п. |
| `InvalidOperationException` | Нет шрифтов (конструктор `FontManager`) |
| `IOException`, `UnauthorizedAccessException` | Нет доступа к входному или выходному файлу |
| `InvalidDataException` | Повреждённый ZIP-пакет DOCX/ODT |

## Программа командной строки

В исходниках есть и программа `docx2pdf`. Она тоже реализована в этой сборке, в классе `DocxToPdf.Program`.
Её можно вызвать из кода:

```csharp
int exitCode = DocxToPdf.Program.Main(["report.docx", "-o", "report.pdf", "--comments", "none"]);
```

Подробное описание поддерживаемых возможностей и ограничений — в `README.md`, справочник API — в `api.md`.

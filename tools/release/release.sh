#!/usr/bin/env bash
# Сборка релиза: архив библиотеки и архив документации.
# Использование: tools/release/release.sh [версия] [каталог-результата]
set -euo pipefail
VER="${1:-1.0.0}"
ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
OUT="${2:-$ROOT/artifacts}"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
mkdir -p "$OUT"

# 1. Библиотека
dotnet build "$ROOT/src/DocxToPdf.Library/DocxToPdf.Library.csproj" -c Release -p:Version="$VER" -o "$STAGE/build" -nologo -v q
LIB="$STAGE/DocxToPdf-$VER"
mkdir -p "$LIB/licenses"
cp "$STAGE/build/DocxToPdf."{dll,xml,pdb,deps.json} "$LIB/"
cp "$ROOT/src/DocxToPdf/Hyphenation/licenses/"* "$LIB/licenses/"
cat > "$LIB/README.txt" <<TXT
DocxToPdf $VER — конвертер DOCX, ODT, RTF, HTML и MHTML в PDF для .NET 10 (без сторонних зависимостей).

DocxToPdf.dll        сборка библиотеки (AnyCPU, net10.0)
DocxToPdf.xml        XML-документация (IntelliSense)
DocxToPdf.pdb        отладочные символы
DocxToPdf.deps.json  описание зависимостей
licenses/            лицензии встроенных словарей переносов (LibreOffice)

Пример:
    var fonts = new DocxToPdf.Fonts.FontManager();
    int pages = DocxToPdf.Converter.Convert("in.docx", "out.pdf", fonts);

Полная документация — в архиве DocxToPdf-$VER-docs.zip (library.html, api.html, README.html).
TXT
(cd "$STAGE" && rm -f "$OUT/DocxToPdf-$VER-lib.zip" && zip -qr "$OUT/DocxToPdf-$VER-lib.zip" "DocxToPdf-$VER")

# 2. Документация
DOC="$STAGE/DocxToPdf-$VER-docs"
mkdir -p "$DOC/compare" "$DOC/tests"
dotnet run --project "$ROOT/tools/release/ApiDocGen/ApiDocGen.csproj" -c Release -- \
    "$STAGE/build/DocxToPdf.dll" "$STAGE/build/DocxToPdf.xml" "$DOC/api.md" >/dev/null
cp "$ROOT/docs/library.md" "$DOC/library.md"
cp "$ROOT/README.md" "$DOC/README.md"
cp "$ROOT/docs/compare/"*.png "$DOC/compare/"
for d in "$ROOT"/tests/*/; do
    [ -f "$d/README.md" ] && cp "$d/README.md" "$DOC/tests/$(basename "$d").md"
done
cp "$ROOT/tools/release/docs.css" "$DOC/docs.css"
NAV='<nav class="top"><a href="index.html">Содержание</a><a href="library.html">Руководство</a><a href="api.html">Справочник API</a><a href="README.html">Возможности и ограничения</a></nav>'
html() {  # html <md> <html> <заголовок> <css> <nav>
    LC_ALL=C.UTF-8 pandoc "$1" -f gfm -t html5 -s --metadata pagetitle="$3" --metadata lang=ru -c "$4" -B <(echo "$5") -o "$2"
}
html "$DOC/library.md" "$DOC/library.html" "DocxToPdf $VER — руководство" docs.css "$NAV"
html "$DOC/api.md" "$DOC/api.html" "DocxToPdf $VER — справочник API" docs.css "$NAV"
html "$DOC/README.md" "$DOC/README.html" "DocxToPdf $VER — возможности и ограничения" docs.css "$NAV"
TNAV="${NAV//href=\"/href=\"../}"
for f in "$DOC"/tests/*.md; do
    html "$f" "${f%.md}.html" "Тесты: $(basename "${f%.md}")" ../docs.css "$TNAV"
done
{
    echo "# DocxToPdf $VER"
    echo
    echo "Конвертер DOCX, ODT, RTF, HTML и MHTML в PDF для .NET 10 без сторонних зависимостей."
    echo
    echo "- [Руководство по библиотеке](library.html): подключение, примеры, настройки, ошибки"
    echo "- [Справочник API](api.html): классы и методы"
    echo "- [Возможности и ограничения](README.html): что поддерживается в каждом формате"
    echo "- Описание тестовых документов:"
    for f in "$DOC"/tests/*.md; do n="$(basename "${f%.md}")"; echo "    - [$n](tests/$n.html)"; done
    echo "- Сравнение с эталонами (Chromium, LibreOffice): каталог \`compare/\`"
    echo
    echo "Каждая страница есть и в формате Markdown (\`.md\`)."
} > "$DOC/index.md"
html "$DOC/index.md" "$DOC/index.html" "DocxToPdf $VER — документация" docs.css "$NAV"
(cd "$STAGE" && rm -f "$OUT/DocxToPdf-$VER-docs.zip" && zip -qr "$OUT/DocxToPdf-$VER-docs.zip" "DocxToPdf-$VER-docs")

echo "$OUT/DocxToPdf-$VER-lib.zip"
echo "$OUT/DocxToPdf-$VER-docs.zip"

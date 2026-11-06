# Тесты flex/grid

- `grid-flex.html` — фон и рамка flex-контейнера при `align-content` (space-between, center), `grid-auto-flow: column`
  (с `grid-auto-columns`), `row` и `row dense`, `subgrid` по строкам (карточки) и по колонкам, `masonry`.
  Эталон — печать в PDF из Chromium.
- `masonry-reference.html` — та же страница, где каменная кладка собрана вручную из flex-колонок
  (Chromium не поддерживает masonry): раскладка элементов по колонкам должна совпасть.
- `masonry-long.html` — длинная кладка на нескольких страницах (A5).

```bash
docx2pdf tests/grid/*.html -o /tmp
```

# Тесты flex/grid

- `grid-flex.html` — фон и рамка flex-контейнера при `align-content` (space-between, center), `grid-auto-flow: column`
  (с `grid-auto-columns`), `row` и `row dense`, `subgrid` по строкам (карточки) и по колонкам, `masonry`.
  Эталон — печать в PDF из Chromium.
- `masonry-reference.html` — та же страница, где каменная кладка собрана вручную из flex-колонок
  (Chromium не поддерживает masonry): раскладка элементов по колонкам должна совпасть.
- `subgrid-radius-gaps.html` — подсетки со скруглёнными углами (в т.ч. разные радиусы), собственным `row-gap`
  (2px при 12px у родителя, рядом обычная колонка) и собственным `column-gap: 0`. Эталон — Chromium.
- `masonry-rows-tracks.html` — кладка по строкам (`grid-template-columns: masonry`), `justify-tracks` по строкам,
  `align-tracks: stretch, end, center, space-evenly` при `height`; `masonry-rows-tracks-reference.html` — та же
  раскладка вручную на flexbox для сравнения в Chromium.
- `masonry-long.html` — длинная кладка на нескольких страницах (A5).

```bash
docx2pdf tests/grid/*.html -o /tmp
```

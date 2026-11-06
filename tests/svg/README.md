# Тесты SVG

- `svg-features.html` (+ `logo.svg`) — встроенные `<svg>`: фигуры, контуры и дуги, градиенты и прозрачность,
  преобразования, `<use>`/`<symbol>`, текст (`tspan`, `textPath`, `text-anchor`), `clipPath`, `<pattern>`, маркеры,
  `<style>` и классы, CSS документа, `preserveAspectRatio`; `<img src="logo.svg">`, `data:image/svg+xml`, SVG во всю ширину.
  Эталон — печать в PDF из Chromium.
- `svg-filters-masks.html` — маски (по яркости с градиентом и узором, с текстом и картинкой, `mask-type: alpha` с
  `maskContentUnits="objectBoundingBox"`) и фильтры: размытие, тень из примитивов и `feDropShadow`, `feColorMatrix`,
  `feComponentTransfer`, `feMorphology` + `feComposite`, `feTurbulence`, `feDisplacementMap`, освещение,
  `feBlend` + `feTile` + `feImage`, CSS-функции фильтра, фильтр с маской и `opacity`, градиент с `stop-opacity`.
  Сравнение с печатью Chromium — `docs/compare/compare-svg-filters.png`.
- `svg-blip.docx` — картинка DOCX с векторным оригиналом `asvg:svgBlip` (растровая копия — серый квадрат;
  в PDF должен быть логотип).

```bash
docx2pdf tests/svg/svg-features.html tests/svg/svg-blip.docx -o /tmp
```

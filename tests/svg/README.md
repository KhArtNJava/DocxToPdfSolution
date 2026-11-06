# Тесты SVG

- `svg-features.html` (+ `logo.svg`) — встроенные `<svg>`: фигуры, контуры и дуги, градиенты и прозрачность,
  преобразования, `<use>`/`<symbol>`, текст (`tspan`, `textPath`, `text-anchor`), `clipPath`, `<pattern>`, маркеры,
  `<style>` и классы, CSS документа, `preserveAspectRatio`; `<img src="logo.svg">`, `data:image/svg+xml`, SVG во всю ширину.
  Эталон — печать в PDF из Chromium.
- `svg-blip.docx` — картинка DOCX с векторным оригиналом `asvg:svgBlip` (растровая копия — серый квадрат;
  в PDF должен быть логотип).

```bash
docx2pdf tests/svg/svg-features.html tests/svg/svg-blip.docx -o /tmp
```

# Словари переносов

Файлы `hyph_*.dic.gz` — словари шаблонов переносов (алгоритм Ф. Лианга, формат библиотеки Hyphen) из репозитория
словарей LibreOffice (https://git.libreoffice.org/dictionaries), сжатые gzip без изменений содержимого.
Это данные, а не код: программа читает их своим собственным разборщиком (`Layout/Hyphenation.cs`).

Лицензии словарей различны (BSD-подобная для английского, LGPL, GPL, MPL, LPPL, CC BY-SA для отдельных языков) —
см. файлы `README_hyph_*.txt` в этой папке и тексты лицензий `COPYING.GPL`, `COPYING.LGPL`, `COPYING.MPL`.
Словарь русского языка (`hyph_ru_RU.dic`) основан на шаблонах ruhyphal (TeX), распространяется по LGPL
в составе словарей LibreOffice.

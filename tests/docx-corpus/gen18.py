from docx import Document
from docx.shared import Pt

# t26: цветные эмодзи (CBDT/NotoColorEmoji): ZWJ-последовательности (семья), флаги (региональные
# индикаторы), вариативный селектор VS16 (✈️), кейкапы, тона кожи, одиночные эмодзи в тексте.
# Проверяется склейка ZWJ-кластеров в один глиф и вывод цветного растра (Type 3, как LibreOffice).

d = Document()
st = d.styles['Normal']; st.font.name = 'Liberation Serif'; st.font.size = Pt(11)

d.add_paragraph('A \U0001F468‍\U0001F469‍\U0001F467 B')          # ZWJ-семья: мужчина+ZWJ+женщина+ZWJ+девочка
d.add_paragraph('C \U0001F468\U0001F469\U0001F467 D')            # те же без ZWJ — три отдельных глифа
d.add_paragraph('E\U0001F600F')                                  # одиночный эмодзи вплотную к тексту
d.add_paragraph('G \U0001F1F7\U0001F1FA H')                      # флаг России (пара RI)
d.add_paragraph('I \u2708\uFE0F J')                              # самолёт + VS16 (текстовый → эмодзи)
d.add_paragraph('K 1\uFE0F\u20E3 L')                             # кейкап 1 (цифра + VS16 + combining enclosing)
d.add_paragraph('M \U0001F44D\U0001F3FD N')                      # thumbs up со средним тоном кожи
d.add_paragraph('\U0001F469‍\U0001F469‍\U0001F466 конец')         # ZWJ-семея в начале строки и русский текст
d.save('t26_emoji.docx')

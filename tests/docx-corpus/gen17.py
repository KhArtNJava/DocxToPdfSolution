from docx import Document
from docx.shared import Pt

# t25: цепочка keepNext+keepLines длиннее страницы (регрессия: раньше дробилась по 2–4 строки
# на почти пустые страницы — 14 абзацев давали 72 страницы; цепочка выше страницы должна
# разрезаться как обычный поток, как в Word/LibreOffice)

d = Document()
st = d.styles['Normal']; st.font.name = 'Liberation Serif'; st.font.size = Pt(11)

# короткая цепочка: заголовок не отрывается от следующего абзаца
h = d.add_paragraph('Заголовок с keepNext')
h.paragraph_format.keep_with_next = True
d.add_paragraph('Абзац после заголовка. ' + 'текст ' * 30)

# длинная цепочка: 14 абзацев по ~10 строк, суммарно ~3 страницы
for i in range(1, 15):
    p = d.add_paragraph(f'keepNext {i}: ' + 'связанный абзац ' * 40)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.keep_together = True
d.save('t25_keepchain.docx')

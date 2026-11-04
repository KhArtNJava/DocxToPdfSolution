import struct, zlib
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# PNG с альфой
def png(w,h):
    raw=b''
    for y in range(h):
        raw+=b'\x00'+b''.join(bytes([int(255*x/w),int(255*y/h),160,255 if (x-w/2)**2+(y-h/2)**2<(w/2)**2 else 0]) for x in range(w))
    def chunk(t,d): return struct.pack('>I',len(d))+t+d+struct.pack('>I',zlib.crc32(t+d)&0xffffffff)
    return b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',w,h,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(raw))+chunk(b'IEND',b'')
open('img.png','wb').write(png(120,80))

d=Document()
def field(par, instr):
    r=par.add_run()
    for t,txt in (('begin',None),(None,instr),('separate',None),(None,'1'),('end',None)):
        if t:
            fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),t); r._r.append(fc)
        elif txt==instr:
            it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text=' '+instr+' '; r._r.append(it)
        else:
            tt=OxmlElement('w:t'); tt.text=txt; r._r.append(tt)
sec=d.sections[0]
hp=sec.header.paragraphs[0]; hp.text='Колонтитул: тестовый документ'; hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
fp=sec.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
fp.add_run('Страница '); field(fp,'PAGE'); fp.add_run(' из '); field(fp,'NUMPAGES')

d.add_heading('Конвертер DOCX → PDF', 0)
d.add_heading('1. Введение', 1)
p=d.add_paragraph('Это ')
p.add_run('жирный').bold=True; p.add_run(', ')
p.add_run('курсив').italic=True; p.add_run(', ')
r=p.add_run('подчёркнутый'); r.underline=True; p.add_run(', ')
r=p.add_run('зачёркнутый'); r.font.strike=True; p.add_run(' и ')
r=p.add_run('красный'); r.font.color.rgb=RGBColor(200,0,0); p.add_run(' текст. Формула: H')
r=p.add_run('2'); r.font.subscript=True; p.add_run('O, E = mc'); r=p.add_run('2'); r.font.superscript=True
p.add_run('. Mixed English text with Ελληνικά and symbols: ≤ ≥ ∑ € ™ ✓.')
lorem=('Съешь же ещё этих мягких французских булок, да выпей чаю. '*6).strip()
p=d.add_paragraph(lorem); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
p.paragraph_format.first_line_indent=Cm(1.25)
p=d.add_paragraph(lorem); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
r=p.runs[0]; r.font.name='Times New Roman'; r.font.size=Pt(12)
p=d.add_paragraph('Выделение маркером, ');
r=p.add_run('жёлтый фон'); r.font.highlight_color=7
d.add_heading('2. Списки', 1)
for t in ['Первый пункт маркированного списка','Второй пункт, достаточно длинный для того чтобы перенестись на следующую строку и проверить выступ','Третий']:
    d.add_paragraph(t, style='List Bullet')
for t in ['Первый','Второй','Третий']:
    d.add_paragraph(t, style='List Number')
d.add_paragraph('Вложенный', style='List Bullet 2')
d.add_heading('3. Таблица', 1)
t=d.add_table(rows=4, cols=3); t.style='Table Grid'
hdr=t.rows[0].cells
for i,s in enumerate(['Название','Количество','Цена, ₽']):
    hdr[i].text=s; hdr[i].paragraphs[0].runs[0].bold=True
data=[('Яблоки','10','120,50'),('Груши — очень длинное название товара, которое переносится','5','99,00'),('Итого','','219,50')]
for r,row in enumerate(data):
    for c,v in enumerate(row): t.rows[r+1].cells[c].text=v
a=t.cell(3,0).merge(t.cell(3,1))
t2=d.add_table(rows=3, cols=2); t2.style='Table Grid'
m=t2.cell(0,0).merge(t2.cell(2,0)); m.text='Объединено по вертикали'
for i in range(3): t2.cell(i,1).text=f'Строка {i+1}'
d.add_heading('4. Изображение', 1)
d.add_paragraph().add_run().add_picture('img.png', width=Cm(5))
d.add_paragraph('Ссылка: ')
d.add_heading('5. Оглавление с табуляцией', 1)
for name,pg in [('Введение','3'),('Основная часть','7'),('Заключение','15')]:
    p=d.add_paragraph(name+'\t'+pg)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
p=d.add_paragraph('Перед разрывом страницы.')
p.add_run().add_break(WD_BREAK.PAGE)
d.add_heading('6. Новая страница', 1)
for i in range(40):
    d.add_paragraph(f'Абзац {i+1}. '+lorem[:150])
d.save('test.docx')

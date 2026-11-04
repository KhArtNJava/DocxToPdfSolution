import copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK, WD_UNDERLINE, WD_COLOR_INDEX
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def rpr(run, xml):
    rPr=run._r.get_or_add_rPr(); rPr.append(parse_xml(f'<{xml.split()[0]} {W} {" ".join(xml.split()[1:])}/>' if not xml.startswith('<') else xml))
def raw(par, xml):
    par._p.append(parse_xml(xml))
LOREM='Съешь же ещё этих мягких французских булок, да выпей чаю. The quick brown fox jumps over the lazy dog. '

# ---------------- 1. Текст ----------------
d=Document()
d.add_heading('1. Форматирование символов',1)
p=d.add_paragraph()
for txt,fn in [('жирный',lambda r:setattr(r,'bold',True)),('курсив',lambda r:setattr(r,'italic',True)),
  ('подчёркнутый',lambda r:setattr(r,'underline',True)),('двойное',lambda r:setattr(r,'underline',WD_UNDERLINE.DOUBLE)),
  ('волнистое',lambda r:setattr(r,'underline',WD_UNDERLINE.WAVY)),('пунктир',lambda r:setattr(r,'underline',WD_UNDERLINE.DOTTED)),
  ('зачёркнутый',lambda r:setattr(r.font,'strike',True)),('двойное зачёркивание',lambda r:setattr(r.font,'double_strike',True)),
  ('ПРОПИСНЫЕ из строчных',lambda r:setattr(r.font,'all_caps',True)),('Капитель Small Caps',lambda r:setattr(r.font,'small_caps',True)),
  ('верхний',lambda r:setattr(r.font,'superscript',True)),('нижний',lambda r:setattr(r.font,'subscript',True)),
  ('красный',lambda r:setattr(r.font.color,'rgb',RGBColor(0xC0,0,0))),('выделение',lambda r:setattr(r.font,'highlight_color',WD_COLOR_INDEX.YELLOW)),
  ('скрытый (не должен быть виден)',lambda r:setattr(r.font,'hidden',True)),('14pt',lambda r:setattr(r.font,'size',Pt(14))),
  ('Courier New',lambda r:setattr(r.font,'name','Courier New')),('Arial',lambda r:setattr(r.font,'name','Arial'))]:
    r=p.add_run(txt); fn(r); p.add_run(' · ')
r=p.add_run('разреженный 3pt'); rpr(r,'w:spacing w:val="60"'); p.add_run(' · ')
r=p.add_run('уплотнённый'); rpr(r,'w:spacing w:val="-20"'); p.add_run(' · ')
r=p.add_run('заливка'); rpr(r,'w:shd w:val="clear" w:color="auto" w:fill="CCE5FF"'); p.add_run(' · ')
r=p.add_run('приподнятый'); rpr(r,'w:position w:val="8"'); p.add_run(' · ')
r=p.add_run('масштаб 200%'); rpr(r,'w:w w:val="200"'); p.add_run(' · ')
r=p.add_run('рамка вокруг'); rpr(r,'w:bdr w:val="single" w:sz="4" w:space="0" w:color="FF0000"'); p.add_run(' · ')
r=p.add_run('контур'); rpr(r,'w:outline'); p.add_run(' · ')
r=p.add_run('эмодзи 😀 и символы ✓ ★ € ∑ ≠ → Ω')
d.add_heading('2. Абзацы',1)
for al in [WD_ALIGN_PARAGRAPH.LEFT,WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_PARAGRAPH.RIGHT,WD_ALIGN_PARAGRAPH.JUSTIFY,WD_ALIGN_PARAGRAPH.DISTRIBUTE]:
    p=d.add_paragraph(f'[{al}] '+LOREM*2); p.alignment=al
p=d.add_paragraph('Отступы: слева 2 см, справа 3 см, первая строка 1.25 см. '+LOREM*2); f=p.paragraph_format
f.left_indent=Cm(2); f.right_indent=Cm(3); f.first_line_indent=Cm(1.25)
p=d.add_paragraph('Выступ 1.5 см. '+LOREM*2); p.paragraph_format.left_indent=Cm(1.5); p.paragraph_format.first_line_indent=Cm(-1.5)
p=d.add_paragraph('Отрицательный отступ слева -1 см (в поле). '+LOREM); p.paragraph_format.left_indent=Cm(-1)
for rule,val,name in [(WD_LINE_SPACING.SINGLE,None,'одинарный'),(WD_LINE_SPACING.ONE_POINT_FIVE,None,'полуторный'),(WD_LINE_SPACING.DOUBLE,None,'двойной'),
                      (WD_LINE_SPACING.EXACTLY,Pt(10),'точно 10pt'),(WD_LINE_SPACING.AT_LEAST,Pt(24),'минимум 24pt'),(WD_LINE_SPACING.MULTIPLE,0.8,'множитель 0.8')]:
    p=d.add_paragraph(f'Интервал {name}. '+LOREM*2); p.paragraph_format.line_spacing_rule=rule
    if val is not None: p.paragraph_format.line_spacing=val
p=d.add_paragraph('Абзац с рамкой и заливкой. '+LOREM)
p._p.get_or_add_pPr().append(parse_xml(f'<w:pBdr {W}><w:top w:val="single" w:sz="12" w:space="4" w:color="2F5496"/><w:left w:val="single" w:sz="12" w:space="4" w:color="2F5496"/><w:bottom w:val="single" w:sz="12" w:space="4" w:color="2F5496"/><w:right w:val="single" w:sz="12" w:space="4" w:color="2F5496"/></w:pBdr>'))
p._p.get_or_add_pPr().append(parse_xml(f'<w:shd {W} w:val="clear" w:color="auto" w:fill="DEEAF6"/>'))
p2=d.add_paragraph('Второй абзац с той же рамкой — рамки должны объединиться в одну. '+LOREM)
p2._p.get_or_add_pPr().append(copy.deepcopy(p._p.pPr.find(qn('w:pBdr')))); p2._p.get_or_add_pPr().append(copy.deepcopy(p._p.pPr.find(qn('w:shd'))))
p=d.add_paragraph('Табуляция:\tпо умолчанию\tещё'); 
p=d.add_paragraph('Слева\tЦентр\tСправа\t1234,56'); ts=p.paragraph_format.tab_stops
from docx.enum.text import WD_TAB_ALIGNMENT as TA, WD_TAB_LEADER as TL
ts.add_tab_stop(Cm(8),TA.CENTER); ts.add_tab_stop(Cm(13),TA.RIGHT,TL.DOTS); ts.add_tab_stop(Cm(16),TA.DECIMAL)
p=d.add_paragraph('Мягкий перенос: сверх­dлинное­слово и неразрывный пробел, а также дефис‑неразрывный.')
p=d.add_paragraph(); r=p.add_run('Строка 1'); r.add_break(); p.add_run('Строка 2 после разрыва строки'); r=p.add_run(); r.add_break(WD_BREAK.LINE)
d.add_paragraph('Длинное_слово_без_пробелов_'+'x'*150)
d.add_paragraph('中文文本测试，日本語のテキスト、한국어 텍스트. Ελληνικά. עברית (RTL). العربية')
d.save('t1_text.docx')

# ---------------- 2. Списки ----------------
d=Document()
num=d.part.numbering_part.element
num.append(parse_xml(f'''<w:abstractNum {W} w:abstractNumId="90"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="792" w:hanging="432"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2.%3."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1224" w:hanging="504"/></w:pPr></w:lvl>
<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%4)"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1728" w:hanging="360"/></w:pPr></w:lvl>
<w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%5."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="2232" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum>'''))
num.append(parse_xml(f'''<w:abstractNum {W} w:abstractNumId="91"><w:multiLevelType w:val="hybridMultilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl>
<w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="2880" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl>
</w:abstractNum>'''))
num.append(parse_xml(f'<w:num {W} w:numId="90"><w:abstractNumId w:val="90"/></w:num>'))
num.append(parse_xml(f'<w:num {W} w:numId="91"><w:abstractNumId w:val="91"/></w:num>'))
num.append(parse_xml(f'<w:num {W} w:numId="92"><w:abstractNumId w:val="90"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'))
def li(text, numid, lvl):
    p=d.add_paragraph(text); p._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {W}><w:ilvl w:val="{lvl}"/><w:numId w:val="{numid}"/></w:numPr>')); return p
d.add_heading('Многоуровневая нумерация',1)
for t,l in [('Первый',0),('Подпункт',1),('Подподпункт',2),('Буквенный',3),('Римский справа',4),('Римский 2',4),('Буква 2',3),('Подпункт 2',1),('Второй',0),('Подпункт после возврата',1),('Третий пункт с очень длинным текстом, который переносится на следующую строку, чтобы проверить выступ и выравнивание текста относительно номера',0)]:
    li(t,90,l)
d.add_paragraph('Обычный абзац между списками.')
li('Продолжение того же списка (должно быть 4.)',90,0)
li('Новый список с startOverride (должно быть 1.)',92,0)
li('и 2.',92,0)
d.add_heading('Маркеры Symbol/Wingdings/Courier',1)
for t,l in [('Точка Symbol',0),('Кружок Courier o',1),('Квадрат Wingdings',2),('Стрелка Wingdings',3),('Снова точка',0)]: li(t,91,l)
d.add_heading('Встроенные стили',1)
for s in ['List Bullet','List Bullet 2','List Number','List Number 2','List Number','List Bullet 3']:
    d.add_paragraph(s, style=s)
d.save('t2_lists.docx')

# ---------------- 3. Таблицы ----------------
d=Document()
d.add_heading('Таблицы',1)
t=d.add_table(rows=4,cols=4); t.style='Light Grid Accent 1'
for i in range(4):
    for j in range(4): t.cell(i,j).text=f'R{i}C{j}'
d.add_paragraph('Таблица со стилем Light Grid Accent 1 (условное форматирование заголовка/полос).')
t=d.add_table(rows=3,cols=3); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
t.cell(0,0).merge(t.cell(0,2)).text='Объединено по всей ширине, по центру'; t.cell(0,0).paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
a=t.cell(1,0).merge(t.cell(2,0)); a.text='Вертикально, по центру'; a.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
t.cell(1,1).text='Заливка'; t.cell(1,1)._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {W} w:val="clear" w:color="auto" w:fill="FFE699"/>'))
t.cell(1,2).text='Снизу'; t.cell(1,2).vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.BOTTOM; t.rows[1].height=Cm(2)
t.cell(2,1).text='Вертикальный текст'; t.cell(2,1)._tc.get_or_add_tcPr().append(parse_xml(f'<w:textDirection {W} w:val="btLr"/>')); t.rows[2].height=Cm(3)
inner=t.cell(2,2).add_table(rows=2,cols=2); inner.style='Table Grid'
for i in range(2):
    for j in range(2): inner.cell(i,j).text=f'вл{i}{j}'
d.add_paragraph()
# ширина в процентах, без tblGrid ширин
t=d.add_table(rows=2,cols=3); t.style='Table Grid'
tblPr=t._tbl.tblPr; tblPr.append(parse_xml(f'<w:tblW {W} w:w="5000" w:type="pct"/>'))
for gc in t._tbl.tblGrid.findall(qn('w:gridCol')): gc.set(qn('w:w'),'0')
for i in range(2):
    for j in range(3): t.cell(i,j).text=('ширина 100% pct ' if j==0 else 'ячейка ')+str(i)+str(j)
d.add_paragraph()
# разные рамки
t=d.add_table(rows=2,cols=2)
t._tbl.tblPr.append(parse_xml(f'<w:tblBorders {W}><w:top w:val="double" w:sz="6" w:color="FF0000"/><w:bottom w:val="thick" w:sz="18" w:color="00AA00"/><w:left w:val="dashed" w:sz="8" w:color="0000FF"/><w:right w:val="dotted" w:sz="8" w:color="0000FF"/><w:insideH w:val="single" w:sz="4" w:color="999999"/><w:insideV w:val="none"/></w:tblBorders>'))
for i in range(2):
    for j in range(2): t.cell(i,j).text=f'рамки {i}{j}'
d.add_paragraph('Длинная таблица с повтором шапки и строкой, которая не должна разрываться:')
t=d.add_table(rows=1,cols=3); t.style='Table Grid'
for j,h in enumerate(['№','Описание','Сумма']): t.cell(0,j).text=h
t.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {W}/>'))
for i in range(45):
    r=t.add_row().cells; r[0].text=str(i+1); r[1].text=(LOREM if i%7==3 else 'Позиция '+str(i+1)); r[2].text=f'{i*13.7:.2f}'
    if i==20:
        t.rows[-1]._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {W}/>')); r[1].text=LOREM*6
d.save('t3_tables.docx')

# ---------------- 4. Разделы, колонтитулы ----------------
d=Document()
s=d.sections[0]; s.different_first_page_header_footer=True
s.first_page_header.paragraphs[0].text='ВЕРХНИЙ КОЛОНТИТУЛ ПЕРВОЙ СТРАНИЦЫ'
s.header.paragraphs[0].text='Обычный верхний колонтитул'
fp=s.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
def field(par, instr, cached='1'):
    r=par.add_run(); r._r.append(parse_xml(f'<w:fldChar {W} w:fldCharType="begin"/>'))
    r=par.add_run(); r._r.append(parse_xml(f'<w:instrText {W} xml:space="preserve"> {instr} </w:instrText>'))
    r=par.add_run(); r._r.append(parse_xml(f'<w:fldChar {W} w:fldCharType="separate"/>'))
    par.add_run(cached)
    r=par.add_run(); r._r.append(parse_xml(f'<w:fldChar {W} w:fldCharType="end"/>'))
fp.add_run('Стр. '); field(fp,'PAGE'); fp.add_run(' из '); field(fp,'NUMPAGES')
d.settings.element.append(parse_xml(f'<w:evenAndOddHeaders {W}/>'))
s.even_page_header.paragraphs[0].text='ЧЁТНАЯ СТРАНИЦА (колонтитул)'
d.add_heading('Раздел 1 (книжный, особый первый колонтитул, чёт/нечет)',1)
for i in range(12): d.add_paragraph(f'{i}. '+LOREM*3)
s2=d.add_section(WD_SECTION.NEW_PAGE); s2.orientation=WD_ORIENT.LANDSCAPE; s2.page_width,s2.page_height=s2.page_height,s2.page_width
s2.header.is_linked_to_previous=False; s2.header.paragraphs[0].text='Колонтитул альбомного раздела'
pg=s2._sectPr; pg.append(parse_xml(f'<w:pgNumType {W} w:start="100"/>'))
d.add_heading('Раздел 2 (альбомный, нумерация с 100)',1)
d.add_paragraph('Поле PAGE в тексте документа: '); field(d.paragraphs[-1],'PAGE','?')
for i in range(3): d.add_paragraph(LOREM*3)
s3=d.add_section(WD_SECTION.CONTINUOUS)
s3._sectPr.append(parse_xml(f'<w:cols {W} w:num="2" w:space="720"/>'))
d.add_heading('Раздел 3 — две колонки (continuous)',2)
for i in range(6): d.add_paragraph(f'Колонка-абзац {i}. '+LOREM*2)
s4=d.add_section(WD_SECTION.NEW_PAGE); s4.orientation=WD_ORIENT.PORTRAIT; s4.page_width,s4.page_height=s4.page_height,s4.page_width
s4.left_margin=Cm(1); s4.right_margin=Cm(1); s4.top_margin=Cm(1)
d.add_heading('Раздел 4 — узкие поля',1); d.add_paragraph(LOREM*5)
d.save('t4_sections.docx')

# ---------------- 5. Объекты ----------------
d=Document()
d.add_heading('Объекты',1)
d.add_paragraph('Картинка в строке: ').add_run().add_picture('img.png',width=Cm(2))
# плавающая картинка с обтеканием square справа
p=d.add_paragraph(LOREM*4)
r=p.add_run(); r.add_picture('img.png',width=Cm(4))
inline=r._r.find('.//'+qn('wp:inline'))
anchor_xml=f'''<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>{inline.find(qn('wp:extent')).xml if False else ''}</wp:anchor>'''
anc=parse_xml(anchor_xml)
for ch in list(inline):
    if ch.tag==qn('wp:docPr'): anc.append(parse_xml('<wp:wrapSquare xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" wrapText="bothSides"/>'))
    anc.append(ch)
inline.getparent().replace(inline,anc)
p.add_run(LOREM*3)
# Надпись (текстовое поле) через VML/wps — упрощённо в mc:AlternateContent Fallback v:textbox
p=d.add_paragraph('Перед надписью. ')
raw(p, f'''<w:r {W} xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape id="tb1" type="#_x0000_t202" style="width:200pt;height:50pt" fillcolor="#FFF2CC" strokecolor="#BF9000"><v:textbox><w:txbxContent><w:p><w:r><w:t>Текст внутри надписи (textbox)</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>''')
p.add_run(' После надписи.')
# Сноски
fn_part=None
p=d.add_paragraph('Текст со сноской')
raw(p, f'<w:r {W}><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>')
p.add_run(' и продолжение.')
# Гиперссылка
p=d.add_paragraph('Ссылка: ')
rid=d.part.relate_to('https://example.com', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
raw(p, f'<w:hyperlink {W} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{rid}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>example.com</w:t></w:r></w:hyperlink>')
# Исправления
p=d.add_paragraph('Исправления: ')
raw(p, f'<w:ins {W} w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>вставлено </w:t></w:r></w:ins>')
raw(p, f'<w:del {W} w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>УДАЛЕНО-НЕ-ПОКАЗЫВАТЬ </w:delText></w:r></w:del>')
p.add_run('конец.')
# Формула OMML
p=d.add_paragraph('Формула: ')
raw(p, '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:f><m:num><m:r><m:t>a+b</m:t></m:r></m:num><m:den><m:r><m:t>c</m:t></m:r></m:den></m:f><m:r><m:t>=</m:t></m:r><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad></m:oMath>')
# Поля: DATE, REF, оглавление
p=d.add_paragraph('Комментарий'); raw(p,f'<w:commentRangeStart {W} w:id="0"/>'); 
# Буквица (framePr dropCap)
p=d.add_paragraph('Б'); p._p.get_or_add_pPr().append(parse_xml(f'<w:framePr {W} w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>'))
p.runs[0].font.size=Pt(40)
d.add_paragraph('уквица (w:framePr dropCap) — следующий абзац должен обтекать букву. '+LOREM*3)
# Колонки текста в одном разделе нет; закладка
d.add_paragraph('Последний абзац.')
d.save('t5_objects.docx')

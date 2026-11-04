from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import parse_xml
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def fld(p, instr, cached):
    for x in [f'<w:r {W}><w:fldChar w:fldCharType="begin"/></w:r>', f'<w:r {W}><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>',
              f'<w:r {W}><w:fldChar w:fldCharType="separate"/></w:r>', f'<w:r {W}><w:t>{cached}</w:t></w:r>', f'<w:r {W}><w:fldChar w:fldCharType="end"/></w:r>']:
        p._p.append(parse_xml(x))
d=Document()
s=d.sections[0]
s._sectPr.append(parse_xml(f'<w:pgNumType {W} w:fmt="upperRoman" w:start="4"/>'))
fp=s.footer.paragraphs[0]; fp.add_run('Страница '); fld(fp,'PAGE','1')
d.add_heading('A. Поля и ссылки',1)
p=d.add_paragraph('Перекрёстная ссылка на страницу: '); fld(p,'PAGEREF _Ref1 \\h','7')
p=d.add_paragraph('Дата (закэшированная): '); fld(p,'DATE \\@ "dd.MM.yyyy"','01.02.2026')
p=d.add_paragraph('Рисунок '); fld(p,'SEQ Рисунок \\* ARABIC','1'); p.add_run(' — подпись.')
p=d.add_paragraph('Символ w:sym: ')
p._p.append(parse_xml(f'<w:r {W}><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>'))
p._p.append(parse_xml(f'<w:r {W}><w:t xml:space="preserve"> и Symbol: </w:t></w:r>'))
p._p.append(parse_xml(f'<w:r {W}><w:sym w:font="Symbol" w:char="F061"/></w:r>'))
d.add_heading('B. Таблица с отступами ячеек и gridBefore',1)
t=d.add_table(rows=3, cols=3); t.style='Table Grid'
t._tbl.tblPr.append(parse_xml(f'<w:tblCellMar {W}><w:top w:w="144" w:type="dxa"/><w:left w:w="288" w:type="dxa"/><w:bottom w:w="144" w:type="dxa"/><w:right w:w="288" w:type="dxa"/></w:tblCellMar>'))
for i in range(3):
    for j in range(3): t.cell(i,j).text=f'ячейка {i}{j}'
c=t.cell(1,1); c._tc.get_or_add_tcPr().append(parse_xml(f'<w:tcMar {W}><w:left w:w="0" w:type="dxa"/></w:tcMar>')); c.text='без левого поля'
d.add_paragraph()
d.add_heading('C. Нумерованные заголовки через стиль',1)
num=d.part.numbering_part.element
num.append(parse_xml(f'''<w:abstractNum {W} w:abstractNumId="77"><w:multiLevelType w:val="multilevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading3"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr></w:lvl></w:abstractNum>'''))
num.append(parse_xml(f'<w:num {W} w:numId="77"><w:abstractNumId w:val="77"/></w:num>'))
st=d.styles['Heading 2'].element; st.get_or_add_pPr().append(parse_xml(f'<w:numPr {W}><w:numId w:val="77"/></w:numPr>'))
st=d.styles['Heading 3'].element; st.get_or_add_pPr().append(parse_xml(f'<w:numPr {W}><w:ilvl w:val="1"/><w:numId w:val="77"/></w:numPr>'))
d.add_heading('Глава',2); d.add_heading('Раздел',3); d.add_heading('Ещё раздел',3); d.add_heading('Вторая глава',2); d.add_heading('Её раздел',3)
d.add_heading('D. Табуляция с выравниванием по разделителю и отступ табуляцией',1)
from docx.enum.text import WD_TAB_ALIGNMENT as TA
for v in ['1,5','123,25','10000,125']:
    p=d.add_paragraph('Сумма:\t'+v); p.paragraph_format.tab_stops.add_tab_stop(Cm(8),TA.DECIMAL)
d.add_heading('E. Межсимвольный кернинг и регистр',1)
p=d.add_paragraph(); r=p.add_run('Скрытый текст '); r.font.hidden=True; p.add_run('видимый текст; ')
r=p.add_run('Emboss'); r._r.get_or_add_rPr().append(parse_xml(f'<w:emboss {W}/>')); p.add_run(' ')
r=p.add_run('ТЕНЬ'); r._r.get_or_add_rPr().append(parse_xml(f'<w:shadow {W}/>'))
d.save('t10.docx')

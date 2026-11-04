from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
d.sections[0].header.paragraphs[0].add_run('Лого: ').add_picture('img.png',width=Cm(1))
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
def fld(instr, res): return f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{res}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
P(R('Оглавление (поля HYPERLINK \\l и PAGEREF):'))
P(fld('HYPERLINK \\l "_Toc1"', 'Глава первая')+'<w:r><w:tab/></w:r>'+fld('PAGEREF _Toc1 \\h','99'),'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8640"/></w:tabs>')
P(R('Распределённое выравнивание jc=distribute:'))
P(R('Ра з д е л'),'<w:jc w:val="distribute"/>')
P(R('Нумерованный список прерывается таблицей и продолжается:'))
P(R('первый'),'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
raw(f'<w:tbl {NS}><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Сноска в таблице</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p></w:tc></w:tr></w:tbl>')
P(R('второй (ожидается номер 2)'),'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>')
for i in range(3): P(R('Заполнитель'))
P('<w:bookmarkStart w:id="5" w:name="_Toc1"/>'+R('Глава первая (цель ссылки) — на стр. 2')+'<w:bookmarkEnd w:id="5"/>','<w:pageBreakBefore/>')
P(R('Ссылка REF на закладку: ')+fld('REF _Toc1 \\h','Глава первая (цель ссылки)'))
d.save('t17.docx')

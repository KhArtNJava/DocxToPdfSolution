from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
# 1 Нумерация строк
sp=b.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
sp.append(parse_xml(f'<w:lnNumType {NS} w:countBy="1" w:distance="283" w:restart="continuous"/>'))
P(R('1. Абзац с нумерацией строк (lnNumType). Каждая строка документа должна иметь номер на левом поле.'))
# 2 Рамка-абзац framePr (не буквица)
P(R('ТекстВРамке справа'),'<w:framePr w:w="2000" w:hSpace="180" w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/><w:pBdr><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/></w:pBdr>')
P(R('2. Этот абзац должен обтекать рамку справа. '*4))
# 3 Узор заливки
P(R('3. Заливка узором pct25 (25 % чёрного на белом).'),'<w:shd w:val="pct25" w:color="000000" w:fill="FFFFFF"/>')
P(R('3б. Узор горизонтальные полосы на жёлтом.'),'<w:shd w:val="horzStripe" w:color="FF0000" w:fill="FFFF00"/>')
# 4 Граница между абзацами
for i in range(3): P(R(f'4. Абзац в группе с границей between {i}'),'<w:pBdr><w:top w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:between w:val="single" w:sz="4" w:color="FF0000"/></w:pBdr>')
# 5 Барная табуляция
P(R('5. Бар')+'<w:r><w:tab/></w:r>'+R('после черты'),'<w:tabs><w:tab w:val="bar" w:pos="2000"/><w:tab w:val="left" w:pos="3000"/></w:tabs>')
# 6 Плавающая таблица
P(R('6. Плавающая таблица справа:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/><w:tblW w:w="2500" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="2500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>ПлавающаяЯчейка</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>Строка 2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('Текст рядом с плавающей таблицей должен её обтекать слева. '*4))
# 7 Картинка с обрезкой и поворотом
d.add_paragraph('7. Картинка: обычная, обрезанная слева на 50%, повёрнутая 90°, отражённая:')
p=d.add_paragraph(); r=p.add_run(); r.add_picture('img.png',width=Cm(3))
for mode in ('crop','rot','flip'):
    r=p.add_run(' '); r.add_picture('img.png',width=Cm(3))
d.save('t13.docx')

from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
sp=b.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
sp.append(parse_xml(f'<w:lnNumType {NS} w:countBy="5" w:restart="newPage"/>'))
for i in range(70):
    p=d.add_paragraph(f'Строка {i+1}' + (' (без номера)' if i==3 else ''))
    if i==3: p._p.get_or_add_pPr().append(parse_xml(f'<w:suppressLineNumbers {NS}/>'))
d.save('t14.docx')

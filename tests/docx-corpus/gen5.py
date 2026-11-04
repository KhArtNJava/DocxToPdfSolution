from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document()
s=d.sections[0]; s.left_margin=Cm(3); s.right_margin=Cm(1.5)
d.settings.element.append(parse_xml(f'<w:mirrorMargins {W}/>'))
p=d.add_paragraph('Оглавление: ')
p._p.append(parse_xml(f'<w:hyperlink {W} w:anchor="chap2"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>перейти к главе 2</w:t></w:r></w:hyperlink>'))
d.add_paragraph('Страница 1 (нечётная): поле слева 3 см, справа 1,5 см. '*8)
d.add_page_break()
d.add_paragraph('Страница 2 (чётная): при зеркальных полях внутреннее поле справа — 3 см, наружное слева — 1,5 см. '*8)
p=d.add_paragraph()
p._p.append(parse_xml(f'<w:bookmarkStart {W} w:id="0" w:name="chap2"/>'))
p.add_run('Глава 2 (закладка chap2)')
p._p.append(parse_xml(f'<w:bookmarkEnd {W} w:id="0"/>'))
d.save('t11.docx')

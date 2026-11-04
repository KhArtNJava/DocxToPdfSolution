from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
WPNS='xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS} {WPNS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('1. Длинное слово без пробелов: ')+R('Ааааааааааааааааааааббббббббббббббббббббввввввввввввввввввввггггггггггггггггггггддддддддддддддддддддееееееееее'))
d.add_paragraph('2. Картинки: CMYK JPEG (верх красный, низ синий), 16-бит PNG (белый|серый), серый с альфой:')
p=d.add_paragraph()
for f in ('ph0.png','ph1.png','ph2.png'): p.add_run().add_picture(f,width=Cm(3)); p.add_run(' ')
P(R('3. Таблица справа налево bidiVisual (первая колонка справа):'))
raw(f'''<w:tbl {NS}><w:tblPr><w:bidiVisual/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>ПЕРВАЯ</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>ВТОРАЯ</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('4. Фигура DrawingML без текста (синий прямоугольник 3×1 см) и VML-прямоугольник (зелёный):'))
P('''<w:r><w:drawing><wp:inline><wp:extent cx="1080000" cy="360000"/><wp:docPr id="50" name="r"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1080000" cy="360000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="2F5597"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'''+R(' и ')+'''<w:r><w:pict><v:rect style="width:60pt;height:20pt" fillcolor="#00B050" strokecolor="black"/></w:pict></w:r>''')
P(R('5. Линия DrawingML (горизонтальная 5 см):'))
P('''<w:r><w:drawing><wp:inline><wp:extent cx="1800000" cy="0"/><wp:docPr id="51" name="l"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1800000" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>''')
P(R('6. Картинка с обтеканием в абсолютной позиции страницы (x=12 см, y=20 см):'))
d.add_paragraph('Текст абзаца рядом.')
d.save('t19.docx')

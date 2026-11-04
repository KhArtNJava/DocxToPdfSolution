import zipfile, re
from docx import Document
from docx.shared import Cm
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
E_=lambda cm: int(cm*360000)
def inline(cx,cy,graphic,did): return f'<w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{did}" name="x{did}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">{graphic}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
def sp(prst,x,y,cx,cy,fill,extra=''): return f'<wps:wsp><wps:cNvPr id="1" name="s"/><wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill></wps:spPr>{extra}<wps:bodyPr/></wps:wsp>'
P(R('1. Группа фигур (красный прямоугольник слева, зелёный эллипс справа, 6×2 см):'))
grp=f'''<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{E_(6)}" cy="{E_(2)}"/><a:chOff x="0" y="0"/><a:chExt cx="600" cy="200"/></a:xfrm></wpg:grpSpPr>{sp("rect",0,0,250,200,"FF0000").replace("wps:wsp>","wps:wsp>",1)}{sp("ellipse",350,0,250,200,"00B050")}</wpg:wgp>'''
P(f'<w:r><w:drawing><wp:inline><wp:extent cx="{E_(6)}" cy="{E_(2)}"/><wp:docPr id="10" name="g"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">{grp}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
P(R('2. Произвольный контур custGeom (синий треугольник 3×2 см):'))
cg=f'''<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{E_(3)}" cy="{E_(2)}"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="100" h="100"><a:moveTo><a:pt x="50" y="0"/></a:moveTo><a:lnTo><a:pt x="100" y="100"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>'''
P(inline(E_(3),E_(2),cg,11))
P(R('3. Картинка с рамкой 3 pt (оранжевой) и ссылкой на example.org:'))
d.save('t21.docx')

from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
E_=lambda cm:int(cm*360000)
def shp(prst,fill,did,cx=E_(3.5),cy=E_(2.5)):
    return f'<w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{did}" name="g{did}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill}</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
def gs(*stops): return '<a:gsLst>'+''.join(f'<a:gs pos="{p}"><a:srgbClr val="{c}"/></a:gs>' for p,c in stops)+'</a:gsLst>'
lin=f'<a:gradFill>{gs((0,"FF0000"),(100000,"0000FF"))}<a:lin ang="0"/></a:gradFill>'
lin3=f'<a:gradFill>{gs((0,"FFFF00"),(50000,"FF0000"),(100000,"000080"))}<a:lin ang="2700000"/></a:gradFill>'
rad=f'<a:gradFill>{gs((0,"FFFFFF"),(100000,"2F5597"))}<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>'
radc=f'<a:gradFill>{gs((0,"FFF2CC"),(60000,"ED7D31"),(100000,"7F3F00"))}<a:path path="circle"><a:fillToRect l="100000" t="0" r="0" b="100000"/></a:path></a:gradFill>'
raw(f'<w:p {NS}><w:r><w:t>Градиенты: линейный, 3 цвета под 45°, радиальный (эллипс), радиальный из угла:</w:t></w:r></w:p>')
raw(f'<w:p {NS}>'+shp('rect',lin,1)+shp('roundRect',lin3,2)+shp('ellipse',rad,3)+shp('rect',radc,4)+'</w:p>')
d.save('grad.docx')

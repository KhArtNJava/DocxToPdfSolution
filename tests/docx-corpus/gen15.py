import zipfile
from docx import Document
from docx.shared import Cm
d=Document(); d.add_paragraph('1. Фигуры: star5, chevron, heart, plus, homePlate:'); d.add_paragraph('SHAPES1')
d.add_paragraph('2. Эллипс с текстом, градиент (красный→синий), accent1 lumMod 50% (тёмно-синий), roundRect adj=50%:'); d.add_paragraph('SHAPES2')
d.add_paragraph('3. Картинка с прозрачностью 40% (alphaModFix) и в оттенках серого (grayscl):'); p=d.add_paragraph()
p.add_run().add_picture('img.png',width=Cm(3)); p.add_run('  '); p.add_run().add_picture('img.png',width=Cm(3))
d.add_paragraph('4. custGeom с дугой arcTo (полукруг):'); d.add_paragraph('SHAPES3')
d.add_paragraph('5. Диаграмма (гистограмма: 3, 5, 2):'); d.add_paragraph('CHART')
d.add_paragraph('Конец.')
d.save('t22.docx')
z=zipfile.ZipFile('t22.docx'); files={i.filename:z.read(i) for i in z.infolist()}; z.close()
E_=lambda cm:int(cm*360000)
def wsp(prst,fill,cx=E_(2),cy=E_(1.5),txt=None,av='',fillxml=None,did=1):
    f=fillxml or f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    tb=f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{txt}</w:t></w:r></w:p></w:txbxContent></wps:txbx>' if txt else ''
    return f'<w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{did}" name="s{did}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="{prst}"><a:avLst>{av}</a:avLst></a:prstGeom>{f}</wps:spPr>{tb}<wps:bodyPr anchor="ctr"/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>'
s1=''.join(wsp(p,'FFC000',did=30+i) for i,p in enumerate(['star5','chevron','heart','plus','homePlate']))
grad='<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>'
scheme='<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'
s2=wsp('ellipse','9DC3E6',cx=E_(3),cy=E_(1.6),txt='В эллипсе',did=40)+wsp('rect',None,fillxml=grad,did=41)+wsp('rect',None,fillxml=scheme,did=42)+wsp('roundRect','70AD47',av='<a:gd name="adj" fmla="val 50000"/>',did=43)
arc=f'<w:r><w:drawing><wp:inline><wp:extent cx="{E_(3)}" cy="{E_(1.5)}"/><wp:docPr id="50" name="a"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{E_(3)}" cy="{E_(1.5)}"/></a:xfrm><a:custGeom><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="100"/></a:moveTo><a:arcTo wR="100" hR="100" stAng="10800000" swAng="10800000"/><a:close/></a:path></a:pathLst></a:custGeom><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
x=files['word/document.xml'].decode()
NSX=' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
x=x.replace('<w:document ','<w:document'+NSX+' ',1)
x=x.replace('<w:r><w:t>SHAPES1</w:t></w:r>',s1).replace('<w:r><w:t>SHAPES2</w:t></w:r>',s2).replace('<w:r><w:t>SHAPES3</w:t></w:r>',arc)
chart=f'<w:r><w:drawing><wp:inline><wp:extent cx="{E_(10)}" cy="{E_(6)}"/><wp:docPr id="60" name="c"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
x=x.replace('<w:r><w:t>CHART</w:t></w:r>',chart)
# alpha / grayscale on 2nd/3rd blips
parts=x.split('<a:blip ')
parts[1]=parts[1].replace('/>','><a:alphaModFix amt="40000"/></a:blip>',1)
parts[2]=parts[2].replace('/>','><a:grayscl/></a:blip>',1)
x='<a:blip '.join(parts)
files['word/document.xml']=x.encode()
files['word/charts/chart1.xml']=('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Продажи</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Ряд 1</c:v></c:pt></c:strCache></c:strRef></c:tx><c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr><c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>Янв</c:v></c:pt><c:pt idx="1"><c:v>Фев</c:v></c:pt><c:pt idx="2"><c:v>Мар</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>5</c:v></c:pt><c:pt idx="2"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="1"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart></c:chartSpace>''').encode()
files['word/_rels/document.xml.rels']=files['word/_rels/document.xml.rels'].decode().replace('</Relationships>','<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>').encode()
files['[Content_Types].xml']=files['[Content_Types].xml'].decode().replace('</Types>','<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>').encode()
out=zipfile.ZipFile('t22.docx','w',zipfile.ZIP_DEFLATED)
for k,v in files.items(): out.writestr(k,v)
out.close()

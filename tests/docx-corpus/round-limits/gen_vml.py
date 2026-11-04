from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner): raw(f'<w:p {NS}>{inner}</w:p>')
def R(t): return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'
tri='<v:shapetype id="_x0000_t5" coordsize="21600,21600" o:spt="5" adj="10800" path="m@0,l,21600r21600,xe"><v:formulas><v:f eqn="val #0"/><v:f eqn="prod #0 1 2"/><v:f eqn="sum @1 10800 0"/></v:formulas></v:shapetype>'
arrow='<v:shapetype id="_x0000_t13" coordsize="21600,21600" o:spt="13" adj="16200,5400" path="m@0,l@0@1,0@1,0@2@0@2@0,21600,21600,10800xe"><v:formulas><v:f eqn="val #0"/><v:f eqn="val #1"/><v:f eqn="sum height 0 #1"/><v:f eqn="sum 10800 0 #1"/><v:f eqn="sum width 0 #0"/><v:f eqn="prod @4 @3 10800"/><v:f eqn="sum width 0 @5"/></v:formulas></v:shapetype>'
P(R('1. VML-фигуры из шаблонов: треугольник, стрелка (adj=12000), произвольная кривая, дуги:'))
P('<w:r><w:pict>'+tri+'<v:shape type="#_x0000_t5" style="width:60pt;height:50pt" fillcolor="#4472c4" strokecolor="black"/></w:pict></w:r>'+R(' ')+
  '<w:r><w:pict>'+arrow+'<v:shape type="#_x0000_t13" adj="12000" style="width:80pt;height:40pt" fillcolor="#ed7d31"/></w:pict></w:r>'+R(' ')+
  '<w:r><w:pict><v:shape coordsize="1000,1000" path="m0,1000c200,0,800,0,1000,1000x e" style="width:70pt;height:50pt" fillcolor="#70ad47"/></w:pict></w:r>'+R(' ')+
  '<w:r><w:pict><v:shape coordsize="1000,1000" path="m0,500qx500,0,1000,500,500,1000,0,500xe m300,300 l700,300,700,700,300,700 x nf e" style="width:50pt;height:50pt" fillcolor="#ffc000" strokecolor="#7030a0" strokeweight="2pt"/></w:pict></w:r>'+R(' ')+
  '<w:r><w:pict><v:shape coordsize="1000,1000" path="ar0,0,1000,1000,1000,500,0,500 l500,500 x e" style="width:50pt;height:50pt" fillcolor="#c00000"/></w:pict></w:r>'+R(' ')+
  '<w:r><w:pict><v:polyline points="0,0 40pt,30pt 80pt,0 120pt,30pt" filled="f" strokecolor="blue" strokeweight="2pt" style="width:120pt;height:30pt"/></w:pict></w:r>')
P(R('2. VML-надпись в строке: до ')+'<w:r><w:pict><v:shape style="width:120pt;height:40pt" fillcolor="#fff2cc" strokecolor="black"><v:textbox><w:txbxContent><w:p><w:r><w:t>Надпись VML в строке</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>'+R(' после.'))
d.save('vml.docx')

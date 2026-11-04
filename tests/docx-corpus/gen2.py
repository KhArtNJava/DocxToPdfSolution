from docx import Document
from docx.shared import Pt, Cm
from docx.oxml import parse_xml
from docx.oxml.ns import qn
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LOREM='Съешь же ещё этих мягких французских булок, да выпей чаю. '
d=Document()
s=d.sections[0]
s._sectPr.append(parse_xml(f'<w:pgBorders {W} w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:left w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:bottom w:val="single" w:sz="12" w:space="24" w:color="000000"/><w:right w:val="single" w:sz="12" w:space="24" w:color="000000"/></w:pgBorders>'))
d.element.insert(0, parse_xml(f'<w:background {W} w:color="FFF8E7"/>'))
d.settings.element.append(parse_xml(f'<w:displayBackgroundShape {W}/>'))
# водяной знак: VML-надпись в колонтитуле
hp=s.header.paragraphs[0]
hp._p.append(parse_xml(f'''<w:r {W} xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:pict><v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m@7,l@8,m@5,21600l@6,21600e"><v:textpath on="t" fitshape="t"/></v:shapetype><v:shape id="PowerPlusWaterMarkObject" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:412pt;height:137pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="ЧЕРНОВИК"/></v:shape></w:pict></w:r>'''))
# Колонтитул с таблицей
ft=s.footer.add_table(rows=1, cols=3, width=Cm(16)); ft.style='Table Grid'
for i,t in enumerate(['Документ','Лист 1','Изм.']): ft.cell(0,i).text=t
d.add_heading('Рамка страницы, фон, водяной знак, docGrid',1)
s._sectPr.append(parse_xml(f'<w:docGrid {W} w:type="lines" w:linePitch="360"/>'))
for i in range(8): d.add_paragraph(LOREM*4)
# современный стиль таблицы Grid Table 4 Accent 1 (как в Word 2013+) — добавляем вручную
st=d.styles.element
st.append(parse_xml(f'''<w:style {W} w:type="table" w:styleId="GridTable4-Accent1"><w:name w:val="Grid Table 4 Accent 1"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="49"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:left w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:right w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="8EAADB"/></w:tblBorders></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:bCs/><w:color w:val="FFFFFF"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:left w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:right w:val="single" w:sz="4" w:space="0" w:color="4472C4"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="lastRow"><w:rPr><w:b/><w:bCs/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4" w:space="0" w:color="4472C4"/></w:tcBorders></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="firstCol"><w:rPr><w:b/><w:bCs/></w:rPr></w:tblStylePr>
<w:tblStylePr w:type="band1Horz"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
</w:style>'''))
t=d.add_table(rows=6, cols=3)
t._tbl.tblPr.append(parse_xml(f'<w:tblStyle {W} w:val="GridTable4-Accent1"/>'))
t._tbl.tblPr.append(parse_xml(f'<w:tblLook {W} w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'))
for i in range(6):
    for j in range(3): t.cell(i,j).text=['Регион','Q1','Q2'][j] if i==0 else (f'Регион {i}' if j==0 else str(i*j*10))
d.save('t8_page.docx')

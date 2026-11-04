import zipfile, re
from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
num=lambda nid,l=0: f'<w:numPr><w:ilvl w:val="{l}"/><w:numId w:val="{nid}"/></w:numPr>'
P(R('A. Два numId с одним abstractNum (должно продолжаться 1,2,3,4):'))
P(R('пункт'),num(101)); P(R('пункт'),num(101)); P(R('другой numId'),num(102)); P(R('другой numId'),num(102))
P(R('B. isLgl и lvlRestart:'))
P(R('Глава'),num(103,0)); P(R('Подпункт (ожидается 1.1 арабскими)'),num(103,1)); P(R('Подпункт'),num(103,1)); P(R('Глава'),num(103,0)); P(R('Подпункт (снова 2.1)'),num(103,1))
P(R('C. Отступ w:start=1440, w:end=1440 (новые имена атрибутов) — длинный текст для проверки правого отступа тоже. '*2),'<w:ind w:start="1440" w:end="1440"/>')
P(R('D. Интервал beforeLines=200 (2 строки) перед этим абзацем'),'<w:spacing w:beforeLines="200" w:before="0"/>')
P(R('E. Выделение darkYellow, darkCyan, lightGray: ')+R('тёмно-жёлтый','<w:highlight w:val="darkYellow"/>')+R(' ')+R('тёмно-бирюзовый','<w:highlight w:val="darkCyan"/>')+R(' ')+R('светло-серый','<w:highlight w:val="lightGray"/>'))
P(R('F. Таблица с tblCellSpacing:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblCellSpacing w:w="100" w:type="dxa"/><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="12"/><w:left w:val="single" w:sz="12"/><w:bottom w:val="single" w:sz="12"/><w:right w:val="single" w:sz="12"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>{"".join(f"<w:tr><w:tc><w:p><w:r><w:t>Я{i}1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Я{i}2</w:t></w:r></w:p></w:tc></w:tr>" for i in range(2))}</w:tbl>''')
P(R('G. Сноска с форматом lowerRoman')+'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'+R(' и вторая')+'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r>')
P(R('H. Таблица: строка с точной высотой 0,5 см и большим текстом (лишнее обрезается):'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="284" w:hRule="exact"/></w:trPr><w:tc><w:p><w:r><w:t>Видимая строка</w:t></w:r><w:r><w:br/><w:t>ОБРЕЗАНО1</w:t></w:r><w:r><w:br/><w:t>ОБРЕЗАНО2</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('I. Конец.'))
sp=b.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
sp.append(parse_xml(f'<w:footnotePr {NS}><w:numFmt w:val="lowerRoman"/></w:footnotePr>'))
# second section, vertically centered title page
d.save('t15.docx')
z=zipfile.ZipFile('t15.docx'); files={i.filename:z.read(i) for i in z.infolist()}; z.close()
nx=files['word/numbering.xml'].decode()
absn=f'''<w:abstractNum w:abstractNumId="90"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:abstractNum w:abstractNumId="91"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="Глава %1"/><w:pPr><w:ind w:left="0" w:firstLine="0"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:isLgl/><w:lvlText w:val="%1.%2"/><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr></w:lvl></w:abstractNum>'''
nx=re.sub(r'(<w:abstractNum )',absn.replace('\\','\\\\')+r'\1',nx,count=1)
nx=nx.replace('</w:numbering>','<w:num w:numId="101"><w:abstractNumId w:val="90"/></w:num><w:num w:numId="102"><w:abstractNumId w:val="90"/></w:num><w:num w:numId="103"><w:abstractNumId w:val="91"/></w:num></w:numbering>')
files['word/numbering.xml']=nx.encode()
fn=f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
<w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Первая сноска (i)</w:t></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Вторая сноска (ii)</w:t></w:r></w:p></w:footnote></w:footnotes>'''
files['word/footnotes.xml']=fn.encode()
files['word/_rels/document.xml.rels']=files['word/_rels/document.xml.rels'].decode().replace('</Relationships>','<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode()
ct=files['[Content_Types].xml'].decode()
files['[Content_Types].xml']=ct.replace('</Types>','<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>').encode()
dx=files['word/document.xml'].decode()
# insert section break: first section vAlign center with title
title=f'<w:p><w:pPr><w:jc w:val="center"/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/><w:vAlign w:val="center"/></w:sectPr></w:pPr><w:r><w:t>ТИТУЛ ПО ЦЕНТРУ СТРАНИЦЫ</w:t></w:r></w:p>'
dx=dx.replace('<w:body>','<w:body>'+title,1)
files['word/document.xml']=dx.encode()
out=zipfile.ZipFile('t15.docx','w',zipfile.ZIP_DEFLATED)
for k,v in files.items(): out.writestr(k,v)
out.close()

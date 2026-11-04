import zipfile, re
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml
from docx.shared import Pt
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document()
# Стиль с нумерацией и символьный стиль bold
st=d.styles.add_style('NumHead',1); st.font.size=Pt(14); st.font.bold=True
cs=d.styles.add_style('BoldChar',2); cs.font.bold=True
d.sections[0].header.paragraphs[0].text='ВЕРХНИЙ КОЛОНТИТУЛ РАЗДЕЛА 1'
d.sections[0].footer.paragraphs[0].text='НИЖНИЙ КОЛОНТИТУЛ'
b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('Заголовок стилем с нумерацией (ожидается «1.»)'),'<w:pStyle w:val="NumHead"/>')
P(R('Обычный текст. ')+R('Жирный стиль + прямой жирный (Word: остаётся жирным)','<w:rStyle w:val="BoldChar"/><w:b/>'))
P(R('Заголовок стилем с нумерацией (ожидается «2.»)'),'<w:pStyle w:val="NumHead"/>')
P(R('Ссылка полем: ')+'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.com/field" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>example.com</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
P(R('Ниже пустой абзац с маркером 36 pt:'))
P('', '<w:rPr><w:sz w:val="72"/></w:rPr>')
P(R('После пустого абзаца.'))
P(R('Таблица с gridBefore=1 и gridAfter=1 во второй строке:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideH w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
<w:tr><w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="2000" w:type="dxa"/><w:gridAfter w:val="1"/><w:wAfter w:w="2000" w:type="dxa"/></w:trPr><w:tc><w:p><w:r><w:t>ТолькоСредняя</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('Перенос с очисткой обтекания: br clear=all.'))
# раздел 2 без своих колонтитулов — должен унаследовать
s2=d.add_section(WD_SECTION.NEW_PAGE)
d.add_paragraph('Раздел 2: колонтитулы должны быть унаследованы от раздела 1.')
d.save('t16.docx')
z=zipfile.ZipFile('t16.docx'); files={i.filename:z.read(i) for i in z.infolist()}; z.close()
nx=files['word/numbering.xml'].decode()
absn='<w:abstractNum w:abstractNumId="95"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="NumHead"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl></w:abstractNum>'
nx=re.sub(r'(<w:abstractNum )',absn+r'\1',nx,count=1).replace('</w:numbering>','<w:num w:numId="150"><w:abstractNumId w:val="95"/></w:num></w:numbering>')
files['word/numbering.xml']=nx.encode()
sx=files['word/styles.xml'].decode()
sx=sx.replace('<w:style w:type="paragraph" w:customStyle="1" w:styleId="NumHead"><w:name w:val="NumHead"/>','<w:style w:type="paragraph" w:customStyle="1" w:styleId="NumHead"><w:name w:val="NumHead"/><w:pPr><w:numPr><w:numId w:val="150"/></w:numPr></w:pPr>')
assert 'w:numId w:val="150"' in sx
files['word/styles.xml']=sx.encode()
out=zipfile.ZipFile('t16.docx','w',zipfile.ZIP_DEFLATED)
for k,v in files.items(): out.writestr(k,v)
out.close()

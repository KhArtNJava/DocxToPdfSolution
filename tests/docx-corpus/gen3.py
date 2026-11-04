from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document()
d.add_paragraph('Смешанный текст: English עברית מימין לשמאל and العربية من اليمين إلى اليسار 123 конец.')
p=d.add_paragraph('שלום עולם! זהו פסקה בכיוון מימין לשמאל עם מספר 2026 ו-English.')
p._p.get_or_add_pPr().append(parse_xml(f'<w:bidi {W}/>'))
p=d.add_paragraph('مرحبا بالعالم. هذه فقرة عربية من اليمين إلى اليسار.')
p._p.get_or_add_pPr().append(parse_xml(f'<w:bidi {W}/>'))
d.add_paragraph('Формула отдельной строкой:')
p=d.add_paragraph()
p._p.append(parse_xml('<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:sSup><m:e><m:r><m:t>e</m:t></m:r></m:e><m:sup><m:r><m:t>iπ</m:t></m:r></m:sup></m:sSup><m:r><m:t>+1=0, </m:t></m:r><m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub></m:e></m:nary><m:r><m:t>, </m:t></m:r><m:f><m:num><m:r><m:t>-b±</m:t></m:r><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:sSup><m:e><m:r><m:t>b</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup><m:r><m:t>-4ac</m:t></m:r></m:e></m:rad></m:num><m:den><m:r><m:t>2a</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>'))
d.save('t9_bidi_math.docx')

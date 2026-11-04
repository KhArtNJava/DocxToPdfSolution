from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('1. Лам-алеф: ')+R('لا إله إلا الله — السلام عليكم','<w:rtl/>'))
P(R('هذا نص (بين قوسين) و [معقوفين] و «اقتباس» 123.','<w:rtl/>'),'<w:bidi/>')
P(R('3. Встраивание RLE: abc ‫שלום world‬ def'))
P(R('4. Переопределение w:bdo: ')+f'<w:bdo w:val="rtl"><w:r><w:t>ABCDEF</w:t></w:r></w:bdo>'+R(' конец'))
P(R('5. Изолят RLI: user ⁧משתמש 12⁩ logged in'))
P(R('6. Иврит с числом и процентами: ')+R('המחיר הוא 25% מתוך 1,234.50 ש"ח','<w:rtl/>'))
d.save('bidi.docx')

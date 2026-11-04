from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
h=d.styles.add_style('BoldPara',1); h.font.bold=True
cs=d.styles.add_style('BoldChar2',2); cs.font.bold=True
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('1. Поле переплёта gutter=1440: левый край текста должен быть на 1 дюйм правее обычного.'))
P(R('2. Иврит с szCs=40 (крупно) и sz=20: ')+R('שלום עולם','<w:rtl/><w:sz w:val="20"/><w:szCs w:val="40"/>')+R(' и bCs: ')+R('مرحبا','<w:rtl/><w:bCs/>'))
P(R('3. Перемещение: ')+'<w:moveFrom w:id="1" w:author="a"><w:r><w:t>ПЕРЕМЕЩЕНО_ОТСЮДА</w:t></w:r></w:moveFrom>'+R('остаток'))
P(R('4. Абзац с удалённым знаком абзаца — '),'<w:rPr><w:del w:id="2" w:author="a"/></w:rPr>')
P(R('продолжение должно быть в той же строке.'))
P(R('5. Таблица без tblGrid (ширины из tcW 1000/3000):'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Узк</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Широкая ячейка</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('6. Знаки ударения w:em dot: ')+R('важно','<w:em w:val="dot"/>'))
P(R('7. Стиль абзаца жирный + символьный жирный (Word: XOR → обычный): ')+R('ОБЫЧНЫЙ','<w:rStyle w:val="BoldChar2"/>'),'<w:pStyle w:val="BoldPara"/>')
P(R('8. Эмодзи и редкие символы: 😀 ✓ → ∑ ≠ ½'))
P(R('9. Отрицательная позиция: базовая ')+R('опущено','<w:position w:val="-8"/>')+R(' текст'))
sp=b.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr')
pm=sp.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pgMar')
pm.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}gutter','1440')
d.save('t18.docx')

from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner, ppr=''): raw(f'<w:p {NS}><w:pPr>{ppr}</w:pPr>{inner}</w:p>')
def R(t, rpr=''): return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('1. SDT блок: ')+'')
raw(f'<w:sdt {NS}><w:sdtPr><w:alias w:val="x"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>ТекстВнутриSDTблока</w:t></w:r></w:p></w:sdtContent></w:sdt>')
P(R('2. SDT строчный: до ')+'<w:sdt><w:sdtContent><w:r><w:t>СтрочныйSDT</w:t></w:r></w:sdtContent></w:sdt>'+R(' после'))
P(R('3. smartTag: ')+'<w:smartTag w:uri="u" w:element="City"><w:r><w:t>ГородМосква</w:t></w:r></w:smartTag>'+'<w:customXml w:element="c"><w:r><w:t> КастомXML</w:t></w:r></w:customXml>')
P(R('4. fldSimple: ')+'<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>ИмяАвтора</w:t></w:r></w:fldSimple>')
P(R('5. Символ sym: ')+'<w:r><w:sym w:font="Wingdings" w:char="F04A"/></w:r>'+R(' конец'))
P(R('6. Капитель: ')+R('Малые Прописные Буквы','<w:smallCaps/>'))
P(R('7. Разрядка: ')+R('Разреженный','<w:spacing w:val="60"/>')+R(' Уплотнённый','<w:spacing w:val="-20"/>'))
P(R('8. ptab:')+'<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="dot"/></w:r>'+R('ПравыйКрай'))
P(R('9. Неразрывный')+'<w:r><w:noBreakHyphen/></w:r>'+R('дефис и ')+'<w:r><w:t>мягкий</w:t></w:r><w:r><w:softHyphen/></w:r><w:r><w:t>перенос</w:t></w:r>')
P(R('10. Вложенное поле IF: ')+'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 "Да" "Нет" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>РезультатПоля</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
P(R('11. Выделение цветом: ')+R('жёлтый','<w:highlight w:val="yellow"/>')+R(' заливка','<w:shd w:val="clear" w:fill="00FF00"/>'))
P(R('12. Двойное зачёркивание','<w:dstrike/>')+R(' волнистое подчёркивание','<w:u w:val="wave" w:color="FF0000"/>'))
P(R('13. Абзац справа налево с отступом'),'<w:ind w:left="2000"/><w:jc w:val="right"/>')
P(R('14. Вертикальный: H')+R('2','<w:vertAlign w:val="subscript"/>')+R('O, x')+R('2','<w:vertAlign w:val="superscript"/>')+R(' поднят','<w:position w:val="10"/>'))
P(R('15. Скрытый текст:')+R('НЕВИДИМЫЙ','<w:vanish/>')+R(' видимый'))
P(R('16. Удалено:')+'<w:del w:id="1" w:author="a"><w:r><w:delText>УДАЛЁННОЕ</w:delText></w:r></w:del>'+'<w:ins w:id="2" w:author="a"><w:r><w:t>ВставленноеСлово</w:t></w:r></w:ins>')
P(R('17. Таблица по центру:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:jc w:val="center"/><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>ЦентрA</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>ЦентрB</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('18. Таблица с отступом tblInd 2 см:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblInd w:w="1134" w:type="dxa"/><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>ОтступТаблицы</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('20. ptab центр:')+'<w:r><w:ptab w:relativeTo="margin" w:alignment="center" w:leader="none"/></w:r>'+R('Середина')+'<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="hyphen"/></w:r>'+R('стр. 7'))
P(R('21. Стили: ')+R('толстое','<w:u w:val="thick" w:color="0000FF"/>')+R(' ')+R('точки','<w:u w:val="dotted"/>')+R(' ')+R('штрих','<w:u w:val="dash" w:color="00AA00"/>')+R(' только слова здесь','<w:u w:val="words"/>')+R(' двойная волна','<w:u w:val="wavyDouble" w:color="FF0000"/>'))
P(R('22. Таблица 50%:'))
raw(f'''<w:tbl {NS}><w:tblPr><w:tblW w:w="50%" w:type="pct"/><w:jc w:val="right"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr>
<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Узкая</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Широкая колонка</w:t></w:r></w:p></w:tc></w:tr></w:tbl>''')
P(R('19. Конец.'))
d.save('t12.docx')

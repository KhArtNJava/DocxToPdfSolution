import zipfile, sys
from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
d=Document(); b=d.element.body
if len(sys.argv)>1:
    from docx.shared import Cm
    d.sections[0].right_margin=Cm(1.5)
def raw(x): b.insert(len(b)-1, parse_xml(x))
def P(inner): raw(f'<w:p {NS}>{inner}</w:p>')
def R(t): return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'
P(R('Исходный текст, ')+'<w:ins w:id="1" w:author="Иван Петров" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">вставленный Иваном, </w:t></w:r></w:ins>'+'<w:del w:id="2" w:author="Иван Петров"><w:r><w:delText xml:space="preserve">удалённый Иваном, </w:delText></w:r></w:del>'+R('и ещё ')+'<w:ins w:id="3" w:author="Мария Смирнова"><w:r><w:t>вставка Марии</w:t></w:r></w:ins>'+R('.'))
P(R('Без исправлений — обычный абзац без черты на поле.'))
P('<w:moveFrom w:id="4" w:author="Иван Петров"><w:r><w:t>Перемещённая фраза.</w:t></w:r></w:moveFrom>'+R(' Остаток абзаца.'))
P(R('Сюда: ')+'<w:moveTo w:id="5" w:author="Иван Петров"><w:r><w:t>Перемещённая фраза.</w:t></w:r></w:moveTo>')
P(R('Здесь ')+'<w:commentRangeStart w:id="10"/>'+R('выделенный фрагмент с примечанием')+'<w:commentRangeEnd w:id="10"/><w:r><w:commentReference w:id="10"/></w:r>'+R(' и ещё одно ')+'<w:commentRangeStart w:id="11"/>'+R('слово')+'<w:commentRangeEnd w:id="11"/><w:r><w:commentReference w:id="11"/></w:r>'+R('.'))
d.save('rev.docx')
z=zipfile.ZipFile('rev.docx'); files={i.filename:z.read(i) for i in z.infolist()}; z.close()
files['word/comments.xml']=f'''<?xml version="1.0" encoding="UTF-8"?><w:comments {NS}><w:comment w:id="10" w:author="Мария Смирнова" w:initials="МС"><w:p><w:r><w:t>Нужно уточнить формулировку и проверить источник.</w:t></w:r></w:p></w:comment><w:comment w:id="11" w:author="Иван Петров" w:initials="ИП"><w:p><w:r><w:t>Согласен.</w:t></w:r></w:p></w:comment></w:comments>'''.encode()
files['word/_rels/document.xml.rels']=files['word/_rels/document.xml.rels'].decode().replace('</Relationships>','<Relationship Id="rIdCm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>').encode()
files['[Content_Types].xml']=files['[Content_Types].xml'].decode().replace('</Types>','<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>').encode()
out=zipfile.ZipFile('rev.docx','w',zipfile.ZIP_DEFLATED)
for k,v in files.items(): out.writestr(k,v)
out.close()

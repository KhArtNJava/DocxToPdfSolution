import zipfile, re
from docx import Document
from docx.shared import Cm
from PIL import Image
files=['shapes.emf','shapes.wmf','rgb_none.tif','rgb_lzw.tif','rgb_deflate.tif','rgb_packbits.tif','rgb_jpeg.tif','bw_g4.tif','bw_g3.tif','bw_none.tif','pal.tif','gray.tif','cmyk.tif','rgba.tif']
d=Document()
d.add_paragraph('EMF и WMF (векторные, из LibreOffice):')
p=d.add_paragraph()
for i,f in enumerate(files):
    Image.new('RGB',(40+i,30),(i*10,0,0)).save(f'ph{i}.png')
for i,f in enumerate(files[:2]): p.add_run().add_picture(f'ph{i}.png',width=Cm(7.5)); p.add_run(' ')
d.add_paragraph('TIFF: без сжатия, LZW, Deflate, PackBits, JPEG, G4, G3, 1 бит, палитра, серый LZW, CMYK, RGBA:')
p=d.add_paragraph()
for i,f in enumerate(files[2:],2): p.add_run().add_picture(f'ph{i}.png',width=Cm(3.6)); p.add_run(' ')
d.save('t23.docx')
z=zipfile.ZipFile('t23.docx'); data={i.filename:z.read(i) for i in z.infolist()}; z.close()
rels=data['word/_rels/document.xml.rels'].decode()
for i,f in enumerate(files):
    ph=open(f'ph{i}.png','rb').read()
    name=[k for k,v in data.items() if k.startswith('word/media/') and v==ph][0]
    ext=f.split('.')[1]
    new=name.rsplit('.',1)[0]+'.'+ext
    del data[name]; data[new]=open(f,'rb').read()
    rels=rels.replace(name.split('word/')[1], new.split('word/')[1])
data['word/_rels/document.xml.rels']=rels.encode()
ct=data['[Content_Types].xml'].decode()
ct=ct.replace('</Types>','<Default Extension="emf" ContentType="image/x-emf"/><Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="tif" ContentType="image/tiff"/></Types>')
data['[Content_Types].xml']=ct.encode()
out=zipfile.ZipFile('t23.docx','w',zipfile.ZIP_DEFLATED)
for k,v in data.items(): out.writestr(k,v)
out.close()

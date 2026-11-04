import subprocess, re, collections, sys
from PIL import Image, ImageDraw
import glob, os
def words(pdf):
    t=subprocess.run(['pdftotext','-layout',pdf,'-'],capture_output=True,text=True).stdout
    return collections.Counter(re.findall(r'[\wЀ-ӿ]+',t.lower()))
def pages(pdf):
    o=subprocess.run(['pdfinfo',pdf],capture_output=True,text=True).stdout
    return int(re.search(r'Pages:\s+(\d+)',o).group(1))
for f in sorted(glob.glob('t*.docx')):
    b=f[:-5]; o=f'ours/{b}.pdf'; l=f'lo/{b}.pdf'
    wo,wl=words(o),words(l)
    miss=[w for w in wl if wl[w]>wo[w] and len(w)>2]; extra=[w for w in wo if wo[w]>wl[w] and len(w)>2]
    print(f'== {b}: pages ours={pages(o)} lo={pages(l)}')
    print('   missing:',sorted(miss)[:40]); print('   extra  :',sorted(extra)[:40])
    for side,pdf in (('o',o),('l',l)):
        for p in glob.glob(f'{b}-{side}-*.png'): os.remove(p)
        subprocess.run(['pdftoppm','-r','50','-png',pdf,f'{b}-{side}'])
    n=max(len(glob.glob(f'{b}-o-*.png')),len(glob.glob(f'{b}-l-*.png')))
    rows=[]
    for i in range(1,n+1):
        def ld(side):
            g=sorted(glob.glob(f'{b}-{side}-*{i}.png'))
            g=[x for x in g if re.search(rf'-0*{i}\.png$',x)]
            return Image.open(g[0]) if g else None
        a,c=ld('o'),ld('l')
        rows.append((a,c))
    W=max((x.width for r in rows for x in r if x),default=100); H=max((x.height for r in rows for x in r if x),default=100)
    for k in range(0,len(rows),3):
        chunk=rows[k:k+3]
        img=Image.new('RGB',(2*W+20,len(chunk)*(H+10)),'gray')
        for j,(a,c) in enumerate(chunk):
            if a: img.paste(a,(0,j*(H+10)))
            if c: img.paste(c,(W+20,j*(H+10)))
        img.save(f'sheet_{b}_{k//3}.png')

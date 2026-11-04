from docx import Document
from docx.oxml import parse_xml
NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
d=Document(); b=d.element.body
def raw(x): b.insert(len(b)-1, parse_xml(x))
def r(t, plain=False): return f'<m:r>{"<m:rPr><m:sty m:val=\"p\"/></m:rPr>" if plain else ""}<m:t xml:space="preserve">{t}</m:t></m:r>'
def e(x): return f'<m:e>{x}</m:e>'
def frac(n,dd): return f'<m:f><m:num>{n}</m:num><m:den>{dd}</m:den></m:f>'
def sup(bb,s): return f'<m:sSup><m:e>{bb}</m:e><m:sup>{s}</m:sup></m:sSup>'
def sub(bb,s): return f'<m:sSub><m:e>{bb}</m:e><m:sub>{s}</m:sub></m:sSub>'
def subsup(bb,s,p): return f'<m:sSubSup><m:e>{bb}</m:e><m:sub>{s}</m:sub><m:sup>{p}</m:sup></m:sSubSup>'
def rad(x,deg=None): return f'<m:rad><m:radPr><m:degHide m:val="{0 if deg else 1}"/></m:radPr><m:deg>{deg or ""}</m:deg><m:e>{x}</m:e></m:rad>'
def nary(ch,s,p,x,loc=None): return f'<m:nary><m:naryPr><m:chr m:val="{ch}"/>{f"<m:limLoc m:val=\"{loc}\"/>" if loc else ""}</m:naryPr><m:sub>{s}</m:sub><m:sup>{p}</m:sup><m:e>{x}</m:e></m:nary>'
def delim(x,b0='(',e0=')'): return f'<m:d><m:dPr><m:begChr m:val="{b0}"/><m:endChr m:val="{e0}"/></m:dPr>{x}</m:d>'
def mat(rows): return '<m:m>'+''.join('<m:mr>'+''.join(e(c) for c in row)+'</m:mr>' for row in rows)+'</m:m>'
def para(x, disp=True):
    if disp: raw(f'<w:p {NS}><m:oMathPara><m:oMath>{x}</m:oMath></m:oMathPara></w:p>')
    else: raw(f'<w:p {NS}><w:r><w:t xml:space="preserve">Встроенная формула </w:t></w:r><m:oMath>{x}</m:oMath><w:r><w:t xml:space="preserve"> в тексте абзаца.</w:t></w:r></w:p>')
para(r('x')+r('=')+frac(r('-b')+r('±')+rad(sup(r('b'),r('2'))+r('-4ac')), r('2a')))
para(nary('∑',r('i=1'),r('n'),sup(r('i'),r('2')))+r('=')+frac(r('n')+delim(e(r('n+1')))+delim(e(r('2n+1'))),r('6')))
para(nary('∫',r('0'),r('∞'),sup(r('e'),r('-')+sup(r('x'),r('2')))+r('dx'))+r('=')+frac(rad(r('π')),r('2')))
para(r('A=')+delim(e(mat([[r('a'),r('b'),r('c')],[r('d'),r('e'),r('f')],[r('g'),r('h'),r('i')]])),'[',']')+r(', ')+'<m:func><m:fName>'+r('sin',True)+'</m:fName><m:e>'+r('α')+'</m:e></m:func>'+r('+')+'<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>'+r('v')+'</m:e></m:acc>'+r('+')+'<m:acc><m:e>'+r('x')+'</m:e></m:acc>')
para(r('f')+delim(e(r('x')))+r('=')+delim(e('<m:eqArr><m:e>'+r('x, x≥0')+'</m:e><m:e>'+r('-x, x&lt;0')+'</m:e></m:eqArr>'),'{',''))
para('<m:limLow><m:e>'+r('lim',True)+'</m:e><m:lim>'+r('n→∞')+'</m:lim></m:limLow>'+delim(e(sup(delim(e(r('1+')+frac(r('1'),r('n')))),r('n'))))+r('=e'))
para(frac(frac(r('a'),r('b')),frac(r('c'),r('d')))+r('+')+rad(r('x+y'),r('3'))+r('+')+subsup(r('x'),r('i'),r('2'))+r('+')+'<m:groupChr><m:e>'+r('a+…+a')+'</m:e></m:groupChr>')
para(frac(r('a+b'),r('c'))+r('+')+sup(r('x'),r('2')), disp=False)
d.save('math.docx')

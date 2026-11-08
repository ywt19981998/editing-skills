# 提取 docx 段落文字，并把 Word 公式（OMML）转成可读的线性文字，公式用 ⟦ ⟧ 包起来；MathType 等 OLE 公式标为 [MathType公式]
import sys, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
def ln(el): return etree.QName(el).localname
def ns(el): return etree.QName(el).namespace

def mtext(el):
    """把一个 OMML 结构转成线性文字"""
    tag = ln(el)
    kids = lambda name: [c for c in el if ns(c) == M and ln(c) == name]
    sub = lambda name: ''.join(mtext(c) for c in kids(name))
    if tag == 't': return el.text or ''
    if tag == 'f':  return f'({sub("num")})/({sub("den")})'
    if tag == 'sSup': return f'{sub("e")}^{{{sub("sup")}}}'
    if tag == 'sSub': return f'{sub("e")}_{{{sub("sub")}}}'
    if tag == 'sSubSup': return f'{sub("e")}_{{{sub("sub")}}}^{{{sub("sup")}}}'
    if tag == 'rad':
        deg = sub('deg'); return f'√[{deg}]({sub("e")})' if deg else f'√({sub("e")})'
    if tag == 'd':
        pr = el.find(f'{{{M}}}dPr'); beg, end = '(', ')'
        if pr is not None:
            b, e = pr.find(f'{{{M}}}begChr'), pr.find(f'{{{M}}}endChr')
            if b is not None: beg = b.get(f'{{{M}}}val', '')
            if e is not None: end = e.get(f'{{{M}}}val', '')
        return beg + ','.join(mtext(c) for c in kids('e')) + end
    if tag == 'nary':
        pr = el.find(f'{{{M}}}naryPr'); ch = '∫'
        if pr is not None and pr.find(f'{{{M}}}chr') is not None: ch = pr.find(f'{{{M}}}chr').get(f'{{{M}}}val')
        return f'{ch}_{{{sub("sub")}}}^{{{sub("sup")}}} {sub("e")}'
    if tag == 'func': return f'{sub("fName")} {sub("e")}'
    if tag == 'limLow': return f'{sub("e")}_{{{sub("lim")}}}'
    if tag == 'limUpp': return f'{sub("e")}^{{{sub("lim")}}}'
    if tag in ('acc', 'bar'): return f'{sub("e")}̂' if tag == 'acc' else f'‾{sub("e")}'
    if tag == 'eqArr': return ' ; '.join(mtext(c) for c in kids('e'))
    if tag == 'm': return '[' + ' ; '.join(' , '.join(mtext(e) for e in r if ln(e) == 'e') for r in el if ln(r) == 'mr') + ']'
    if tag.endswith('Pr') or tag == 'ctrlPr': return ''
    return ''.join(mtext(c) for c in el)

def para(p):
    out = []
    for ch in p:
        if ns(ch) == M and ln(ch) in ('oMath', 'oMathPara'):
            out.append('⟦' + mtext(ch) + '⟧')
        elif ln(ch) == 'r':
            for x in ch.iter():
                t = ln(x)
                if t == 't' and ns(x) == W: out.append(x.text or '')
                elif t == 'tab': out.append('\t')
                elif t == 'noBreakHyphen': out.append('-')
                elif t == 'OLEObject': out.append('[MathType公式]')
                elif t in ('drawing',) : out.append('[图片]')
                elif t == 'imagedata' and not any(ln(a) == 'OLEObject' for a in ch.iter()): out.append('[图片]')
        elif ln(ch) in ('hyperlink', 'smartTag', 'ins'):
            out.append(para(ch))
    return ''.join(out)

z = zipfile.ZipFile(sys.argv[1]); body = etree.fromstring(z.read('word/document.xml')).find(f'{{{W}}}body')
n = 0
for el in body:
    if ln(el) == 'p':
        n += 1; t = para(el)
        if t.strip(): print(f'P{n}: {t}')
    elif ln(el) == 'tbl':
        for tr in el.iter(f'{{{W}}}tr'):
            n += 1
            print(f'T{n}: ' + ' | '.join(''.join(para(p) for p in tc.iter(f'{{{W}}}p')) for tc in tr.findall(f'{{{W}}}tc')))

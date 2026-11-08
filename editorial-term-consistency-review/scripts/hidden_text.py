# 列出 docx 中不可见的文字：白色字体、隐藏属性(vanish)、极小字号
import sys, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
z = zipfile.ZipFile(sys.argv[1]); doc = etree.fromstring(z.read('word/document.xml'))
for i, p in enumerate(doc.iter(q('p'))):
    hid = []
    for r in p.iter(q('r')):
        txt = ''.join(t.text or '' for t in r.iter(q('t')))
        if not txt.strip(): continue
        rpr = r.find(q('rPr')); why = []
        if rpr is not None:
            c = rpr.find(q('color'))
            if c is not None and (c.get(q('val')) or '').upper() in ('FFFFFF', 'FEFEFE'): why.append('白色')
            if rpr.find(q('vanish')) is not None: why.append('隐藏')
            sz = rpr.find(q('sz'))
            if sz is not None and int(sz.get(q('val'), '20')) <= 2: why.append('极小字号')
        if why: hid.append(f'{"/".join(why)}:{txt[:30]}')
    if hid: print(f'p{i}:', hid[:4], f'(共{len(hid)}段)' if len(hid) > 4 else '')

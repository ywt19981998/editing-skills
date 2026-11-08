# 提取 docx 段落文字，标出修订（插入/删除）和批注
import sys, zipfile, re
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': W}
z = zipfile.ZipFile(sys.argv[1])
doc = etree.fromstring(z.read('word/document.xml'))
comments = {}
if 'word/comments.xml' in z.namelist():
    c = etree.fromstring(z.read('word/comments.xml'))
    for cm in c.findall('w:comment', ns):
        comments[cm.get(f'{{{W}}}id')] = ''.join(cm.itertext())
def run_text(el):
    out = []
    for node in el.iter():
        tag = etree.QName(node).localname
        if tag == 't': out.append(node.text or '')
        elif tag == 'delText': out.append(node.text or '')
        elif tag == 'tab': out.append('\t')
        elif tag == 'noBreakHyphen': out.append('-')
        elif tag in ('drawing', 'pict'): out.append('[图片]')
    return ''.join(out)

def is_yellow(r):
    rpr = r.find('w:rPr', ns)
    if rpr is None: return False
    h, sh = rpr.find('w:highlight', ns), rpr.find('w:shd', ns)
    return (h is not None and h.get(f'{{{W}}}val') == 'yellow') or (sh is not None and (sh.get(f'{{{W}}}fill') or '').upper() == 'FFFF00')
body = doc.find('w:body', ns)
n = 0
def para(p):
    s = []
    for ch in p:
        tag = etree.QName(ch).localname
        if tag == 'ins': s.append('[+' + run_text(ch) + '+]')
        elif tag == 'del': s.append('[-' + run_text(ch) + '-]')
        elif tag == 'commentRangeStart': s.append('[批注' + ch.get(f'{{{W}}}id') + '>')
        elif tag == 'commentRangeEnd': s.append('<批注' + ch.get(f'{{{W}}}id') + ']')
        elif tag == 'r':
            t = run_text(ch)
            s.append(f'【黄:{t}】' if t and is_yellow(ch) else t)
        elif tag in ('hyperlink', 'smartTag', 'fldSimple'): s.append(run_text(ch))
    return ''.join(s).replace('】【黄:', '')
for el in body:
    tag = etree.QName(el).localname
    if tag == 'p':
        n += 1
        t = para(el)
        if t.strip(): print(f'P{n}: {t}')
    elif tag == 'tbl':
        for r_i, tr in enumerate(el.iter(f'{{{W}}}tr')):
            cells = [''.join(para(p) for p in tc.iter(f'{{{W}}}p')) for tc in tr.findall('w:tc', ns)]
            n += 1
            print(f'T{n}: ' + ' | '.join(cells))
for k, v in comments.items():
    print(f'批注{k}: {v}')

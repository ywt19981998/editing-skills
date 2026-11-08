# 列出某些写法在稿中的全部出现位置，并给出全文唯一的 ctx，用来写 B 方案的修订
# 用法: find_term.py 稿件.docx 词1 [词2 ...] [--exclude 含此词的不算 ...]
# 例：find_term.py 第3章.docx 经纬网格 --exclude 经纬网格化
import sys, zipfile, json
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
args = sys.argv[2:]
terms = args[:args.index('--exclude')] if '--exclude' in args else args
excl = args[args.index('--exclude') + 1:] if '--exclude' in args else []
doc = etree.fromstring(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))
paras = [''.join(t.text or '' for t in p.iter(f'{{{W}}}t')) for p in doc.iter(f'{{{W}}}p')]
# P 号与 dump_docx.py 一致：正文段落和表格的每一行各占一个号
label, n = [], 0
for el in doc.find(f'{{{W}}}body'):
    if el.tag == f'{{{W}}}p':
        n += 1; label += [f'P{n}'] * (1 + sum(1 for _ in el.iter(f'{{{W}}}p')) - 1)
    elif el.tag == f'{{{W}}}tbl':
        for tr in el.iter(f'{{{W}}}tr'):
            n += 1; label += [f'P{n}'] * sum(1 for _ in tr.iter(f'{{{W}}}p'))

def ctx(t, k, n):
    for pad in (4, 8, 12, 20, 40):
        s = t[max(0, k - pad): k + n + pad]
        if sum(s in x for x in paras) == 1 and t.find(s) == max(0, k - pad): return s
    return t
hits = []
for pi, t in enumerate(paras):
    for term in terms:
        k = t.find(term)
        while k >= 0:
            if not any(x in t[max(0, k - len(x)): k + len(term) + len(x)] and term in x for x in excl):
                hits.append(dict(para=label[pi] if pi < len(label) else f'#{pi}', term=term, ctx=ctx(t, k, len(term)), around=t[max(0, k - 15): k + len(term) + 15]))
            k = t.find(term, k + len(term))
for h in hits: print(json.dumps(h, ensure_ascii=False))
print(f'共 {len(hits)} 处', file=sys.stderr)

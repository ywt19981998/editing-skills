# 打印指定段落的 run 切分，判断目标字串是否落在单个 run 内
import sys, zipfile
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
z=zipfile.ZipFile(sys.argv[1]); doc=etree.fromstring(z.read('word/document.xml'))
body=doc.find(f'{{{W}}}body')
targets=sys.argv[2:]
for i,p in enumerate(body.iter(f'{{{W}}}p')):
    txt=''.join(t.text or '' for t in p.iter(f'{{{W}}}t'))
    for tg in targets:
        if tg in txt:
            runs=[''.join(t.text or '' for t in r.iter(f'{{{W}}}t')) for r in p.iter(f'{{{W}}}r')]
            hit=[k for k,r in enumerate(runs) if tg in r]
            print(f'p{i} 目标「{tg}」 单run命中:{hit} 共{len(runs)}个run; 相关run:', [r for r in runs if any(c in r for c in tg[:2])][:6])
print('parts:', [n for n in z.namelist() if 'comment' in n or n.endswith('.rels') or 'Content_Types' in n])

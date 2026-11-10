# 列出同一段里挂了两条及以上批注的段落（查重复批注）
import sys,zipfile,collections
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'; q=lambda t:'{%s}%s'%(W,t)
for f in sys.argv[1:]:
    z=zipfile.ZipFile(f)
    if 'word/comments.xml' not in z.namelist(): continue
    cm={c.get(q('id')):''.join(c.itertext()) for c in etree.fromstring(z.read('word/comments.xml')).iter(q('comment'))}
    doc=etree.fromstring(z.read('word/document.xml'))
    n=0; out=[]
    for p in doc.iter(q('p')):
        ids=[e.get(q('id')) for e in p.iter(q('commentRangeStart'))]
        if len(ids)>1:
            out.append((''.join(t.text or '' for t in p.iter(q('t')))[:20],[cm.get(i,'')[:45] for i in ids]))
    print('==',f.split('/')[-1],len(cm),'条批注；多批注段',len(out))
    for t,cs in out:
        print('  ',t); [print('      ',c) for c in cs]

# 删除 docx 中以指定文字开头的批注（连同正文中的批注锚点）；用法: remove_comments.py 输入 输出 "批注开头1" ["批注开头2" ...]
import sys, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
src, dst, prefixes = sys.argv[1], sys.argv[2], sys.argv[3:]
z = zipfile.ZipFile(src)
doc = etree.fromstring(z.read('word/document.xml')); cm = etree.fromstring(z.read('word/comments.xml'))
drop = []
for c in cm.findall(q('comment')):
    if any(''.join(c.itertext()).startswith(p) for p in prefixes):
        drop.append(c.get(q('id'))); cm.remove(c)
for tag in ('commentRangeStart', 'commentRangeEnd'):
    for el in list(doc.iter(q(tag))):
        if el.get(q('id')) in drop: el.getparent().remove(el)
for ref in list(doc.iter(q('commentReference'))):
    if ref.get(q('id')) in drop: r = ref.getparent(); r.getparent().remove(r)
out = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for it in z.infolist():
    if it.filename == 'word/document.xml': data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    elif it.filename == 'word/comments.xml': data = etree.tostring(cm, xml_declaration=True, encoding='UTF-8', standalone=True)
    else: data = z.read(it.filename)
    out.writestr(it, data)
out.close(); print(f'删除批注 {len(drop)} 条：编号 {drop}')

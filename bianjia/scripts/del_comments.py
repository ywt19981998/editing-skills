# 按批注编号删除批注（范围标记、引用 run、comments.xml 里的条目一起删），正文和修订不动
# 用法: del_comments.py 输入.docx 输出.docx 2,37,3 ...
import sys, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q = lambda t: '{%s}%s' % (W, t)
src, dst, ids = sys.argv[1], sys.argv[2], set(sys.argv[3].split(','))
z = zipfile.ZipFile(src)
doc = etree.fromstring(z.read('word/document.xml'))
cms = etree.fromstring(z.read('word/comments.xml'))
n = {'start': 0, 'end': 0, 'ref': 0, 'body': 0}
for tag, k in (('commentRangeStart', 'start'), ('commentRangeEnd', 'end')):
    for e in list(doc.iter(q(tag))):
        if e.get(q('id')) in ids: e.getparent().remove(e); n[k] += 1
for e in list(doc.iter(q('commentReference'))):
    if e.get(q('id')) in ids:
        r = e.getparent(); r.getparent().remove(r); n['ref'] += 1
for c in list(cms.iter(q('comment'))):
    if c.get(q('id')) in ids:
        print('删：', ''.join(c.itertext())[:50]); cms.remove(c); n['body'] += 1
print(n)
assert n['body'] == len(ids), '有编号没找到'
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
    for it in z.infolist():
        if it.filename == 'word/document.xml': data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
        elif it.filename == 'word/comments.xml': data = etree.tostring(cms, xml_declaration=True, encoding='UTF-8', standalone=True)
        else: data = z.read(it.filename)
        zo.writestr(it, data)

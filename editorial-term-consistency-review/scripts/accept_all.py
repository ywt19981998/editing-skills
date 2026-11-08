# 接受 docx 中的全部修订（删去 w:del、展开 w:ins、清除格式修订标记），批注保留；用法: accept_all.py 输入 输出
import sys, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'

src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)

def accept(root):
    n = 0
    for d in list(root.iter(q('del'), q('moveFrom'))):
        par = d.getparent()
        if par is not None and etree.QName(par).localname == 'rPr':   # 段落标记被删：合并到下一段
            p = par.getparent().getparent()
            nxt = p.getnext()
            par.remove(d)
            if nxt is not None and nxt.tag == q('p'):
                for c in [c for c in nxt if c.tag != q('pPr')]: p.append(c)
                nxt.getparent().remove(nxt)
        else:
            par.remove(d)
        n += 1
    for i in list(root.iter(q('ins'), q('moveTo'))):
        par = i.getparent()
        if etree.QName(par).localname == 'rPr': par.remove(i)        # 段落标记的插入标记
        else:
            idx = par.index(i)
            for k, c in enumerate(list(i)): par.insert(idx + k, c)
            par.remove(i)
        n += 1
    for tag in ('rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange'):
        for c in list(root.iter(q(tag))): c.getparent().remove(c); n += 1
    return n

zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
total = 0
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename in ('word/document.xml', 'word/footnotes.xml', 'word/endnotes.xml') or \
       (item.filename.startswith('word/header') or item.filename.startswith('word/footer')) and item.filename.endswith('.xml'):
        root = etree.fromstring(data); total += accept(root)
        data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    zout.writestr(item, data)
zout.close()
print(f'接受修订 {total} 处 → {dst}')

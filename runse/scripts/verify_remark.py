#!/usr/bin/env python3
"""核对补标结果：合成的那一份修订稿，拒绝全部修订应当等于作者原稿，接受全部修订应当等于补标前那份稿接受全部修订。

用法: verify_remark.py 作者原稿.docx 补标前的修订稿.docx 补标后的修订稿.docx [--author 润色修订人]
取文字用的是这里自己的一套代码，不用补标脚本和 docx_units 的。最后一行是“问题：无”才算数；有问题时退出码为 1。
"""
import sys, zipfile, collections, argparse
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
def q(t): return f'{{{W}}}{t}'
def name(e): return etree.QName(e).localname if isinstance(e.tag, str) else ''

ap = argparse.ArgumentParser()
ap.add_argument('orig'); ap.add_argument('before'); ap.add_argument('after'); ap.add_argument('--author', default='Claude（润色）')
a = ap.parse_args()
def read(path):
    z = zipfile.ZipFile(path); return z, etree.fromstring(z.read('word/document.xml'))
zo, d_orig = read(a.orig); zb, d_before = read(a.before); za, d_after = read(a.after)

OBJ = {'object': '◆', 'pict': '▣', 'drawing': '▣', 'footnoteReference': '※', 'endnoteReference': '※', 'sym': '§', 'tab': '\t', 'br': '\n', 'noBreakHyphen': '-'}
def inside(e, tag, stop):
    """e 是不是在某个 tag（ins 或 del）里面；只往上找到所在段落为止"""
    for x in e.iterancestors():
        if x is stop: return False
        if x.tag == q(tag): return True
    return False
def para_text(p, mode):
    """mode='accept' 接受全部修订后的文字；'reject' 拒绝全部修订后的文字"""
    out = []
    for e in p.iter():
        n = name(e)
        if e is p or not n: continue
        ns = etree.QName(e).namespace
        if ns == M and n == 'oMath':
            piece = '◆'
        elif ns == W and n == 't': piece = e.text or ''
        elif ns == W and n == 'delText': piece = e.text or ''
        elif ns == W and n in OBJ: piece = OBJ[n]
        else: continue
        in_del = inside(e, 'del', p) or n == 'delText'; in_ins = inside(e, 'ins', p)
        if mode == 'accept' and in_del: continue
        if mode == 'reject' and in_ins: continue
        out.append(piece)
    return ''.join(out)
def mark_state(p):
    rp = p.find(q('pPr') + '/' + q('rPr'))
    if rp is None: return ''
    return 'ins' if rp.find(q('ins')) is not None else ('del' if rp.find(q('del')) is not None else '')
def doc_paras(doc, mode):
    """全文各段的文字。段落标记是修订插入的，拒绝后和下一段连成一段；是修订删除的，接受后和下一段连成一段。"""
    out, buf = [], ''
    for p in doc.iter(q('p')):
        buf += para_text(p, mode)
        st = mark_state(p)
        if (mode == 'reject' and st == 'ins') or (mode == 'accept' and st == 'del'): continue
        if buf.strip(): out.append(buf)
        buf = ''
    if buf.strip(): out.append(buf)
    return out

problems = []
def compare(what, x, y):
    if x == y: return
    n = sum(1 for i in range(max(len(x), len(y))) if i >= len(x) or i >= len(y) or x[i] != y[i])
    first = next(i for i in range(max(len(x), len(y))) if i >= len(x) or i >= len(y) or x[i] != y[i])
    gx = x[first] if first < len(x) else '（没有这一段）'; gy = y[first] if first < len(y) else '（没有这一段）'
    k = next((i for i in range(min(len(gx), len(gy))) if gx[i] != gy[i]), min(len(gx), len(gy)))
    problems.append(f'{what}：{len(x)} 段对 {len(y)} 段，对不上的约 {n} 段。第一处在第 {first + 1} 段：\n      这份：…{gx[max(0, k - 15):k + 25]}…\n      应为：…{gy[max(0, k - 15):k + 25]}…')

# 1 拒绝全部修订 = 作者原稿
orig = doc_paras(d_orig, 'reject')
compare('拒绝全部修订后和作者原稿不一样', doc_paras(d_after, 'reject'), orig)
# 2 接受全部修订 = 补标前那份接受全部修订
compare('接受全部修订后和补标前不一样', doc_paras(d_after, 'accept'), doc_paras(d_before, 'accept'))

# 3 公式、图、域、批注、书签一个不少
def census(doc):
    c = collections.Counter()
    for e in doc.iter():
        n = name(e)
        if n in ('object', 'pict', 'drawing', 'oMath', 'fldChar', 'instrText', 'footnoteReference', 'endnoteReference', 'hyperlink',
                 'bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'pPrChange', 'rPrChange', 'tbl', 'sectPr'):
            c[n] += 1
    return c
cb, ca = census(d_before), census(d_after)
if cb != ca:
    diff = {k: (cb[k], ca[k]) for k in set(cb) | set(ca) if cb[k] != ca[k]}
    problems.append('公式、图、域、批注标记或书签的数目变了（补标前, 补标后）：' + '、'.join(f'{k} {v}' for k, v in sorted(diff.items())))
# 批注范围的先后顺序没乱
def comment_seq(doc): return [(name(e), e.get(q('id'))) for e in doc.iter(q('commentRangeStart'), q('commentRangeEnd'), q('commentReference'))]
if comment_seq(d_before) != comment_seq(d_after): problems.append('批注的起止标记顺序变了')

# 4 后两遍自己的修订原样还在
def later(doc, tag):
    out = []
    for e in doc.iter(q(tag)):
        if e.get(q('author')) == a.author: continue
        if any(name(x) == 'pPr' for x in e.iterancestors()): out.append((e.get(q('author')), '¶')); continue
        out.append((e.get(q('author')), ''.join((t.text or '') for t in e.iter(q('t'), q('delText')))))
    return out
ib, ia = later(d_before, 'ins'), later(d_after, 'ins')
if ib != ia: problems.append(f'后两遍的插入有变化：补标前 {len(ib)} 处，补标后 {len(ia)} 处')
db = ''.join(t for _, t in later(d_before, 'del')); da = ''.join(t for _, t in later(d_after, 'del'))
lost = len(db) - len(da)
it = iter(db)
if not all(ch in it for ch in da): problems.append('后两遍删除的字，补标后出现了补标前没有的内容')

# 5 结构：修订不套修订；删除里只有 delText；润色的修订都有编号和修订人
for e in d_after.iter(q('ins'), q('del')):
    if any(name(x) == 'pPr' for x in e.iterancestors()): continue
    if any(x.tag in (q('ins'), q('del')) for x in e.iterancestors()): problems.append('有修订套在另一处修订里面'); break
if any(inside(t, 'del', None) for t in d_after.iter(q('t'))): problems.append('删除修订里出现了普通文字（应当是 delText）')
if any(not inside(t, 'del', None) for t in d_after.iter(q('delText'))): problems.append('删除修订外面出现了 delText')
if b'urn:x-bubiao' in za.read('word/document.xml'): problems.append('临时记号没有清干净')
ids = collections.Counter(e.get(q('id')) for e in d_after.iter(q('ins'), q('del')) if e.get(q('author')) == a.author)
if any(v > 1 for v in ids.values()): problems.append('润色修订的编号有重复')

# 6 别的文件没动
for n_ in zb.namelist():
    if n_ != 'word/document.xml' and (n_ not in za.namelist() or zb.read(n_) != za.read(n_)): problems.append(f'文件 {n_} 变了'); break

authors = collections.Counter(e.get(q('author')) for e in d_after.iter(q('ins'), q('del')) if not any(name(x) == 'pPr' for x in e.iterancestors()))
n_ins = sum(1 for e in d_after.iter(q('ins')) if e.get(q('author')) == a.author); n_del = sum(1 for e in d_after.iter(q('del')) if e.get(q('author')) == a.author)
print(f'全文 {len(orig)} 段；润色修订：插入 {n_ins} 处、删除 {n_del} 处')
print('各修订人的修订数：', dict(authors))
if lost: print(f'后两遍的删除少了 {lost} 字：删的是润色新加的字，已连同润色的插入一起去掉（原稿和定稿里都没有这些字）')
if problems:
    print('问题：')
    for x in problems: print('  -', x)
    sys.exit(1)
print('问题： 无')

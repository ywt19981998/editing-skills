"""批注跳转链接：批注里写“2.2.1节“原文””“图3-1“原文””这类位置时，把正文里那句原文框成书签，
批注里的位置文字改成指向它的链接。在 Word 里按住 ⌘（Windows 按 Ctrl）点链接，跳过去并选中那句话。

用法：comment_links.py 输入.docx [输出.docx]（不给输出就原地改）
- 识别的位置写法：x.x节／x.x.x节、图x-x、表x-x、式（x-x）；中文编号的前言、第一章、第一章第二节、第一节一（二）等；位置后可隔几个字（章题、首段、写、作），再跟引号“……”里的原文；
- 中文编号：写了“第X章”只在该章找，“前言”只在前言找，只写“第X节”的优先在批注所在的那一章找；
- 引文里有“……”省略时，取最长的一段（至少 4 字）去找；
- “x.x.x节”只在该节（含下级小节）里找，图、表、式在全文找；找不到或位置文字跨 run 的不加链接，最后列出来；
- 已加过链接的批注跳过，可重复运行。只加书签和链接，不改正文文字。
"""
import sys, re, copy, zipfile, os, shutil, tempfile
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; q = lambda t: f'{{{W}}}{t}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
CN = '一二三四五六七八九十'
# 中文编号的位置：前言、第一章、第一章第二节、第一节一（二）、第六章章题等
CNLABEL = rf'(?:前言|后记|第[{CN}]+章(?:第[{CN}]+节)?(?:[{CN}]+(?:（[{CN}]+）)?)?|第[{CN}]+节(?:[{CN}]+(?:（[{CN}]+）)?)?)'
LABEL = r'(\d+\.\d+(?:\.\d+)*节|[图表]\s?\d+[-－.]\d+|式\s?[（(]\d+[-－.]\d+[a-z]?[)）]|' + CNLABEL + ')'
# 位置和引号之间可隔几个字（“章题”“首段”“写”“作”“小标题是”）
REF = re.compile(LABEL + r'[^“”，。；、]{0,5}“([^”]{2,})”')
CHAP = re.compile(rf'^\s*(第[{CN}]+章|前\s*言|后\s*记)(?![^\s　]{{20,}})')


def cn2int(s):
    if '十' in s:
        a, b = s.split('十', 1)
        return (CN.index(a) + 1 if a else 1) * 10 + (CN.index(b) + 1 if b else 0)
    return CN.index(s) + 1
HEAD = re.compile(r'^(\d+\.\d+(?:\.\d+)*)[\s　]')


def runs_of(p):
    # 现文的 run：直接在段里或在插入修订里；删除修订里的字在 delText，不算
    return [r for r in p.iter(q('r')) if r.find(q('t')) is not None and r.getparent().tag != q('del')]


def ptext(p):
    return ''.join(r.find(q('t')).text or '' for r in runs_of(p))


def split_run(r, k):
    """把 run 在第 k 个字处拆开，返回后一半"""
    t = r.find(q('t')); a, b = t.text[:k], t.text[k:]
    r2 = copy.deepcopy(r); t.text = a; t.set(XS, 'preserve')
    t2 = r2.find(q('t')); t2.text = b; t2.set(XS, 'preserve'); r.addnext(r2); return r2


def bookmark(p, phrase, name, bid):
    rs = runs_of(p); full = ''.join(r.find(q('t')).text or '' for r in rs)
    i = full.find(phrase); j = i + len(phrase); pos = 0; first = last = None
    for r in rs:
        s = r.find(q('t')).text or ''; a, b = pos, pos + len(s); pos = b
        if b <= i or a >= j: continue
        if a < i: r = split_run(r, i - a); a = i
        if b > j: split_run(r, j - a)
        first = first if first is not None else r; last = r
    st = etree.Element(q('bookmarkStart')); st.set(q('id'), str(bid)); st.set(q('name'), name)
    en = etree.Element(q('bookmarkEnd')); en.set(q('id'), str(bid))
    first.addprevious(st); last.addnext(en)


def link(t, start, end, anchor):
    """把批注里某个 w:t 的 [start, end) 这几个字包成指向书签的链接"""
    r = t.getparent(); text = t.text
    t.text = text[:start]; t.set(XS, 'preserve')
    h = etree.Element(q('hyperlink')); h.set(q('anchor'), anchor); h.set(q('history'), '1')
    r2 = copy.deepcopy(r)
    for x in list(r2):
        if x.tag != q('rPr'): r2.remove(x)
    rpr = r2.find(q('rPr'))
    if rpr is None: rpr = etree.Element(q('rPr')); r2.insert(0, rpr)
    for tag, val in (('color', '0563C1'), ('u', 'single')):
        el = etree.SubElement(rpr, q(tag)); el.set(q('val'), val)
    nt = etree.SubElement(r2, q('t')); nt.text = text[start:end]; nt.set(XS, 'preserve'); h.append(r2)
    r3 = copy.deepcopy(r)
    for x in list(r3):
        if x.tag == q('t'): r3.remove(x)
    t3 = etree.SubElement(r3, q('t')); t3.text = text[end:]; t3.set(XS, 'preserve')
    r.addnext(h); h.addnext(r3)
    return t3


def main(src, dst):
    z = zipfile.ZipFile(src); files = {n: z.read(n) for n in z.namelist()}
    if 'word/comments.xml' not in files: print('批注链接：没有批注'); return
    doc = etree.fromstring(files['word/document.xml']); cx = etree.fromstring(files['word/comments.xml'])
    body = doc.find(q('body'))
    paras = [p for p in body.iter(q('p'))]
    # 每段所在小节：章首节目（下一段紧跟标题，或后一个标题编号不大于它）不算小节起点
    texts = [ptext(p) for p in paras]
    heads = [(i, tuple(map(int, m.group(1).split('.'))), m.group(1)) for i, t in enumerate(texts) for m in [HEAD.match(t.strip())] if m]
    real = {}
    for k, (i, num, s) in enumerate(heads):
        nxt = next((j for j in range(i + 1, len(texts)) if texts[j].strip()), None)
        if nxt is not None and HEAD.match(texts[nxt].strip()): continue
        if k + 1 < len(heads) and heads[k + 1][1] <= num: continue
        real[i] = s
    sec, cur = [], ''
    for i in range(len(paras)):
        cur = real.get(i, cur); sec.append(cur)
    # 每段属于哪一章（中文章题“第一章”记作 1，前言、后记各记一档）
    chap, cc = [], None
    for i, tx in enumerate(texts):
        m = CHAP.match(tx) if len(tx.strip()) < 40 else None
        if m:
            k = re.sub(r'\s', '', m.group(1)); cc = cn2int(k[1:-1]) if k.startswith('第') else k
        chap.append(cc)
    # 批注挂在哪一段
    anchor = {}
    for i, p in enumerate(paras):
        for s in p.iter(q('commentRangeStart')): anchor.setdefault(s.get(q('id')), i)
    bid = max([int(b.get(q('id'))) for b in doc.iter(q('bookmarkStart'))] + [0]) + 1
    done = miss = 0; report = []
    for c in cx.iter(q('comment')):
        if c.find('.//' + q('hyperlink')) is not None: continue
        cid = c.get(q('id')); n = 0
        for t in list(c.iter(q('t'))):
            while t is not None and t.text:
                m = REF.search(t.text)
                if not m: break
                label, quote = m.group(1), m.group(2)
                phrase = max(re.split(r'…+|\.{3,}', quote), key=len).strip()
                hit = None
                if len(phrase) >= 4:
                    cands = [i for i in range(len(paras)) if phrase in texts[i]]
                    if not cands:
                        # 批注里套引号常写成单引号‘’，正文是双引号“”；换成双引号再找
                        alt = phrase.replace('‘', '“').replace('’', '”')
                        cands = [i for i in range(len(paras)) if alt in texts[i]]
                        if cands: phrase = alt
                    allc = list(cands)
                    if re.match(r'\d', label):
                        want = label[:-1] if label.endswith('节') else None
                        cands = [i for i in cands if want is None or sec[i] == want or sec[i].startswith(want + '.')]
                    else:
                        cm = re.match(rf'第([{CN}]+)章', label)
                        if cm: cands = [i for i in cands if chap[i] == cn2int(cm.group(1))]
                        elif label in ('前言', '后记'): cands = [i for i in cands if chap[i] == label]
                        else:
                            # 只写“第一节……”时，优先批注所在的那一章
                            own = chap[anchor[cid]] if cid in anchor else None
                            cands = [i for i in cands if chap[i] == own] or cands
                    # 按位置筛不到、但全书只有一处的，也认（批注里的节号偶有写偏）
                    hit = cands[0] if cands else (allc[0] if len(allc) == 1 else None)
                if hit is None:
                    miss += 1; report.append(f'批注{int(cid) + 1}：{label}“{quote[:20]}”没找到')
                    rest = t.text[m.end():]
                    # 跳过这一处，继续找同一段文字里后面的位置
                    nxt = REF.search(rest)
                    if not nxt: break
                    t = link_skip(t, m.end()); continue
                name = f'_批注{cid}_{n}'; n += 1
                bookmark(paras[hit], phrase, name, bid); bid += 1
                texts[hit] = ptext(paras[hit])
                t = link(t, m.start(1), m.end(1), name); done += 1
    files['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    files['word/comments.xml'] = etree.tostring(cx, xml_declaration=True, encoding='UTF-8', standalone=True)
    fd, tmp = tempfile.mkstemp(suffix='.docx'); os.close(fd)
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as o:
        for name_, data in files.items(): o.writestr(name_, data)
    shutil.move(tmp, dst)
    print(f'批注链接：加了 {done} 处，没对上 {miss} 处')
    for r in report: print('  ' + r)


def link_skip(t, k):
    """没找到的位置不加链接：把 w:t 在 k 处拆开，后半段放进同格式的新 run，继续往后找"""
    r = t.getparent(); text = t.text
    t.text = text[:k]; t.set(XS, 'preserve')
    r2 = copy.deepcopy(r)
    for x in list(r2):
        if x.tag == q('t'): r2.remove(x)
    t2 = etree.SubElement(r2, q('t')); t2.text = text[k:]; t2.set(XS, 'preserve'); r.addnext(r2)
    return t2


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else sys.argv[1])

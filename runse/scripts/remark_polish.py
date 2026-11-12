#!/usr/bin/env python3
"""补标：把润色的改动标回“后两遍的修订稿”，合成一份修订稿。

后两遍（一致性审查、通读）的修订是写在润色后底稿上的，稿子里只看得到它们自己的改动。本脚本在这份稿上把润色的改动补标出来：
  润色加的字（现在是普通正文）  → 标成插入，修订人“Claude（润色）”
  润色删的字（稿子里已经没有）  → 按原稿的字和格式放回原位，标成删除
  润色加的字又被后两遍删掉的    → 原稿没有、定稿也没有，直接去掉，不留痕迹
后两遍自己的修订、批注、格式修订一概不动。结果：拒绝全部修订 = 作者原稿，接受全部修订 = 定稿。

用法: remark_polish.py [--author 修订人] 作者原稿.docx 后两遍的修订稿.docx 输出.docx 正文.txt 终稿.json [更多.json]
  正文.txt、终稿.json 是润色时的那两份（工作目录里）。后两遍的修订稿可以已经跑过体例脚本（分项拆段、顶格等）。
  有哪一段标不回去，整份都不写，列出这些段。写完必须再跑 verify_remark.py 核对。
"""
import sys, os, zipfile, re, json, copy, datetime, difflib, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
import jieba
jieba.setLogLevel(60)
from docx_units import XSP, q, ln, nsp, W, load, has_revisions, plain, clean_para, build_units, unit_text, body_rows, read_slines, align

ap = argparse.ArgumentParser()
ap.add_argument('orig'); ap.add_argument('merged'); ap.add_argument('dst'); ap.add_argument('txt'); ap.add_argument('results', nargs='+')
ap.add_argument('--author', default='Claude（润色）'); ap.add_argument('--terms')
a = ap.parse_args()
AUTHOR = a.author
DATE = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
TMP = 'urn:x-bubiao'                      # 干活时临时挂的记号，写文件前全部去掉
LDEL = f'{{{TMP}}}ldel'                   # 挂在 run 上：这个 run 原来在后两遍的第几个删除里
LINS = f'{{{TMP}}}lins'                   # 占位元素：后两遍的第几个插入原来在这里

zo_, doc_o = load(a.orig)
if has_revisions(doc_o): sys.exit('停：作者原稿里有没接受的修订，这不是润色用的那份原稿。')
zin, doc = load(a.merged)
if any(e.get(q('author')) == AUTHOR for e in doc.iter(q('ins'), q('del'))):
    sys.exit(f'停：这份稿里已经有“{AUTHOR}”的修订，是补标过的；要重做请从没补标的那份开始。')
if next(doc.iter(q('moveFrom'), q('moveTo')), None) is not None: sys.exit('停：稿子里有“移动”修订，脚本处理不了。')
if a.terms:
    for w in open(a.terms, encoding='utf-8'):
        if w.strip(): jieba.add_word(w.strip())
_ids = [int(e.get(q('id'))) for e in doc.iter() if (e.get(q('id')) or '').isdigit()]
_rid = [max(_ids + [0]) + 2000]
def nid():
    _rid[0] += 1; return str(_rid[0])
def mark(tag):
    e = etree.Element(q(tag)); e.set(q('id'), nid()); e.set(q('author'), AUTHOR); e.set(q('date'), DATE); return e

class Unwritable(Exception): pass

# ---------- 润色前后的文字 ----------
slines = read_slines(a.txt)
res = {}
for f in a.results:
    for it in json.load(open(f, encoding='utf-8')):
        if isinstance(it, dict) and it.get('S'): res[it['S']] = it
def polished(sid, text):
    it = res.get(sid, {})
    new = (it.get('new') or '').strip() if it.get('changed') else ''
    return new if new and new != text.strip() else None
body = [(sid, text.strip(), polished(sid, text)) for sid, tag, text in slines if not tag]
para_o = align(body_rows(doc_o), slines)          # 段号 → 作者原稿里的段落

# ---------- 把后两遍的修订暂时“拒绝”掉，露出润色后底稿的样子（记号留着，最后恢复） ----------
ldel_attrs, lins_els = [], []
def strip_later(children_owner):
    """children_owner：一个 w:p。把它直接下面的后两遍删除拆开（里面的 run 挂上记号、删除的字变回普通字），
       后两遍的插入换成占位元素。"""
    for ch in list(children_owner):
        if nsp(ch) != W: continue
        tag = ln(ch)
        if tag == 'del':
            k = len(ldel_attrs); ldel_attrs.append(dict(ch.attrib)); at = ch
            for x in list(ch):
                if ln(x) == 'r':
                    x.set(LDEL, str(k))
                    for t in x.iter(q('delText')): t.tag = q('t')
                at.addnext(x); at = x
            children_owner.remove(ch)
        elif tag == 'ins':
            k = len(lins_els); m = etree.Element(LINS); m.set('k', str(k)); m.tail = ch.tail
            children_owner.replace(ch, m); ch.tail = None; lins_els.append(ch)
def mark_inserted(p):
    """这一段的段落标记是不是后两遍插入的（拆段、插入新段）：是的话，拒绝修订后它和下一段是同一段"""
    rp = p.find(q('pPr') + '/' + q('rPr'))
    return rp is not None and rp.find(q('ins')) is not None
def kids(group):
    return [ch for p in group for ch in p if not (nsp(ch) == W and ln(ch) == 'pPr')]
def base_text(group):
    """一组段落（通常就一段）去掉后两遍修订后的文字。在副本上算，不动原件。"""
    g = [copy.deepcopy(p) for p in group]
    n1, n2 = len(ldel_attrs), len(lins_els)
    for p in g: strip_later(p); clean_para(p)
    del ldel_attrs[n1:]; del lins_els[n2:]
    try: return ''.join(unit_text(u) for u in build_units(kids(g)))
    except ValueError: return None

groups, cur = [], []
for el in doc.find(q('body')):
    if ln(el) != 'p': continue
    cur.append(el)
    if not mark_inserted(el): groups.append(cur); cur = []
if cur: groups.append(cur)
texts = [base_text(g) for g in groups]

# 正文各段按顺序在稿子里找到自己（去掉后两遍修订后，应当正好是润色后的文字）
where, gi, lost = {}, 0, []
for sid, old, new in body:
    want = new or old; j = gi
    while j < len(groups) and (texts[j] or '').strip() != want: j += 1
    if j == len(groups): lost.append(sid); continue
    where[sid] = j; gi = j + 1
if lost:
    sys.exit(f'停：有 {len(lost)} 段在这份稿里对不上（去掉后两遍的修订后，应当等于润色后的文字）：{"、".join(lost[:12])}'
             + '。\n确认这份修订稿是在这一版润色后底稿上做的；润色终稿改过的话，后两遍要在新底稿上重出。')

# ---------- 按词对齐的改动（和写回脚本同一套切法），两边的位置都留着 ----------
PUNC = set('，。；：、！？（）“”‘’《》—…,.;:!?()[] \t')
def toks(s): return [w for w in jieba.cut(s, HMM=False) if w]
def offs(ts):
    o = [0]
    for w in ts: o.append(o[-1] + len(w))
    return o
def hunks(old, new):
    """old：稿子里现在的字（润色后）；new：原稿的字。返回 [[a0, a1, b0, b1]]：old[a0:a1] 是润色加的，new[b0:b1] 是润色删的"""
    ta, tb = toks(old), toks(new); oa, ob = offs(ta), offs(tb); sa, sb = set(oa), set(ob)
    out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes():
        if op == 'equal': continue
        a0, a1, b0, b1 = oa[i1], oa[i2], ob[j1], ob[j2]
        k = 0
        while k < min(a1 - a0, b1 - b0) and old[a0 + k] == new[b0 + k]: k += 1
        while k and not (a0 + k in sa and b0 + k in sb): k -= 1
        a0 += k; b0 += k
        k = 0
        while k < min(a1 - a0, b1 - b0) and old[a1 - 1 - k] == new[b1 - 1 - k]: k += 1
        while k and not (a1 - k in sa and b1 - k in sb): k -= 1
        a1 -= k; b1 -= k
        if a0 == a1 and b0 == b1: continue
        out.append([a0, a1, b0, b1])
    merged = []
    for h in out:
        if merged:
            gap = old[merged[-1][1]:h[0]]
            if len(gap) == 1 and gap not in PUNC and new[merged[-1][3]:h[2]] == gap:
                merged[-1] = [merged[-1][0], h[1], merged[-1][2], h[3]]; continue
        merged.append(h)
    return merged

def run_text(r): return r.find(q('t')).text or ''
def split(seg, pos):
    acc = 0
    for i, r in enumerate(seg):
        t = r.find(q('t')); n = len(t.text or '')
        if acc < pos < acc + n:
            k = pos - acc; r2 = copy.deepcopy(r); r2.tail = None
            t.text, r2.find(q('t')).text = t.text[:k], t.text[k:]
            for x in (t, r2.find(q('t'))): x.set(XSP, 'preserve')
            r.addnext(r2); seg.insert(i + 1, r2); return
        acc += n
def slice_runs(seg, x, y):
    """原稿一块文字里第 x～y 个字，按原来的 run 和格式抄一份出来"""
    out, acc = [], 0
    for r in seg:
        s = run_text(r); n = len(s); lo, hi = max(x, acc), min(y, acc + n)
        if lo < hi:
            r2 = copy.deepcopy(r); r2.tail = None; t = r2.find(q('t')); t.text = s[lo - acc:hi - acc]; t.set(XSP, 'preserve'); out.append(r2)
        acc += n
    return out

stat = {'ins': 0, 'del': 0, 'drop': 0, 'drop_chars': 0, 'hunks': 0}
dropped = []
def apply_hunk(seg, a0, a1, o_runs, left, right):
    """seg：这一块里还没动过的文字 run（稿子里的）；[a0, a1) 是润色加的字；o_runs：润色删掉的字（从原稿抄来的 run）"""
    split(seg, a0); split(seg, a1)
    acc, mid, prev, nxt = 0, [], None, None
    for r in seg:
        n = len(run_text(r))
        if not n: continue
        if a0 < a1 and a0 <= acc and acc + n <= a1: mid.append(r)
        if acc + n == a0: prev = r
        if acc == a1 and nxt is None: nxt = r
        acc += n
    if o_runs:
        d = mark('del')
        for r in o_runs:
            r.find(q('t')).tag = q('delText'); d.append(r)
        if mid: mid[0].addprevious(d)
        elif prev is not None: prev.addnext(d)
        elif nxt is not None: nxt.addprevious(d)
        elif left is not None: left.addnext(d)
        elif right is not None: right.addprevious(d)
        else: raise Unwritable('找不到放回原文的位置')
        stat['del'] += 1
    wrap = None
    for r in mid:
        if r.get(LDEL) is not None:                       # 润色加的字，后两遍又删了：两边都没有它，去掉
            stat['drop'] += 1; stat['drop_chars'] += len(run_text(r)); dropped.append(run_text(r))
            r.getparent().remove(r); wrap = None; continue
        if wrap is not None and wrap.getnext() is r: wrap.append(r)
        else:
            wrap = mark('ins'); r.addprevious(wrap); wrap.append(r); stat['ins'] += 1
    for r in mid: seg.remove(r)

def segments(units):
    """按不能动的块切开：返回 (各块的文字 run 列表, 各块两边的元素, 不能动的块的样子)"""
    segs, bounds, cur_seg, left, atoms = [], [], [], None, []
    for u in units:
        if u['k'] == 'atom' and u['disp']:
            segs.append(cur_seg); bounds.append((left, u['first'])); cur_seg = []; left = u['last']; atoms.append(u['disp'])
        elif u['k'] == 'txt': cur_seg.append(u['run'])
    segs.append(cur_seg); bounds.append((left, None))
    return segs, bounds, atoms

def remark(sid, group, old, new):
    for p in group: strip_later(p); clean_para(p)
    left_over = [e for p in group for e in p.iter(q('ins'), q('del'))
                 if not any(nsp(x) == W and ln(x) == 'pPr' for x in e.iterancestors())]
    if left_over: raise Unwritable('后两遍的修订嵌在超链接之类的块里，脚本够不着')
    try: units = build_units(kids(group))
    except ValueError as e: raise Unwritable(str(e))
    full = ''.join(unit_text(u) for u in units)
    if full.strip() != new: raise Unwritable('去掉后两遍的修订后，和润色后的文字对不上')
    po = copy.deepcopy(para_o[sid]); clean_para(po)
    o_units = build_units(po); o_full = ''.join(unit_text(u) for u in o_units)
    if o_full.strip() != old: raise Unwritable('作者原稿这一段和正文.txt 不一致')
    segs, bounds, atoms = segments(units); o_segs, _, o_atoms = segments(o_units)
    if atoms != o_atoms: raise Unwritable('公式、图等不能动的块，和原稿对不上')
    for k, (seg, o_seg) in enumerate(zip(segs, o_segs)):
        s_new = ''.join(run_text(r) for r in seg); s_old = ''.join(run_text(r) for r in o_seg)
        if s_new == s_old: continue
        hs = hunks(s_new, s_old); stat['hunks'] += len(hs)
        for a0, a1, b0, b1 in reversed(hs):
            apply_hunk(seg, a0, a1, slice_runs(o_seg, b0, b1), bounds[k][0], bounds[k][1])

def restore(group):
    """把后两遍的修订原样恢复：挂着记号的 run 重新包进删除，占位元素换回原来的插入"""
    for p in group:
        seen = set(); wrap = None; wrap_k = None
        for r in [x for x in p.iter(q('r')) if x.get(LDEL) is not None]:
            k = int(r.get(LDEL)); del r.attrib[LDEL]
            for t in r.iter(q('t')): t.tag = q('delText')
            if wrap is not None and wrap_k == k and wrap.getnext() is r: wrap.append(r); continue
            wrap = etree.Element(q('del'))
            for kk, vv in ldel_attrs[k].items(): wrap.set(kk, vv)
            if k in seen: wrap.set(q('id'), nid())            # 一个删除被拆成几截时，后面几截换新编号
            seen.add(k); wrap_k = k; r.addprevious(wrap); wrap.append(r)
        for m in list(p.iter(LINS)):
            el = lins_els[int(m.get('k'))]; el.tail = m.tail; m.getparent().replace(m, el)

# ---------- 逐段补标 ----------
failed = []; n_p = 0
for sid, old, new in body:
    if not new: continue
    group = groups[where[sid]]; backup = [copy.deepcopy(p) for p in group]
    n1, n2 = len(ldel_attrs), len(lins_els); snap = dict(stat); nd = len(dropped)
    try:
        remark(sid, group, old, new); restore(group); n_p += 1
    except Unwritable as e:
        failed.append(f'{sid}　{e}')
        for p, b in zip(group, backup): p.getparent().replace(p, b)
        del ldel_attrs[n1:]; del lins_els[n2:]; stat.update(snap); del dropped[nd:]
if failed:
    print(f'有 {len(failed)} 段标不回去，整份没有写：')
    for x in failed: print('  -', x)
    sys.exit(1)

xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
if TMP.encode() in xml:                                    # 临时记号没清干净就不写
    etree.cleanup_namespaces(doc); xml = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    if TMP.encode() in xml: sys.exit('内部错误：临时记号没有清干净，没有写文件。')
zout = zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    zout.writestr(item, xml if item.filename == 'word/document.xml' else zin.read(item.filename))
zout.close()
print(f'补标 {n_p} 段、{stat["hunks"]} 处改动：标成插入 {stat["ins"]} 处，放回并标成删除 {stat["del"]} 处，修订人“{AUTHOR}”')
if stat['drop']:
    print(f'润色加的字被后两遍删掉的 {stat["drop"]} 处、{stat["drop_chars"]} 字，已直接去掉：' + '、'.join(f'“{x}”' for x in dropped[:8]) + ('……' if len(dropped) > 8 else ''))
print('→', a.dst)

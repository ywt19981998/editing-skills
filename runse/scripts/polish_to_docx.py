#!/usr/bin/env python3
"""把润色终稿写成 Word 修订（删除线＋插入）：公式、图、脚注号、域（交叉引用、文献序号）原样不动，只改它们之间的文字。

用法: polish_to_docx.py [--dry] [--terms 术语表.txt] [--author 修订人] 原稿.docx 正文.txt 输出.docx 终稿.json [更多.json]
  原稿须没有未接受的修订；正文.txt 是准备脚本从这份原稿导出的。--dry 只打印每段怎么改，不写文件。
  有哪一段写不回去（不能动的块被改了等），整份都不写，列出这些段：先在终稿里把它们改好或退回原文再跑。
写完必须再跑 verify_polish_docx.py 核对。
"""
import sys, os, zipfile, re, json, copy, datetime, difflib, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
import jieba
jieba.setLogLevel(60)
from docx_units import M, XSP, W, q, ln, nsp, load, has_revisions, plain, clean_para, build_units, unit_text, body_rows, read_slines, align

ap = argparse.ArgumentParser()
ap.add_argument('src'); ap.add_argument('txt'); ap.add_argument('dst'); ap.add_argument('results', nargs='+')
ap.add_argument('--dry', action='store_true'); ap.add_argument('--terms'); ap.add_argument('--author', default='Claude（润色）')
a = ap.parse_args()
DRY, AUTHOR = a.dry, a.author
DATE = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')

zin, doc = load(a.src)
if has_revisions(doc): sys.exit('停：原稿里还有没接受的修订，不能在上面写润色修订。')
if a.terms:
    for w in open(a.terms, encoding='utf-8'):
        if w.strip(): jieba.add_word(w.strip())
_ids = [int(e.get(q('id'))) for e in doc.iter() if (e.get(q('id')) or '').isdigit()]
_rid = [max(_ids + [0]) + 1000]
def nid():
    _rid[0] += 1; return str(_rid[0])

class Unwritable(Exception): pass

def merge_runs(p):
    """合并段内相邻、格式完全相同的普通 run（与 revise_docx.py 做法一致）"""
    prev = None
    for r in list(p.findall(q('r'))):
        if prev is not None and plain(prev) and plain(r) and r.getprevious() is prev:
            x, y = prev.find(q('rPr')), r.find(q('rPr'))
            if (etree.tostring(x) if x is not None else b'') == (etree.tostring(y) if y is not None else b''):
                prev.find(q('t')).text = (prev.find(q('t')).text or '') + (r.find(q('t')).text or '')
                prev.find(q('t')).set(XSP, 'preserve'); p.remove(r); continue
        prev = r

def run_text(r): return r.find(q('t')).text or ''
INFIELD = set()      # 落在域里的文字 run

# ---------- 段号对到段落 ----------
slines = read_slines(a.txt)
para_of = align(body_rows(doc), slines)
old_of = {sid: text for sid, tag, text in slines if not tag}
res = {}
for f in a.results:
    for it in json.load(open(f, encoding='utf-8')):
        if isinstance(it, dict) and it.get('S'): res[it['S']] = it
changed = [sid for sid, tag, text in slines if not tag and res.get(sid, {}).get('changed')
           and (res[sid].get('new') or '').strip() and res[sid]['new'].strip() != text.strip()]

# ---------- 一块文字里的改动：按词对齐，只在两边都是词边界的地方切 ----------
PUNC = set('，。；：、！？（）“”‘’《》—…,.;:!?()[] \t')
def toks(s): return [w for w in jieba.cut(s, HMM=False) if w]
def offs(ts):
    o = [0]
    for w in ts: o.append(o[-1] + len(w))
    return o
def hunks(old, new):
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
        out.append([a0, a1, new[b0:b1]])
    # 两处改动之间只隔一个非标点的字，并成一处，免得修订碎成一个字一个字的
    merged = []
    for h in out:
        if merged:
            gap = old[merged[-1][1]:h[0]]
            if len(gap) == 1 and gap not in PUNC:
                merged[-1] = [merged[-1][0], h[1], merged[-1][2] + gap + h[2]]; continue
        merged.append(h)
    return merged

# ---------- 把一处改动写成修订 ----------
def has_vert(r):
    rp = r.find(q('rPr'))
    return rp is not None and rp.find(q('vertAlign')) is not None
def cjk_n(s): return len(re.findall(r'[一-鿿]', s))
def split(seg, pos):
    acc = 0
    for i, r in enumerate(seg):
        t = r.find(q('t')); n = len(t.text or '')
        if acc < pos < acc + n:
            k = pos - acc; r2 = copy.deepcopy(r); r2.tail = None
            t.text, r2.find(q('t')).text = t.text[:k], t.text[k:]
            for x in (t, r2.find(q('t'))): x.set(XSP, 'preserve')
            r.addnext(r2); seg.insert(i + 1, r2)
            if r in INFIELD: INFIELD.add(r2)
            return
        acc += n
def mark(tag):
    e = etree.Element(q(tag)); e.set(q('id'), nid()); e.set(q('author'), AUTHOR); e.set(q('date'), DATE); return e
def apply_hunk(p, seg, x, y, new_text, left, right, fallback):
    """seg：这一块里还没动过的文字 run；x、y：要删掉的范围（块内字符位置）；left/right：块两边不动的元素"""
    split(seg, x); split(seg, y)
    acc, dels, prev, nxt = 0, [], None, None
    for r in seg:
        n = len(run_text(r))
        if not n: continue
        if x < y and x <= acc and acc + n <= y: dels.append(r)
        if acc + n == x: prev = r
        if acc == y and nxt is None: nxt = r
        acc += n
    if dels: base = max(dels, key=lambda r: (not has_vert(r), cjk_n(run_text(r)), len(run_text(r))))
    else:
        c = [r for r in (prev, nxt) if r is not None] or [fallback]
        base = max(c, key=lambda r: (not has_vert(r), r is prev))
    if base is None: raise Unwritable('这一段没有可以借用格式的文字')
    base = copy.deepcopy(base); base.tail = None
    for r in dels:
        if r in INFIELD and re.search(r'[^，。；：、,.;:\s]', run_text(r)): warns.append(f'{cur_sid[0]} 域里的文字“{run_text(r)}”被改了，交叉引用或引注的链接可能丢，要人看')
    for r in dels:
        d = mark('del'); r.addprevious(d); d.append(r); r.find(q('t')).tag = q('delText')
    if new_text:
        ins = mark('ins'); base.find(q('t')).text = new_text; base.find(q('t')).set(XSP, 'preserve'); ins.append(base)
        if dels: dels[-1].getparent().addnext(ins)
        elif prev is not None: prev.addnext(ins)
        elif nxt is not None: nxt.addprevious(ins)
        elif left is not None: left.addnext(ins)
        elif right is not None: right.addprevious(ins)
        else: p.append(ins)
        out_of_field(p, ins, new_text)
    for r in dels: seg.remove(r)

def visible(el):
    """段落的直接子元素有没有看得见的内容（已删除的不算）"""
    if nsp(el) == M: return True
    if nsp(el) != W or ln(el) == 'del': return False
    if ln(el) == 'r': return bool(''.join(x.text or '' for x in el if ln(x) == 't')) or any(el.find(q(k)) is not None for k in ('object', 'pict', 'drawing', 'tab'))
    return ln(el) in ('ins', 'hyperlink', 'smartTag', 'fldSimple', 'sdt')
def fld_type(el):
    fc = el.find(q('fldChar')) if nsp(el) == W and ln(el) == 'r' else None
    return fc.get(q('fldCharType')) if fc is not None else None
def out_of_field(p, ins, new_text):
    """新加的字如果落在域（如文献引注）显示文字的头上或尾上，挪到域外：域一更新，域里的字会被冲掉"""
    kids = list(p); idx = kids.index(ins); stack = []
    for ch in kids[:idx]:
        ty = fld_type(ch)
        if ty == 'begin': stack.append(ch)
        elif ty == 'end' and stack: stack.pop()
    if not stack: return
    begin, depth, end = stack[0], len(stack), None
    for ch in kids[idx + 1:]:
        ty = fld_type(ch)
        if ty == 'begin': depth += 1
        elif ty == 'end':
            depth -= 1
            if depth == 0: end = ch; break
    if end is None: return
    if not any(visible(c) for c in kids[idx + 1:kids.index(end)]): end.addnext(ins)
    elif not any(visible(c) for c in kids[kids.index(begin) + 1:idx]): begin.addprevious(ins)
    else: warns.append(f'{cur_sid[0]} 新加的“{new_text}”落在域的显示文字中间，要人看')

def show(old, hs):
    out, pos = [], 0
    for x, y, n in hs:
        out.append(old[pos:x])
        if y > x: out.append('[-' + old[x:y] + '-]')
        if n: out.append('{+' + n + '+}')
        pos = y
    out.append(old[pos:]); return ''.join(out)

def write_para(sid, p):
    """把一段的润色写成修订，返回 (改动处数, 空跑时的标注文字)"""
    clean_para(p); merge_runs(p)
    try: units = build_units(p, INFIELD)
    except ValueError as e: raise Unwritable(str(e))
    full = ''.join(unit_text(u) for u in units)
    if full.strip() != old_of[sid]: raise Unwritable('原稿段落和正文.txt 不一致')
    lead, trail = full[:len(full) - len(full.lstrip())], full[len(full.rstrip()):]
    new_full = lead + res[sid]['new'].strip() + trail
    atoms = [u for u in units if u['k'] == 'atom' and u['disp']]
    # 在润色后的文字里依次找到每个不动的块，块与块之间就是要比对的文字
    cuts, pos = [], 0
    for u in atoms:
        i = new_full.find(u['disp'], pos)
        if i < 0: raise Unwritable(f'润色后的文字里找不到原有的“{u["disp"].strip() or "制表符"}”（公式、引注或不能拆的一块被动了）')
        cuts.append((i, i + len(u['disp']))); pos = i + len(u['disp'])
    # 交叉核对：整段按字对齐，不动的块应当整块落在“没变”的区间里，位置和上面找到的一致
    blocks = difflib.SequenceMatcher(None, full, new_full, autojunk=False).get_matching_blocks(); o = 0; k = 0
    for u in units:
        t = unit_text(u)
        if u['k'] == 'atom' and u['disp']:
            hit = [bl for bl in blocks if bl.a <= o and o + len(t) <= bl.a + bl.size]
            if not hit or hit[0].b + (o - hit[0].a) != cuts[k][0]: warns.append(f'{sid} “{t}” 的位置两种算法不一致，要人看')
            k += 1
        o += len(t)
    new_segs, pos = [], 0
    for i0, i1 in cuts: new_segs.append(new_full[pos:i0]); pos = i1
    new_segs.append(new_full[pos:])
    old_segs, bounds, cur_seg, left = [], [], [], None      # 每块：run 列表、左边元素、右边元素
    for u in units:
        if u['k'] == 'atom' and u['disp']:
            old_segs.append(cur_seg); bounds.append((left, u['first'])); cur_seg = []; left = u['last']
        elif u['k'] == 'txt': cur_seg.append(u['run'])
    old_segs.append(cur_seg); bounds.append((left, None))
    if len(old_segs) != len(new_segs): raise Unwritable('文字块数对不上')
    all_txt = [u['run'] for u in units if u['k'] == 'txt']
    lines = []; cnt = 0
    for k, (seg, new_s) in enumerate(zip(old_segs, new_segs)):
        old_s = ''.join(run_text(r) for r in seg)
        hs = hunks(old_s, new_s) if old_s != new_s else []
        # 空格不改：只增删空格的改动丢掉
        hs = [h for h in hs if re.sub(r'\s', '', old_s[h[0]:h[1]]) != re.sub(r'\s', '', h[2])]
        lines.append(show(old_s, hs)); cnt += len(hs)
        if DRY: continue
        fb = (seg or all_txt or [None])[0]
        if seg == [] and all_txt:                               # 这一块原来没有文字：格式取它前面最近的文字
            Lb = bounds[k][0]; before = [r for r in all_txt if Lb is not None and r in list(Lb.itersiblings(preceding=True))]
            fb = before[-1] if before else all_txt[0]
        for x, y, n in reversed(hs): apply_hunk(p, seg, x, y, n, bounds[k][0], bounds[k][1], fb)
    marked = ''
    for k, s in enumerate(lines): marked += s + (atoms[k]['disp'] if k < len(atoms) else '')
    return cnt, marked.strip()

n_h = n_p = 0; warns = []; cur_sid = ['']; failed = []
for sid in changed:
    cur_sid[0] = sid; p = para_of[sid]; backup = copy.deepcopy(p)
    try:
        cnt, marked = write_para(sid, p)
    except Unwritable as e:
        failed.append(f'{sid}　{e}')
        p.getparent().replace(p, backup); continue
    n_h += cnt; n_p += 1
    if DRY: print(f'{sid}（{cnt} 处）{marked}\n')
for w in warns: print('注意：', w)
if failed:
    print(f'有 {len(failed)} 段写不回去，整份没有写：')
    for x in failed: print('  -', x)
    print('先在终稿里把这些段改好，或者退回原文（"changed": false），再跑。')
    sys.exit(1)
print(f'{n_p} 段，共 {n_h} 处修订，修订人“{AUTHOR}”' + ('（空跑，没有写文件）' if DRY else ''))
if DRY: sys.exit(0)
zout = zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True) if item.filename == 'word/document.xml' else zin.read(item.filename)
    zout.writestr(item, data)
zout.close()
print('→', a.dst)

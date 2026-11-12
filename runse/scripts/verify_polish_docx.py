#!/usr/bin/env python3
"""核对写回的 Word：接受全部修订后应等于润色终稿，拒绝全部修订后应等于原稿；公式、图、脚注号、域一个不少，也没有落进修订里；没改的段落原样。

用法: verify_polish_docx.py 原稿.docx 输出.docx 正文.txt 终稿.json [更多.json]
最后一行是“问题：无”才算写对了；有问题时退出码为 1。
（本脚本故意不用写回脚本和 docx_units.py 的任何代码，文字是另写的一套办法取的，两套对得上才可信。）
"""
import sys, zipfile, re, json, collections
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
O = 'urn:schemas-microsoft-com:office:office'
def q(t): return f'{{{W}}}{t}'
def name(e):
    if not isinstance(e.tag, str): return None, None
    qn = etree.QName(e); return qn.namespace, qn.localname

base_f, out_f, txt_f = sys.argv[1:4]; res_files = sys.argv[4:]
zb, zo = zipfile.ZipFile(base_f), zipfile.ZipFile(out_f)
bad = zo.testzip()
print('压缩包完整：', '是' if bad is None else f'否，坏在 {bad}')
for n in zo.namelist():
    if n.endswith('.xml') or n.endswith('.rels'): etree.fromstring(zo.read(n))
print('各部件 XML 都能解析；部件清单与原稿一致：', sorted(zb.namelist()) == sorted(zo.namelist()))
others = [n for n in zb.namelist() if n != 'word/document.xml' and n in zo.namelist() and zb.read(n) != zo.read(n)]
print('除正文外有变化的部件：', others or '无')
db, do = etree.fromstring(zb.read('word/document.xml')), etree.fromstring(zo.read('word/document.xml'))

def text(p, mode):
    """mode='acc' 接受修订后的文字；'rej' 拒绝修订后的文字；公式、图、脚注号、特殊符号写成占位"""
    out = []; cnt = collections.Counter()
    def ph(k): cnt[k] += 1; out.append(f'〔{k}{cnt[k]}〕')
    def walk(el):
        for ch in el:
            ns, tag = name(ch)
            if ns == M and tag in ('oMath', 'oMathPara'): ph('式'); continue
            if ns != W: continue
            if tag == 'r':
                if ch.find(q('fldChar')) is not None: continue
                for x in ch:
                    xn, xt = name(x)
                    if xn == W and xt in ('t', 'delText'): out.append(x.text or '')
                    elif xn == W and xt == 'tab': out.append('\t')
                    elif xn == W and xt == 'noBreakHyphen': out.append('-')
                    elif xn == W and xt == 'object':
                        ole = x.find(f'{{{O}}}OLEObject'); prog = (ole.get('ProgID') or '') if ole is not None else ''
                        ph('式' if (ole is None or prog.startswith('Equation') or 'MathType' in prog) else '图')
                    elif xn == W and xt in ('pict', 'drawing'): ph('图')
                    elif xn == W and xt in ('footnoteReference', 'endnoteReference'): ph('注')
                    elif xn == W and xt == 'sym': ph('符')
                    elif xt == 'AlternateContent': ph('图')
            elif tag == 'ins':
                if mode == 'acc': walk(ch)
            elif tag == 'del':
                if mode == 'rej': walk(ch)
            elif tag in ('hyperlink', 'smartTag', 'fldSimple', 'sdt', 'sdtContent', 'customXml'): walk(ch)
    walk(p); return ''.join(out)

pb, po = list(db.iter(q('p'))), list(do.iter(q('p')))
print('段落数：原稿', len(pb), '输出', len(po))
if len(pb) != len(po): print('问题：\n  - 段落数变了，写回脚本不该增删段落'); sys.exit(1)

# 正文.txt 的段号对到原稿段落：按顺序整串去对，只能有一个位置对得上
S = []
for line in open(txt_f, encoding='utf-8'):
    m = re.match(r'^(S\d+)(【[^】]*】)?\s?(.*)$', line.rstrip('\n'))
    if m: S.append((m.group(1), m.group(2) or '', m.group(3)))
rows = []
for el in db.find(q('body')):
    ns, tag = name(el)
    if tag == 'p':
        t = text(el, 'acc')
        if t.strip(): rows.append(('p', t.strip(), el))
    elif tag == 'tbl':
        for tr in el.iter(q('tr')): rows.append(('t', None, tr))
def same(row, s): return row[0] == 't' if s[1].startswith('【表格') else (row[0] == 'p' and row[1] == s[2])
starts = [k for k in range(len(rows) - len(S) + 1) if all(same(rows[k + i], s) for i, s in enumerate(S))]
print('正文.txt 在原稿里对得上的位置：', len(starts), '个')
if len(starts) != 1: print('问题：\n  - 正文.txt 和原稿对不上，没法核对'); sys.exit(1)
idx_of = {id(p): i for i, p in enumerate(pb)}
sid_at = {idx_of[id(rows[starts[0] + i][2])]: s[0] for i, s in enumerate(S) if not s[1]}
old_of = {s[0]: s[2] for s in S}
res = {}
for f in res_files:
    for it in json.load(open(f, encoding='utf-8')):
        if isinstance(it, dict) and it.get('S'): res[it['S']] = it
want = {sid: res[sid]['new'].strip() for sid in old_of if res.get(sid, {}).get('changed') and (res[sid].get('new') or '').strip()
        and res[sid]['new'].strip() != old_of[sid].strip() and sid in sid_at.values()}

def guard(p):
    """公式对象、图、脚注号、域标记的清单，以及它们有没有落进修订里"""
    items = []; inside = 0
    for e in p.iter():
        ns, tag = name(e); key = None
        if ns == M and tag in ('oMath', 'oMathPara'): key = 'oMath:' + ''.join(e.itertext())
        elif ns == W and tag == 'object': key = 'object:' + ','.join(sorted(v for x in e.iter() for k, v in x.attrib.items() if k.endswith('}id')))
        elif ns == W and tag == 'fldChar': key = 'fldChar:' + (e.get(q('fldCharType')) or '')
        elif ns == W and tag == 'instrText': key = 'instr:' + (e.text or '')
        elif ns == W and tag in ('drawing', 'pict'): key = tag
        elif ns == W and tag in ('footnoteReference', 'endnoteReference'): key = tag + ':' + (e.get(q('id')) or '')
        elif ns == W and tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference', 'bookmarkStart'): key = tag + ':' + (e.get(q('id')) or '')
        if key:
            items.append(key)
            if any(name(x)[1] in ('ins', 'del') for x in e.iterancestors()) and not key.startswith(('comment', 'bookmark')): inside += 1
    return items, inside

n_changed = n_same = 0; errs = []; seen = set()
for i, (b, o) in enumerate(zip(pb, po)):
    tb = text(b, 'acc'); gb, _ = guard(b); go, inside = guard(o); sid = sid_at.get(i); tag = sid or tb.strip()[:20]
    if gb != go: errs.append(f'{tag}：公式、图、域、批注或书签的清单变了')
    if inside: errs.append(f'{tag}：有 {inside} 个公式、图或域落进了修订里')
    if text(o, 'rej') != tb: errs.append(f'{tag}：拒绝修订后和原稿不一样')
    if sid in want:
        seen.add(sid); n_changed += 1; got = text(o, 'acc')
        if got.strip() != want[sid]: errs.append(f'{sid}：接受修订后和终稿不一样\n   应为 {want[sid]}\n   实为 {got.strip()}')
        if got[:len(got) - len(got.lstrip())] != tb[:len(tb) - len(tb.lstrip())]: errs.append(f'{sid}：段首空白变了')
    else:
        n_same += 1
        if etree.tostring(b) != etree.tostring(o): errs.append(f'{tag}：没改的段落被动了')
miss = [sid for sid in want if sid not in seen]
if miss: errs.append(f'这些段没写进去：{"、".join(miss)}')
ins, dels = list(do.iter(q('ins'))), list(do.iter(q('del')))
rev_ids = [e.get(q('id')) for e in ins + dels]
base_ids = set(e.get(q('id')) for e in db.iter() if e.get(q('id')))
if len(rev_ids) != len(set(rev_ids)): errs.append('修订编号有重复')
if set(rev_ids) & base_ids: errs.append('修订编号和原稿里已有的编号相撞')
empty = [e for e in ins + dels if not ''.join(e.itertext())]
if empty: errs.append(f'有 {len(empty)} 处空的修订')
print(f'改动段 {n_changed}，原样段 {n_same}；插入 {len(ins)} 处、删除 {len(dels)} 处')
print('修订人：', dict(collections.Counter(e.get(q('author')) for e in ins + dels)))
print('问题：', '无' if not errs else '')
for e in errs: print('  -', e)
sys.exit(1 if errs else 0)

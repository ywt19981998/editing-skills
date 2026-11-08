# 编号与引用检查：图号、表号、公式号的连续、重号、格式统一，引用是否存在、是否先见文后见图；文献序号的首引顺序、有引无文、有文无引
# 用法: scan_numbering.py 书稿.docx 输出目录
# 输出: 输出目录/numbering-issues.json（候选问题，交 AI 逐条判断）、numbering-report.md（清单与问题，给人看）
import sys, os, re, json, zipfile
from collections import defaultdict, Counter
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
def q(t): return f'{{{W}}}{t}'
def qm(t): return f'{{{M}}}{t}'

# ---------- 编号写法 ----------
SEP = r'[-‐‑‒–—―－.．]'
NUM = rf'(?:([A-Z]|\d{{1,2}})\s*({SEP}{{1,2}})\s*(\d{{1,3}})|(\d{{1,3}}))(?![\d])'   # 章-序 或 单序号
LP, RP = r'[（(]', r'[)）]'
CONN = r'(?:、|，|,|和|与|及|或|以及|至|到|~|～|—|–)'
RANGE = set('至到~～—–')
LABEL_BEFORE_EQ = re.compile(r'(式|公式|方程|见|和|与|及|、|至|到|~|～|—|–)\s*$')
CAPTION_NOT = re.compile(r'^\s*(所示|给出|表明|中|为|是|的|和|与|及|、|，|,|展示|描述|说明|可|显示|反映|列出|即|分别|对比|里)')
KIND_NAME = {'图': '图', '表': '表', '式': '式'}

def num_of(m, base):
    """从匹配对象取编号 (章, 序, 分隔符)；base 为 NUM 在整个正则里的分组起点"""
    ch, sep, seq, single = m.group(base), m.group(base + 1), m.group(base + 2), m.group(base + 3)
    if single: return (None, int(single), '')
    return (int(ch) if ch.isdigit() else ch, int(seq), sep)

def fmt(n):
    ch, seq = n[0], n[1]
    return f'{seq}' if ch is None else f'{ch}-{seq}'

CN = {'零': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
def cn2int(s):
    if s.isdigit(): return int(s)
    if s == '十': return 10
    if '十' in s:
        a, b = s.split('十')
        return (CN.get(a, 1) if a else 1) * 10 + (CN.get(b, 0) if b else 0)
    return CN.get(s)

# ---------- 读取段落 ----------
def load(path):
    z = zipfile.ZipFile(path)
    doc = etree.fromstring(z.read('word/document.xml'))
    styles = {}
    if 'word/styles.xml' in z.namelist():
        st = etree.fromstring(z.read('word/styles.xml'))
        for s in st.findall(q('style')):
            n = s.find(q('name'))
            styles[s.get(q('styleId'))] = (n.get(q('val')) if n is not None else '') or ''
    # P 号与 dump_docx.py 一致：正文段落和表格的每一行各占一个号
    plabel, n = {}, 0
    for el in doc.find(q('body')):
        if el.tag == q('p'):
            n += 1
            for x in el.iter(q('p')): plabel[x] = n
        elif el.tag == q('tbl'):
            for tr in el.iter(q('tr')):
                n += 1
                for x in tr.iter(q('p')): plabel[x] = n
    paras = []
    for p in doc.find(q('body')).iter(q('p')):
        def own(el):
            a = next(el.iterancestors(q('p')), None)
            return a is p
        # 目录、图目录不参与检查
        ps = p.find(f'{q("pPr")}/{q("pStyle")}')
        sid = ps.get(q('val')) if ps is not None else ''
        sname = styles.get(sid, sid)
        sdt_toc = any(g.get(q('val'), '').lower().startswith('table of') for a in p.iterancestors(q('sdt'))
                      for g in a.iter(q('docPartGallery')))
        instr = ''.join(t.text or '' for t in p.iter(q('instrText')) if own(t))
        is_toc = sdt_toc or re.match(r'(?i)^(toc|目录|table of figures|图表目录)', sname or '') or 'PAGEREF' in instr
        wtext, full, ctxt = [], [], []   # ctxt 只含 w:t，与修订脚本定位段落所用文字一致
        for el in p.iter():
            if not own(el): continue
            tag = el.tag
            if tag in (q('t'), q('tab'), q('noBreakHyphen')):
                r = next(el.iterancestors(q('r')), None)
                if r is not None and r.find(f'{q("rPr")}/{q("vanish")}') is not None: continue
                s = el.text or '' if tag == q('t') else ('\t' if tag == q('tab') else '-')
                wtext.append(s); full.append(s)
                if tag == q('t'): ctxt.append(s)
            elif tag == qm('t'):
                full.append(el.text or '')
        has_math = any(own(x) for x in p.iter(qm('oMath')))
        has_ole = any(own(x) for x in p.iter(q('object')))
        has_pic = any(own(x) for x in p.iter(q('drawing'))) or any(own(x) for x in p.iter(q('pict')))
        has_num = p.find(f'{q("pPr")}/{q("numPr")}') is not None
        paras.append(dict(i=len(paras) + 1, P=plabel.get(p, 0), wtext=''.join(wtext), full=''.join(full), ctxt=''.join(ctxt), style=sname,
                          toc=bool(is_toc), math=has_math, ole=has_ole, pic=has_pic, numpr=has_num))
    return paras

# ---------- 识别 ----------
RE_CHAPTER = re.compile(r'^\s*第\s*([0-9一二三四五六七八九十两]+)\s*章')
RE_APPENDIX = re.compile(r'^\s*附\s*录\s*([A-Z])')
RE_CAPTION = re.compile(rf'^\s*(续)?\s*(图|表)\s*{NUM}(?P<rest>.*)$')
RE_EQNO = re.compile(rf'{LP}\s*{NUM}\s*{RP}\s*$')
RE_FIGTAB_REF = re.compile(rf'(?<![续代发列])(图|表)\s*{NUM}')
RE_EQ_REF = re.compile(rf'(?:不等式|等式|公式|方程|式)\s*{LP}\s*{NUM}\s*{RP}')
RE_EQ_BARE = re.compile(rf'(?<![式程A-Za-z0-9]){LP}\s*([A-Z]|\d{{1,2}})\s*({SEP})\s*(\d{{1,3}})\s*{RP}')
RE_CONT_FT = re.compile(rf'\s*({CONN})\s*(?:图|表)?\s*{NUM}')
RE_CONT_EQ = re.compile(rf'\s*({CONN})\s*(?:式)?\s*{LP}\s*{NUM}\s*{RP}')
RE_HEADING_NUM = re.compile(r'^\s*(\d+(?:\.\d+){0,3})\s+\S')

def is_caption(p, m):
    if re.search(r'(?i)caption|题注', p['style'] or ''): return True
    t, rest = p['wtext'].strip(), m.group('rest')
    if len(t) > 80 or re.search(r'[。；]', t): return False
    if re.match(r'^\d', rest): return False
    if re.match(r'^[ \t\u3000]+\S', rest): return True   # 编号后有空格：题注写法
    return not CAPTION_NOT.match(rest)

def scan(paras):
    chapter, heading = None, ''
    defs, refs, notes = [], [], []
    for p in paras:
        if p['toc']: continue
        t = p['wtext'].strip()
        mc = RE_CHAPTER.match(t)
        if mc and len(t) < 50:
            chapter = cn2int(mc.group(1)); heading = t[:30]
        ma = RE_APPENDIX.match(t)
        if ma and len(t) < 50:
            chapter = ma.group(1); heading = t[:30]
        mh = RE_HEADING_NUM.match(t)
        if mh and len(t) < 50 and not mc: heading = t[:30]
        p['chapter'], p['heading'] = chapter, heading
        caption_span = None
        # 图题、表题
        m = RE_CAPTION.match(p['wtext'])
        if m and is_caption(p, m):
            n = num_of(m, 3)
            defs.append(dict(kind=m.group(2), num=n, cont=bool(m.group(1)), para=p['i'], pos=m.start(2),
                             raw=m.group(0)[:m.start('rest') - m.start(0)].strip(), chapter=chapter,
                             heading=heading, text=t[:60]))
            caption_span = (0, m.start('rest'))
        elif re.search(r'(?i)caption|题注', p['style'] or '') and p['numpr'] and t:
            notes.append(f"P{p['P']}「{t[:30]}」是题注样式但编号由 Word 列表自动生成，脚本读不到编号，需人工核对")
        # 公式编号（编号可能在公式内部，用含公式文字的 full 判断）
        eq_span = None
        f = p['full'].rstrip()
        me = RE_EQNO.search(f)
        if me and not m:
            before = f[:me.start()]
            formula_like = p['math'] or p['ole'] or before.endswith(('\t', ' ', '　')) or '=' in before
            # MathType/Word 公式对象后面只跟“制表符+编号”时，编号前没有文字
            if formula_like and (before.strip() or p['math'] or p['ole']) and not LABEL_BEFORE_EQ.search(before):
                n = num_of(me, 1)
                defs.append(dict(kind='式', num=n, cont=False, para=p['i'], pos=10**6, raw=me.group(0).strip(),
                                 chapter=chapter, heading=heading, text=f.strip()[-60:],
                                 in_math=me.group(0).strip() not in p['wtext']))
                eq_span = me.group(0).strip()
        # 引用
        w = p['wtext']
        def add_ref(kind, n, pos, raw, bare=False):
            refs.append(dict(kind=kind, num=n, para=p['i'], pos=pos, raw=raw, bare=bare,
                             chapter=chapter, heading=heading))
        for r in RE_FIGTAB_REF.finditer(w):
            if caption_span and r.start() < caption_span[1]: continue
            kind, n = r.group(1), num_of(r, 2)
            add_ref(kind, n, r.start(), r.group(0))
            end, prev = r.end(), n
            while True:   # “图1-2、图1-3和1-4”“图2-1～2-5”
                c = RE_CONT_FT.match(w, end)
                if not c: break
                n2 = num_of(c, 2)
                if n2[0] != prev[0]: break
                if c.group(1) in RANGE and n2[1] > prev[1] + 1:
                    for s in range(prev[1] + 1, n2[1]): add_ref(kind, (n2[0], s, n2[2]), c.start(), f'{kind}{fmt(prev)}~{fmt(n2)}（范围）')
                add_ref(kind, n2, c.start(), c.group(0).strip())
                end, prev = c.end(), n2
        eq_taken = []
        for r in RE_EQ_REF.finditer(w):
            n = num_of(r, 1)
            add_ref('式', n, r.start(), r.group(0)); eq_taken.append((r.start(), r.end()))
            end, prev = r.end(), n
            while True:
                c = RE_CONT_EQ.match(w, end)
                if not c: break
                n2 = num_of(c, 2)
                if n2[0] != prev[0]: break
                if c.group(1) in RANGE and n2[1] > prev[1] + 1:
                    for s in range(prev[1] + 1, n2[1]): add_ref('式', (n2[0], s, n2[2]), c.start(), f'式{fmt(prev)}~{fmt(n2)}（范围）')
                add_ref('式', n2, c.start(), c.group(0).strip()); eq_taken.append((c.start(), c.end()))
                end, prev = c.end(), n2
        for r in RE_EQ_BARE.finditer(w):
            if any(a <= r.start() < b for a, b in eq_taken): continue
            if eq_span and r.group(0).strip() == eq_span and r.end() >= len(w.rstrip()) - 1: continue
            ch = r.group(1)
            add_ref('式', (int(ch) if ch.isdigit() else ch, int(r.group(3)), r.group(2)), r.start(), r.group(0), bare=True)
    return defs, refs, notes

# ---------- 上下文定位（给修订脚本用的 ctx：在全文唯一） ----------
def make_ctx(paras, pi, target):
    texts = [p['ctxt'] for p in paras]
    w = texts[pi - 1]
    k = w.find(target)
    if k < 0: return w.strip()[:20]
    for pad in (4, 8, 12, 20, 40, 80):
        s = w[max(0, k - pad): k + len(target) + pad]
        if sum(s in x for x in texts) == 1: return s
    return w

# ---------- 检查 ----------
def check(paras, defs, refs):
    issues = []
    def add(check, kind, sev, item, msg, target=None, fix=None):
        p = paras[item['para'] - 1]
        tgt = target or item.get('raw', '')
        editable = bool(tgt) and tgt in p['ctxt']
        issues.append(dict(check=check, kind=kind, severity=sev, chapter=item.get('chapter'), heading=item.get('heading'),
                           para=item['para'], P=f"P{p['P']}", target=tgt, ctx=make_ctx(paras, item['para'], tgt) if editable else None,
                           editable=editable, message=msg, fix=fix, context=p['wtext'].strip()[:200]))

    main = [d for d in defs if not d['cont']]
    by_kind = defaultdict(list)
    for d in main: by_kind[d['kind']].append(d)

    # 1. 编号格式（分隔符、标签与数字之间是否空格、公式编号括号全半角）
    for kind in ('图', '表', '式'):
        items = [d for d in defs if d['kind'] == kind] + [r for r in refs if r['kind'] == kind and not r['bare'] and '范围' not in r['raw']]
        items = [x for x in items if x['num'][0] is not None]
        if len(items) < 3: continue
        seps = Counter(x['num'][2] for x in items)
        major, cnt = seps.most_common(1)[0]
        for x in items:
            s = x['num'][2]
            if s != major:
                new = x['raw'].replace(s, major, 1) if s in x['raw'] else None
                sev = '确定' if cnt / len(items) >= 0.8 else '待确认'
                add('编号格式', kind, sev, x, f'{kind}号分隔符“{s}”与全书多数写法“{major}”（{cnt}/{len(items)}）不一致',
                    target=x['raw'], fix=dict(old=x['raw'], new=new) if new else None)
        if kind == '式':
            eqd = [d for d in defs if d['kind'] == '式']
            if eqd:
                half = [d for d in eqd if d['raw'].startswith('(')]
                fullw = [d for d in eqd if d['raw'].startswith('（')]
                if half and fullw:
                    minor, maj = (half, '全角') if len(half) < len(fullw) else (fullw, '半角')
                    for d in minor:
                        add('编号格式', '式', '待确认', d, f'公式编号括号全半角混用：多数为{maj}括号', target=d['raw'])

    # 2. 题注章号与所在章不符；3. 跳号、重号、不从 1 开始、顺序颠倒
    for kind, items in by_kind.items():
        wrong = []
        for d in items:
            if d['chapter'] is not None and d['num'][0] is not None and d['num'][0] != d['chapter']:
                add('章号不符', kind, '确定', d, f'{kind}{fmt(d["num"])}位于第{d["chapter"]}章，章号不符')
                wrong.append(d)
        groups = defaultdict(list)
        for d in items:
            if d in wrong: continue   # 放错章的不参与连续性检查，免得连带误报
            groups[d['num'][0] if d['num'][0] is not None else '全书'].append(d)
        for g, ds in groups.items():
            seen = {}
            expect = 1
            for d in ds:
                seq = d['num'][1]
                if seq in seen:
                    add('重号', kind, '确定', d, f'{kind}{fmt(d["num"])}重号（另见 P{paras[seen[seq]["para"] - 1]["P"]}「{seen[seq]["text"][:20]}」）')
                    continue
                if seq != expect:
                    if seq > expect:
                        miss = '、'.join(fmt((d['num'][0], s)) for s in range(expect, seq))
                        add('跳号', kind, '确定' if kind != '式' else '待确认', d,
                            f'{kind}{fmt(d["num"])}之前缺少{kind}{miss}（按出现顺序应为{kind}{fmt((d["num"][0], expect))}）')
                    else:
                        add('顺序颠倒', kind, '确定', d, f'{kind}{fmt(d["num"])}出现在更大编号之后，编号顺序颠倒')
                seen[seq] = d
                expect = max(expect, seq + 1)

    # 4. 引用的编号不存在
    defined = {(d['kind'], d['num'][0], d['num'][1]) for d in defs}
    chap_mode = {k for k in '图表式' if any(d['kind'] == k and d['num'][0] is not None for d in defs)}
    for r in refs:
        if (r['kind'], r['num'][0], r['num'][1]) not in defined:
            if '范围' in r['raw']: continue
            if r['num'][0] is None and r['kind'] in chap_mode:   # 全书按“章-序”编号时，“图3”多半不是引用
                add('引用格式', r['kind'], '待确认', r, f'“{r["raw"]}”不是“章-序”编号，确认是否为引用、是否写错', target=r['raw'])
                continue
            sev = '待确认' if r['bare'] else '确定'
            hint = '（无“式”字的括号编号，可能不是公式引用）' if r['bare'] else ''
            add('引用不存在', r['kind'], sev, r, f'引用的{r["kind"]}{fmt(r["num"])}在全书题注/编号中找不到{hint}', target=r['raw'])

    # 5. 图表未被引用；6. 先见图表后见文；8. 公式在定义之前被引用
    first_ref = {}
    for r in sorted(refs, key=lambda x: (x['para'], x['pos'])):
        first_ref.setdefault((r['kind'], r['num'][0], r['num'][1]), r)
    for d in main:
        key = (d['kind'], d['num'][0], d['num'][1])
        fr = first_ref.get(key)
        if d['kind'] in '图表':
            if not fr:
                add('未引用', d['kind'], '确定', d, f'{d["kind"]}{fmt(d["num"])}在正文中没有引用（应先见文后见{d["kind"]}）')
            elif fr['para'] > d['para']:
                add('先见图后见文', d['kind'], '确定', d,
                    f'{d["kind"]}{fmt(d["num"])}的首次引用在 P{paras[fr["para"] - 1]["P"]}，位于{d["kind"]}题之后（P{paras[d["para"] - 1]["P"]}），应先见文后见{d["kind"]}')
        elif fr and fr['para'] < d['para'] and not fr['bare']:
            add('公式先引后出', '式', '待确认', fr, f'式{fmt(d["num"])}在 P{paras[d["para"] - 1]["P"]} 才出现，此处提前引用，确认是否为“下文将推出”的有意前指',
                target=fr['raw'])

    # 7. 图表编号顺序应与首次引用顺序一致
    for kind in '图表':
        groups = defaultdict(list)
        for d in by_kind.get(kind, []):
            fr = first_ref.get((kind, d['num'][0], d['num'][1]))
            if fr: groups[d['num'][0]].append((fr['para'], fr['pos'], d))
        for g, xs in groups.items():
            xs.sort(key=lambda x: (x[0], x[1]))
            mx = 0
            for _, _, d in xs:
                if d['num'][1] < mx:
                    add('引用顺序', kind, '待确认', d, f'{kind}{fmt(d["num"])}的首次引用晚于{kind}{fmt((g, mx))}的首次引用，编号顺序与引用顺序不一致')
                mx = max(mx, d['num'][1])

    unref_eq = [d for d in by_kind.get('式', []) if (d['kind'], d['num'][0], d['num'][1]) not in first_ref]
    return issues, unref_eq

# ---------- 文献序号 ----------
RE_REF_HEAD = re.compile(r'^\s*参\s*考\s*文\s*献\s*$')
RE_CITE = re.compile(r'\[(\d{1,3}(?:\s*[-–~～,，、]\s*\d{1,3})*)\]')
RE_ENTRY = re.compile(r'^\s*[\[［(（]?(\d{1,3})[\]］)）.．、]\s*\S')

def cite_nums(s):
    out = []
    for part in re.split(r'[,，、]', s):
        ab = re.split(r'[-–~～]', part)
        a, b = int(ab[0]), int(ab[-1])
        out += list(range(a, b + 1)) if a <= b <= a + 50 else [a, b]
    return out

def check_citations(paras):
    """每个“参考文献”标题管它前面的引用（章末文献）；序号按首次出现递增，有引有文对应"""
    issues, notes, cur, i = [], [], [], 0
    def add(check, sev, p, msg, target=''):
        editable = bool(target) and target in p['ctxt']
        issues.append(dict(check=check, kind='文献', severity=sev, chapter=p.get('chapter'), heading=p.get('heading'), para=p['i'], P=f"P{p['P']}",
                           target=target, ctx=make_ctx(paras, p['i'], target) if editable else None, editable=editable,
                           message=msg, fix=None, context=p['wtext'].strip()[:200]))
    def close(entries, head):
        seen, last = set(), 0
        for n, p, raw in cur:
            if n not in seen:
                if last and n != last + 1:   # 只报断点，断点之后接着连续的不再报；开头不算（单章稿可能接着全书序号）
                    add('文献序号', '待确认', p, f'首次引用的序号从[{last}]断到[{n}]，序号应按首次引用的先后连续', raw)
                seen.add(n); last = n
        if entries is None:
            if cur: notes.append(f'{len(seen)} 个文献序号之后没有“参考文献”标题，只查了首引顺序；有引无文、有文无引要等文献表齐了再看')
            return
        ids = [n for n, _ in entries]
        for n, p in entries:
            if ids.count(n) > 1: add('文献表', '确定', p, f'文献表里序号[{n}]重复')
            elif n not in seen: add('文献表', '待确认', p, f'文献表第[{n}]条在正文里没有引用（引用写在公式、图片或脚注里时脚本看不到）')
        for n in sorted(seen - set(ids)):
            p, raw = next((p, raw) for m, p, raw in cur if m == n)
            add('文献表', '确定', p, f'正文引用[{n}]，文献表只到[{max(ids or [0])}]或没有这一条；若是取值区间则不算', raw)
        notes.append(f'{head}：文献表 {len(ids)} 条，正文引用 {len(seen)} 个序号')
    while i < len(paras):
        p = paras[i]; t = p['wtext'].strip()
        if p['toc']: i += 1; continue
        if RE_REF_HEAD.match(t):
            entries, j = [], i + 1
            while j < len(paras):
                e = paras[j]; s = e['wtext'].strip()
                if not s: j += 1; continue
                if RE_CHAPTER.match(s) or RE_APPENDIX.match(s) or re.match(r'^\d+\.\d+', s): break
                m = RE_ENTRY.match(s)
                if m: entries.append((int(m.group(1)), e))
                elif e['numpr']: entries.append((len(entries) + 1, e))   # Word 自动编号，按段序算
                else: break
                j += 1
            close(entries, p.get('heading') or f"P{p['P']}"); cur.clear(); i = j
            continue
        w = p['wtext']
        for m in RE_CITE.finditer(w):
            nums = cite_nums(m.group(1))
            before = w[max(0, m.start() - 8): m.start()]
            if 0 in nums or re.search(r'[A-Za-z\])]$|区间|范围|取值|∈|=', before): continue   # 下标、取值区间
            cur += [(n, p, m.group(0)) for n in nums]
        i += 1
    close(None, '')
    return issues, notes

# ---------- 输出 ----------
def main():
    src, out = sys.argv[1], sys.argv[2]
    os.makedirs(out, exist_ok=True)
    paras = load(src)
    defs, refs, notes = scan(paras)
    issues, unref_eq = check(paras, defs, refs)
    cite_issues, cite_notes = check_citations(paras)
    issues += cite_issues; notes += cite_notes
    issues.sort(key=lambda x: (x['para'], x['check']))
    for k, it in enumerate(issues, 1): it['id'] = f'N-{k:03d}'
    json.dump(dict(source=os.path.abspath(src), issues=issues, notes=notes,
                   inventory=[dict(kind=d['kind'], num=fmt(d['num']), cont=d['cont'], para=d['para'], text=d['text']) for d in defs]),
              open(os.path.join(out, 'numbering-issues.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

    L = [f'# 编号与引用检查\n\n稿件：`{os.path.basename(src)}`\n']
    cnt = Counter(d['kind'] for d in defs if not d['cont'])
    L.append(f'识别到 图 {cnt["图"]} 个、表 {cnt["表"]} 个、编号公式 {cnt["式"]} 个；引用 {len(refs)} 处。'
             f'候选问题 {len(issues)} 条（确定 {sum(i["severity"]=="确定" for i in issues)}，待确认 {sum(i["severity"]=="待确认" for i in issues)}）。\n')
    L.append('> 脚本结果是候选问题，须经逐条判断后才写入修订。\n')
    if notes: L += ['## 需人工核对', *[f'- {n}' for n in notes], '']
    L.append('## 候选问题\n\n| 编号 | 位置 | 类型 | 程度 | 说明 | 可直接修订 |\n|---|---|---|---|---|---|')
    for i in issues:
        loc = f'第{i["chapter"]}章 {i["P"]}' if i['chapter'] is not None else i['P']
        fx = f'{i["fix"]["old"]}→{i["fix"]["new"]}' if i['fix'] and i['fix'].get('new') else ('否' if not i['editable'] else '需判断')
        L.append(f'| {i["id"]} | {loc} | {i["check"]} | {i["severity"]} | {i["message"]} | {fx} |')
    if unref_eq:
        L.append(f'\n## 提示：未被引用的编号公式（{len(unref_eq)} 个）\n\n按体例要求决定是否保留编号：' +
                 '、'.join(f'({fmt(d["num"])})' for d in unref_eq))
    L.append('\n## 编号清单\n')
    for kind in ('图', '表', '式'):
        ds = [d for d in defs if d['kind'] == kind]
        if ds: L.append(f'**{kind}**：' + '、'.join(('续' if d['cont'] else '') + fmt(d['num']) for d in ds) + '\n')
    open(os.path.join(out, 'numbering-report.md'), 'w', encoding='utf-8').write('\n'.join(L))
    print(f'图{cnt["图"]} 表{cnt["表"]} 式{cnt["式"]} 引用{len(refs)} → 候选问题 {len(issues)} 条；输出在 {out}')

if __name__ == '__main__':
    main()

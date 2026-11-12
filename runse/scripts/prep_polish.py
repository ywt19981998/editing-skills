#!/usr/bin/env python3
"""润色前的准备：把 Word 原稿导成带段号的正文（公式、图、脚注号用占位标出），把 PDF 页面转成图片，给出分组建议。

只做不用判断的事。拿不准的（哪些短段是标题、页码对不上）列在《准备报告》里，由主会话看过再派单。

用法: prep_polish.py 原稿.docx 输出目录 [--from 正则] [--to 正则] [--pdf 文件 --pages 202-218] [--dpi 110]
                     [--body S012,S020-S025] [--skip S030-S040] [--group 4500] [--group-paras 30] [--force]
  --from/--to  只取一部分：从第一个匹配 --from 的段起，到其后第一个匹配 --to 的段之前（一个文件装两章时用）。
               同一句话既在章首节目列表里、又是正文节题时，取用了标题样式的那一个。不给 --to 时到“参考文献”为止。
  --pdf/--pages 能看到公式的 PDF（排版稿，或从 WPS 导出的）和本章所在的页码范围（PDF 的实际页码）。
  --body/--skip 纠正标记：--body 里的段当正文润色，--skip 里的段不润色。
  --group/--group-paras  每组正文的字数、段数上限，默认 4500 字、30 段（试跑时每组不超过这个量）。
产物: 输出目录/正文.txt、准备报告.md、页面/pNNN.png
"""
import sys, os, re, argparse, collections
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from docx_units import q, ln, PH, load, has_revisions, body_rows, build_units

ap = argparse.ArgumentParser()
ap.add_argument('docx'); ap.add_argument('out')
ap.add_argument('--from', dest='start'); ap.add_argument('--to', dest='stop')
ap.add_argument('--pdf'); ap.add_argument('--pages'); ap.add_argument('--dpi', type=int, default=110)
ap.add_argument('--body', default=''); ap.add_argument('--skip', default='')
ap.add_argument('--group', type=int, default=4500); ap.add_argument('--group-paras', type=int, default=30)
ap.add_argument('--force', action='store_true')
a = ap.parse_args()
os.makedirs(a.out, exist_ok=True)

z, doc = load(a.docx)
if has_revisions(doc):
    sys.exit('停：这份稿子里有还没接受的修订。润色要在干净的稿子上做，先问编辑这些修订是接受还是拒绝，处理完再来。')

# ---------- 样式：哪些是标题、图表题、公式 ----------
styles = {}
if 'word/styles.xml' in z.namelist():
    for st in etree.fromstring(z.read('word/styles.xml')).iter(q('style')):
        name = st.find(q('name')); ppr = st.find(q('pPr')); bo = st.find(q('basedOn'))
        ol = ppr.find(q('outlineLvl')) if ppr is not None else None
        styles[st.get(q('styleId'))] = ((name.get(q('val')) if name is not None else '') or '',
                                         int(ol.get(q('val'))) if ol is not None else None,
                                         bo.get(q('val')) if bo is not None else None)
def style_of(p):
    """返回 (样式名, 大纲级别)；大纲级别沿“基于”一路找上去"""
    ppr = p.find(q('pPr')); sid = None; ol = None
    if ppr is not None:
        if ppr.find(q('pStyle')) is not None: sid = ppr.find(q('pStyle')).get(q('val'))
        if ppr.find(q('outlineLvl')) is not None: ol = int(ppr.find(q('outlineLvl')).get(q('val')))
    name = styles.get(sid, ('', None, None))[0]; seen = set()
    while ol is None and sid in styles and sid not in seen:
        seen.add(sid); ol = styles[sid][1]; sid = styles[sid][2]
    return name, ol

# ---------- 每一行是什么 ----------
EQNO = re.compile(r'[（(]\d+([-－—.．]\d+)*[a-zA-Z]?[)）]')
END = '。；！？;!?'
def cjk_len(s): return len(re.findall(r'[一-鿿]', s))
def classify(kind, text, sname, ol):
    """返回 (类别, 标题层级或 None)"""
    if kind == 't': return '表格', None
    s = text.strip(); core = re.sub(r'[\s　]', '', PH.sub('', s))
    rest = re.sub(r'[，。；：、,.;:]', '', EQNO.sub('', core))
    if not rest:
        return ('独立公式' if '〔式' in s else '图' if '〔图' in s else '独立公式'), None
    sl = sname.lower(); L = len(core); ends = s[-1] in END
    if re.search(r'toc|目录', sl) or re.search(r'\t\d+$', s): return '目录', None
    if ol is not None and ol < 9: return '标题', ol + 1
    if re.search(r'heading|标题|title', sl): return '标题', 1
    if re.search(r'图题|表题|图说|表说|图注|表注|题注|caption', sl): return '图表题', None
    if '公式' in sname and cjk_len(rest) < 6: return '独立公式', None
    if re.match(r'^续表', s) and L <= 12: return '图表题', None
    m = re.match(r'^(图|表|算法|代码|程序)\s*[\dA-Za-z]+([-－—.．]+\d+)*', s)
    if m and not ends:                      # “图3-1  题名”是图表题；“表4-1列出了……”这种接着说下去的是正文
        after = s[m.end():]
        if (after[:1] in ' 　\t' and L <= 60) or (L <= 30 and not re.search(r'[，。；]', after)): return '图表题', None
    if re.match(r'^[（(][a-zA-Z][)）]', s) and L <= 40 and not ends: return '图表题', None
    if re.match(r'^(资料来源|数据来源|来源)[:：]', s): return '图表题', None
    if re.match(r'^第[一二三四五六七八九十百零〇\d]+[章篇部]', s) and L <= 40: return '标题', 1
    if re.match(r'^第[一二三四五六七八九十百零〇\d]+节', s) and L <= 40: return '标题', 2
    m = re.match(r'^(\d+(?:\.\d+)+)[\s　]*\S', s)
    if m and L <= 45 and not ends: return '标题', m.group(1).count('.') + 1
    if re.match(r'^[一二三四五六七八九十]+、', s) and L <= 30 and not ends: return '标题', 2
    if re.match(r'^（[一二三四五六七八九十]+）', s) and L <= 30 and not ends: return '标题', 3
    if re.match(r'^[（(]\d+[)）]', s) and L <= 18: return '标题', 4
    return '正文', None

rows = body_rows(doc)
info = []
for kind, text, el in rows:
    sname, ol = style_of(el) if kind == 'p' else ('', None)
    cat, lvl = classify(kind, text, sname, ol)
    info.append({'kind': kind, 'text': text.strip() if kind == 'p' else text, 'el': el, 'cat': cat, 'lvl': lvl, 'styled': ol is not None, 'note': ''})

# ---------- 范围 ----------
def find(pat, lo):
    hits = [i for i in range(lo, len(info)) if info[i]['kind'] == 'p' and re.search(pat, PH.sub('', info[i]['text']).strip())]
    if not hits: sys.exit(f'停：没有哪一段匹配 {pat}')
    styled = [i for i in hits if info[i]['styled']]
    return (styled or hits)[0]
REFS = ('参考文献', '本章参考文献', '主要参考文献', '参考资料', 'References', 'REFERENCES')
lo = find(a.start, 0) if a.start else 0
hi = find(a.stop, lo + 1) if a.stop else len(info)
ref = next((i for i in range(lo, hi) if info[i]['kind'] == 'p' and re.sub(r'\s', '', PH.sub('', info[i]['text'])) in REFS), None)
if ref is not None: hi = ref          # 参考文献不润色，也不用给子代理当上下文
sel = info[lo:hi]
if not sel: sys.exit('停：范围里没有内容')
width = 3 if len(sel) < 1000 else 4
for n, r in enumerate(sel, 1): r['S'] = f'S{n:0{width}d}'

# 写不回去的段（域跨段等）：标成不改
for r in sel:
    if r['kind'] == 'p':
        try: build_units(r['el'])
        except ValueError:
            if r['cat'] == '正文': r['cat'] = '不改'; r['note'] = '段里的域没有配对，脚本写不回去'

# ---------- 人工纠正 ----------
def parse_ids(s):
    out = set()
    for part in filter(None, re.split(r'[,，\s]+', s)):
        m = re.match(r'^S?(\d+)(?:[-~～]S?(\d+))?$', part)
        if not m: sys.exit(f'停：看不懂的段号 {part}')
        x, y = int(m.group(1)), int(m.group(2) or m.group(1))
        out.update(range(x, y + 1))
    return out
to_body, to_skip = parse_ids(a.body), parse_ids(a.skip)
for n, r in enumerate(sel, 1):
    if n in to_body and r['kind'] == 'p': r['cat'] = '正文'; r['lvl'] = None
    if n in to_skip and r['cat'] == '正文': r['cat'] = '不改'
TAG = {'正文': '', '标题': '【标题，不改】', '图表题': '【图表题，不改】', '独立公式': '【独立公式】', '表格': '【表格，不改】',
       '图': '【图，不改】', '目录': '【目录，不改】', '不改': '【不改】'}
with open(os.path.join(a.out, '正文.txt'), 'w', encoding='utf-8') as f:
    for r in sel: f.write(f"{r['S']}{TAG[r['cat']]}{' ' if not TAG[r['cat']] else ''}{r['text']}\n")

def chars(r): return len(re.sub(r'[\s　]', '', PH.sub('', r['text'])))
body = [r for r in sel if r['cat'] == '正文']
stat = collections.Counter(r['cat'] for r in sel)
n_formula = sum(1 for r in body if '〔式' in r['text'])
n_ph = sum(len(PH.findall(r['text'])) for r in body)

# ---------- PDF 页面 ----------
page_note = []; pages = []
if a.pdf:
    import pymupdf
    pdf = pymupdf.open(a.pdf)
    if a.pages:
        m = re.match(r'^(\d+)\s*[-~～]\s*(\d+)$', a.pages) or re.match(r'^(\d+)()$', a.pages)
        if not m: sys.exit(f'停：--pages 写成 202-218 这样，看不懂 {a.pages}')
        p0, p1 = int(m.group(1)), int(m.group(2) or m.group(1))
    else: p0, p1 = 1, len(pdf)
    if p0 < 1 or p1 > len(pdf) or p0 > p1: sys.exit(f'停：页码 {p0}-{p1} 超出 PDF 的 1-{len(pdf)} 页')
    pages = list(range(p0, p1 + 1))
    pdir = os.path.join(a.out, '页面'); os.makedirs(pdir, exist_ok=True)
    norm = lambda s: re.sub(r'[^一-鿿A-Za-z0-9]', '', s)
    ptext = {}
    for n in pages:
        pg = pdf[n - 1]; ptext[n] = norm(pg.get_text())
        png = os.path.join(pdir, f'p{n:03d}.png')
        if a.force or not os.path.exists(png): pg.get_pixmap(dpi=a.dpi).save(png)
    # 每个正文段在哪一页：取几小段连续的字去各页里找
    for r in body:
        bits = [norm(x) for x in PH.split(r['text'])]
        snips = [b[i:i + 10] for b in bits if len(b) >= 10 for i in (0, max(0, len(b) // 2 - 5), len(b) - 10)][:6]
        votes = collections.Counter(n for sn in set(snips) for n in pages if sn in ptext[n])
        r['page'] = votes.most_common(1)[0][0] if votes else None
    found = sum(1 for r in body if r.get('page'))
    page_note.append(f'页面 {len(pages)} 张（p{p0:03d}～p{p1:03d}，{a.dpi} dpi）→ 页面/；{len(body)} 个正文段里有 {found} 个在 PDF 文字里对上了页码')
    if body and found < len(body) * 0.5:
        page_note.append('**对上的不到一半**：PDF 可能是扫描件，或者和这份稿子差得多。派单里就不要写各组的页码范围，让子代理把整章页面都看一遍。')
elif n_ph:
    page_note.append(f'**没有给 PDF，而正文段里有 {n_ph} 个公式等占位**：子代理看不到公式内容，公式旁边的话容易改错。请编辑给一份能看到公式的 PDF（排版稿，或在 WPS 里“输出为 PDF”），再带 --pdf 重跑。')

# ---------- 分组建议：在节的边界上切，每组不超过 --group 字、--group-paras 段 ----------
def n_body(items): return sum(1 for r in items if r['cat'] == '正文')
def n_chars(items): return sum(chars(r) for r in items if r['cat'] == '正文')
def sections(items, lvl):
    """在层级不低于 lvl 的标题前切开；连着的几个标题不拆开"""
    secs = [[]]
    for r in items:
        if r['cat'] == '标题' and r['lvl'] and r['lvl'] <= lvl and n_body(secs[-1]): secs.append([])
        secs[-1].append(r)
    return secs
def too_big(items, k=1.0): return n_chars(items) > a.group * k or n_body(items) > a.group_paras * k
def pack(secs):
    groups = [[]]
    for sec in secs:
        if n_body(groups[-1]) and too_big(groups[-1] + sec): groups.append([])
        groups[-1] += sec
    return groups
order = sorted({r['lvl'] for r in sel if r['cat'] == '标题' and r['lvl']})
def split(items, levels):
    """先按最高一级标题切；某一节自己就超了，再按下一级切；没有标题可切时按段数硬切"""
    for i, lvl in enumerate(levels):
        secs = sections(items, lvl)
        if len(secs) < 2: continue
        out = []
        for g in pack(secs):
            out += split(g, levels[i + 1:]) if too_big(g, 1.3) else [g]
        return out
    if not too_big(items, 1.3): return [items]
    out = [[]]
    for r in items:
        if r['cat'] == '正文' and n_body(out[-1]) >= a.group_paras: out.append([])
        out[-1].append(r)
    return out
groups = [g for g in split(sel, order) if n_body(g)]
# 太小的组并到前一组（并不进去就并到后一组），免得为两三段话单派一个子代理
i = 0
while i < len(groups) and len(groups) > 1:
    g = groups[i]
    if n_chars(g) < a.group * 0.25 and n_body(g) < a.group_paras * 0.25:
        if i > 0 and not too_big(groups[i - 1] + g, 1.3): groups[i - 1] += g; del groups[i]; continue
        if i + 1 < len(groups) and not too_big(g + groups[i + 1], 1.3): groups[i + 1] = g + groups[i + 1]; del groups[i]; continue
    i += 1

# ---------- 报告 ----------
L = ['# 润色准备报告', '',
     f'- 原稿：`{os.path.abspath(a.docx)}`',
     f'- 范围：{sel[0]["S"]}～{sel[-1]["S"]}（“{sel[0]["text"][:20]}”到“{sel[-1]["text"][:20]}”）' + ('，“参考文献”以后没有导出' if ref is not None else ''),
     f'- 要润色的正文段 {len(body)} 个，共 {sum(chars(r) for r in body)} 字；其中带公式的 {n_formula} 段，公式、图、脚注号等占位共 {n_ph} 个',
     '- 不润色的行：' + '、'.join(f'{k} {v}' for k, v in stat.items() if k != '正文')]
ncom = len(list(doc.iter(q('commentReference'))))
if ncom: L.append(f'- 原稿自带批注 {ncom} 条，写回修订后原样保留')
L += [f'- {x}' for x in page_note]
L += ['', '## 分组建议', '', '每组派一个润色子代理、一个复查子代理。', '', '| 组 | 范围 | 正文段 | 字数 | 页面 | 从哪到哪 |', '|---|---|---|---|---|---|']
for k, g in enumerate(groups, 1):
    b = [r for r in g if r['cat'] == '正文']; pg = sorted(r['page'] for r in b if r.get('page'))
    heads = [re.sub(r'〔图\d+〕', '', r['text']).strip()[:18] for r in g if r['cat'] == '标题' and r['lvl'] and r['lvl'] <= 3]
    heads = [h for h in heads if h]
    span = f'p{pg[0]:03d}～p{pg[-1]:03d}' if pg else ('全部' if pages else '—')
    where = (heads[0] + (' …… ' + heads[-1] if len(heads) > 1 else '')) if heads else '（接上一组）'
    L.append(f'| {k} | {g[0]["S"]}～{g[-1]["S"]} | {len(b)} | {n_chars(g)} | {span} | {where} |')
# 标记可能标错的
doubt = []
for r in sel:
    if r['cat'] == '正文' and chars(r) <= 15 and r['text'][-1] not in '：。；，:.;,': doubt.append(f'{r["S"]} 当正文了，但很短、句末没有标点（像标题）：{r["text"][:30]}')
    if r['cat'] in ('标题', '图表题') and chars(r) > 40: doubt.append(f'{r["S"]} 当{r["cat"]}了，但很长：{r["text"][:30]}……')
    if r['note']: doubt.append(f'{r["S"]} 标成不改：{r["note"]}')
L += ['', '## 标记要看一眼的', '']
L += [f'- {x}' for x in doubt] or ['- 无']
L += ['', '标错了就加 `--body 段号` 或 `--skip 段号` 重跑（段号可以写 S012 或 S020-S025，逗号隔开）。算法步骤、代码这类脚本认不出来的，派单里已经交代子代理不改。', '']
open(os.path.join(a.out, '准备报告.md'), 'w', encoding='utf-8').write('\n'.join(L))
print(f'{sel[0]["S"]}～{sel[-1]["S"]}：正文段 {len(body)}（{sum(chars(r) for r in body)} 字，带公式 {n_formula} 段），' + '、'.join(f'{k} {v}' for k, v in stat.items() if k != '正文'))
for x in page_note: print(re.sub(r'\*\*', '', x))
print(f'建议分 {len(groups)} 组；要看一眼的标记 {len(doubt)} 处 → {os.path.join(a.out, "准备报告.md")}')

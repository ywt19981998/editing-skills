#!/usr/bin/env python3
"""机器核对润色结果：漏段、公式等占位、数值、引注、外文是不是原样；名词有没有换；标点；改动幅度。并出逐段对照。

硬问题（占位、数值、引注、外文有增减，漏段，不能动的块被改）必须解决才能写回 Word；
其余是给复查看的线索，不一定是错。

用法: check_polish.py 原稿.docx 正文.txt 输出目录 结果1.json [结果2.json ...]
                      [--terms 术语表.txt] [--range S038-S078] [--prefix 终稿-]
  --terms   一行一个术语（有本书名词表时导一份），分词更准，术语被换时单独列出
  --range   只核对这一段范围（只润色了一部分时用）
  --prefix  输出文件名前缀：第二遍核对终稿时用，免得盖掉第一遍的
产物: 输出目录/[前缀]核对单.md、对照.md、对照.html；有硬问题时退出码为 1
"""
import sys, os, re, json, argparse, collections, difflib, html
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import jieba, jieba.posseg as pseg
jieba.setLogLevel(60)
from docx_units import PH, load, body_rows, read_slines, align, para_text, fixed_parts

ap = argparse.ArgumentParser()
ap.add_argument('docx'); ap.add_argument('txt'); ap.add_argument('out'); ap.add_argument('results', nargs='+')
ap.add_argument('--terms'); ap.add_argument('--range', dest='rng'); ap.add_argument('--prefix', default='')
a = ap.parse_args()
os.makedirs(a.out, exist_ok=True)

slines = read_slines(a.txt)
z, doc = load(a.docx)
el_of = align(body_rows(doc), slines)
text_of = {sid: text for sid, tag, text in slines}; tag_of = {sid: tag for sid, tag, text in slines}
order = [sid for sid, tag, text in slines]
num = lambda sid: int(sid[1:])
lo, hi = num(order[0]), num(order[-1])
if a.rng:
    m = re.match(r'^S?(\d+)\s*[-~～]\s*S?(\d+)$', a.rng)
    if not m: sys.exit(f'--range 写成 S038-S078 这样，看不懂 {a.rng}')
    lo, hi = int(m.group(1)), int(m.group(2))
in_rng = [sid for sid in order if lo <= num(sid) <= hi]
body = [sid for sid in in_rng if not tag_of[sid]]

terms = []
if a.terms:
    terms = sorted({w.strip() for w in open(a.terms, encoding='utf-8') if w.strip()}, key=len, reverse=True)
    for w in terms: jieba.add_word(w, tag='nz')

res, dup, bad_files, stray = {}, [], [], []
for f in a.results:
    try: data = json.load(open(f, encoding='utf-8'))
    except Exception as e: bad_files.append(f'{os.path.basename(f)}：{e}'); continue
    if not isinstance(data, list): bad_files.append(f'{os.path.basename(f)}：最外层不是列表'); continue
    for it in data:
        S = it.get('S') if isinstance(it, dict) else None
        if not S: continue
        if S in res: dup.append(S)
        res[S] = it
for S in res:
    if S not in text_of or tag_of[S]: stray.append(S)

def strip_ph(s): return PH.sub('', s)
def nums(s): return collections.Counter(re.findall(r'\d+(?:\.\d+)?', strip_ph(s)))
def cites(s): return collections.Counter(re.sub(r'\s', '', x) for x in re.findall(r'\[[\d,，\-–~～\s]+\]', s))
def latin(s): return collections.Counter(re.findall(r'[A-Za-z][A-Za-z0-9+]*(?:-[A-Za-z0-9+]+)*', strip_ph(s)))
def term_counts(s):
    s = strip_ph(s); c = collections.Counter()
    for w in terms:
        n = s.count(w)
        if n: c[w] = n; s = s.replace(w, '□' * len(w))
    return c
def nouns(s):
    out = collections.Counter()
    for w, fl in pseg.cut(strip_ph(s)):
        if len(w) >= 2 and (fl.startswith('n') or fl in ('vn', 'eng')) and re.search(r'[一-鿿A-Za-z]', w): out[w] += 1
    return out
CJK = r'[一-鿿〔〕]'
def punct_lint(s):
    p = []; t = PH.sub('式', s)
    if re.search(CJK + r'[,;:!?]' + r'|[,;:!?]' + CJK, t): p.append('中文里夹半角标点')
    if re.search(r'[，。；：、！？]{2,}', re.sub(r'[。！？][”’）]', '', t)): p.append('标点连用')
    for x, y in (('（', '）'), ('“', '”'), ('《', '》'), ('(', ')'), ('[', ']')):
        if t.count(x) != t.count(y): p.append(f'{x}{y}不成对')
    if t and t[-1] not in '。！？：；…”）)': p.append('段末没有标点')
    if re.search(r'、[^，。；：]{0,12}、[^，。；：]{0,30}(?<!，)以及', t): p.append('“以及”前缺逗号（疑似）')
    return p
def toks(s): return [w for w in jieba.cut(s, HMM=False) if w]
def inline_diff(old, new, fmt):
    ta, tb = toks(old), toks(new); out = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes():
        x, y = ''.join(ta[i1:i2]), ''.join(tb[j1:j2])
        if op == 'equal': out.append(fmt('=', x))
        else:
            if x: out.append(fmt('-', x))
            if y: out.append(fmt('+', y))
    return ''.join(out)
md_fmt = lambda k, s: s if k == '=' else ('[-' + s + '-]' if k == '-' else '{+' + s + '+}')
html_fmt = lambda k, s: html.escape(s) if k == '=' else (f'<del>{html.escape(s)}</del>' if k == '-' else f'<ins>{html.escape(s)}</ins>')

missing = [S for S in body if S not in res]
changed = []
for S in body:
    it = res.get(S)
    if it and it.get('changed') and (it.get('new') or '').strip() and it['new'].strip() != text_of[S].strip(): changed.append(S)
hard = collections.defaultdict(list)      # 段号 → 硬问题
soft = collections.defaultdict(list)      # 段号 → 线索
term_rows, noun_rows = [], []
for S in changed:
    old, new = text_of[S], res[S]['new'].strip()
    if '\n' in new: hard[S].append('润色稿里有换行（一段只能是一行）')
    if PH.findall(old) != PH.findall(new):
        hard[S].append(f'占位变了：原 {"".join(PH.findall(old)) or "无"} → 改后 {"".join(PH.findall(new)) or "无"}')
    else:
        full = para_text(el_of[S]); lead, trail = full[:len(full) - len(full.lstrip())], full[len(full.rstrip()):]
        new_full = lead + new + trail; pos = 0
        atoms, fields = fixed_parts(el_of[S])
        for d in atoms:
            i = new_full.find(d, pos)
            if i < 0:
                hard[S].append(f'“{d.strip() or "制表符"}”在原稿里是不能拆开改的一块（超链接、带制表符或换行的文字），润色稿里找不到原样的了'); break
            pos = i + len(d)
        gone = [x for x in fields if x not in new]
        if gone: soft[S].append('交叉引用或引注的显示文字被改了（链接可能丢）：' + '、'.join(f'“{x}”' for x in gone))
    for name, fn in (('数值', nums), ('引注序号', cites), ('外文、缩写', latin)):
        x, y = fn(old), fn(new)
        if x != y:
            less, more = dict(x - y), dict(y - x)
            hard[S].append(f'{name}有增减：' + '；'.join(filter(None, [f'少了 {"、".join(less)}' if less else '', f'多了 {"、".join(more)}' if more else ''])))
    if terms:
        x, y = term_counts(old), term_counts(new)
        if x != y:
            term_rows.append((S, dict(x - y), dict(y - x)))
            soft[S].append('术语表里的词有变化：' + '；'.join(filter(None, [f'少了 {"、".join(x - y)}' if x - y else '', f'多了 {"、".join(y - x)}' if y - x else ''])))
    na, nb = nouns(old), nouns(new)
    gone = [w for w in na if w not in nb]; come = [w for w in nb if w not in na]
    if gone or come:
        noun_rows.append((S, gone, come))
        soft[S].append('名词（分词粗查）：' + '；'.join(filter(None, [f'消失 {"、".join(gone)}' if gone else '', f'新出现 {"、".join(come)}' if come else ''])))
    lo_, ln_ = punct_lint(old), punct_lint(new)
    if ln_: soft[S].append('标点疑点：' + '、'.join(ln_))
    a_, b_ = strip_ph(old), strip_ph(new)
    ratio = difflib.SequenceMatcher(None, a_, b_, autojunk=False).ratio()
    if ratio < 0.5: soft[S].append(f'整段基本重写（和原文相同的部分不到一半），逐句对意思')
    if a_ and not 0.75 <= len(b_) / len(a_) <= 1.3: soft[S].append(f'长度从 {len(a_)} 字变成 {len(b_)} 字，看有没有增删内容')
    if PH.search(old) and not res[S].get('formulas'): soft[S].append('改了带公式的段，却没写明各个占位是什么公式')

def esc(s): return s.replace('\n', ' ')
n_hard = sum(len(v) for v in hard.values())
old_chars = sum(len(strip_ph(text_of[S])) for S in changed); all_chars = sum(len(strip_ph(text_of[S])) for S in body)
L = ['# 润色核对单', '',
     f'- 范围 {in_rng[0]}～{in_rng[-1]}：正文 {len(body)} 段，交回 {len([S for S in body if S in res])} 段，其中改了 {len(changed)} 段、没改 {len([S for S in body if S in res]) - len(changed)} 段',
     f'- 改动段的原文合计 {old_chars} 字（范围内正文共 {all_chars} 字）',
     f'- **硬问题 {n_hard + len(missing) + len(bad_files)} 个**' + ('（要先解决，才能写回 Word）' if n_hard or missing or bad_files else ''),
     f'- 给复查的线索：{len(soft)} 段', '']
L += ['## 一、硬问题', '']
if bad_files: L += [f'- 结果文件读不了：{x}' for x in bad_files]
if missing: L.append(f'- 漏交 {len(missing)} 段：{"、".join(missing)}')
for S in changed:
    for x in hard.get(S, []): L.append(f'- {S}　{x}')
if not (n_hard or missing or bad_files): L.append('- 无：占位、数值、引注序号、外文缩写全部原样，没有漏段。')
if dup: L.append(f'- （备查）重复交了的段，按后交的算：{"、".join(sorted(set(dup)))}')
if stray: L.append(f'- （备查）多交的条目，是标题、公式、表格或不存在的段号，写回时不会用：{"、".join(stray)}')
L += ['', '## 二、给复查的线索（不一定是错）', '']
for S in changed:
    for x in soft.get(S, []): L.append(f'- {S}　{x}')
if not soft: L.append('- 无')
asks = [(S, res[S]['why']) for S in body if S in res and not (res[S].get('changed')) and '读不懂' in (res[S].get('why') or '')]
if asks:
    L += ['', '## 三、润色时读不懂、没有改的', '']
    L += [f'- {S}　{esc(w)}' for S, w in asks]
L.append('')
open(os.path.join(a.out, a.prefix + '核对单.md'), 'w', encoding='utf-8').write('\n'.join(L))

# 对照（给复查和编辑看）：范围内每一行都列，改了的段列原、改、标出改动和线索
D = ['# 润色对照', '', '改了的段列出原文、润色后、标出改动（[-删掉的-]{+新加的+}）和脚本提示；没改的段和标题、公式、表格只列原文，当上下文看。', '']
for S in in_rng:
    if tag_of[S]: D.append(f'{S}{tag_of[S]}{text_of[S]}'); D.append(''); continue
    if S in changed:
        it = res[S]; new = it['new'].strip()
        D += [f'## {S}　改了', '', f'- 原：{text_of[S]}', f'- 改：{esc(new)}', f'- 标出改动：{inline_diff(text_of[S], esc(new), md_fmt)}']
        if it.get('formulas'): D.append('- 公式：' + '；'.join(f'{k} 是 {v}' for k, v in it['formulas'].items()))
        if it.get('why'): D.append(f'- 润色说明：{esc(str(it["why"]))}')
        for x in hard.get(S, []): D.append(f'- **硬问题**：{x}')
        for x in soft.get(S, []): D.append(f'- 脚本提示：{x}')
        if it.get('review'): D.append(f'- 复查：{esc(str(it["review"]))}')
        D.append('')
    else:
        note = '漏交' if S not in res else ('复查后退回原文' if res[S].get('review') else '没改')
        D += [f'## {S}　{note}', '', text_of[S], '']
open(os.path.join(a.out, a.prefix + '对照.md'), 'w', encoding='utf-8').write('\n'.join(D))

H = ['<!doctype html><meta charset="utf-8"><title>润色对照</title>',
     '<style>body{font-family:-apple-system,"PingFang SC",sans-serif;max-width:980px;margin:24px auto;padding:0 16px;line-height:1.75;color:#222}'
     'del{color:#b00020;background:#fde7ea;text-decoration:line-through}ins{color:#0a6b2d;background:#e3f6e8;text-decoration:none}'
     '.p{margin:14px 0;padding:10px 14px;border:1px solid #ddd;border-radius:8px}.s{font-weight:600;color:#555}.w{color:#777;font-size:13px}'
     '.same{color:#999}.tag{color:#aaa;font-size:13px;margin:6px 0}.hard{color:#b00020;font-size:13px}</style>',
     f'<h2>润色对照（{in_rng[0]}～{in_rng[-1]}）</h2><p class="w">红色划线是删掉的字，绿色是新加的字。正文 {len(body)} 段，改了 {len(changed)} 段。</p>']
for S in in_rng:
    if tag_of[S]: H.append(f'<div class="tag">{S}{html.escape(tag_of[S])}{html.escape(text_of[S][:80])}</div>'); continue
    if S in changed:
        extra = ''.join(f'<br><span class="hard">硬问题：{html.escape(x)}</span>' for x in hard.get(S, []))
        if res[S].get('review'): extra += f'<br><span class="w">复查：{html.escape(str(res[S]["review"]))}</span>'
        H.append(f'<div class="p"><span class="s">{S}</span>　<span class="w">{html.escape(str(res[S].get("why", "")))}</span><br>{inline_diff(text_of[S], esc(res[S]["new"].strip()), html_fmt)}{extra}</div>')
    else:
        H.append(f'<div class="p same"><span class="s">{S}</span>　<span class="w">{"漏交" if S not in res else "没改"}</span><br>{html.escape(text_of[S])}</div>')
open(os.path.join(a.out, a.prefix + '对照.html'), 'w', encoding='utf-8').write('\n'.join(H))

print(f'正文 {len(body)} 段，改 {len(changed)} 段；硬问题 {n_hard + len(missing) + len(bad_files)} 个（漏段 {len(missing)}）；有线索的 {len(soft)} 段'
      + (f'；术语有变化 {len(term_rows)} 段' if terms else '') + f'；名词有增减 {len(noun_rows)} 段 → {os.path.join(a.out, a.prefix + "核对单.md")}')
sys.exit(1 if (n_hard or missing or bad_files) else 0)

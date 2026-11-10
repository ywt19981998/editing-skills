#!/usr/bin/env python3
"""名词清单 → 带全部位置、相近写法排在一起的清单（给“归组判断”那一步用）。

模型列的清单只负责“全”；本脚本负责三件机械的事：
  1. 清理：去掉纯数字、“文献[n]”、开头的引出语；核对每条是否原样出现在所标段落（走样的列出来）。
  2. 找位置：每种写法在正文范围内全部出现的段落和次数（穷尽搜索，不靠模型记）。
  3. 排在一起：含同一个外文词的排一组，中文按末尾中心词排一组；另列“只差一两个字”的写法对。
不判断是不是同一个东西——那是模型下一步的事。

用法: group_terms.py 名词清单.md 段落.txt [图片文字.md] --end P124 -o 输出.md
"""
import sys, re, argparse, difflib, collections
ap = argparse.ArgumentParser()
ap.add_argument('lst'); ap.add_argument('paras'); ap.add_argument('fig', nargs='?'); ap.add_argument('--end'); ap.add_argument('-o', required=True)
a = ap.parse_args()
src = open(a.paras, encoding='utf-8').read()
clean = lambda v: re.sub(r'\[批注\d+>|<批注\d+\]|\[公式\]|\[图片\]', '', v)
paras = collections.OrderedDict((m.group(1), clean(m.group(2))) for m in re.finditer(r'^([PT]\d+): (.*)$', src, flags=re.M))
end = int(a.end.lstrip('PT')) if a.end else max(int(k[1:]) for k in paras)
paras = collections.OrderedDict((k, v) for k, v in paras.items() if int(k[1:]) <= end)
fig = open(a.fig, encoding='utf-8').read() if a.fig else ''
figsec = collections.OrderedDict()
if fig:
    cur = None
    for l in fig.split('\n'):
        m = re.search(r'图\s*\d+[-－]\d+', l) if l.startswith('#') else None
        if m: cur = re.sub(r'\s', '', m.group(0)); figsec[cur] = ''
        elif cur: figsec[cur] += l + '\n'
LEAD = re.compile(r'^(一种名为|一种新型?的?|一种用于|一种|一个新的|一个|名为|所谓|该类|这种|这些)')
items, bad = [], []
for l in open(a.lst, encoding='utf-8').read().split('\n'):
    if '｜' not in l: continue
    k, *ws = [x.strip() for x in l.split('｜')]
    for w in ws:
        if not w or w == '无' or re.fullmatch(r'[（(]?[0-9.\-+%=～~×·, ]+[)）]?', w) or re.fullmatch(r'文献\s*\[[\d,\-–、～ ]+\]', w): continue
        w2 = LEAD.sub('', w).strip()
        if len(w2) < 2: continue
        if k in paras and w2 not in paras[k] and re.sub(r'\s', '', w2) not in re.sub(r'\s', '', paras[k]): bad.append((k, w))
        items.append(w2)
terms = list(dict.fromkeys(items))

def positions(t):
    pos = [(k, v.count(t)) for k, v in paras.items() if t in v]
    figs = [f for f, txt in figsec.items() if t in txt]
    return pos, figs
info = {}
for t in terms:
    pos, figs = positions(t); n = sum(c for _, c in pos)
    if n == 0 and not figs:        # 清单里有、正文搜不到：多半是带了空格差异，按去空格再搜
        t0 = re.sub(r'\s', '', t); pos = [(k, re.sub(r'\s', '', v).count(t0)) for k, v in paras.items() if t0 in re.sub(r'\s', '', v)]; n = sum(c for _, c in pos)
    info[t] = (n, pos, figs)
def fmt(t):
    n, pos, figs = info[t]
    s = ' '.join(k + (f'×{c}' if c > 1 else '') for k, c in pos)
    if figs: s += ('；' if s else '') + '、'.join(figs)
    return n, s or '（正文里没搜到，核对写法）'

GENERIC = set('方法 模型 算法 数据 特征 结构 问题 能力 效果 结果 信息 过程 方式 部分 技术 策略 机制 任务 性能 表示 关系 类型 水平 指标 分析 学习 网络 矩阵 模块 数据集 损失 函数 向量 空间'.split())
def key(t):
    en = re.findall(r'[A-Za-z][A-Za-z0-9\-+]{1,}', t)
    if en: return '外文：' + max(en, key=len).upper()
    zh = re.sub(r'[（(][^）)]*[）)]|[^一-鿿]', '', t)
    for L in (3, 2):
        if len(zh) >= L and zh[-L:] in GENERIC: return '中心词：' + zh[-L:]
    return '中心词：' + (zh[-2:] if len(zh) >= 2 else zh)
groups = collections.OrderedDict()
for t in terms: groups.setdefault(key(t), []).append(t)

# 只差一两个字的写法对（不同组的也列，防止中心词不同被分开）
near = []
zh_terms = [t for t in terms if 3 <= len(t) <= 24]
for i, x in enumerate(zh_terms):
    for y in zh_terms[i + 1:]:
        if abs(len(x) - len(y)) > 3 or x in y or y in x and abs(len(x) - len(y)) > 2: continue
        r = difflib.SequenceMatcher(None, x, y).ratio()
        if r >= 0.8 and x != y:
            ops = [o for o in difflib.SequenceMatcher(None, x, y).get_opcodes() if o[0] != 'equal']
            # 只留“换了一两个字”或“中间多/少一两个字”的；一头多出一截的多是同一短语的长短截取，不算
            if sum(max(o[2] - o[1], o[4] - o[3]) for o in ops) <= 2 and not (x.startswith(y) or y.startswith(x) or x.endswith(y) or y.endswith(x)): near.append((r, x, y))
near.sort(reverse=True)

out = [f'# 名词清单（带位置）', '',
       f'模型列出 {len(items)} 条，去重后 {len(terms)} 种写法；位置由脚本在正文 P1–P{end} 里搜全，不靠记忆。',
       '同一组只表示“长得像、排在一起方便比”，是不是同一个东西要你判断。', '']
multi = [(k, v) for k, v in groups.items() if len(v) > 1]; single = [(k, v) for k, v in groups.items() if len(v) == 1]
multi.sort(key=lambda kv: (not kv[0].startswith('外文'), -len(kv[1])))
out += [f'## 一、排在一起的写法（{len(multi)} 组，{sum(len(v) for _, v in multi)} 种写法）', '']
for k, v in multi:
    out += [f'### {k}（{len(v)} 种写法）', '| 写法 | 次数 | 位置 |', '|---|---|---|']
    for t in sorted(v, key=lambda t: (len(t), t)):
        n, s = fmt(t); out.append(f'| {t} | {n} | {s} |')
    out.append('')
out += [f'## 二、只差一两个字的写法对（{len(near)} 对）', '', '| 写法甲 | 写法乙 |', '|---|---|']
for r, x, y in near: out.append(f'| {x}（{fmt(x)[1]}） | {y}（{fmt(y)[1]}） |')
out += ['', f'## 三、单独一种写法的（{len(single)} 种，没有相近写法；仍可能和别的词指同一个东西）', '', '| 写法 | 次数 | 位置 |', '|---|---|---|']
for k, v in single:
    n, s = fmt(v[0]); out.append(f'| {v[0]} | {n} | {s} |')
if bad:
    out += ['', f'## 四、清单里标的段落里找不到原文的（{len(bad)} 条，可能抄走样）', ''] + [f'- {k}｜{w}' for k, w in bad]
open(a.o, 'w', encoding='utf-8').write('\n'.join(out) + '\n')
print(f'清单 {len(items)} 条 → {len(terms)} 种写法；排在一起 {len(multi)} 组；只差一两个字 {len(near)} 对；单独 {len(single)} 种；走样 {len(bad)} 条 → {a.o}')

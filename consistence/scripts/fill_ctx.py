#!/usr/bin/env python3
"""把“位置＋旧＋新”的短格式改法补成完整修订规则（自动取 ctx），写完跑预检。

用法: fill_ctx.py <底稿docx> <改法1.json> [改法2.json ...] -o <规则.json> [--before 已有规则.json ...]
  --before  先于本批写入的完整规则（机械规则、上一步的修订规则），按顺序给；取 ctx 时会避开它们改过的字

短格式（子代理写）:
  {"edits":    [{"p": "P58", "old": "紧致性", "new": "紧凑性", "row": "C12"}],
   "comments": [{"p": "P77", "target": "68.76", "text": "【请作者核实】……"}]}
  p       docx段落.txt 里的段落号（P58）或表格行号（T43）。
          同一种改法出现在多段时写成列表 ["P12", "P15", "P20"]：这些段里的 old 每一处都改
  old     要改的原文。p 是单个段落号时，old 在该段里只能出现一次，出现多次就往前后多带几个字；
          写入时只标出真正改动的字
  target  批注挂在哪几个字上，在该段里只出现一次
  row     本章名词表的行号（可不写），用来对上是哪一行的结论
  可选键照抄到规则里: "decision"（D 项）、"whole"（专名、数值整词替换）
  已经带 "ctx" 的条目原样保留、排在最前（脚注等要手写的情况）；其余按段落先后排好再补
  整段文字在稿里不止一处（章首目录里的节题、几张表里同名的行）时，ctx 取整段并加 "exact"，按出现顺序落位：
  前面同文的段落要么本批也改，要么这一条手写

退出码: 有条目没补上、或预检有错误为 1。
"""
import sys, os, json, zipfile, subprocess
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
PY = os.path.expanduser('~/.venvs/book-review/bin/python')
CHECK = os.path.expanduser('~/.claude/skills/tongdu/scripts/check_rules.py')
MAXLEN = 40


def q(t):
    return f'{{{W}}}{t}'


def load_doc(path):
    doc = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    label, n = {}, 0
    for el in doc.find(q('body')):          # 编号方式与 dump_docx.py、check_rules.py 相同
        if el.tag == q('p'):
            n += 1
            for x in el.iter(q('p')):
                label[x] = n
        elif el.tag == q('tbl'):
            for tr in el.iter(q('tr')):
                n += 1
                for x in tr.iter(q('p')):
                    label[x] = n
    ps = list(doc.iter(q('p')))
    full = [''.join(t.text or '' for t in p.iter(q('t'))) for p in ps]
    direct = [''.join(r.find(q('t')).text or '' for r in p.findall(q('r')) if len(r.findall(q('t'))) == 1) for p in ps]
    by_no = {}
    for k, p in enumerate(ps):
        by_no.setdefault(label.get(p), []).append(k)
    return full, direct, by_no


def changed_span(old, new, whole):
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]:
        a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1 - b] == new[-1 - b]:
        b += 1
    return (0, 0) if whole else (a, b)


class Sim:
    """按 check_rules 的办法模拟：修订写入后，删掉和插入的字都不再算在段落文字里"""

    def __init__(self, full, direct):
        self.F, self.D = full, list(direct)
        self.cuts = {}                       # 段落下标 -> 已改动处在当前文字里的位置

    def apply(self, k, j, old, new, whole=False):
        a, b = changed_span(old, new, whole)
        t = self.D[k]
        gone = len(old) - a - b
        self.D[k] = t[:j + a] + t[j + len(old) - b:]
        cs = [c - gone if c > j + a else c for c in self.cuts.get(k, [])]
        self.cuts[k] = cs + [j + a]

    def apply_full(self, e):
        """已有完整规则：找到就模拟，找不到不报（交给预检）"""
        for k, t in enumerate(self.D):
            ci = t.find(e['ctx'])
            if ci < 0:
                continue
            j = t.find(e['old'], ci)
            if 0 <= j <= ci + len(e['ctx']):
                self.apply(k, j, e['old'], e['new'], e.get('whole', False))
            return

    def pick_ctx(self, texts, k, j, n, cuts=()):
        """在 texts[k] 里取包含 [j, j+n) 的最短上下文：全稿只出现一次，尽量不跨已改动处，优先向左取"""
        t = texts[k]
        for avoid in (True, False):
            for extra in range(4, MAXLEN):
                for left in range(extra, -1, -1):
                    s, e = max(0, j - left), min(len(t), j + n + extra - left)
                    c = t[s:e]
                    if len(c) < min(len(t), n + 4):
                        continue
                    if avoid and any(s < x < e and not (j <= x <= j + n) for x in cuts):
                        continue
                    if sum(x.count(c) for x in texts) != 1 or sum(x.count(c) for x in self.F) > 1:
                        continue
                    return c
        return None


def para_no(p):
    p = str(p).strip()
    if not p or p[0] not in 'PT' or not p[1:].isdigit():
        return None
    return int(p[1:])


def main():
    args = sys.argv[1:]
    if len(args) < 2 or '-o' not in args:
        print(__doc__)
        sys.exit(2)
    i = args.index('-o'); out = args[i + 1]; del args[i:i + 2]
    before = []
    if '--before' in args:
        i = args.index('--before'); before = args[i + 1:]; del args[i:]
    src, shorts = args[0], args[1:]
    full, direct, by_no = load_doc(src)
    sim = Sim(full, direct)
    for f in before:
        for e in json.load(open(f, encoding='utf-8')).get('edits', []):
            sim.apply_full(e)

    res = {'edits': [], 'comments': []}
    errs, todo, same = [], [], 0
    for f in shorts:
        spec = json.load(open(f, encoding='utf-8'))
        name = os.path.basename(f)
        flat = []                              # p 是列表的拆成每段一条，带 each 标记（该段每一处都改）
        for e in spec.get('edits', []):
            if isinstance(e.get('p'), list):
                flat += [dict(e, p=x, each=True) for x in e['p']]
            else:
                flat.append(e)
        spec['edits'] = flat
        for n_, c in enumerate(spec.get('comments', []), 1):
            tag = f'{name} comment#{n_}（{c.get("p", "")}「{c.get("target", "")[:12]}」）'
            if 'ctx' in c:
                res['comments'].append(c); continue
            no = para_no(c.get('p', ''))
            if no is None or no not in by_no:
                errs.append(f'{tag}：p 写成 P58 或 T43 这样的段落号，且稿里要有这一段'); continue
            hits = [(k, full[k].find(c['target'])) for k in by_no[no] if c.get('target') and c['target'] in full[k]]
            total = sum(full[k].count(c['target']) for k in by_no[no]) if c.get('target') else 0
            if total == 0:
                errs.append(f'{tag}：target 在该段找不到'); continue
            if total > 1:
                errs.append(f'{tag}：target 在该段出现 {total} 次，往前后多带几个字'); continue
            k, j = hits[0]
            ctx = sim.pick_ctx(full, k, j, len(c['target']))
            if ctx is None:
                errs.append(f'{tag}：取不到全稿唯一的上下文，这条手写 ctx'); continue
            r = {'ctx': ctx}; r.update(c)
            res['comments'].append(r)
        for n_, e in enumerate(spec.get('edits', []), 1):
            todo.append((f'{name} edit#{n_}（{e.get("p", "")}「{e.get("old", "")[:12]}」→「{e.get("new", "")[:12]}」）', e))
    for tag, e in [x for x in todo if 'ctx' in x[1]]:
        sim.apply_full(e); res['edits'].append(e)
    rest = [x for x in todo if 'ctx' not in x[1]]
    rest.sort(key=lambda x: para_no(x[1].get('p', '')) or 0)          # 稳定排序：同一段内保持原来的先后
    if True:
        for tag, e in rest:
            no = para_no(e.get('p', ''))
            if no is None or no not in by_no:
                errs.append(f'{tag}：p 写成 P58 或 T43 这样的段落号，且稿里要有这一段'); continue
            old, new = e.get('old', ''), e.get('new', '')
            if not old:
                errs.append(f'{tag}：old 不能为空，纯插入时把插入处相邻的几个字带进 old 和 new'); continue
            if old == new:
                errs.append(f'{tag}：old 与 new 相同'); continue
            each = e.get('each')
            total = sum(sim.D[k].count(old) for k in by_no[no])
            if total == 0:
                why = '前面的规则已经改过这里' if any(old in direct[k] for k in by_no[no]) else \
                      ('落在超链接或原有修订里，改不了，换成批注' if any(old in full[k] for k in by_no[no]) else '该段没有这几个字')
                errs.append(f'{tag}：old 找不到（{why}）'); continue
            if total > 1 and not each:
                errs.append(f'{tag}：old 在该段出现 {total} 次，往前后多带几个字，带到本段唯一'); continue
            a, b = changed_span(old, new, e.get('whole', False))
            for k in [k for k in by_no[no] if old in sim.D[k]]:
                pos = 0
                while True:
                    j = sim.D[k].find(old, pos)
                    if j < 0:
                        break
                    ctx = sim.pick_ctx(sim.D, k, j, len(old), sim.cuts.get(k, ()))
                    r = {k_: v for k_, v in e.items() if k_ not in ('exact', 'each')}
                    if ctx is None:                                   # 整段同文：取整段、exact，按出现顺序落位
                        ctx = sim.D[k]
                        if any(sim.D[k2] == ctx for k2 in range(k)) or sim.D[k].count(old) > 1:
                            errs.append(f'{tag}：整段文字在前面还有相同的一处没改，这条手写 ctx（或把前面那处一并改）'); break
                        r['exact'] = True; same += 1
                    res['edits'].append(dict({'ctx': ctx}, **r))
                    sim.apply(k, j, old, new, e.get('whole', False))
                    pos = j + a + b                                   # 跳过 old 里没动的头尾，接着找本段下一处

    json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'补 ctx：修订 {len(res["edits"])} 条、批注 {len(res["comments"])} 条 → {out}；没补上 {len(errs)} 条'
          + (f'；整段同文按顺序落位 {same} 条' if same else ''))
    for x in errs:
        print('没补上：' + x)
    r = subprocess.run([PY, CHECK, src] + before + [out], capture_output=True, text=True)
    lines = (r.stdout + r.stderr).strip().splitlines()
    print('预检：' + ('通过' if r.returncode == 0 else '有错误'))
    print('\n'.join(l for l in lines if l.startswith(('错误', '提醒', '修订', '疑似', '    '))))
    sys.exit(1 if (errs or r.returncode) else 0)


if __name__ == '__main__':
    main()

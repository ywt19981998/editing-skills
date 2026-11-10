#!/usr/bin/env python3
"""把“已定的统一替换”写成机械规则：每处自动取全文唯一的 ctx，写完跑预检。

哪些词该换、哪几处不该换，由主会话看过《准备报告》第四节后决定；本脚本只负责把决定写成规则。
用法: mech_rules.py <书的工作目录> <章号> --add 旧=新 [--add 旧=新 ...] [--whole] [--only P41,T43] [--skip P94,P56] [-o 规则.json]
  --add    正文范围内该词每一处都换（配合 --only / --skip 限定段落）
  --whole  修订显示为整词整值替换（专名、数值、编号用）
  默认追加到 <书>/第NN章-机械规则.json（没有就新建）
单处的特殊改法（旧、新不是同一个词）直接在 json 里手写一条 {"ctx","old","new"}。
"""
import sys, os, re, json, argparse, subprocess
PY = os.path.expanduser('~/.venvs/book-review/bin/python')
S2 = os.path.expanduser('~/.claude/skills/tongdu/scripts')
ap = argparse.ArgumentParser()
ap.add_argument('book'); ap.add_argument('chapter'); ap.add_argument('--add', action='append', default=[])
ap.add_argument('--whole', action='store_true'); ap.add_argument('--only', default=''); ap.add_argument('--skip', default=''); ap.add_argument('-o')
a = ap.parse_args()
book = os.path.abspath(a.book); n = int(a.chapter); NN = f'{n:02d}'
cfg = json.load(open(os.path.join(book, '本书配置.json'), encoding='utf-8'))
src = (cfg.get('原稿覆盖') or {}).get(NN) or cfg['原稿'].replace('{n}', str(n)).replace('{NN}', NN)
txt = open(os.path.join(book, 'source', f'第{NN}章-docx段落.txt'), encoding='utf-8').read()
paras = [(m.group(1), re.sub(r'\[批注\d+>|<批注\d+\]', '', m.group(2))) for m in re.finditer(r'^([PT]\d+): (.*)$', txt, flags=re.M)]
ref = next((int(k[1:]) for k, v in paras if v.strip() in ('参考文献', '本章参考文献')), None)
if ref: paras = [(k, v) for k, v in paras if int(k[1:]) < ref]
full = '\n'.join(v for _, v in paras)
only = set(x for x in a.only.split(',') if x); skip = set(x for x in a.skip.split(',') if x)
pairs = [x.split('=', 1) for x in a.add]
olds = [p[0] for p in pairs]
new_rules, fail = [], []
for old, new in pairs:
    for k, v in paras:
        if (only and k not in only) or k in skip: continue
        for m in re.finditer(re.escape(old), v):
            i, j = m.start(), m.end(); ctx = None
            for L in range(2, 24):                      # 由短到长找全文唯一、且不夹带别的待换词的上下文
                for left in range(L, -1, -1):           # 优先向左取：左边的字不会被本批规则改掉
                    s, e = max(0, i - left), min(len(v), j + (L - left)); c = v[s:e]
                    if full.count(c) != 1: continue
                    rest = c[:i - s] + '\0' + c[j - s:]
                    if any(o in rest for o in olds): continue
                    ctx = c; break
                if ctx: break
            if ctx is None: fail.append(f'{k}: …{v[max(0, i - 10):j + 10]}…'); continue
            r = {'ctx': ctx, 'old': old, 'new': new}
            if a.whole: r['whole'] = True
            new_rules.append((k, r))
out = a.o or os.path.join(book, f'第{NN}章-机械规则.json')
d = json.load(open(out, encoding='utf-8')) if os.path.exists(out) else {'edits': [], 'comments': []}
have = {(e['ctx'], e['old']) for e in d['edits']}
added = [(k, r) for k, r in new_rules if (r['ctx'], r['old']) not in have]
d['edits'] += [r for _, r in added]
json.dump(d, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
for k, r in added: print(f'{k}  {r["old"]}→{r["new"]}  ctx「{r["ctx"]}」')
print(f'新增 {len(added)} 条，文件共 {len(d["edits"])} 条 → {out}')
if fail: print('没取到唯一上下文、要手写的：\n  ' + '\n  '.join(fail))
r = subprocess.run([PY, f'{S2}/check_rules.py', src, out], capture_output=True, text=True)
msg = [l for l in (r.stdout + r.stderr).splitlines() if 'ctx 不唯一' not in l]
print('预检：' + ('通过' if r.returncode == 0 else '有错误') + '\n' + '\n'.join(msg[-12:]))
sys.exit(1 if (r.returncode or fail) else 0)

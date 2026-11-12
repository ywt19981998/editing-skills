#!/usr/bin/env python3
"""把复查结论并进润色结果，出终稿：通过的照用；有问题的按复查给的最小修正改；修不过来的整段退回原文。

用法: apply_review.py 正文.txt 终稿.json --polish 结果1.json [结果2.json ...] --review 复查1.json [复查2.json ...]
                      [--allow-unreviewed]
产物: 终稿.json（格式同润色结果，直接给核对、写回脚本用）和同目录下的 复查汇总.md
      有修正没对上、或有改动段没有复查结论时退出码为 1（加 --allow-unreviewed 时没复查的段照润色稿用）
"""
import sys, os, json, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx_units import read_slines

ap = argparse.ArgumentParser()
ap.add_argument('txt'); ap.add_argument('out')
ap.add_argument('--polish', nargs='+', required=True); ap.add_argument('--review', nargs='+', required=True)
ap.add_argument('--allow-unreviewed', action='store_true')
a = ap.parse_args()

slines = read_slines(a.txt)
text_of = {sid: text for sid, tag, text in slines}
body = [sid for sid, tag, text in slines if not tag]
def load(files):
    d = {}
    for f in files:
        for it in json.load(open(f, encoding='utf-8')):
            if isinstance(it, dict) and it.get('S'): d[it['S']] = it
    return d
polish, review = load(a.polish), load(a.review)

out, errs, fixed, reverted, asks = [], [], [], [], []
n_changed = n_pass = 0
for S in body:
    it = polish.get(S)
    if it is None: continue
    old = text_of[S].strip(); new = (it.get('new') or '').strip()
    if not (it.get('changed') and new and new != old):
        e = {'S': S, 'changed': False}
        if it.get('why'): e['why'] = it['why']
        out.append(e); continue
    n_changed += 1
    rv = review.get(S)
    if rv is None:
        if not a.allow_unreviewed: errs.append(f'{S} 改了，但没有复查结论')
        out.append(dict(it, review='没有复查')); continue
    if rv.get('ask'): asks.append((S, str(rv['ask'])))
    if rv.get('revert'):
        why = str(rv.get('why') or '')
        e = {'S': S, 'changed': False, 'review': '整段退回原文：' + why, 'polished': new}
        if rv.get('ask'): e['ask'] = rv['ask']
        out.append(e); reverted.append((S, why)); continue
    cur = new; applied = []
    for fx in rv.get('fixes') or []:
        o = str(fx.get('old') or ''); n = '' if fx.get('new') is None else str(fx['new'])
        c = cur.count(o) if o else 0
        if c != 1:
            errs.append(f'{S} 的一处修正没对上：“{o}”在润色稿里出现 {c} 次（要正好 1 次）。把复查文件里这条的 old 改成能唯一找到的原话再跑'); continue
        cur = cur.replace(o, n); applied.append((o, n, str(fx.get('why') or '')))
    if rv.get('ok') is False and not rv.get('fixes'):
        errs.append(f'{S} 复查说有问题，却没给修正，也没说退回')
    if cur.strip() == old:
        e = {'S': S, 'changed': False, 'review': '修正后与原文相同，等于退回', 'polished': new}
        reverted.append((S, '修正后与原文相同'))
    else:
        e = dict(it); e['new'] = cur
        if applied:
            e['review'] = f'修了 {len(applied)} 处：' + '；'.join(f'“{o}”→“{n}”（{w}）' for o, n, w in applied)
            fixed += [(S, o, n, w) for o, n, w in applied]
        else:
            e['review'] = '通过'; n_pass += 1
    if rv.get('ask'): e['ask'] = rv['ask']
    out.append(e)
extra = [S for S in review if S not in polish or not polish[S].get('changed')]

json.dump(out, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
L = ['# 复查汇总', '',
     f'- 润色改了 {n_changed} 段：复查通过 {n_pass} 段，修了 {len(fixed)} 处（{len({x[0] for x in fixed})} 段），整段退回原文 {len(reverted)} 段',
     f'- 要编辑看一眼的 {len(asks)} 处；没处理掉的问题 {len(errs)} 个', '']
if errs: L += ['## 没处理掉的（要先解决）', ''] + [f'- {x}' for x in errs] + ['']
if fixed: L += ['## 复查修掉的', '', '| 段 | 润色稿 | 改成 | 原因 |', '|---|---|---|---|'] + [f'| {S} | {o} | {n or "（删去）"} | {w} |' for S, o, n, w in fixed] + ['']
hand = [x for x in fixed if '误译' in x[3] or '一致性' in x[3]]
if hand: L += ['## 留给一致性审查的名词（润色想改、已恢复原词）', '', '| 段 | 原词（已恢复） | 润色想改成 | 说明 |', '|---|---|---|---|'] + [f'| {S} | {n} | {o} | {w} |' for S, o, n, w in hand] + ['']
if reverted: L += ['## 整段退回原文的', ''] + [f'- {S}　{w}' for S, w in reverted] + ['']
if asks: L += ['## 要编辑看一眼的', ''] + [f'- {S}　{w}' for S, w in asks] + ['']
if extra: L += ['## 复查里多出来的条目（没用上）', '', '、'.join(extra), '']
open(os.path.join(os.path.dirname(os.path.abspath(a.out)), '复查汇总.md'), 'w', encoding='utf-8').write('\n'.join(L))
print(f'润色改 {n_changed} 段：通过 {n_pass}，修 {len(fixed)} 处，退回 {len(reverted)} 段，要编辑看 {len(asks)} 处；没处理掉的问题 {len(errs)} 个 → {a.out}')
for x in errs: print('  -', x)
sys.exit(1 if errs else 0)

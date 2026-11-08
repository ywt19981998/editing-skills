# 决策单：把待定决策表和各章修订规则汇总成一张“按决策拍板”的单子，每个决策列出涉及几处、改成什么样
# 用法: decision_sheet.py 待定决策.md 输出.md 第02章-修订规则.json 第03章-修订规则.json ...
# 用户看完只需回复“D-12 选B”“其余按建议”，再用 apply_decisions.py 重新生成修订稿
import sys, os, re, json
from collections import defaultdict

md_path, out_path, rule_paths = sys.argv[1], sys.argv[2], sys.argv[3:]

def cells(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]

rows = []
for line in open(md_path, encoding='utf-8'):
    if re.match(r'^\|\s*D-\d+', line):
        c = cells(line)
        if len(c) >= 6:
            rows.append(dict(id=c[0], q=c[1], a=c[2], b=c[3], scope=c[4], status=c[5]))

def chap_of(path):
    m = re.search(r'(第\d+章|文前)', os.path.basename(path))
    return m.group(1) if m else os.path.basename(path)

uses = defaultdict(list)       # 决策编号 -> [(章, edit)]
notes = defaultdict(list)      # 决策编号 -> [(章, 批注)]
for rp in rule_paths:
    spec = json.load(open(rp, encoding='utf-8'))
    for e in spec.get('edits', []):
        if e.get('decision'): uses[e['decision']].append((chap_of(rp), e))
    for c in spec.get('comments', []):
        m = re.match(r'^【建议·(D-\d+)】', c.get('text', ''))
        if m: notes[m.group(1)].append((chap_of(rp), c))

def decided(r):
    return bool(re.search(r'(?<!待)用户确认|用户已要求', r['status']))

def show(e):
    old, new = e['old'], e['new']
    if len(old) > 40 or len(new) > 40:
        return f'「{old[:30]}…」→「{new[:30]}…」' if new else f'删除「{old[:30]}…」'
    return f'「{old}」→「{new}」' if new else f'删除「{old}」'

pending = [r for r in rows if not decided(r)]
done = [r for r in rows if decided(r)]
L = [f'# 决策单\n',
     f'来源：`{os.path.basename(md_path)}`；修订规则：' + '、'.join(f'`{os.path.basename(p)}`' for p in rule_paths) + '\n',
     f'**待你拍板 {len(pending)} 项**，已确认 {len(done)} 项。\n',
     '回复方式：直接说“D-12 选B”“D-16 按建议”，或者“其余都按建议”。选 A 的，相关修订改为确定修订，修订人从“Claude（建议）”变成“Claude”，【建议·D-xx】批注删掉；'
     '选 B 的，撤回 A 方案的修订，按 B 方案在每一处重新写修订。\n',
     '> 拍板后再开始通读修订稿。通读到一半再改选的，只能在你手上那份稿子上补改。\n']

def block(r, k):
    us = uses.get(r['id'], [])
    by_ch = defaultdict(int)
    for ch, _ in us: by_ch[ch] += 1
    L.append(f'## {k}. {r["id"]}　{r["q"]}\n')
    if us:
        where = '、'.join(f'{ch} {n} 处' for ch, n in by_ch.items())
        L.append(f'- **A（建议，已写入修订稿）**：{r["a"]}')
        L.append(f'  - 已写成修订：{where}，共 {len(us)} 处。例：' + '；'.join(f'{ch} {show(e)}' for ch, e in us[:3]) + ('……' if len(us) > 3 else ''))
    else:
        L.append(f'- **A（建议）**：{r["a"]}')
        L.append('  - 目前只有批注，没有写成修订')
    L.append(f'- **B**：{r["b"]}')
    L.append(f'- 涉及：{r["scope"]}')
    L.append(f'- 状态：{r["status"]}\n')

for k, r in enumerate(pending, 1): block(r, k)
if done:
    L.append('## 已确认\n\n| 编号 | 问题 | 结论 |\n|---|---|---|')
    for r in done:
        L.append(f'| {r["id"]} | {r["q"]} | {r["status"]} |')
orphan = sorted(set(uses) - {r['id'] for r in rows})
if orphan:
    L.append('\n## 注意：修订规则里有、待定决策表里没登记的编号\n\n' + '、'.join(orphan))
open(out_path, 'w', encoding='utf-8').write('\n'.join(L) + '\n')
print(f'待拍板 {len(pending)} 项，已确认 {len(done)} 项 → {out_path}')

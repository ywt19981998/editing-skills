# 按用户拍板的决策重新生成修订稿：从底稿 + 修订规则 + 决策选择重新写一遍，原规则文件不动
# 用法: apply_decisions.py 底稿.docx 修订规则.json 决策选择.json 输出.docx [--chapter 第03章]
# 决策选择.json 格式：
#   {"D-16": {"choice": "A"},
#    "D-12": {"choice": "B", "edits": [{"chapter": "第03章", "ctx": "...", "old": "...", "new": "..."}],
#             "comments": [{"chapter": "第03章", "ctx": "...", "target": "...", "text": "..."}]}}
# 选 A：该决策的修订去掉 decision 标记（修订人从“Claude（建议）”变为“Claude”），删去【建议·D-xx】批注
# 选 B：撤回该决策的全部修订和【建议·D-xx】批注，换成 B 方案的修订（只取 chapter 与本章相符或未注明章的条目）
# 未出现在决策选择里的决策保持原样。合成后的规则另存为 输出.rules.json，再调用一致性技能的 revise_docx.py 写入
import sys, os, re, json, zipfile, subprocess, argparse
from lxml import etree

ap = argparse.ArgumentParser()
ap.add_argument('base'); ap.add_argument('rules'); ap.add_argument('choices'); ap.add_argument('out')
ap.add_argument('--chapter', help='本章标识，如“第03章”；默认从规则文件名取')
a = ap.parse_args()

REVISE = os.path.expanduser('~/.claude/skills/editorial-term-consistency-review/scripts/revise_docx.py')
spec = json.load(open(a.rules, encoding='utf-8'))
choices = json.load(open(a.choices, encoding='utf-8'))
m = re.search(r'(第\d+章|文前)', os.path.basename(a.rules))
chapter = a.chapter or (m.group(1) if m else None)

def mine(x):
    return not x.get('chapter') or x['chapter'] == chapter

edits, comments, log = [], [], []
for e in spec.get('edits', []):
    d = e.get('decision')
    c = choices.get(d, {}).get('choice') if d else None
    if c == 'A':
        e = {k: v for k, v in e.items() if k != 'decision'}
    elif c == 'B':
        continue
    edits.append(e)
for cm in spec.get('comments', []):
    mm = re.match(r'^【建议·(D-\d+)】', cm.get('text', ''))
    if mm and choices.get(mm.group(1), {}).get('choice') in ('A', 'B'):
        continue
    comments.append(cm)
for d, ch in choices.items():
    n_old = sum(1 for e in spec.get('edits', []) if e.get('decision') == d)
    if ch.get('choice') == 'B':
        be = [{k: v for k, v in e.items() if k != 'chapter'} for e in ch.get('edits', []) if mine(e)]
        bc = [{k: v for k, v in c.items() if k != 'chapter'} for c in ch.get('comments', []) if mine(c)]
        edits += be; comments += bc
        log.append(f'{d} 选B：撤回 {n_old} 处建议修订，按 B 方案写入 {len(be)} 处修订、{len(bc)} 条批注')
    elif ch.get('choice') == 'A':
        log.append(f'{d} 选A：{n_old} 处修订改为确定修订')

# ---- 预检：每条修订都能在底稿里定位，且同一段内的修改不重叠 ----
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
doc = etree.fromstring(zipfile.ZipFile(a.base).read('word/document.xml'))
paras = [''.join(t.text or '' for t in p.iter(f'{{{W}}}t')) for p in doc.iter(f'{{{W}}}p')]
problems, spans = [], {}
for k, e in enumerate(edits):
    pi = next((i for i, t in enumerate(paras) if e['ctx'] in t), None)
    if pi is None:
        problems.append(f'找不到段落：ctx=「{e["ctx"][:30]}」'); continue
    t = paras[pi]; ci = t.find(e['ctx'])
    i = t.find(e['old'], ci)
    if i < 0 or i > ci + len(e['ctx']): i = t.find(e['old'])
    if i < 0:
        problems.append(f'段落里找不到要改的字：「{e["old"][:30]}」（ctx=「{e["ctx"][:20]}」）'); continue
    for (j, s, t2, o) in spans.get(pi, []):
        if i < t2 and s < i + len(e['old']):
            problems.append(f'修改重叠：「{e["old"][:20]}」与「{o[:20]}」改的是同一段文字')
    spans.setdefault(pi, []).append((k, i, i + len(e['old']), e['old']))
for c in comments:
    if not any(c['ctx'] in t for t in paras):
        problems.append(f'批注找不到段落：ctx=「{c["ctx"][:30]}」')
if problems:
    print('预检未通过，没有生成文件：'); [print(' -', p) for p in problems]; sys.exit(1)

out_rules = os.path.splitext(a.out)[0] + '.rules.json'
merged = dict(spec); merged['edits'] = edits; merged['comments'] = comments
json.dump(merged, open(out_rules, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
for x in log: print(x)
r = subprocess.run([sys.executable, REVISE, a.base, a.out, out_rules])
sys.exit(r.returncode)

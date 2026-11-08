# 修订规则预检（可合并多份规则）：写入 Word 之前查 ctx 是否唯一、old/target 能否找到、前后修改是否互相踩到
# 用法: check_rules.py 底稿.docx 规则1.json [规则2.json ...] [--merge 合并规则.json]
# 例：check_rules.py 第3章底稿.docx 第03章-修订规则.json 通读-修订规则.json --merge 第03章-合并规则.json
# 按 revise_docx 的顺序模拟：先放全部批注（看原文），再逐条修订（前一条改过的字后一条就看不到了）
# 有错误时退出码为 1，不写合并文件
# 带 "part": "footnote" 的规则在脚注里查
import sys, json, zipfile
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
args = sys.argv[1:]
merge = None
if '--merge' in args:
    i = args.index('--merge'); merge = args[i + 1]; del args[i:i + 2]
src, files = args[0], args[1:]
zf = zipfile.ZipFile(src)
doc = etree.fromstring(zf.read('word/document.xml'))
fn = etree.fromstring(zf.read('word/footnotes.xml')) if 'word/footnotes.xml' in zf.namelist() else None
plabel, n = {}, 0
for el in doc.find(q('body')):
    if el.tag == q('p'):
        n += 1
        for x in el.iter(q('p')): plabel[x] = n
    elif el.tag == q('tbl'):
        for tr in el.iter(q('tr')):
            n += 1
            for x in tr.iter(q('p')): plabel[x] = n
ps = list(doc.iter(q('p')))
full = [''.join(t.text or '' for t in p.iter(q('t'))) for p in ps]          # 批注定位所用文字
direct = [''.join(r.find(q('t')).text or '' for r in p.findall(q('r')) if len(r.findall(q('t'))) == 1) for p in ps]   # 修订定位所用文字
fps = list(fn.iter(q('p'))) if fn is not None else []
ffull = [''.join(t.text or '' for t in p.iter(q('t'))) for p in fps]
fdirect = [''.join(r.find(q('t')).text or '' for r in p.findall(q('r')) if len(r.findall(q('t'))) == 1) for p in fps]
def pick(x):
    """按规则的 part 取正文或脚注的段落文字"""
    return (ffull, fdirect, '脚注') if x.get('part') == 'footnote' else (full, direct, '')
TAGS = ('【明确问题】', '【待确认】', '【请作者核实】', '【请作者提供出处】', '【说明】', '【建议', '【还原原稿】')

errs, warns = [], []
out = {'edits': [], 'comments': [], 'inserts': []}
specs = []
for f in files:
    spec = json.load(open(f, encoding='utf-8'))
    for kind in ('edits', 'comments', 'inserts'): out[kind] += spec.get(kind, [])
    specs.append((f, spec))
for f, spec in specs:
    for i, c in enumerate(spec.get('comments', [])):
        F, _, where = pick(c)
        tag = f'{f} comment#{i + 1}{where}「{c["ctx"][:16]}」'
        hits = [k for k, t in enumerate(F) if c['ctx'] in t]
        if not hits: errs.append(f'{tag}：ctx 在稿中找不到'); continue
        if sum(t.count(c['ctx']) for t in F) > 1: warns.append(f'{tag}：ctx 不唯一，批注会落在第一处')
        if c['target'] not in F[hits[0]]: errs.append(f'{tag}：target「{c["target"][:16]}」不在该段')
        if not c['text'].startswith(TAGS): warns.append(f'{tag}：批注没有以【明确问题】【待确认】【请作者核实】【请作者提供出处】【说明】开头')
for f, spec in specs:
    for i, e in enumerate(spec.get('edits', [])):
        F, D, where = pick(e)
        tag = f'{f} edit#{i + 1}{where}「{e["ctx"][:16]}」'
        hits = [k for k, t in enumerate(D) if e['ctx'] in t]
        if not hits:
            errs.append(f'{tag}：ctx 在稿中找不到（或落在超链接、修订痕迹里）'); continue
        if sum(t.count(e['ctx']) for t in F) > 1: warns.append(f'{tag}：ctx 不唯一，修订会落在第一处' + ('' if where else f'（P{plabel.get(ps[hits[0]], "?")}）'))
        k = hits[0]; t = D[k]; ci = t.find(e['ctx'])
        j = t.find(e['old'], ci)
        if j < 0 or j > ci + len(e['ctx']):
            j = t.find(e['old'])
            errs.append(f'{tag}：old「{e["old"]}」不在 ctx 里' + ('，会改到本段第一处' if j >= 0 else '，本段也找不到')); continue
        if e['old'] == e['new']: warns.append(f'{tag}：old 与 new 相同')
        a = 0
        while a < min(len(e['old']), len(e['new'])) and e['old'][a] == e['new'][a]: a += 1
        b = 0
        while b < min(len(e['old']), len(e['new'])) - a and e['old'][-1 - b] == e['new'][-1 - b]: b += 1
        if e.get('whole'): a = b = 0
        D[k] = t[:j + a] + t[j + len(e['old']) - b:]   # 写入后删除、插入的字都不再算在段落文字里

for w in warns: print('提醒：' + w)
for e in errs: print('错误：' + e)
print(f'修订 {len(out["edits"])} 处、批注 {len(out["comments"])} 条；错误 {len(errs)}，提醒 {len(warns)}')
if errs: sys.exit(1)
if merge:
    if not out['inserts']: del out['inserts']
    scopes = {spec.get('scope_from') for _, spec in specs if spec.get('scope_from')}
    if len(scopes) == 1: out['scope_from'] = scopes.pop()   # 各规则同一范围起点时保留
    elif len(scopes) > 1: print('提醒：各规则 scope_from 不同，合并稿未写 scope_from，请手动处理')
    json.dump(out, open(merge, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'已合并 → {merge}')

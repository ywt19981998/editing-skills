# 把通读发现（段号 + 原文 + 改法/批注）生成 revise_docx 规则：ctx 自动取全文唯一，并避开一致性审查要改的文字
# 用法: make_rules.py 底稿.docx 通读发现.json 输出规则.json [--avoid 一致性修订规则.json]
#
# 通读发现.json 是一个列表，每条一处：
#   直接修改：{"P": "P24", "old": "表明，可以看到风速", "new": "表明，风速", "note": "“表明”与“可以看到”杂糅，删后者。", "cat": "病句"}
#   只批注：  {"P": "P24", "target": "垂直风切变", "text": "【请作者核实】……", "cat": "知识性疑点"}
#   可选字段：nth（本段第几个，从 0 起，同段同一文字出现多次时用）
#   可选字段：whole（true 时修订显示为整词整值替换，用于数值、专有名词、编号，如 GÉANT→Geant、9.56e-07→9.56×10⁻⁷）
# 2026-11-08 起默认不出【说明】批注（编辑：直接改的看修订就够了）；note 只留在通读发现.json 和判断记录里。
# 加 --with-notes 才在同一位置加“【说明】+note”批注（旧做法）。
import sys, json, zipfile, re, argparse
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
q = lambda t: f'{{{W}}}{t}'
ap = argparse.ArgumentParser()
ap.add_argument('base'); ap.add_argument('findings'); ap.add_argument('out'); ap.add_argument('--avoid'); ap.add_argument('--with-notes', action='store_true')
a = ap.parse_args()

doc = etree.fromstring(zipfile.ZipFile(a.base).read('word/document.xml'))
plist = list(doc.iter(q('p')))
texts = [''.join(t.text or '' for t in p.iter(q('t'))) for p in plist]
label, n = {}, 0      # P 号与 dump_docx.py 一致：正文段落和表格的每一行各占一个号
for el in doc.find(q('body')):
    if el.tag == q('p'):
        n += 1
        for x in el.iter(q('p')): label[x] = n
    elif el.tag == q('tbl'):
        for tr in el.iter(q('tr')):
            n += 1
            for x in tr.iter(q('p')): label[x] = n
# 合并写入时一致性修订先做，它要改的字不能落在通读规则的 ctx 里，否则定位会失败
forbid = [e['old'] for e in json.load(open(a.avoid, encoding='utf-8'))['edits']] if a.avoid else []

def locate(P, s, nth=0):
    num = int(str(P).lstrip('PpTt'))
    hit = 0
    for p, t in zip(plist, texts):
        if label.get(p) != num: continue
        ks = [m.start() for m in re.finditer(re.escape(s), t)]
        if hit + len(ks) <= nth: hit += len(ks); continue
        k = ks[nth - hit]
        bad = [(m.start(), m.end()) for f in forbid if f and f not in s for m in re.finditer(re.escape(f), t)]
        for pad in (0, 3, 6, 10, 16, 24, 40):
            for lo, hi in ((max(0, k - pad), k + len(s) + pad), (max(0, k - pad), k + len(s)), (k, k + len(s) + pad)):
                c = t[lo:hi]
                if c and sum(c in x for x in texts) == 1 and t.find(c) == lo and t.find(s, lo) == k \
                   and not any(x < hi and lo < y for x, y in bad):
                    return c
        return None
    return None

edits, comments, errs = [], [], []
for i, f in enumerate(json.load(open(a.findings, encoding='utf-8')), 1):
    tag = f"#{i} {f.get('P')}「{(f.get('old') or f.get('target') or '')[:20]}」"
    if 'old' in f:
        c = locate(f['P'], f['old'], f.get('nth', 0))
        if not c: errs.append(f'{tag}：段内找不到这段原文，或找不到不与一致性修订重叠的唯一定位'); continue
        edits.append(dict(ctx=c, old=f['old'], new=f['new'], category=f.get('cat', ''), source='通读', P=f['P']))
        if f.get('whole'): edits[-1]['whole'] = True
        if f.get('note') and a.with_notes:
            comments.append(dict(ctx=c, target=f['old'], text='【说明】' + f['note'], category=f.get('cat', ''), source='通读', P=f['P']))
    else:
        c = locate(f['P'], f['target'], f.get('nth', 0))
        if not c: errs.append(f'{tag}：段内找不到批注落点'); continue
        comments.append(dict(ctx=c, target=f['target'], text=f['text'], category=f.get('cat', ''), source='通读', P=f['P']))
if errs:
    print('有问题，没有生成规则：'); [print(' -', e) for e in errs]; sys.exit(1)
json.dump(dict(edits=edits, comments=comments), open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(f'修订 {len(edits)} 处，批注 {len(comments)} 条（其中【说明】{sum(c["text"].startswith("【说明】") for c in comments)} 条）→ {a.out}')

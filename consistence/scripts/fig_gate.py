#!/usr/bin/env python3
"""逐图核对的门禁：图片转写稿里的每个标签，都要在图文核对表里有一行（或写进“不核对”）。

用法:
  fig_gate.py labels <图片文字.md>                        # 只列出每张图抽到的标签，供核对时逐个过
  fig_gate.py gate <图片文字.md> <第NN章-图文核对.md> [--rules 第NN章-图文改法.json]

gate 查四件事，都要到 0:
  1. 图没核：转写稿里有、核对表里没有“## 图N-M”一节的图。
  2. 标签没核：转写稿里抽到的标签，在该图那一节（表格或“不核对：”一行）里找不到。
     不是标签的（描述性文字、无字的图像）写进该节的“不核对：”一行即可。
  3. 结论没写对：表格行的结论不是“一致／改正文／请改图／D-xx／待确认／正文未出现／范围外”开头。
  4. 判了没落：结论里有“→改”的行，图文改法里没有带这一行号（row）的修订；有“→批注”的行，没有带这一行号的批注。
退出码：任一项不为 0 时为 1。

标签怎么抽（转写稿按图片转写模板写，每张图一节“## 图N-M（P号，…）图题：…”）:
  图题、图内标题、各条“- 框／节点／标注／箭头／连线／图例／纵轴／横轴……：”下的每一项，
  取每项“（”之前的文字；“→”连起来的流程逐段拆开；“、”列举的逐个拆开；“无”“无图例”跳过。
"""
import sys, re, json

SKIP_KEYS = ('子图标号', '看不清', '坐标轴', '横轴刻度', '纵轴刻度', '其他标注', '输入图像')
OK_START = ('一致', '改正文', '请改图', 'D-', '待确认', '正文未出现', '范围外')


def clean(s):
    s = re.split(r'[（(]', s, maxsplit=1)[0]
    s = re.sub(r'^[^：:“”]{1,8}[：:]', '', s.strip())          # 去掉“左上小框：”“上行：”这类位置前缀
    s = s.strip(' 　-—“”"\'。；;，,')
    return s


def tidy(s):
    s = re.sub(r'(?<=[一-鿿])\s+(?=[一-鿿])', '', s.strip())     # 转写时按折行留的空格（“时序 编码器”）去掉
    return s


def split_items(text):
    out = []
    q = re.findall(r'“([^”]+)”', text)
    if q:                                                             # 有引号的，引号里才是图上的字
        return [tidy(x) for x in q if not re.fullmatch(r'[\d.,%~～\s+-]+', x)]
    text = re.sub(r'（[^（）]*）', lambda m: m.group(0).replace('、', '，').replace('；', '，').replace(';', '，'), text)   # 括注里的顿号不拆
    for part in re.split(r'→', text):
        for p in re.split(r'[、；;]', part):
            p = tidy(clean(p))
            if re.fullmatch(r'[\d.,%~～\s+-]+', p or '0'): continue          # 纯数字（刻度、柱顶数值）不逐个核
            if p and p not in ('无', '无图例') and not p.startswith('无'):
                out.append(p)
    return out


def labels(path):
    figs = {}
    cur = None; key = None
    for line in open(path, encoding='utf-8'):
        line = line.rstrip('\n')
        m = re.match(r'^##\s*(图\d+[-－—.]\d+).*?图题[:：]\s*(?:图\d+[-－—.]\d+\s*)?(.*)$', line)
        if m:
            cur = m.group(1).replace('－', '-').replace('—', '-')
            figs.setdefault(cur, [])
            title = re.sub(r'（第.张子图）$', '', m.group(2)).strip()
            if title: figs[cur].append(title)
            key = None
            continue
        if not cur: continue
        m = re.match(r'^-\s*([^：:]+)[：:]\s*(.*)$', line)
        if m:
            key, rest = m.group(1), m.group(2)
            if any(k in key for k in SKIP_KEYS): key = 'SKIP'; continue
            if rest: figs[cur] += split_items(key + '：' + rest if '标题' in key else rest)
            continue
        m = re.match(r'^\s+-\s*(.*)$', line)
        if m and key != 'SKIP':
            t = m.group(1)
            if '无文字' in t and not re.search(r'[一-鿿A-Za-z]{2,}（', t.split('无文字')[0]): continue
            figs[cur] += split_items(t)
    for f in figs:
        figs[f] = [x for x in dict.fromkeys(figs[f]) if len(x) >= 1]
    return figs


def sections(path):
    secs = {}; cur = None
    for line in open(path, encoding='utf-8'):
        m = re.match(r'^##\s*(图\d+[-－—.]\d+)', line)
        if m:
            cur = m.group(1).replace('－', '-').replace('—', '-'); secs.setdefault(cur, [])
            continue
        if line.startswith('## '): cur = None; continue
        if cur: secs[cur].append(line.rstrip('\n'))
    return secs


def norm(s):
    return re.sub(r'\s+', '', s)


def main():
    a = sys.argv[1:]
    if not a or a[0] not in ('labels', 'gate'):
        print(__doc__); sys.exit(2)
    figs = labels(a[1])
    if a[0] == 'labels':
        for f, ls in figs.items():
            print(f'{f}（{len(ls)} 个）：' + '｜'.join(ls))
        return
    secs = sections(a[2])
    rules = {'edits': [], 'comments': []}
    if '--rules' in a:
        for p in a[a.index('--rules') + 1:]:
            if p.startswith('--'): break
            d = json.load(open(p, encoding='utf-8'))
            rules['edits'] += d.get('edits', []); rules['comments'] += d.get('comments', [])
    e_rows = {e.get('row') for e in rules['edits']}; c_rows = {c.get('row') for c in rules['comments']}
    miss_fig, miss_lab, bad, undone = [], [], [], []
    for f, ls in figs.items():
        if f not in secs: miss_fig.append(f); continue
        body = norm('\n'.join(secs[f]))
        for x in ls:
            if norm(x) not in body: miss_lab.append((f, x))
    for f, lines in secs.items():
        for l in lines:
            if not l.startswith('|') or set(l.replace('|', '').strip()) <= set('-: '): continue
            c = [x.strip() for x in l.strip().strip('|').split('|')]
            if len(c) < 4 or c[0] in ('#', '序号'): continue
            rid, concl = c[0], c[-1]
            if not concl.startswith(OK_START): bad.append((f, rid, concl[:40]))
            if '→改' in concl and rid not in e_rows: undone.append((f, rid, '修订'))
            if '→批注' in concl and rid not in c_rows: undone.append((f, rid, '批注'))
    print(f'图 {len(figs)} 张、标签 {sum(len(v) for v in figs.values())} 个；'
          f'【1 图没核】{len(miss_fig)}　【2 标签没核】{len(miss_lab)}　【3 结论没写对】{len(bad)}　【4 判了没落】{len(undone)}')
    for f in miss_fig: print(f'- 图没核：{f}')
    for f, x in miss_lab: print(f'- 标签没核：{f}「{x}」')
    for f, r, cc in bad: print(f'- 结论没写对：{f} {r}「{cc}」')
    for f, r, k in undone: print(f'- 判了没落：{f} {r} 缺{k}')
    sys.exit(1 if (miss_fig or miss_lab or bad or undone) else 0)


if __name__ == '__main__':
    main()

"""回忆提示：按阅读顺序，给每一段列出“读到这里该往前翻哪儿”。只提醒，不判断。

用法：~/.venvs/book-review/bin/python ~/.claude/skills/reader-consistency/scripts/recall_hints.py 段落文件.txt [--figs 图片文字.md] [--memory 全书名词表.md] -o 回忆提示.md

- 段落文件：dump_docx.py 生成的“P号: 文字”，T 行为表格；修订中删除的文字不算现文。
- 图片文字（可选）：照原图转写的 md，每张图一个“## 图x-y  图题（P28，……）”标题，下面是标签；
  图按标题里的 P 号插到那一段后面，像读者翻到图那样读。
- 全书名词表（可选）：表格两栏“采用名称｜要改成它的写法”里的词，当作“前面见过的词”，提示里显示为“全书名词表”。

提示只挂在后出现的那一处（读者读到后面才开始回忆），分三种：
1. 名字像：这里第一次出现的词，和前文或全书名词表里的词差一个字、多一个字少一个字、末字换字、换了字序；
2. 有列举：这一段列了几项（“A、B 和 C”“（1）……（2）……”“三种”），前文列过同一后缀或有同项的列举也一并给出；
3. 有指代：该／本／上述／前文／第×章……，提醒回看它指的那个东西。
"""
import argparse
import collections
import re
import sys
from pathlib import Path

# 差一字、增删字、末字换字的找法直接用正式版的 variant_pairs.py，不另写一份
sys.path.insert(0, str(Path.home() / '.claude/skills/editorial-term-consistency-review/scripts'))
import variant_pairs as vp  # noqa: E402
import jieba  # noqa: E402

CN_NUM = '一二两三四五六七八九十'
# 以这些动词开头的词组是句子成分，不是名称（掌握电网／掌握电力）
VERBS = ('掌握 了解 熟悉 理解 包括 包含 具有 具备 通过 利用 采用 使用 保证 保障 实现 进行 提高 降低 满足 确保 '
         '经过 需要 可以 能够 用于 作为 成为 存在 提供 产生 生产 引起 造成 导致').split()


def clean(t):
    t = re.sub(r'\[-.*?-\]', '', t)
    return re.sub(r'【黄:|】|\[\+|\+\]|\[图片\]', '', t).strip()


def load_units(para_path, fig_path):
    """阅读顺序的单元：[(标签, 文字)]；图插在标题注明的 P 号之后"""
    units = []
    for line in open(para_path, encoding='utf-8'):
        m = re.match(r'^([PT]\d+):\s?', line)
        if m:
            units.append((m.group(1), clean(line[m.end():])))
    if not fig_path:
        return units
    figs = collections.defaultdict(list)   # P号 -> [(图号, 文字)]
    cur = None
    for line in open(fig_path, encoding='utf-8'):
        h = re.match(r'^##\s*((?:图|表)\d+-\d+)\s*(.*)', line)
        if h:
            at = re.search(r'[（(](P\d+)', h.group(2))
            title = re.split(r'[（(]P\d+', h.group(2))[0].strip()
            cur = [h.group(1), title]
            figs[at.group(1) if at else 'END'].append(cur)
        elif line.startswith('## '):
            cur = None
        elif cur is not None and line.strip():
            cur.append(line.strip().lstrip('-').strip())
    out = []
    for lab, t in units:
        out.append((lab, t))
        for f in figs.get(lab, []):
            out.append((f[0], '；'.join(f[1:])))
    for f in figs.get('END', []):
        out.append((f[0], '；'.join(f[1:])))
    return out


def load_memory(path):
    """全书名词表里的写法：表格两栏（采用名称｜要改成它的写法）里的词，括号里的说明不算"""
    terms = set()
    if not path:
        return terms
    for line in open(path, encoding='utf-8'):
        if not line.startswith('|') or set(line.strip()) <= set('|-: '):
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        if cells and cells[0] == '采用名称':
            continue
        for c in cells:
            c = re.sub(r'[（(][^）)]*[）)]', '', c)
            for w in re.split(r'[、，,；;／/\s]+', c):
                if re.fullmatch(r'[一-鿿A-Za-z0-9\-]{2,12}', w):
                    terms.add(w)
    return terms


def first_seen(units, memory):
    """每个整词第一次出现的位置；全书名词表里的词算在最前面"""
    vp.ALIGNED.clear(); vp.TAIL2.clear()
    occ = vp.ngrams(units + [('全书名词表', t) for t in sorted(memory)])
    order = {lab: i for i, (lab, _) in enumerate(units)}
    order['全书名词表'] = -1
    first = {g: min(ws, key=lambda w: order[w]) for g, ws in occ.items()}
    return occ, first, order


def swaps(occ):
    """换了字序的整词（风电场／电风场）；两个词整体对调（特征融合／融合特征）多是动宾和偏正，不算"""
    by = collections.defaultdict(list)
    for g in occ:
        if len(g) >= 3 and g in vp.ALIGNED and g[0] not in vp.SOFT and g[-1] not in vp.SOFT:
            by[''.join(sorted(g))].append(g)
    found = {}
    for gs in by.values():
        for i in range(len(gs)):
            for j in range(i + 1, len(gs)):
                a, b = gs[i], gs[j]
                ta = [w for w in jieba.lcut(a, HMM=False)]
                tb = [w for w in jieba.lcut(b, HMM=False)]
                if len(ta) == 2 and ta[::-1] == tb:
                    continue
                if b in a + a:      # 整体轮转是 n 元切片切出来的，不是换序
                    continue
                found[tuple(sorted((a, b)))] = '换序'
    return found


def tokens_of(units):
    """每个单元按 jieba 精确模式切成词（不猜新词），去掉标点"""
    return {lab: [w for w in jieba.lcut(t, HMM=False) if re.search(r'[一-鿿A-Za-z]', w)] for lab, t in units}


def terms_of(toks, order):
    """词和两词组合（用电＋负荷、输电＋线路）：返回 {写法: [出现的单元…]}，按阅读顺序"""
    occ = collections.defaultdict(list)
    for lab in sorted(toks, key=lambda l: order[l]):
        ws = toks[lab]
        for w in ws:
            if re.fullmatch(r'[一-鿿]{2,8}', w):
                occ[w].append(lab)
        for x, y in zip(ws, ws[1:]):
            g = x + y
            if re.fullmatch(r'[一-鿿]{3,6}', g) and len(x) >= 2 and len(y) >= 1 \
                    and not any(x.startswith(v) for v in VERBS) and y not in vp.SOFT:
                occ[g].append(lab)
    return occ


def trunc_pairs(terms):
    """去首字或去尾字的短称（发电厂／电厂、输电线路／输电线）：两个都要在原文里作为词或词组出现过"""
    found = {}
    for g in terms:
        if len(g) < 3:
            continue
        for sw, kind in ((g[1:], '去首字'), (g[:-1], '去尾字')):
            # 稳定／稳定性、标准／标准化、稳定性／不稳定性是构词，不是短称
            if (kind == '去尾字' and g[-1] in '性化到级们者率量式') or (kind == '去首字' and g[0] in '不无非跨'):
                continue
            # 去尾字剩两个字的多是动词＋后缀（发电厂／发电、输电线／输电），不算短称
            if kind == '去尾字' and len(sw) == 2:
                continue
            if sw in terms and len(sw) >= 2 and sw[0] not in vp.SOFT and sw[-1] not in vp.SOFT \
                    and not any(sw.startswith(v) for v in VERBS):
                found[(g, sw)] = kind
    return found


def head_pairs(terms):
    """中心词相同、限定词不同（用电负荷／电力负荷、电力设备／电气设备）：
    限定词、中心词都是两字词，限定词至少有一个字相同，且有一方本章只出现一两次"""
    by_head = collections.defaultdict(list)
    for g, labs in terms.items():
        if len(g) == 4 and g[:2] != g[2:] and jieba.dt.FREQ.get(g[:2]) and jieba.dt.FREQ.get(g[2:]) \
                and g[0] not in vp.SOFT and not any(g.startswith(v) for v in VERBS):
            by_head[g[2:]].append((g[:2], len(labs)))
    found = {}
    for y, mods in by_head.items():
        for i in range(len(mods)):
            for j in range(i + 1, len(mods)):
                (ma, na), (mb, nb) = mods[i], mods[j]
                if not (set(ma) & set(mb)) or min(na, nb) > 2:
                    continue
                found[tuple(sorted((ma + y, mb + y)))] = '同中心词'
    return found


ENUM_SPLIT = re.compile(r'[、]|(?<=[一-鿿])(?:和|及|与|以及)(?=[一-鿿])')
COUNT_RE = re.compile(rf'([二两三四五六七八九十]+|[2-9]\d*)\s*(个|种|类|项|方面|部分|阶段|步骤|方式|要求|特点|特性|性质|环节|层次|级)')
NUM_HEAD = re.compile(r'^\s*(?:(?P<p>[（(]\s*(?P<pn>\d+)\s*[)）])|(?P<h>(?P<hn>\d+)\s*[)）])|(?P<d>(?P<dn>\d+)[．.、](?!\d))|(?P<c>[①-⑩]))\s*')
HEADING = re.compile(r'^\s*(第[0-9一二三四五六七八九十]+章|\d+\.\d+)')


def enums_in(text):
    """一段里的“A、B 和 C”式列举：返回 [项列表]"""
    found = []
    for clause in re.split(r'[，。；：！？,;:]', text):
        if clause.count('、') < 1:
            continue
        items = [x.strip() for x in ENUM_SPLIT.split(clause) if x.strip()]
        if len(items) < 3:
            continue
        # 首项常带着句子成分（“包括发电厂”），取它最后 2～8 个汉字
        head = re.split(r'包括|包含|主要有|分为|具备|具有|如|有|由|是|为|驱动着', items[0])[-1]
        items[0] = head[-10:]
        tail = re.split(r'等|的基本|组成', items[-1])[0]
        items[-1] = tail or items[-1]
        items = [i[:12] for i in items]
        found.append(items)
    return found


def numbered_groups(units):
    """同一种编号从 1 起连续的段合成一组（中间夹正文不断，遇到节题重开）：返回 [[(标签, 项名)]]"""
    groups, cur = [], {}
    def close(style):
        if len(cur.get(style, [])) >= 2:
            groups.append(cur[style])
        cur.pop(style, None)
    for lab, t in units:
        if not lab.startswith('P'):
            continue
        if HEADING.match(t):
            for st in list(cur):
                close(st)
            continue
        m = NUM_HEAD.match(t)
        if not m:
            continue
        style = next(k for k in ('p', 'h', 'd', 'c') if m.group(k))
        n = '①②③④⑤⑥⑦⑧⑨⑩'.index(m.group('c')) + 1 if style == 'c' else int(m.group(style + 'n'))
        name = re.split(r'[：:。，,]', t[m.end():], maxsplit=1)[0][:12]
        if n == 1:
            close(style)
            cur[style] = [(lab, name, n)]
        elif cur.get(style) and cur[style][-1][2] == n - 1:
            cur[style].append((lab, name, n))
    for st in list(cur):
        close(st)
    return [[(l, nm) for l, nm, _ in g] for g in groups]


CUE_RE = re.compile(r'((?<![基根成日资版样书剧原文])本|(?<![因如彼])此|该|上述|前述|所述|前文|上文|如前所述|前面|上一节|上节|第[0-9一二三四五六七八九十]+章|[0-9]+\.[0-9]+(?:\.[0-9]+)?节)'
                    r'(方法|模型|算法|框架|系统|网络|模块|过程|步骤|阶段|特性|特点|问题|方式|设备|装置|电网|线路|变电所|变电站|发电厂|要求|原则|指标|章|节)?')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('para')
    ap.add_argument('--figs')
    ap.add_argument('--memory')
    ap.add_argument('-o', '--out', required=True)
    a = ap.parse_args()

    units = load_units(a.para, a.figs)
    memory = load_memory(a.memory)
    occ, first, order = first_seen(units, memory)
    text_of = dict(units)
    toks = tokens_of(units)
    hints = collections.defaultdict(list)   # 标签 -> [提示]

    # 1. 名字像：挂在两种写法中后出现的那一处
    pairs = vp.prune(vp.pairs(occ), occ)
    pairs.update(swaps(occ))
    terms = terms_of(toks, order)
    hp = head_pairs(terms)
    n_name = 0
    for (x, y), kind in pairs.items():
        # 只看本章里第一次出现的位置；全书名词表里有、本章没出现的一方算“前面见过”
        cx = [w for w in occ[x] if w != '全书名词表']; cy = [w for w in occ[y] if w != '全书名词表']
        if not cx and not cy:
            continue
        fx = min(cx, key=lambda w: order[w]) if cx else '全书名词表'
        fy = min(cy, key=lambda w: order[w]) if cy else '全书名词表'
        later, early, lw, ew = (fx, fy, x, y) if order[fx] >= order[fy] else (fy, fx, y, x)
        if later == '全书名词表':
            continue
        if any(lw.startswith(v) for v in VERBS) or any(ew.startswith(v) for v in VERBS):
            continue
        where = '全书名词表' if early == '全书名词表' else early
        if early == later:
            where = '同一段'
        hints[later].append(f'名字像（{kind}）：“{lw}”← 前面 {where}“{ew}”')
        n_name += 1

    # 1b. 去首字、去尾字的短称和同中心词：按词和词组比，挂在后出现的那个第一次出现处
    extra = dict(trunc_pairs(terms)); extra.update(hp)
    for (g, sw), kind in extra.items():
        fg, fs = terms[g][0], terms[sw][0]
        later, early, lw, ew = (fs, fg, sw, g) if order[fs] >= order[fg] else (fg, fs, g, sw)
        # 同中心词出现在同一段的，多是并列列举（发电环节、输电环节……），不提示
        if kind == '同中心词' and later == early:
            continue
        hints[later].append(f'名字像（{kind}）：“{lw}”← 前面 {"同一段" if later == early else early}“{ew}”')
        n_name += 1

    # 2. 有列举：“A、B 和 C”和连续编号段；同后缀或有同项的前文列举一并给出
    lists = []   # (标签, 项, 来源)
    for lab, t in units:
        for items in enums_in(t):
            lists.append((lab, items, '顿号'))
    for g in numbered_groups(units):
        lists.append((g[0][0], [n for _, n in g], f'编号段 {g[0][0]}–{g[-1][0]}'))
    for lab, t in units:
        for m in COUNT_RE.finditer(t):
            hints[lab].append(f'说了数量：“{m.group(0)}”——对一下后文实际列了几项')
    n_enum = 0
    for k, (lab, items, src) in enumerate(lists):
        line = f'列举（{src}，{len(items)} 项）：' + '、'.join(items)
        rel = []
        for lab2, items2, _ in lists[:k]:
            same = set(items) & set(items2)
            suf = {i[-1] for i in items if len(i) >= 2} & {i[-1] for i in items2 if len(i) >= 2}
            if same or (suf and all(i.endswith(tuple(suf)) for i in items) and all(i.endswith(tuple(suf)) for i in items2)):
                rel.append(f'{lab2}（{len(items2)} 项：' + '、'.join(items2) + '）')
        if rel:
            line += '　← 前面也列过：' + '；'.join(rel[-3:])
        hints[lab].append(line)
        n_enum += 1

    # 3. 有指代
    n_cue = 0
    for lab, t in units:
        if HEADING.match(t):
            continue
        cues = sorted({m.group(0) for m in CUE_RE.finditer(t) if m.group(2) or m.group(1) in ('上述', '前述', '前文', '上文', '如前所述', '前面') or '章' in m.group(1) or '节' in m.group(1)})
        if cues:
            hints[lab].append('有指代：' + '、'.join(f'“{c}”' for c in cues) + '——回看它指的是前面哪个东西，叫法对不对得上')
            n_cue += len(cues)

    out = [f'# 回忆提示：{Path(a.para).name}', '',
           f'阅读单元 {len(units)} 个（含图 {sum(1 for l, _ in units if not l[0] in "PT")} 张）；'
           f'名字像 {n_name} 条、列举 {n_enum} 处、指代 {n_cue} 处。只是提醒，判断由读者做；没列到的想起来了照样要翻。', '']
    for lab, _ in units:
        if hints.get(lab):
            snippet = text_of[lab][:30].replace('\n', ' ')
            out.append(f'## {lab}　{snippet}')
            out += [f'- {h}' for h in hints[lab]]
            out.append('')
    Path(a.out).write_text('\n'.join(out), encoding='utf-8')
    print(f'阅读单元 {len(units)}；名字像 {n_name}、列举 {n_enum}、指代 {n_cue}；有提示的单元 {sum(1 for v in hints.values() if v)}；写入 {a.out}')


if __name__ == '__main__':
    main()

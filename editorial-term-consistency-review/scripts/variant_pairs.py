"""找出本章中“只差一个字”的写法配对，作为一致性审查的候选清单。

用法：~/.venvs/book-review/bin/python variant_pairs.py 段落文件.txt [图表文字.txt] [输出.md]
- 段落文件：dump_docx.py 生成的 “P号: 文字” 格式；
- 图表文字（可选）：每行 “图1-2: 文字”，由审读者看 PDF 原页抄录；
差异只算一个字的替换、增加或删除，且差异不在词组的首尾（首尾不同只是上下文不同）。
配对两端都须是整词（首尾落在 jieba 词边界上），过滤“世界先/世界领先”这类拆坏的半截词。
"""
import sys, re, collections
import jieba
jieba.setLogLevel(60)

FUNC = set('的了和与及或在是对等为将其各该此这那并而被把从向于以由也都就还又所之上下中内外个种些')
# 副词、量词、虚化动词：作为增删字或词组首尾时多是语句差异，不是名称变体
SOFT = FUNC | set('很较更最极已仍再均皆仅只远曾将会能可要需使让给有到着过一两三几每某其大小多少高低新')
MINLEN, MAXLEN = 3, 8


def load(path, label_re):
    """读入文本，返回 [(位置标签, 文字)]，去掉修订、批注、黄色标记"""
    out = []
    for line in open(path, encoding='utf-8'):
        m = re.match(label_re, line)
        if not m: continue
        t = line[m.end():]
        t = re.sub(r'\[-.*?-\]', '', t)          # 删除的文字不算现文
        t = re.sub(r'【黄:|】|\[\+|\+\]|\[图片\]', '', t)
        out.append((m.group(1), t))
    return out


ALIGNED = set()   # 至少有一次出现时首尾都落在 jieba 词边界上的子串（不是拆坏的半截词）
TAIL2 = set()     # 整词子串中以两字词结尾的（如“混沌特性”的“特性”），用于末字换字


def ngrams(units):
    """每个 CJK 连续片段内取 3～8 字的子串，记录出现位置；同时用 jieba 标出整词子串"""
    occ = collections.defaultdict(list)
    for where, t in units:
        cuts = {0, len(t)}
        # 关闭 HMM 新词猜测（否则会切出“层来”“层中”）；搜索模式让长词内部的子词也算边界（如“滑动|模型”）
        for _, a, b in jieba.tokenize(t, mode='search', HMM=False):
            cuts.add(a); cuts.add(b)
        for m in re.finditer(r'[一-鿿]+', t):
            seg, base = m.group(), m.start()
            for n in range(MINLEN, MAXLEN + 1):
                for i in range(len(seg) - n + 1):
                    g = seg[i:i + n]
                    occ[g].append(where)
                    if base + i in cuts and base + i + n in cuts:
                        ALIGNED.add(g)
                        if base + i + n - 2 in cuts: TAIL2.add(g)
    return occ


def pairs(occ):
    """找差一个字的配对：替换用通配键，增删用删字键"""
    found = {}
    # 替换：把内部某一位换成 *，同键不同字即为替换配对
    sub = collections.defaultdict(set)
    for g in occ:
        if g[0] in SOFT or g[-1] in SOFT: continue
        for i in range(1, len(g) - 1):
            sub[(g[:i], g[i + 1:])].add(g)
    for (l, r), gs in sub.items():
        gs = sorted(g for g in gs if g in ALIGNED)   # 两端都要是整词子串
        for a in range(len(gs)):
            for b in range(a + 1, len(gs)):
                ca, cb = gs[a][len(l)], gs[b][len(l)]
                if ca in SOFT or cb in SOFT: continue
                found[(gs[a], gs[b])] = '换字'
    # 末字换字：末尾两字词只差后一个字（特性/特征、方法/方式、校正/修正之外的“前同后异”），前面至少两个字相同
    tail = collections.defaultdict(set)
    for g in TAIL2:
        if len(g) >= 4 and g[0] not in SOFT and g[-1] not in SOFT: tail[g[:-1]].add(g)
    for pre, gs in tail.items():
        gs = sorted(gs)
        for a in range(len(gs)):
            for b in range(a + 1, len(gs)):
                x, y = gs[a], gs[b]
                # 末尾两字都须是词典里的词（排除“模型的不/模型的局”这类半截）
                # 不再要求一方是少数写法（≤2处）：两种都常用的（异常信号5处/异常信息7处）更该比，
                # 放开后每章只多0～2对（第13、15、16、17章实测，11-15）
                if not (jieba.dt.FREQ.get(x[-2:]) and jieba.dt.FREQ.get(y[-2:])): continue
                nx, ny = len(occ[x]), len(occ[y])
                if nx + ny < 3: continue
                found[(x, y)] = '末字换字'
    # 增删：长串删去一个内部字后等于短串
    for g in occ:
        if len(g) < MINLEN + 1 or g[0] in SOFT or g[-1] in SOFT: continue
        for i in range(1, len(g) - 1):
            s = g[:i] + g[i + 1:]
            # 叠字（远远、仅仅）不算
            if s in occ and g in ALIGNED and s in ALIGNED and g[i] not in SOFT and g[i] != g[i - 1] and g[i] != g[i + 1]:
                found[tuple(sorted((g, s)))] = '增删字'
    return found


def prune(found, occ):
    """去掉被更长配对完全覆盖的短配对，只保留上下文最完整的那一对"""
    # 短串的每次出现都落在长串里（出现次数相同），说明它只是长配对的一部分
    keep = {}
    items = sorted(found.items(), key=lambda kv: -len(kv[0][0]) - len(kv[0][1]))
    for (a, b), kind in items:
        covered = any(
            (a in A and b in B and len(occ[a]) == len(occ[A]) and len(occ[b]) == len(occ[B])) or
            (a in B and b in A and len(occ[a]) == len(occ[B]) and len(occ[b]) == len(occ[A]))
            for (A, B) in keep)
        if not covered: keep[(a, b)] = kind
    return keep


def main():
    para = load(sys.argv[1], r'^(P\d+):\s?')
    figs = load(sys.argv[2], r'^(图\d+-\d+|表\d+-\d+):\s?') if len(sys.argv) > 2 and sys.argv[2].endswith('.txt') else []
    out = sys.argv[-1] if sys.argv[-1].endswith('.md') else None
    occ = ngrams(para + figs)
    res = prune(pairs(occ), occ)
    rows = []
    for (a, b), kind in sorted(res.items(), key=lambda kv: kv[0]):
        fa = ' '.join(dict.fromkeys(occ[a])); fb = ' '.join(dict.fromkeys(occ[b]))
        rows.append(f'| {a}（{fa}） | {b}（{fb}） | {kind} |')
    text = '| 写法甲（位置） | 写法乙（位置） | 差异 |\n|---|---|---|\n' + '\n'.join(rows) + '\n'
    if out: open(out, 'w', encoding='utf-8').write(text)
    print(f'候选配对 {len(rows)} 对，清单 {len(text)} 字')


if __name__ == '__main__':
    main()

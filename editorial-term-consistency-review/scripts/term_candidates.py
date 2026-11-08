"""按中心词抽取本章的命名对象候选，供建本章名词表时逐组核对。

用法：~/.venvs/book-review/bin/python term_candidates.py 段落文件.txt [图表文字.md] [--heads 词1,词2,...] [--search 写法1,写法2,...] [-o 输出.md]
- 段落文件：dump_docx.py 生成的“P号: 文字”（T 行为表格），修订中删除的文字不算现文；
- 图表文字（可选）：每行“图1-2: 文字”或“表1-2: 文字”；
- --heads：追加本书的中心词（默认表见 HEADS，按书增删，文科书可加 理论,学派,制度,政策,事件,王朝 等）；
- --search：总名词表里的“必搜写法”（缺字写法），逐个报出现位置。

输出三部分：
1. 按中心词分组的词组：同一组里“监督分类模型／监督学习模型／无监督模型”摆在一起，一眼看出差异；
2. 英文缩写及其前后称谓：以缩写为中心，列出它前面的“××方法／模型／算法”和后面的“该方法／该模型”；
3. 必搜写法命中位置。
只是候选，噪音要人判；没出现在候选里的对象照样要在逐段精读时登记。
"""
import collections
import re
import sys

import jieba
jieba.setLogLevel(60)

HEADS = ('方法 模型 算法 框架 策略 机制 方案 问题 子问题 阶段 步骤 模块 流程 范式 '
         '数据集 指标 函数 矩阵 张量 向量 网络 系统 平台 工具 任务 场景 '
         '理论 学说 观点 概念 原则 制度 政策 法 条例 办法 规定 标准 规范 报告 纲要 规划 '
         '企业 机构 部门 组织 协会 委员会 '
         '攻击 威胁 异常 误差 损失 约束 正则 编码 表示 学习 检测 预测 分类 聚类 分解 恢复 优化 '
         '过程 结构 关系 成分 部分 类别 类型 方式 形式 特性 特征').split()
# 指代词＋中心词（该方法、本模型、所提算法）单独成组，用来查类别词前后是否一致
DEMON = ('该', '本', '此', '这种', '该种', '所提', '上述', '所提出的', '提出的')
FUNC = set('的了和与及或在是对等为将其各该此这那并而被把从向于以由也都就还又所之上下中内外个种些'
           '很较更最极已仍再均皆仅只远曾会能可要需使让给有到着过一两三几每某大小多少高低新')
PUNCT = set('，。；：、！？“”‘’（）《》〈〉【】—…·,.;:!?()[]<>"\' \t')
# 向前拼词组时碰到这些词停下（它们是句子成分，不是名称的一部分）
STOP = set('使得 进行 引发 有助于 后来 近年来 多数 这种 这些 那些 采用 利用 使用 通过 基于 提出 构建 设计 '
           '引入 结合 实现 得到 作为 成为 存在 具有 包括 包含 称为 称作 视为 记为 表示 定义 如下 所示'.split())
MAXLEN = 12


def load(path, label_re):
    out = []
    for line in open(path, encoding='utf-8'):
        m = re.match(label_re, line)
        if not m:
            continue
        t = line[m.end():]
        t = re.sub(r'\[-.*?-\]', '', t)
        t = re.sub(r'【黄:|】|\[\+|\+\]|\[图片\]', '', t)
        out.append((m.group(1), t.rstrip('\n')))
    return out


def phrases_by_head(units, heads):
    """对每个中心词出现处，向前按 jieba 词边界拼出词组（最长 MAXLEN 字，碰到虚词、标点停）"""
    groups = collections.defaultdict(lambda: collections.defaultdict(list))
    heads_sorted = sorted(heads, key=len, reverse=True)
    for where, t in units:
        toks = list(jieba.tokenize(t, HMM=False))
        for i, (w, a, b) in enumerate(toks):
            head = next((h for h in heads_sorted if w.endswith(h) and (w == h or len(w) <= len(h) + 4)), None)
            if not head:
                continue
            # 向前拼接
            phrase = w
            j = i - 1
            while j >= 0:
                pw = toks[j][0]
                if pw in FUNC or pw in STOP or any(c in PUNCT for c in pw) or re.fullmatch(r'\d+', pw):
                    break
                if len(pw) + len(phrase) > MAXLEN:
                    break
                phrase = pw + phrase
                j -= 1
            if phrase == head and len(head) <= 2:
                # 光秃秃的“方法”“模型”只在前面是指代词时登记（该方法／本模型／所提算法），归入“指代＋类别词”组
                prev = toks[i - 1][0] if i > 0 else ''
                if prev in DEMON:
                    groups['指代＋类别词'][prev + head].append(where)
                continue
            groups[head][phrase].append(where)
    return groups


def abbreviations(units):
    """英文缩写及其前后称谓"""
    ctx = collections.defaultdict(list)
    for where, t in units:
        for m in re.finditer(r'(?<![A-Za-z])[A-Z][A-Za-z0-9]*(?:-[A-Za-z0-9]+)*(?![a-z])', t):
            ab = m.group(0)
            if len(ab) < 2 or sum(c.isupper() for c in ab) < 2:
                continue
            pre = t[max(0, m.start() - 14):m.start()]
            post = t[m.end():m.end() + 8]
            pre = re.split(r'[，。；：、！？（）]', pre)[-1]
            post = re.split(r'[，。；：、！？（）]', post)[0]
            ctx[ab].append((where, pre, post))
    return ctx


def search_terms(units, terms):
    hits = collections.defaultdict(list)
    for where, t in units:
        for term in terms:
            for m in re.finditer(re.escape(term), t):
                hits[term].append((where, t[max(0, m.start() - 10):m.end() + 10]))
    return hits


def main():
    args = sys.argv[1:]
    heads = list(HEADS)
    search = []
    out = None
    files = []
    i = 0
    while i < len(args):
        if args[i] == '--heads':
            heads += [h for h in args[i + 1].split(',') if h]
            i += 2
        elif args[i] == '--search':
            search = [s for s in args[i + 1].split(',') if s]
            i += 2
        elif args[i] == '-o':
            out = args[i + 1]
            i += 2
        else:
            files.append(args[i])
            i += 1
    if not files:
        print(__doc__)
        sys.exit(2)
    for h in heads:
        jieba.add_word(h)
    units = load(files[0], r'^([PT]\d+):\s?')
    if len(files) > 1:
        units += load(files[1], r'^(图\d+-\d+|表\d+-\d+|算法\d+\.\d+):\s?')

    lines = ['# 命名对象候选（脚本抽取，须人判）', '']
    groups = phrases_by_head(units, set(heads))
    lines.append('## 一、按中心词分组的词组')
    lines.append('')
    n_phr = 0
    for head in sorted(groups, key=lambda h: -sum(len(v) for v in groups[h].values())):
        phr = groups[head]
        if sum(len(v) for v in phr.values()) < 1:
            continue
        lines.append(f'### {head}（{len(phr)} 种写法）')
        lines.append('| 写法 | 次数 | 位置 |')
        lines.append('|---|---|---|')
        for p in sorted(phr, key=lambda x: (x[-len(head):], x)):
            pos = ' '.join(dict.fromkeys(phr[p]))
            lines.append(f'| {p} | {len(phr[p])} | {pos} |')
            n_phr += 1
        lines.append('')

    ab = abbreviations(units)
    lines.append('## 二、英文缩写及其前后称谓')
    lines.append('')
    for a in sorted(ab, key=lambda x: -len(ab[x])):
        lines.append(f'### {a}（{len(ab[a])} 处）')
        lines.append('| 位置 | 前文 | 后文 |')
        lines.append('|---|---|---|')
        for where, pre, post in ab[a]:
            lines.append(f'| {where} | {pre} | {post} |')
        lines.append('')

    if search:
        hits = search_terms(units, search)
        lines.append('## 三、必搜写法命中')
        lines.append('')
        for term in search:
            hs = hits.get(term, [])
            lines.append(f'- **{term}**：{len(hs)} 处' + ('' if not hs else '；' + '；'.join(f'{w}“{c}”' for w, c in hs[:12])))
        lines.append('')

    text = '\n'.join(lines)
    if out:
        open(out, 'w', encoding='utf-8').write(text)
    print(f'中心词 {len(groups)} 组、词组 {n_phr} 种、缩写 {len(ab)} 个' + (f'，已写 {out}' if out else ''))
    if not out:
        print(text)


if __name__ == '__main__':
    main()

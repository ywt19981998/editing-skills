"""本章名词表门禁：找出“有多种写法却没有给出结论”的行。

用法：~/.venvs/book-review/bin/python check_glossary.py 第NN章-名词表.md [--strip 标记]
  --strip 标记   判断栏从该标记起截掉再检查（如 --strip 【复查推翻】，用于拿旧表做试验）

检查规则（对应 SKILL.md 第五节“判断栏三选一”）：
1. “一、命名对象”及同格式表中，写法列有两种以上写法的行，判断栏必须给出结论：
   “统一为 X → 问题编号”“D-xx”或“不同对象：＜内容区别＋两端原文＞”。
   只写“一致”“保留”“描述性”“同位语”“概括”“叙述”“各有所指”“泛称”而没有写明内容区别的，列为未决。
   判断栏按“；”分段逐段看：一段里出现放过词、又没有“区别”“不同对象：”字样，这一段就是未决。
   把几个不同对象放在一行的汇总行（小节题列表、图表各项、参数）要写“逐项核对：……”，否则也算未决；
   一行只放一个对象是常态，写法列只放写法，不放“全称未交代”这类备注（备注会被数成一种写法）。
2. “二、非同一对象”表中，区分依据必须含原文位置（P号、节号、图表号、式号）或引号里的原文；
   只写“标签词不同”“上下位”“不同词”这类一句话的，列为依据不足。
退出码：有未决行为 1，没有为 0。主代理收到子代理结果后再跑一遍，不为 0 就打回。
"""
import re
import sys

BAN = ['保留', '描述性', '同位语', '概括', '叙述', '各有所指', '泛称', '不收', '不改', '维持']
# “逐项”用于把几个不同对象放在一行的汇总行（节题列表、图表各项、参数表）：判断栏写“逐项核对：……”才算有结论
OK_MARK = ['区别', '不同对象：', '不同对象:', '逐项']
POS_RE = re.compile(r'(P\d+|T\d+|\d+\.\d+(\.\d+)?节|图\d+-\d+|表\d+-\d+|式（\d+-\d+）|算法\d+\.\d+|定义\d+\.\d+|p\.\d+|[“「][^”」]{2,}[”」])')


def split_cells(line):
    cells = [c.strip() for c in line.strip().strip('|').split('|')]
    return cells


def count_writings(cell):
    """写法列：按“；”分，去掉位置括注后数不同写法"""
    items = []
    for part in re.split(r'[；;]', cell):
        part = re.sub(r'[（(][^（）()]*[）)]\s*$', '', part.strip()).strip()
        part = re.sub(r'[（(][^（）()]*[）)]', lambda m: m.group(0) if re.search(r'[A-Za-z]', m.group(0)) else '', part).strip()
        if part:
            items.append(part)
    return list(dict.fromkeys(items))


def judge_segments(judge):
    """返回未决的判断段落列表"""
    bad = []
    segs = [s.strip() for s in re.split(r'[；;]', judge) if s.strip()]
    for s in segs:
        s_chk = s.replace('不一致', '').replace('图文不符', '')
        has_ok = any(m in s for m in OK_MARK)
        has_ban = any(b in s_chk for b in BAN) or re.search(r'(^|[^不])一致', s_chk)
        if has_ban and not has_ok:
            bad.append(s)
    return bad


def main():
    args = sys.argv[1:]
    strip = None
    if '--strip' in args:
        i = args.index('--strip')
        strip = args[i + 1]
        del args[i:i + 2]
    if not args:
        print(__doc__)
        sys.exit(2)
    text = open(args[0], encoding='utf-8').read().splitlines()

    section = None
    pending = []      # (编号, 对象, 写法数, 未决段)
    weak = []         # 非同一对象表依据不足
    rows_multi = 0
    rows_total = 0
    for line in text:
        if line.startswith('## '):
            section = line
            continue
        if not line.startswith('|') or set(line.replace('|', '').strip()) <= set('-: '):
            continue
        cells = split_cells(line)
        if section and '非同一对象' in section:
            if len(cells) < 3 or cells[0] in ('对象甲',):
                continue
            basis = cells[2]
            if not POS_RE.search(basis) or len(basis) < 8:
                weak.append((cells[0], cells[1], basis))
            continue
        # 命名对象表（含“复查补收”等同格式表）：# | 所指对象 | 写法及位置 | 依据 | 采用写法 | 判断
        if len(cells) < 6 or cells[0] in ('#',):
            continue
        rows_total += 1
        writings = count_writings(cells[2])
        judge = cells[5]
        if strip and strip in judge:
            judge = judge.split(strip)[0]
        if len(writings) < 2:
            continue
        rows_multi += 1
        has_resolution = ('→' in judge) or re.search(r'D-\d+', judge) or any(m in judge for m in OK_MARK)
        bad = judge_segments(judge)
        if not has_resolution:
            pending.append((cells[0], cells[1], len(writings), '多写法无结论：' + judge))
        elif bad:
            for b in bad:
                pending.append((cells[0], cells[1], len(writings), '放过未说明区别：' + b))

    print(f'命名对象 {rows_total} 行，多写法 {rows_multi} 行，未决 {len(pending)} 条；非同一对象依据不足 {len(weak)} 条')
    if pending:
        print('\n## 未决（判断栏必须改成：统一为X→编号 / D-xx / 不同对象：内容区别＋两端原文）')
        for cid, obj, n, why in pending:
            print(f'- {cid} {obj}（{n} 种写法）：{why}')
    if weak:
        print('\n## 非同一对象表依据不足（要写内容上的区别和两端原文位置）')
        for a, b, basis in weak:
            print(f'- {a} ↔ {b}：{basis}')
    sys.exit(1 if (pending or weak) else 0)


if __name__ == '__main__':
    main()

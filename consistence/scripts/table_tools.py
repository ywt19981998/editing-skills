#!/usr/bin/env python3
"""本章名词表的门禁、复核筛选和统计。

用法:
  table_tools.py gate   <第NN章-名词表.md> [--rules 改法.json ...] [--variants 变体候选.md]
  table_tools.py review <第NN章-名词表.md> -o <第NN章-待复核.md>
  table_tools.py stats  <第NN章-名词表.md> [--rules 改法或规则.json ...]

gate（名称子代理交活前跑到 0 条，主代理收活时再跑一遍）查三件事:
  1. 结论没判完：多写法行的结论不是“统一为／不同对象：／逐项核对：／D-xx／待确认”之一，
     或以“保留、描述性、同位语”之类放过却没写内容上的区别（沿用 check_glossary.py）。
  2. 判了没落：结论是“统一为”或 D 项的行，没列进采用写法栏（带引号）的写法，它所在的段落
     （写法后括号里的 P 号、T 号）在改法文件里没有对应的修订或批注。
     认可的写法带引号写进采用写法栏；个别位置确实不用动的，在结论里点出段落号并说明。
     “逐项核对”的汇总行不查这一项。
  3. 候选没看：变体候选里的配对，两种写法没有都出现在本表里。是同一对象的补进“一、命名对象”；
     不是名称变体的，写进“三、候选排除”。
review 把结论是“不同对象”“逐项核对”的行筛出来，供复核子代理重判。
stats 数出章末汇报用的条数。
"""
import sys, os, re, json, subprocess

PY = os.path.expanduser('~/.venvs/book-review/bin/python')
OLD_GATE = os.path.expanduser('~/.claude/skills/editorial-term-consistency-review/scripts/check_glossary.py')
POS = re.compile(r'[PT]\d+')


def cells_of(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]


def rows_of(path):
    """返回“命名对象”格式的行：(编号, 对象, 写法及位置, 依据, 采用写法, 结论)"""
    out, section = [], ''
    for line in open(path, encoding='utf-8').read().splitlines():
        if line.startswith('## '):
            section = line; continue
        if not line.startswith('|') or set(line.replace('|', '').strip()) <= set('-: '):
            continue
        c = cells_of(line)
        if len(c) < 6 or c[0] in ('#',) or '候选排除' in section:
            continue
        out.append(c[:6])
    return out


def writings(cell):
    """写法列按“；”分开，返回 [(写法, [段落号])]；位置取写法后最后一组括号里的 P 号、T 号"""
    res = []
    for part in re.split(r'[；;]', cell):
        part = part.strip()
        if not part:
            continue
        i = part.rfind('（')
        if i > 0 and part.endswith('）'):
            res.append((part[:i].strip(), POS.findall(part[i:])))
        else:
            res.append((part, []))
    return res


def quoted(text):
    return set(re.findall(r'“([^”]+)”', text))


def rule_places(files):
    ps = set()
    n_edit = n_com = n_dec = 0
    for f in files:
        d = json.load(open(f, encoding='utf-8'))
        for kind in ('edits', 'comments'):
            for e in d.get(kind, []):
                p = e.get('p')
                for x in (p if isinstance(p, list) else [p]):
                    if x:
                        ps.add(str(x))
                k = len(p) if isinstance(p, list) else 1
                if kind == 'edits':
                    n_edit += k; n_dec += k if e.get('decision') else 0
                else:
                    n_com += k
    return ps, n_edit, n_com, n_dec


def kind_of(judge):
    """按结论开头的词分类（“【复核改判】”这类前缀先去掉）；汇总行“逐项核对”里即使有“统一为”也算逐项核对"""
    j = re.sub(r'^(【[^】]*】\s*)+', '', judge.strip())
    for key, name in ((r'D-\d+', 'D项'), ('统一为', '统一'), ('不同对象', '不同对象'), ('逐项核对', '逐项核对'), ('待确认', '待确认')):
        if re.match(key, j):
            return name
    for key, name in ((r'D-\d+', 'D项'), ('统一为', '统一'), ('不同对象', '不同对象'), ('逐项核对', '逐项核对'), ('待确认', '待确认')):
        if re.search(key, j):
            return name
    return '其他'


def opt(args, name, many=False):
    if name not in args:
        return [] if many else None
    i = args.index(name)
    if many:
        j = i + 1
        while j < len(args) and not args[j].startswith('-'):
            j += 1
        vals = args[i + 1:j]; del args[i:j]
        return vals
    val = args[i + 1]; del args[i:i + 2]
    return val


def gate(table, rules, variants):
    bad = 0
    r = subprocess.run([PY, OLD_GATE, table], capture_output=True, text=True)
    print('【1 结论没判完】' + (r.stdout.strip() or r.stderr.strip()))
    bad += 1 if r.returncode else 0

    rows = rows_of(table)
    if rules:
        places, *_ = rule_places(rules)
        miss = []
        for cid, obj, wcell, _, adopt, judge in rows:
            if kind_of(judge) not in ('统一', 'D项'):
                continue
            ok = quoted(adopt) | quoted(judge) | {adopt.strip('“” ')}
            said = set(POS.findall(judge))
            for w, ps in writings(wcell):
                if w.strip('“” ') in ok:
                    continue
                left = [p for p in ps if p not in places and p not in said]
                if left:
                    miss.append(f'- {cid} {obj}：“{w}”在 {"、".join(left)} 没有修订或批注')
        print(f'\n【2 判了没落】{len(miss)} 条' + ('（补上改法；认可的写法带引号写进采用写法栏，个别位置不用动的在结论里点出段落号）' if miss else ''))
        print('\n'.join(miss))
        bad += 1 if miss else 0
    else:
        print('\n【2 判了没落】没给 --rules，未查')

    if variants:
        text = open(table, encoding='utf-8').read()
        miss = []
        for line in open(variants, encoding='utf-8').read().splitlines():
            if not line.startswith('|') or set(line.replace('|', '').strip()) <= set('-: '):
                continue
            c = cells_of(line)
            if len(c) < 2 or c[0].startswith('写法甲'):
                continue
            a, b = (re.sub(r'（[^（）]*）\s*$', '', x).strip() for x in c[:2])
            if a and b and not (a in text and b in text):
                miss.append(f'- {c[0]} ／ {c[1]}')
        print(f'\n【3 候选没看】{len(miss)} 对' + ('（同一对象的补进第一节；不是名称变体的写进“三、候选排除”）' if miss else ''))
        print('\n'.join(miss))
        bad += 1 if miss else 0
    else:
        print('\n【3 候选没看】没给 --variants，未查')
    sys.exit(1 if bad else 0)


def review(table, out):
    rows = [r for r in rows_of(table) if kind_of(r[5]) in ('不同对象', '逐项核对')]
    lines = ['# 待复核（结论是“不同对象”“逐项核对”的行）', '',
             '| # | 所指对象 | 本章出现的全部写法及位置 | 结论 |', '|---|---|---|---|']
    lines += [f'| {r[0]} | {r[1]} | {r[2]} | {r[5]} |' for r in rows]
    open(out, 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
    print(f'待复核 {len(rows)} 行 → {out}')


def stats(table, rules):
    rows = rows_of(table)
    cnt = {}
    for r in rows:
        k = kind_of(r[5]) if len(writings(r[2])) > 1 else '单写法'
        cnt[k] = cnt.get(k, 0) + 1
    ds = sorted({d for r in rows for d in re.findall(r'D-\d+', r[5])}, key=lambda x: int(x[2:]))
    over = sum(1 for r in rows if '复核改判' in r[5])
    print(f'命名对象 {len(rows)} 行：' + '、'.join(f'{k} {v}' for k, v in cnt.items()))
    print(f'D 项 {len(ds)} 个：{"、".join(ds) or "无"}；复核改判 {over} 行')
    if rules:
        _, n_edit, n_com, n_dec = rule_places(rules)
        print(f'修订 {n_edit} 条（其中 D 项 {n_dec} 条）、批注 {n_com} 条')

def main():
    args = sys.argv[1:]
    if len(args) < 2 or args[0] not in ('gate', 'review', 'stats'):
        print(__doc__); sys.exit(2)
    cmd = args.pop(0)
    rules = opt(args, '--rules', many=True)
    variants = opt(args, '--variants')
    out = opt(args, '-o')
    table = args[0]
    if cmd == 'gate':
        gate(table, rules, variants)
    elif cmd == 'review':
        if not out:
            print('review 要给 -o 输出文件'); sys.exit(2)
        review(table, out)
    else:
        stats(table, rules)


if __name__ == '__main__':
    main()

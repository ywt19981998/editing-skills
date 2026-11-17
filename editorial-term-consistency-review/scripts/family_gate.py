"""名称族门禁：核心名称定了方向后，带不带后缀（系统、电压、技术……）的写法都要改到。

用法：
  列出各写法出现在哪（开审前问用户方向时用，不判对错）：
    ~/.venvs/book-review/bin/python family_gate.py 名词表.md 段落文件或docx [更多段落文件或docx…] --list
  门禁（写入副本前跑，未覆盖须为 0）：
    ~/.venvs/book-review/bin/python family_gate.py 名词表.md 段落文件或docx --rules 规则1.json [规则2.json …]

名称族表写在总名词表“〇-1、名称族”一节，列：族 | 采用核心 | 变体核心 | 允许保留 | 状态
  - 变体核心用“；”或“、”分隔，如“工厂供配电；工厂配电”。只写核心部分，不写“系统”等后缀——
    脚本按核心搜，后面接什么都算命中。
  - 允许保留：用“；”分隔，每项写引号里的一小段原文，可在前面加段号，如 P9“亦称工厂配电”。
    命中处落在这段原文里就算保留。书名号《》里的命中自动保留（书名、篇名不改）。
  - 状态写“用户确认”才算方向已定；没定的，门禁直接不过。

门禁怎么判：把各规则文件里的 edits 按顺序模拟写进段落（ctx 定位段落，替换 old 的第一次出现），
再在改后的文字里搜变体核心；还剩下、又不在允许保留里的，就是没覆盖。只挂批注不算覆盖。
段落文件是 dump_docx.py 的输出（每行“P号: 文字”）；直接给 .docx 时脚本自己导出。
退出码：有未覆盖或方向未定为 1，否则 0。
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SPLIT = re.compile(r'[；;、，,]')


def load_paras(path):
    if path.lower().endswith('.docx'):
        out = subprocess.run([sys.executable, os.path.join(HERE, 'dump_docx.py'), path],
                             capture_output=True, text=True, check=True).stdout
        lines = out.splitlines()
    else:
        lines = open(path, encoding='utf-8').read().splitlines()
    paras = []
    for ln in lines:
        m = re.match(r'^(P\d+):\s?(.*)$', ln)
        if m:
            # 去掉原稿已有修订的标记：删除的字不算现文，插入的字算
            txt = re.sub(r'\[-.*?-\]', '', m.group(2))
            txt = re.sub(r'\[\+(.*?)\+\]', r'\1', txt)
            paras.append([m.group(1), txt])
    return paras


def load_families(glossary):
    fams = []
    section = None
    for line in open(glossary, encoding='utf-8').read().splitlines():
        if line.startswith('## '):
            section = line
            continue
        if not section or '名称族' not in section or not line.startswith('|'):
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        if not re.match(r'^F\d+', cells[0]) or len(cells) < 5:
            continue
        variants = [v.strip() for v in SPLIT.split(cells[2]) if v.strip()]
        keeps = []
        for k in re.split(r'[；;]', cells[3]):
            k = k.strip()
            if not k or k in ('无', '—', '-'):
                continue
            m = re.match(r'^(P\d+)?\s*[“"「](.+?)[”"」]', k)
            if m:
                keeps.append((m.group(1), m.group(2)))
            else:
                keeps.append((None, k))
        fams.append({'id': cells[0], 'adopt': cells[1], 'variants': variants,
                     'keeps': keeps, 'status': cells[4]})
    return fams


def apply_rules(paras, rule_files):
    """按顺序把 edits 模拟写进段落；返回没对上的规则"""
    miss = []
    for rf in rule_files:
        data = json.load(open(rf, encoding='utf-8'))
        for e in data.get('edits', []):
            if e.get('part') == 'footnote':
                continue
            ctx, old, new = e.get('ctx', ''), e.get('old', ''), e.get('new', '')
            hit = [p for p in paras if ctx and ctx in p[1]]
            if len(hit) != 1 or old not in hit[0][1]:
                # ctx 可能已被前面的规则改掉：退一步按 old 找唯一段
                cand = [p for p in paras if old and old in p[1]]
                if len(cand) == 1 and not hit:
                    hit = cand
                else:
                    miss.append((os.path.basename(rf), ctx, old))
                    continue
            p = hit[0]
            i = p[1].find(old, max(0, p[1].find(ctx)) if ctx in p[1] else 0)
            if i < 0:
                i = p[1].find(old)
            p[1] = p[1][:i] + new + p[1][i + len(old):]
    return miss


def find_hits(paras, fam):
    """在段落里找变体核心；落在采用核心里面的不算（如变体是采用核心的一部分）"""
    alts = sorted(fam['variants'], key=len, reverse=True)
    if not alts:
        return []
    pat = re.compile('|'.join(re.escape(v) for v in alts))
    hits = []
    for pid, txt in paras:
        adopt_spans = [(m.start(), m.end()) for m in re.finditer(re.escape(fam['adopt']), txt)]
        for m in pat.finditer(txt):
            if any(a <= m.start() and m.end() <= b for a, b in adopt_spans):
                continue
            hits.append((pid, txt, m.start(), m.end(), m.group(0)))
    return hits


def kept(hit, fam):
    pid, txt, s, e, _ = hit
    # 书名号里
    left = txt.rfind('《', 0, s)
    if left >= 0 and txt.find('》', left) >= e and txt.rfind('》', 0, s) < left:
        return '书名号内'
    for kp, snip in fam['keeps']:
        if kp and kp != pid:
            continue
        for m in re.finditer(re.escape(snip), txt):
            if m.start() <= s and e <= m.end():
                return f'允许保留“{snip}”'
    return None


def snippet(txt, s, e, w=12):
    return txt[max(0, s - w):s] + '【' + txt[s:e] + '】' + txt[e:e + w]


def main():
    args = sys.argv[1:]
    if not args or '-h' in args or '--help' in args:
        print(__doc__)
        sys.exit(2)
    list_mode = '--list' in args
    if list_mode:
        args.remove('--list')
    rules = []
    if '--rules' in args:
        i = args.index('--rules')
        rules = args[i + 1:]
        args = args[:i]
    glossary, sources = args[0], args[1:]
    fams = load_families(glossary)
    if not fams:
        print('总名词表里没有“名称族”一节或其中没有 F 开头的行。书名、章题核心名称必须建族。')
        sys.exit(1)

    bad = 0
    for src in sources:
        paras = load_paras(src)
        name = os.path.basename(src)
        if list_mode:
            for fam in fams:
                print(f'\n## {name}  {fam["id"]}（采用核心：{fam["adopt"]}；状态：{fam["status"]}）')
                alla = [fam['adopt']] + fam['variants']
                pat = re.compile('|'.join(re.escape(v) for v in sorted(alla, key=len, reverse=True)))
                by = {}
                for pid, txt in paras:
                    for m in pat.finditer(txt):
                        # 连带后面最多 6 个汉字，看出接的是什么后缀
                        tail = re.match(r'[一-鿿]{0,6}', txt[m.end():]).group(0)
                        by.setdefault(m.group(0), []).append(f'{pid}“{snippet(txt, m.start(), m.end() + len(tail), 6)}”')
                for core in alla:
                    locs = by.get(core, [])
                    print(f'- {core}：{len(locs)} 处' + ('' if not locs else '｜' + '；'.join(locs)))
            continue

        miss = apply_rules(paras, rules)
        print(f'\n# {name}')
        if miss:
            print(f'（{len(miss)} 条 edit 在本文件里没对上，可能不是这个文件的规则，已跳过）')
        for fam in fams:
            confirmed = '用户确认' in fam['status'] and '待' not in fam['status']
            hits = find_hits(paras, fam)
            keep_list, open_list = [], []
            for h in hits:
                why = kept(h, fam)
                (keep_list if why else open_list).append((h, why))
            print(f'## {fam["id"]} 采用“{fam["adopt"]}”，变体 {"、".join(fam["variants"])}：'
                  f'改后还剩 {len(hits)} 处，允许保留 {len(keep_list)}，未覆盖 {len(open_list)}')
            if not confirmed:
                print(f'  ！方向未经用户确认（状态：{fam["status"]}）——开审前必须问用户定')
                bad += 1
            for (pid, txt, s, e, core), why in keep_list:
                print(f'  保留 {pid} {snippet(txt, s, e)}（{why}）')
            for (pid, txt, s, e, core), _ in open_list:
                print(f'  未覆盖 {pid} {snippet(txt, s, e)}')
            bad += len(open_list)
    if not list_mode:
        print(f'\n合计未覆盖（含方向未定）{bad}')
        sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()

"""一致性审查 test：同一批变体候选，分别由 OpenAI Decisions 和大模型判断“是否同一对象”，比较结论和用时。

用法（Python 一律用 ~/.venvs/book-review/bin/python）：
  1. 准备  decisions_compare.py prepare 第NN章-段落.txt 变体候选.md 试验目录
           读 variant_pairs.py 的候选表，给每对补两端原句和所在标题，写出 candidates.json / candidates.md、llm-judgments.json 模板
  2. Decisions  decisions_compare.py decisions 试验目录 --confirm-live [--hi 0.8 --lo 0.05]
           从登录钥匙串读 bianjia-decisions 的密钥（不打印、不落盘），分批发送，写 decisions-results.json，记录用时
  3. 计时  decisions_compare.py stamp 试验目录 llm-start | llm-end
           主代理在派判断子代理前后各记一次，算大模型用时
  4. 对比  decisions_compare.py compare 试验目录
           写 对比报告.md：一致率、混淆表、分歧逐条、两边用时
三种结论：同一对象（应统一）／不同对象（或普通叙述，不改）／拿不准。
"""
import argparse, json, re, sys, time, importlib.util, hashlib
from collections import Counter
from datetime import datetime
from pathlib import Path

LABELS = ('同一对象', '不同对象', '拿不准')
BACKEND = Path.home() / '.codex/skills/bianjia-decisions/scripts'
HEAD = re.compile(r'^(第[一二三四五六七八九十\d]+章|\d+(\.\d+)+)\s*\S')
BATCH = 200          # 每次请求的题数（接口上限 400）
CTX = 2              # 每种写法最多给几处原句


def load_paras(path):
    """读 dump_docx.py 的“P号: 文字”，去掉修订标记，只留现文"""
    paras = {}
    for line in open(path, encoding='utf-8'):
        m = re.match(r'^(P\d+):\s?(.*)$', line.rstrip('\n'))
        if not m: continue
        t = re.sub(r'\[-.*?-\]', '', m.group(2))
        paras[m.group(1)] = re.sub(r'【黄:|】|\[\+|\+\]|\[图片\]', '', t)
    return paras


def heading_of(paras, pid):
    """往前找最近的节标题"""
    keys = list(paras); i = keys.index(pid) if pid in keys else -1
    while i >= 0:
        t = paras[keys[i]].strip()
        if HEAD.match(t) and len(t) < 40: return t
        i -= 1
    return ''


def sentence(text, word, width=40):
    """取包含 word 的一句（找不到句界就取前后 width 字）"""
    i = text.find(word)
    if i < 0: return text[:width * 2]
    a = max(text.rfind(c, 0, i) for c in '。；！？') + 1
    ends = [j for j in (text.find(c, i + len(word)) for c in '。；！？') if j >= 0]
    b = min(ends) + 1 if ends else len(text)
    if b - a > width * 3: a, b = max(0, i - width), min(len(text), i + len(word) + width)
    return text[a:b].strip()


def contexts(paras, word, where):
    out = []
    for loc in where:
        if loc in paras and word in paras[loc]:
            out.append({'loc': loc, 'heading': heading_of(paras, loc), 'text': sentence(paras[loc], word)})
        elif not loc.startswith('P'):
            out.append({'loc': loc, 'heading': '', 'text': '（图表内文字）'})
        if len(out) >= CTX: break
    return out


def cmd_prepare(a):
    paras = load_paras(a.paras)
    rows = []
    for line in open(a.pairs, encoding='utf-8'):
        m = re.match(r'^\| (.+?)（(.*?)） \| (.+?)（(.*?)） \| (.+?) \|$', line.strip())
        if not m or m.group(1) == '写法甲': continue
        rows.append(m.groups())
    if not rows: sys.exit('候选表里没有配对（或格式不是 variant_pairs.py 的输出）')
    cands = []
    for k, (wa, la, wb, lb, kind) in enumerate(rows, 1):
        cands.append({'id': f'C{k:03d}', 'a': wa, 'b': wb, 'kind': kind,
                      'where_a': la.split(), 'where_b': lb.split(),
                      'ctx_a': contexts(paras, wa, la.split()), 'ctx_b': contexts(paras, wb, lb.split())})
    out = Path(a.dir); out.mkdir(parents=True, exist_ok=True)
    (out / 'candidates.json').write_text(json.dumps({'source': str(a.paras), 'pairs': cands}, ensure_ascii=False, indent=1), 'utf8')
    (out / 'candidates.md').write_text(render(cands), 'utf8')
    tpl = out / 'llm-judgments.json'
    if not tpl.exists():
        tpl.write_text(json.dumps({c['id']: {'label': '', 'reason': ''} for c in cands}, ensure_ascii=False, indent=1), 'utf8')
    print(f'候选 {len(cands)} 对 → {out}；材料 {len(render(cands))} 字（发送给 Decisions 的就是 candidates.md 的内容）')


def render(cands):
    """两边看到的材料完全相同：Decisions 的 input 和大模型读的 candidates.md 是同一段文字"""
    out = ['以下是一本书稿某一章中脚本找出的“只差一个字”的写法配对。每对给出两种写法各自所在的原句和所在小节。', '']
    for c in cands:
        out.append(f'【{c["id"]}】写法甲“{c["a"]}”（出现于 {" ".join(c["where_a"])}）；写法乙“{c["b"]}”（出现于 {" ".join(c["where_b"])}）')
        for tag, ctx in (('甲', c['ctx_a']), ('乙', c['ctx_b'])):
            for x in ctx:
                out.append(f'  {tag} {x["loc"]}{("〔" + x["heading"] + "〕") if x["heading"] else ""}：{x["text"]}')
        out.append('')
    return '\n'.join(out)


QUESTION = ('判断候选 {id}：在这本书稿里，写法甲“{a}”和写法乙“{b}”是否指同一个对象（同一术语、方法、模型或名称的两种写法，应统一成一种）？'
            '如果两者是不同概念、上下位概念、函数与数值、有意对比的两个对象，或只是普通叙述用词而非名称，回答否。')


def stamp(out, name, value=None):
    f = out / 'timing.json'
    t = json.loads(f.read_text('utf8')) if f.exists() else {}
    t[name] = value if value is not None else datetime.now().isoformat(timespec='seconds')
    f.write_text(json.dumps(t, ensure_ascii=False, indent=1), 'utf8')


def cmd_decisions(a):
    out = Path(a.dir); cands = json.loads((out / 'candidates.json').read_text('utf8'))['pairs']
    if not a.confirm_live:
        print(f'未发送。共 {len(cands)} 对、{(len(cands) + BATCH - 1) // BATCH} 次请求、材料约 {len(render(cands))} 字；'
              '用户确认发送范围和费用后加 --confirm-live。'); return
    sys.path.insert(0, str(BACKEND))
    spec = importlib.util.spec_from_file_location('dec_backend', BACKEND / 'decisions.py'); dec = importlib.util.module_from_spec(spec); spec.loader.exec_module(dec)
    from keychain import load_key
    results, t0, calls = {}, time.time(), []
    for i in range(0, len(cands), BATCH):
        part = cands[i:i + BATCH]
        text = render(part)
        qs = [{'type': 'predicate', 'name': c['id'], 'instructions': QUESTION.format(**c)} for c in part]
        s = time.time()
        key = load_key()
        try:
            r = dec.create_decision(text, qs, key)
        finally:
            del key
        calls.append({'pairs': len(part), 'chars': len(text), 'seconds': round(time.time() - s, 2),
                      'sha256': hashlib.sha256(text.encode()).hexdigest()[:16]})
        for ans in r['answers']:
            if ans['type'] == 'refusal':
                results[ans['name']] = {'label': '拿不准', 'probability': None, 'refusal': True}; continue
            p = ans['probability']
            lab = '同一对象' if p >= a.hi else '不同对象' if p <= a.lo else '拿不准'
            results[ans['name']] = {'label': lab, 'probability': round(p, 4)}
    sec = round(time.time() - t0, 2)
    (out / 'decisions-results.json').write_text(json.dumps({'model': dec.MODEL, 'hi': a.hi, 'lo': a.lo, 'seconds': sec,
                                                             'calls': calls, 'results': results}, ensure_ascii=False, indent=1), 'utf8')
    stamp(out, 'decisions-seconds', sec)
    c = Counter(v['label'] for v in results.values())
    print(f'Decisions 完成：{len(results)} 对，用时 {sec} 秒；' + '、'.join(f'{k} {c[k]}' for k in LABELS))


def cmd_stamp(a):
    stamp(Path(a.dir), a.name); print(f'已记 {a.name}')


def cmd_compare(a):
    out = Path(a.dir)
    cands = {c['id']: c for c in json.loads((out / 'candidates.json').read_text('utf8'))['pairs']}
    d = json.loads((out / 'decisions-results.json').read_text('utf8'))
    llm = json.loads((out / 'llm-judgments.json').read_text('utf8'))
    t = json.loads((out / 'timing.json').read_text('utf8')) if (out / 'timing.json').exists() else {}
    bad = [k for k in cands if llm.get(k, {}).get('label') not in LABELS]
    if bad: sys.exit(f'大模型判断有 {len(bad)} 对未填或标签不对（只能是 {"／".join(LABELS)}）：{" ".join(bad[:20])}')
    miss = [k for k in cands if k not in d['results']]
    if miss: sys.exit(f'Decisions 结果缺 {len(miss)} 对：{" ".join(miss[:20])}')
    llm_sec = None
    if 'llm-start' in t and 'llm-end' in t:
        llm_sec = (datetime.fromisoformat(t['llm-end']) - datetime.fromisoformat(t['llm-start'])).total_seconds()
    conf = Counter((llm[k]['label'], d['results'][k]['label']) for k in cands)
    same = sum(v for (x, y), v in conf.items() if x == y)
    firm = [k for k in cands if d['results'][k]['label'] != '拿不准' and llm[k]['label'] != '拿不准']
    firm_same = sum(1 for k in firm if d['results'][k]['label'] == llm[k]['label'])
    # 最要紧的漏放：大模型认为同一对象（该改），Decisions 却判不同对象（会被直接放过）
    leak = [k for k in cands if llm[k]['label'] == '同一对象' and d['results'][k]['label'] == '不同对象']
    n = len(cands)
    md = [f'# 变体候选判断对比：Decisions 与大模型', '',
          f'- 候选 {n} 对；材料来源 {json.loads((out / "candidates.json").read_text("utf8"))["source"]}',
          f'- Decisions：模型 {d["model"]}，阈值 同一对象 ≥{d["hi"]}、不同对象 ≤{d["lo"]}，其余拿不准；用时 **{d["seconds"]} 秒**（{len(d["calls"])} 次请求）',
          f'- 大模型：' + (f'用时 **{round(llm_sec)} 秒**（{t["llm-start"]} → {t["llm-end"]}）' if llm_sec is not None else '未记录用时')
          + (f'；{t["llm-note"]}' if t.get('llm-note') else ''),
          f'- 两边结论完全一致 {same}/{n}（{same / n:.0%}）；两边都给了确定结论的 {len(firm)} 对中一致 {firm_same}（{(firm_same / len(firm)) if firm else 0:.0%}）',
          f'- **漏放风险**（大模型判同一对象、Decisions 判不同对象，按“Decisions 放行”流程会漏改）：{len(leak)} 对', '',
          '## 结论分布', '', '| | ' + ' | '.join(f'Decisions {x}' for x in LABELS) + ' |', '|---|' + '---|' * 3]
    for x in LABELS:
        md.append(f'| 大模型 {x} | ' + ' | '.join(str(conf[(x, y)]) for y in LABELS) + ' |')
    md += ['', '## 分歧逐条（先列漏放风险）', '', '| 编号 | 写法甲 | 写法乙 | 大模型 | 理由 | Decisions（概率） |', '|---|---|---|---|---|---|']
    diff = leak + [k for k in cands if k not in leak and llm[k]['label'] != d['results'][k]['label']]
    for k in diff:
        c, r = cands[k], d['results'][k]
        cell = lambda s: str(s).replace('|', '｜').replace('\n', ' ')
        md.append(f'| {k} | {cell(c["a"])} | {cell(c["b"])} | {llm[k]["label"]} | {cell(llm[k].get("reason", ""))} | {r["label"]}（{r["probability"]}） |')
    (out / '对比报告.md').write_text('\n'.join(md) + '\n', 'utf8')
    print(f'一致 {same}/{n}，漏放风险 {len(leak)} 对，Decisions {d["seconds"]} 秒，大模型 {round(llm_sec) if llm_sec is not None else "未记"} 秒 → {out / "对比报告.md"}')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    p = sub.add_parser('prepare'); p.add_argument('paras'); p.add_argument('pairs'); p.add_argument('dir'); p.set_defaults(f=cmd_prepare)
    p = sub.add_parser('decisions'); p.add_argument('dir'); p.add_argument('--confirm-live', action='store_true')
    p.add_argument('--hi', type=float, default=0.8); p.add_argument('--lo', type=float, default=0.05); p.set_defaults(f=cmd_decisions)
    p = sub.add_parser('stamp'); p.add_argument('dir'); p.add_argument('name', choices=['llm-start', 'llm-end']); p.set_defaults(f=cmd_stamp)
    p = sub.add_parser('compare'); p.add_argument('dir'); p.set_defaults(f=cmd_compare)
    a = ap.parse_args(); a.f(a)


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""每章合并：把固定顺序的命令一次跑完，遇到要判断的地方停下，列清楚交给主会话。

三段，按顺序各跑各的（前一段没过不会进下一段）：
  rules    门禁 → 拼避让 → 通读发现转规则 → 预检并合并 → 写入 Word（s1.docx）
  style    出稿后脚本（配置“出稿后”）→ 合并稿 → 重复批注清单、改动句复读清单、统计
  deliver  查锁文件、查用户有没有改过 → 旧版移废纸篓 → 放进交付目录 → 记指纹

用法: merge_chapter.py <书的工作目录> <章号> [rules|style|deliver|all] [选项]
  默认 all = rules + style（不含 deliver；交付要单独说）
  --remake            重新从通读发现生成 通读规则.json（会盖掉对它的手工删改）
  --set 键=值         本章给出稿后脚本的参数，如 --set 式下段开头='其中，X表示|其中，Y是'；记进 合并/本章参数.json
  --drop-comments 0,31  出稿后按批注编号删掉重复批注
  --force-deliver     交付目录里那份和上次放的不一样也照样替换（用户明确说可以才用）

主会话只需要改这三个文件，改完重跑即可（脚本不会盖掉它们）：
  合并/A规则-调整.json   第一轮规则按复核、用户意见调整后的版本（首次自动从 A 规则复制）
  合并/通读规则.json     通读规则（删掉与 A 重复的条目）
  合并/通读补-手写.json  通读与 A 改到同一处时，手写的合并改法
要在 s1.docx 上手工补修订的（如段首插“其中，”），存成 合并/s1b.docx，style 段会接着它做。
"""
import sys, os, re, json, shutil, subprocess, argparse, hashlib, datetime, difflib, zipfile, collections

PY = os.path.expanduser('~/.venvs/book-review/bin/python')
S1 = os.path.expanduser('~/.claude/skills/editorial-term-consistency-review/scripts')
S2 = os.path.expanduser('~/.claude/skills/tongdu/scripts')
S3 = os.path.dirname(os.path.abspath(__file__))

ap = argparse.ArgumentParser()
ap.add_argument('book'); ap.add_argument('chapter'); ap.add_argument('stage', nargs='?', default='all', choices=['rules', 'style', 'deliver', 'all'])
ap.add_argument('--remake', action='store_true'); ap.add_argument('--set', action='append', default=[])
ap.add_argument('--drop-comments', default=''); ap.add_argument('--force-deliver', action='store_true')
a = ap.parse_args()
book = os.path.abspath(a.book); n = int(a.chapter); NN = f'{n:02d}'
cfg = json.load(open(os.path.join(book, '本书配置.json'), encoding='utf-8'))
fill = lambda s: s.replace('{n}', str(n)).replace('{NN}', NN)
src = (cfg.get('原稿覆盖') or {}).get(NN) or fill(cfg['原稿'])
M = os.path.join(book, fill(cfg.get('合并目录', '通读/第{NN}章/合并'))); os.makedirs(M, exist_ok=True)
B = lambda name: os.path.join(book, name)
Mf = lambda name: os.path.join(M, name)
load = lambda p: json.load(open(p, encoding='utf-8'))
def save(p, d): json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
def run(cmd, env=None):
    r = subprocess.run(cmd, capture_output=True, text=True, cwd=book, env=env)
    return r.returncode, (r.stdout + r.stderr)
def stop(msg, code=2):
    print('\n■ 停在这里：' + msg); sys.exit(code)
def dump(docx):
    c, o = run([PY, f'{S1}/dump_docx.py', docx]); return o
def para_map(txt):
    return collections.OrderedDict((m.group(1), re.sub(r'\[批注\d+>|<批注\d+\]', '', m.group(2))) for m in re.finditer(r'^([PT]\d+): (.*)$', txt, flags=re.M))

f_mech, f_A, f_rev = B(f'第{NN}章-机械规则.json'), B(f'第{NN}章-修订规则.json'), B(f'第{NN}章-复核规则.json')
f_find = B(fill(cfg.get('通读发现', '通读/第{NN}章/通读发现.json')))
f_Aadj, f_td, f_hand, f_merged = Mf('A规则-调整.json'), Mf('通读规则.json'), Mf('通读补-手写.json'), Mf('合并规则.json')
final = Mf(f'第{NN}章合并稿.docx')

# =====================================================================
def stage_rules():
    print(f'【rules】第{NN}章')
    for p, what in ((src, '原稿'), (f_A, 'A 一致性规则'), (f_find, '通读发现')):
        if not os.path.exists(p): stop(f'缺{what}：{p}')
    # 1 门禁
    gl = B(f'第{NN}章-名词表.md')
    if os.path.exists(gl):
        c, o = run([PY, f'{S1}/check_glossary.py', gl])
        tail = (o.strip().splitlines() or [''])[-1]
        if c: stop(f'名词表门禁没过（{tail}）。打回 A 子代理补判断栏，过了再合并。')
        print(f'  门禁：通过（{tail}）')
    # 2 避让
    mech = load(f_mech) if os.path.exists(f_mech) else {'edits': [], 'comments': []}
    A = load(f_A)
    save(Mf('避让-机械+A.json'), {'edits': mech['edits'] + A['edits'], 'comments': mech.get('comments', []) + A.get('comments', [])})
    if not os.path.exists(f_Aadj): shutil.copy(f_A, f_Aadj); print('  已从 A 规则复制出 A规则-调整.json')
    if not os.path.exists(f_hand): save(f_hand, {'edits': [], 'comments': []})
    # 3 通读发现 → 规则
    if a.remake or not os.path.exists(f_td):
        base = Mf('通读发现-调整.json') if os.path.exists(Mf('通读发现-调整.json')) else f_find
        finds = load(base)
        c, o = run([PY, f'{S2}/make_rules.py', src, base, f_td, '--avoid', Mf('避让-机械+A.json')])
        if c:
            idx = sorted({int(x) for x in re.findall(r'^ - #(\d+) ', o, flags=re.M)})
            if not idx: stop('make_rules 失败：\n' + o[-800:])
            keep = [f for i, f in enumerate(finds, 1) if i not in idx]
            save(Mf('通读发现-自动.json'), keep)
            c2, o2 = run([PY, f'{S2}/make_rules.py', src, Mf('通读发现-自动.json'), f_td, '--avoid', Mf('避让-机械+A.json')])
            if c2: stop('剔掉定位不了的条目后 make_rules 仍失败：\n' + o2[-800:])
            paras = para_map(open(B(f'source/第{NN}章-docx段落.txt'), encoding='utf-8').read())
            av = load(Mf('避让-机械+A.json'))
            L = [f'# 第{NN}章 通读与 A 改到同一处的条目（{len(idx)} 条）', '',
                 '这些通读发现和一致性、机械规则落在同一处，脚本定位不了，已从自动规则里剔掉。',
                 '逐条看：A 已经改到的就不用管；A 没管到的，把合并后的改法手写进 `通读补-手写.json`（ctx 避开 A 要删的字）。', '']
            for i in idx:
                f = finds[i - 1]; P = f.get('P'); pt = paras.get(P, '')
                L.append(f'## #{i} {P}　' + (f'「{f.get("old")}」→「{f.get("new")}」' if 'old' in f else f'批注@「{f.get("target")}」：{f.get("text", "")[:60]}'))
                if f.get('note'): L.append(f'- 通读理由：{f["note"]}')
                key = f.get('old') or f.get('target') or ''; k0 = pt.find(key)
                def close(e):                      # 只列落点挨着这条发现的规则（前后 12 字内）
                    c0 = pt.find(e['ctx'])
                    if c0 < 0: return False
                    o0 = c0 + max(0, e['ctx'].find(e['old'])); o1 = o0 + len(e['old'])
                    return k0 < 0 or (o0 <= k0 + len(key) + 12 and k0 - 12 <= o1)
                near = [e for e in av['edits'] if close(e)]
                for e in near: L.append(f'- 同段已有规则：「{e["old"]}」→「{e["new"]}」（ctx「{e["ctx"]}」）')
                if not near:
                    same = sorted({e['old'] for e in av['edits'] if e['old'] and k0 >= 0 and e['old'] in pt[max(0, k0 - 3):k0 + len(key) + 3]})
                    if same: L.append('- 这几个字里含别处规则要改的词（' + '、'.join(f'「{w}」' for w in same) + '），脚本怕撞上没自动定位；这一处多半没人改，手写一条即可（ctx 取长一点）')
                    else: L.append('- 附近没有别的规则：可能是原文对不上（公式、空格），看原段：' + pt[:80])
                L.append('')
            open(Mf('待处理-通读与A同处.md'), 'w', encoding='utf-8').write('\n'.join(L))
            print(f'  通读发现 {len(finds)} 条，{len(idx)} 条与 A 同处已剔出 → 待处理-通读与A同处.md')
        t = load(f_td); print(f'  通读规则：修订 {len(t["edits"])} 处、批注 {len(t["comments"])} 条')
    else:
        print('  通读规则.json 已有，未重做（要重做加 --remake）')
    # 4 预检 + 合并
    files = [p for p in (f_mech, f_Aadj, f_rev, f_td, f_hand) if os.path.exists(p)]
    c, o = run([PY, f'{S2}/check_rules.py', src] + files + ['--merge', f_merged])
    lines = o.strip().splitlines()
    errs = [l for l in lines if l.startswith('错误')]
    print('  预检：' + next((l for l in reversed(lines) if l.startswith('修订 ')), ''))
    if c or errs:
        print('\n'.join('    ' + e.replace(M + os.sep, '').replace(book + os.sep, '') for e in errs))
        stop('预检有错误。按顺序是 机械 → A规则-调整 → 复核 → 通读规则 → 通读补-手写；\n'
             '  “ctx 在稿中找不到”= 前面的规则已经改到这段字：重复的删通读那条，不重复的改写 ctx 避开被改的字，或挪到 A规则-调整.json 最前面；\n'
             '  “疑似重复批注”= 删一条留更具体的，确是两件事加 "dup_ok": true。改完重跑本脚本。'
             + ('\n  另有待处理：' + Mf('待处理-通读与A同处.md') if os.path.exists(Mf('待处理-通读与A同处.md')) else ''))
    # 5 写入
    c, o = run([PY, f'{S1}/revise_docx.py', src, Mf('s1.docx'), f_merged])
    if c: stop('revise_docx 失败：\n' + o[-800:])
    miss = [l for l in o.splitlines() if '没找到' in l or '未找到' in l]
    print('  写入：' + next((l for l in o.splitlines() if l.startswith('修订 ')), '完成').split(' → ')[0] + (f'；批注跳转没对上 {len(miss)} 处（不影响修订）' if miss else ''))
    if os.path.exists(Mf('s1b.docx')) and os.path.getmtime(Mf('s1b.docx')) < os.path.getmtime(Mf('s1.docx')):
        print('  注意：合并/s1b.docx 比新的 s1.docx 旧，是上一次手工补的。要么在新 s1 上重做存成 s1b，要么删掉它。')
    if os.path.exists(Mf('待处理-通读与A同处.md')) and not load(f_hand)['edits']:
        print('  提醒：有“通读与 A 同处”的条目，而 通读补-手写.json 还是空的，确认都不用补。')

# =====================================================================
def stage_style():
    print(f'【style】第{NN}章')
    cur = Mf('s1b.docx') if os.path.exists(Mf('s1b.docx')) else Mf('s1.docx')
    if not os.path.exists(cur): stop('还没有 s1.docx，先跑 rules。')
    if os.path.basename(cur) == 's1b.docx': print('  从手工补过的 s1b.docx 接着做')
    pf = Mf('本章参数.json'); params = load(pf) if os.path.exists(pf) else {}
    for kv in a.set:
        k, v = kv.split('=', 1); params[k] = v
    if a.set: save(pf, params)
    skipped = []
    for i, st in enumerate(cfg.get('出稿后') or [], 1):
        script = B(st['脚本']) if not os.path.isabs(st['脚本']) else st['脚本']
        args = []
        if st.get('参数键'):
            v = params.get(st['参数键'])
            if not v: skipped.append(f'{st["名"]}（没给“{st["参数键"]}”）'); continue
            args.append(v)
        args += st.get('附加', [])
        nxt = Mf(f's{i + 1}-{os.path.splitext(os.path.basename(script))[0]}.docx')
        c, o = run([PY, script, cur, nxt] + args)
        if c or not os.path.exists(nxt): stop(f'出稿后脚本“{st["名"]}”失败：\n' + o[-600:])
        print(f'  {st["名"]}：' + (o.strip().splitlines() or ['完成'])[-1][:90]); cur = nxt
    if skipped: print('  没跑：' + '；'.join(skipped) + '。本章没有这类情况就不用管；有的话加 --set 重跑 style。')
    drop = a.drop_comments or params.get('删批注', '')
    if a.drop_comments: params['删批注'] = a.drop_comments; save(pf, params)
    if drop:
        before = Mf(f'第{NN}章合并稿-删重前.docx'); shutil.copy(cur, before)
        c, o = run([PY, f'{S3}/del_comments.py', before, final, drop])
        if c: stop('删批注失败：\n' + o[-400:])
        print(f'  删重复批注 {drop}：' + (o.strip().splitlines() or [''])[-1][:80])
    else:
        shutil.copy(cur, final)
    # 统计
    z = zipfile.ZipFile(final); x = z.read('word/document.xml').decode('utf-8', 'ignore')
    cm = z.read('word/comments.xml').decode('utf-8', 'ignore') if 'word/comments.xml' in z.namelist() else ''
    authors = collections.Counter(re.findall(r'<w:comment [^>]*w:author="([^"]*)"', cm))
    stat = f'插入 {len(re.findall("<w:ins ", x))}、删除 {len(re.findall("<w:del ", x))}、格式修订 {len(re.findall("<w:pPrChange ", x))}；批注 {sum(authors.values())} 条（' + '、'.join(f'{k} {v}' for k, v in authors.items()) + '）'
    # 重复批注
    c, dup = run([PY, f'{S3}/dupcom.py', final])
    ndup = int((re.search(r'多批注段 (\d+)', dup) or [0, 0])[1])
    # 改动句复读
    open(Mf('final.txt'), 'w', encoding='utf-8').write(dump(final))
    acc = Mf('接受后.docx'); run([PY, f'{S1}/accept_all.py', final, acc])
    new = list(para_map(dump(acc)).values()); old = list(para_map(dump(src)).values())
    os.remove(acc)
    sm = difflib.SequenceMatcher(None, old, new, autojunk=False); ch = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal': continue
        if tag == 'delete':
            ch.append('〔整段删去〕原：' + ' ／ '.join(o_[:60] for o_ in old[i1:i2])); continue
        for j in range(j1, j2):
            o_ = old[i1 + (j - j1)] if tag == 'replace' and i1 + (j - j1) < i2 else ''
            ch.append(new[j] + (('\n  > 原：' + o_) if o_ and len(o_) <= 40 else ''))
    L = [f'# 第{NN}章 改动句复读（{len(ch)} 段有改动）', '', '下面是接受全部修订后的样子（标题附原文）。读一遍：句子通不通、有没有改重了、改半截、前后句接不上。要看某段的修订痕迹，在 final.txt 里搜。', ''] + [f'{i}. {c_}\n' for i, c_ in enumerate(ch, 1)]
    open(Mf('改动句复读.md'), 'w', encoding='utf-8').write('\n'.join(L))
    rep = [f'# 第{NN}章 合并报告（{datetime.datetime.now():%m-%d %H:%M}）', '', f'- 合并稿：`{final}`', f'- {stat}',
           f'- 改动 {len(ch)} 段 → `改动句复读.md`（要读）', f'- 同段多条批注 {ndup} 段' + ('（看下面有没有重复，重复的用 --drop-comments 编号 重跑 style）' if ndup else ''),
           '- 没跑的出稿后脚本：' + ('；'.join(skipped) if skipped else '无'), '', '## 同段多条批注', '', '```', dup.strip(), '```', '']
    open(Mf('合并报告.md'), 'w', encoding='utf-8').write('\n'.join(rep))
    print(f'  {stat}')
    print(f'  改动 {len(ch)} 段 → 改动句复读.md；同段多条批注 {ndup} 段 → 合并报告.md')
    print('  接下来：读 改动句复读.md；有重复批注就删；都好了跑 deliver。')

# =====================================================================
def stage_deliver():
    print(f'【deliver】第{NN}章')
    if not os.path.exists(final): stop('还没有合并稿，先跑 rules 和 style。')
    ddir = cfg.get('交付目录'); name = fill(cfg.get('交付文件名', '第{NN}章-修订稿.docx'))
    if not ddir or not os.path.isdir(ddir): stop(f'配置里的交付目录不存在：{ddir}')
    tgt = os.path.join(ddir, name)
    locks = [f for f in os.listdir(ddir) if (f.startswith('.~') or f.startswith('~$')) and (f.lstrip('.~$') in name or name.endswith(f[2:]))]
    if locks: stop(f'交付目录里有锁文件 {locks}：用户正开着这份稿。请用户关掉再放；用户在改稿时只列问题、不替换文件。')
    md5 = lambda p: hashlib.md5(open(p, 'rb').read()).hexdigest()
    rec_f = B('交付记录.json'); rec = load(rec_f) if os.path.exists(rec_f) else {}
    if os.path.exists(tgt):
        if md5(tgt) == md5(final): print('  交付目录里已经是这一版，不用再放。'); return
        if rec.get(name, {}).get('md5') != md5(tgt) and not a.force_deliver:
            stop(f'交付目录里的 `{name}` 和我们上次放的不一样（没有记录，或用户改过）。\n  先问用户：是否已在这份上改过。用户说可以替换，再加 --force-deliver。')
        c, o = run(['osascript', '-e', f'tell application "Finder" to delete POSIX file "{tgt}"'])
        if c: stop('旧版移废纸篓失败：' + o[-300:])
        print('  旧版已移到废纸篓')
    shutil.copy(final, tgt); os.chmod(tgt, 0o644)
    if md5(tgt) != md5(final): stop('复制后两份不一致，请检查磁盘。')
    rec[name] = {'md5': md5(tgt), '时间': f'{datetime.datetime.now():%Y-%m-%d %H:%M}', '来源': os.path.relpath(final, book)}
    save(rec_f, rec)
    print(f'  已放进：{tgt}')

if a.stage in ('rules', 'all'): stage_rules()
if a.stage in ('style', 'all'): stage_style()
if a.stage == 'deliver': stage_deliver()

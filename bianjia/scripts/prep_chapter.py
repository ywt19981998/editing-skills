#!/usr/bin/env python3
"""每章准备：把固定命令一次跑完，写出《准备报告》。

只做不需要判断的事（导出、搜索、数数、对页码）；需要判断的列在报告里交给主会话。
遇到不对劲的情况（文件缺、页数对不上、图片全黑、原稿变了）记进“要处理”，不猜、不绕过。
已有的产物默认不重做（免得盖掉后来手工补的内容），加 --force 才重做。

用法: prep_chapter.py <书的工作目录> <章号> [--search 词1,词2,...] [--out 输出目录] [--force]
  --search  本章临时加搜的写法（主会话看过章题后补的）
  --out     产物写到别处（试跑、比对用），不动工作目录
配置: <书的工作目录>/本书配置.json（见 bianjia/references/本书配置-说明.md）
"""
import sys, os, re, json, subprocess, argparse, hashlib

PY = os.path.expanduser('~/.venvs/book-review/bin/python')
S1 = os.path.expanduser('~/.claude/skills/editorial-term-consistency-review/scripts')
S2 = os.path.expanduser('~/.claude/skills/tongdu/scripts')

ap = argparse.ArgumentParser()
ap.add_argument('book'); ap.add_argument('chapter')
ap.add_argument('--search', default=''); ap.add_argument('--out'); ap.add_argument('--force', action='store_true')
a = ap.parse_args()
book = os.path.abspath(a.book); n = int(a.chapter); NN = f'{n:02d}'
cfg = json.load(open(os.path.join(book, '本书配置.json'), encoding='utf-8'))
out = os.path.abspath(a.out) if a.out else os.path.join(book, 'source')
os.makedirs(out, exist_ok=True)

todo, done, info = [], [], []          # 要处理 / 已做 / 备查
def run(cmd, **kw):
    r = subprocess.run(cmd, capture_output=True, text=True, **kw)
    return r.returncode, r.stdout, r.stderr
def fill(s): return s.replace('{n}', str(n)).replace('{NN}', NN)

# ---------- 0 原稿 ----------
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from book_src import chapter_src
src, author_src, polish_dir = chapter_src(cfg, book, n)      # 润色过的章，底稿是润色后底稿
if not os.path.exists(src):
    sys.exit(f'停：找不到原稿 {src}')
if polish_dir:
    info.append(f'底稿是润色后底稿（`{os.path.relpath(src, book)}`），作者原稿 `{author_src}`；一致性、通读都在润色后底稿上做，合并时脚本会把润色改动补标回去')
    hand = os.path.join(polish_dir, '复查汇总.md')
    if os.path.exists(hand) and '留给一致性审查的名词' in open(hand, encoding='utf-8').read():
        todo.append(f'润色留下了“留给一致性审查的名词”（`{os.path.relpath(hand, book)}`）：派 A 一致性时写进派单，当变体候选的线索。')
else:
    todo.append('这一章还没有润色（没找到润色后底稿）。完整流程第一步是润色（技能 runse）；编辑说这一章不润色的，要告诉编辑：通读不查句子和标点，这两样就没人查了。')
ch = (cfg.get('章') or {}).get(str(n)) or (cfg.get('章') or {}).get(NN) or {}

# ---------- 1 段落文字 ----------
f_para = os.path.join(out, f'第{NN}章-docx段落.txt')
code, txt, err = run([PY, f'{S1}/dump_docx.py', src])
if code: sys.exit(f'停：dump_docx 失败\n{err[-500:]}')
if os.path.exists(f_para) and open(f_para, encoding='utf-8').read() != txt and not a.force:
    todo.append(f'原稿与已有的 `{os.path.basename(f_para)}` 不一致（原稿可能换过版本）。确认后加 --force 重导。')
    txt = open(f_para, encoding='utf-8').read()
else:
    open(f_para, 'w', encoding='utf-8').write(txt); done.append(f'段落文字 → {os.path.basename(f_para)}')
paras = [(m.group(1), m.group(2)) for m in re.finditer(r'^([PT]\d+): (.*)$', txt, flags=re.M)]
clean = lambda v: re.sub(r'\[批注\d+>|<批注\d+\]', '', v)
ref = next((int(k[1:]) for k, v in paras if clean(v).strip() in ('参考文献', '本章参考文献')), None)
last = int(paras[-1][0][1:]) if paras else 0
end = (ref - 1) if ref else last
if not ref: todo.append('没找到“参考文献”段，正文末尾按最后一段算，请核对。')
body = [(k, clean(v)) for k, v in paras if int(k[1:]) <= end]
ncom = len(re.findall(r'^批注\d+', txt, flags=re.M))
info.append(f'段落 P1–P{last}；正文到 P{end}' + (f'，参考文献从 P{ref} 起' if ref else ''))
if ncom: info.append(f'原稿自带批注 {ncom} 条（写入修订后会原样保留；派单里提醒子代理别当成我们的批注）')
nrev = len(re.findall(r'\[(?:插入|删除)', txt))
if nrev: todo.append(f'原稿里已有修订标记约 {nrev} 处：revise_docx 读不到已有修订里的文字，先问编辑是否接受后再审。')

# ---------- 2 排版 PDF 分页 ----------
f_pdf = os.path.join(out, f'第{NN}章-PDF分页.txt')
pdf = cfg.get('排版PDF'); rng = ch.get('pdf')
if not pdf or not rng:
    todo.append('配置里没有本章的排版 PDF 页码（“章”→ pdf: [起, 止]），PDF 分页没导。')
elif not os.path.exists(pdf):
    todo.append(f'找不到排版 PDF：{pdf}')
else:
    want = rng[1] - rng[0] + 1
    have = len(re.findall(r'^\[p\.\d+ \| PDF第\d+页\]', open(f_pdf, encoding='utf-8').read(), flags=re.M)) if os.path.exists(f_pdf) else 0
    if have == want and not a.force:
        done.append(f'PDF 分页已有（{have} 页），未重导')
    else:
        import fitz
        d = fitz.open(pdf); off = cfg.get('PDF页码偏移', 0); buf = []
        for p in range(rng[0], rng[1] + 1):
            buf.append(f'[p.{p} | PDF第{p + off}页]\n' + d[p + off - 1].get_text())
        open(f_pdf, 'w', encoding='utf-8').write('\n'.join(buf))
        done.append(f'PDF 分页 → {os.path.basename(f_pdf)}（p.{rng[0]}–{rng[1]}，{want} 页' + (f'；原先只有 {have} 页，已重导' if have else '') + '）')
    pg = open(f_pdf, encoding='utf-8').read()
    title = next((v for k, v in body[:8] if v.strip() and '[图片]' not in v and not re.match(r'^第\s*\d+\s*章$', v.strip())), '')
    squash = lambda s: re.sub(r'\s+', '', s)
    if title and squash(title)[:8] not in squash(pg[:1500]):
        todo.append(f'PDF 起始页里没找到章题“{title[:20]}”，页码范围可能不对，请核对配置。')

# ---------- 3 不可见文字 ----------
code, o, err = run([PY, f'{S1}/hidden_text.py', src])
hid = [l for l in o.splitlines() if l.strip()]
info.append('不可见文字：' + (f'{len(hid)} 段（' + '；'.join(hid[:3]) + ('…' if len(hid) > 3 else '') + '）' if hid else '无'))

# ---------- 4 图片 ----------
figdir = os.path.join(out, f'第{NN}章-图片'); fmd = os.path.join(figdir, 'figures.md')
if os.path.exists(fmd) and not a.force:
    done.append('图片已导出过，未重导（figures.md 可能有手工补的内容）')
else:
    code, o, err = run([PY, f'{S2}/extract_figures.py', src, figdir, '--from', '1', '--to', str(end)])
    if code: todo.append(f'extract_figures 失败：{(err or o)[-300:]}')
    else: done.append('图片 → ' + (o.strip().splitlines() or [''])[-1])
if os.path.exists(fmd):
    ft = open(fmd, encoding='utf-8').read()
    rows = [l for l in ft.splitlines() if re.match(r'^\| P\d+', l)]
    bad = [l for l in rows if not re.search(r'\|\s*图\s*\d+[-－—.]\d+', l)]
    nums = [int(m.group(1)) for l in rows for m in [re.search(r'图\s*\d+[-－—.](\d+)', l)] if m]
    miss = sorted(set(range(1, max(nums) + 1)) - set(nums)) if nums else []
    info.append(f'图片 {len(rows)} 张；图题见 figures.md')
    if bad: todo.append('这些图没配上图题或图题不像“图N-M”，请看是否配错：' + '；'.join(re.sub(r'\s*\|\s*', ' ', b).strip()[:40] for b in bad[:6]))
    if miss: todo.append(f'图号不连续，缺 图{n}-' + f'、图{n}-'.join(map(str, miss)) + '（可能是 Word 原生图表或图题配错）')
    if '原生图表' in ft: todo.append('有 Word 原生图表导不出：看排版页面截图，数据从 word/charts/chartN.xml 抽（做法见流程文档）。')
    if '转不了' in ft and not re.search(r'转不了 0 张', ft): todo.append('有 EMF/WMF 图转不了，看排版页面。')
    try:
        from PIL import Image, ImageStat
        dark = []
        for fn in sorted(os.listdir(os.path.join(figdir, '图片'))):
            if fn.lower().endswith('.png'):
                if ImageStat.Stat(Image.open(os.path.join(figdir, '图片', fn)).convert('L')).mean[0] < 25: dark.append(fn)
        if dark: todo.append('这些图导出后几乎全黑，看不见字：' + '、'.join(dark))
    except Exception as e:
        info.append(f'（图片亮度没查：{e}）')

# ---------- 5 候选 ----------
gl = os.path.join(book, '名词表.md'); must = []
if os.path.exists(gl):
    for m in re.finditer(r'必搜[^：:|]{0,12}[：:]([^|\n]*)', open(gl, encoding='utf-8').read()):
        seg = re.split(r'[）)；;。]', m.group(1))[0]
        for w in re.split(r'[、，,]', seg):
            w = re.sub(r'（[^）]*$|\*+', '', w).strip()
            if 2 <= len(w) <= 20 and not re.search(r'\d+\s*处', w): must.append(w)
words = []
for w in must + (cfg.get('常搜') or []) + [s['搜'] for s in cfg.get('全书替换') or []] + [w for w in a.search.split(',') if w.strip()]:
    if w.strip() and w.strip() not in words: words.append(w.strip())
f_cand = os.path.join(out, f'第{NN}章-候选.md'); f_var = os.path.join(out, f'第{NN}章-变体候选.md')
figtxt = os.path.join(out, f'第{NN}章-图片文字.md')
cmd = [PY, f'{S1}/term_candidates.py', f_para] + ([figtxt] if os.path.exists(figtxt) else []) + (['--search', ','.join(words)] if words else []) + ['-o', f_cand]
code, o, err = run(cmd)
if code: todo.append(f'term_candidates 失败：{(err or o)[-300:]}')
else: done.append(f'候选 → {os.path.basename(f_cand)}（必搜 {len(words)} 个：名词表 {len(must)}、配置和临时 {len(words) - len(set(must))}）')
code, o, err = run([PY, f'{S1}/variant_pairs.py', f_para, f_var])
if code: todo.append(f'variant_pairs 失败：{(err or o)[-300:]}')
else: done.append(f'变体候选 → {os.path.basename(f_var)}')

# ---------- 6 全书替换与残留：只列位置，改不改主会话判断 ----------
def hits(pat, is_re=False, width=14):
    res = []
    for k, v in body:
        for m in (re.finditer(pat, v) if is_re else re.finditer(re.escape(pat), v)):
            s, e = max(0, m.start() - width), min(len(v), m.end() + width)
            res.append(f'{k}　…{v[s:m.start()]}【{m.group(0)}】{v[m.end():e]}…')
    return res
sec = []
for it in cfg.get('全书替换') or []:
    h = hits(it['搜'])
    sec.append(f'### “{it["搜"]}”：{len(h)} 处　（{it.get("说明", "")}）\n' + ('\n'.join(f'- {x}' for x in h) if h else '- 无'))
builtin = [
    ('图表、算法、式的编号不带章号（论文改书的残留）', rf'(?<![\d\-－.])(?:图|表|算法|式|公式)\s?[（(]?\d{{1,2}}[)）]?(?![\d\-－—.]|\s*[-－—.]\s*\d)'),
    ('数字中间夹句号（小数点误作句号）', r'\d。\d'),
]
for name, pat in builtin:
    h = hits(pat, True)
    if h: sec.append(f'### {name}：{len(h)} 处\n' + '\n'.join(f'- {x}' for x in h))

# ---------- 7 编号接着从哪起 ----------
def last_no(fn, pat):
    p = os.path.join(book, fn)
    if not os.path.exists(p): return None
    v = [int(x) for x in re.findall(pat, open(p, encoding='utf-8').read(), flags=re.M)]
    return max(v) if v else None
T, D, R = last_no('名词表.md', r'^\| T(\d+)'), last_no('待定决策.md', r'^\| D-(\d+)'), last_no('回查清单.md', r'^\| R(\d+)')
nxt = '；'.join(f'{a_} 从 {a_}{"-" if a_ == "D" else ""}{(b_ or 0) + 1} 起' for a_, b_ in (('T', T), ('D', D), ('R', R)))

# ---------- 报告 ----------
rep = [f'# 第{NN}章 准备报告', '',
       f'原稿：`{src}`', f'原稿指纹：{hashlib.md5(open(src, "rb").read()).hexdigest()[:12]}', '',
       '## 一、要处理（' + (f'{len(todo)} 项' if todo else '无') + '）', ''] + ([f'{i}. {t}' for i, t in enumerate(todo, 1)] or ['没有异常。']) + ['',
       '## 二、已做', ''] + [f'- {d}' for d in done] + ['',
       '## 三、备查', ''] + [f'- {i}' for i in info] + [f'- 本章新编号：{nxt}；问题号 {n}-01 起', '',
       '## 四、要你判断的替换（脚本只列位置；该改的写进 `第' + NN + '章-机械规则.json`，可用 mech_rules.py 生成）', ''] + (sec or ['配置里没有“全书替换”，也没搜到残留。']) + ['',
       '## 五、接下来', '',
       '1. 处理第一节各项；看 figures.md 的图题有没有配错。',
       '2. 判断第四节各处，写机械规则并预检。',
       '3. 读章首和各节标题，写派单里的“本章重点防范”；派单给用户看过再派。', '']
f_rep = os.path.join(out, f'第{NN}章-准备报告.md')
open(f_rep, 'w', encoding='utf-8').write('\n'.join(rep))
print('\n'.join(rep[:10 + len(todo) + len(done) + len(info) + 6]))
print(f'\n（第四节共 {len(sec)} 组，全文见 {f_rep}）')
sys.exit(2 if todo else 0)

#!/usr/bin/env python3
"""从润色修订稿生成“润色后底稿”：把润色的修订全部接受，得到一份不带修订的干净稿，给后面的一致性审查和通读当底稿。

用法: accept_polish.py 润色修订稿.docx 润色后底稿.docx 正文.txt 终稿.json [更多.json]
  后两样用来核对：润色后底稿里每一段的文字应当等于终稿（没改的段等于原文）。对不上就不算数，退出码为 1。
  只认修订人是“Claude（润色）”的修订；稿子里有别人的修订时停下（润色修订稿不该有）。
"""
import sys, os, zipfile, json, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from docx_units import q, load, has_revisions, body_rows, read_slines

ap = argparse.ArgumentParser()
ap.add_argument('src'); ap.add_argument('dst'); ap.add_argument('txt'); ap.add_argument('results', nargs='+')
ap.add_argument('--author', default='Claude（润色）')
a = ap.parse_args()

zin, doc = load(a.src)
others = {e.get(q('author')) for e in doc.iter(q('ins'), q('del'))} - {a.author}
if others or next(doc.iter(q('moveFrom'), q('moveTo')), None) is not None:
    sys.exit(f'停：稿子里有不是润色写的修订（{"、".join(sorted(x or "无名" for x in others)) or "移动修订"}），不能自动接受。')
n_del = n_ins = 0
for d in list(doc.iter(q('del'))):                 # 删除的内容：整个去掉
    d.getparent().remove(d); n_del += 1
for i in list(doc.iter(q('ins'))):                 # 插入的内容：去掉修订外壳，留下里面的字
    par = i.getparent(); k = par.index(i)
    for ch in list(i): par.insert(k, ch); k += 1
    par.remove(i); n_ins += 1
assert not has_revisions(doc)

# 核对：每段文字等于终稿
slines = read_slines(a.txt)
res = {}
for f in a.results:
    for it in json.load(open(f, encoding='utf-8')):
        if isinstance(it, dict) and it.get('S'): res[it['S']] = it
want = []
for sid, tag, text in slines:
    it = res.get(sid, {})
    new = (it.get('new') or '').strip() if (not tag and it.get('changed')) else ''
    want.append((sid, tag, new or text))
rows = body_rows(doc)
def same(row, s): return row[0] == 't' if s[1].startswith('【表格') else (row[0] == 'p' and row[1].strip() == s[2].strip())
starts = [k for k in range(len(rows) - len(want) + 1) if all(same(rows[k + j], s) for j, s in enumerate(want))]
if len(starts) != 1:
    bad = []
    if len(rows) >= len(want):                     # 找出是哪几段对不上，方便查
        best = max(range(len(rows) - len(want) + 1), key=lambda k: sum(same(rows[k + j], s) for j, s in enumerate(want)))
        bad = [s[0] for j, s in enumerate(want) if not same(rows[best + j], s)][:10]
    sys.exit(f'核对没过：接受修订后的文字和终稿对不上（对不上的段：{"、".join(bad) or "整体错位"}）。没有写文件。')

zout = zipfile.ZipFile(a.dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True) if item.filename == 'word/document.xml' else zin.read(item.filename)
    zout.writestr(item, data)
zout.close()
changed = sum(1 for sid, tag, text in slines if not tag and res.get(sid, {}).get('changed') and (res[sid].get('new') or '').strip())
print(f'接受了 {n_ins} 处插入、{n_del} 处删除；{len(want)} 行逐行等于终稿（其中润色改过的 {changed} 段）→ {a.dst}')

# 导出本章图片，供逐张看图内文字（错别字、单位写法）：只读底稿，不改稿子
# 用法: extract_figures.py 底稿.docx 输出目录 [--from P233] [--to P288]
# 输出: 输出目录/图片/*.png、figures.md（段号、图题、文件名）；公式小图跳过，EMF/WMF 转不了的、Word 原生图表导不出的列出来，要看排版页面；透明底垫白
import sys, os, re, io, zipfile, posixpath
from hashlib import md5
from lxml import etree
from PIL import Image

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V = 'urn:schemas-microsoft-com:vml'
def q(t): return f'{{{W}}}{t}'
MIN_EMU = 360000 * 1.5     # 显示尺寸小于 1.5 厘米的多是公式或图标
MAX_SIDE = 1600

args = sys.argv[1:]
def opt(name):
    if name in args:
        i = args.index(name); v = args[i + 1]; del args[i:i + 2]; return int(v.lstrip('PpTt'))
p_from, p_to = opt('--from'), opt('--to')
src, out = args[0], args[1]
os.makedirs(os.path.join(out, '图片'), exist_ok=True)

zf = zipfile.ZipFile(src)
doc = etree.fromstring(zf.read('word/document.xml'))
rels = {r.get('Id'): r.get('Target') for r in etree.fromstring(zf.read('word/_rels/document.xml.rels'))}

# P 号与 dump_docx.py 一致：正文段落和表格的每一行各占一个号
plabel, n = {}, 0
for el in doc.find(q('body')):
    if el.tag == q('p'):
        n += 1
        for x in el.iter(q('p')): plabel[x] = n
    elif el.tag == q('tbl'):
        for tr in el.iter(q('tr')):
            n += 1
            for x in tr.iter(q('p')): plabel[x] = n
plist = list(doc.iter(q('p')))
text = lambda p: ''.join('-' if t.tag == q('noBreakHyphen') else t.text or '' for t in p.iter(q('t'), q('noBreakHyphen'))).strip()

def caption(i):
    for j in list(range(i + 1, min(i + 4, len(plist)))) + list(range(i - 1, max(i - 3, -1), -1)):
        t = text(plist[j])
        if re.match(r'^图\s*[\dA-Z]', t) and len(t) < 80: return t
    return ''

seen, rows, skipped, failed, same, first = set(), [], 0, [], [], {}
charts = []
for i, p in enumerate(plist):
    P = plabel.get(p, 0)
    if (p_from and P < p_from) or (p_to and P > p_to): continue
    for draw in list(p.iter(q('drawing'))) + list(p.iter(q('pict'))):
        ext = next(draw.iter(f'{{{WP}}}extent'), None)
        gd = next(draw.iter(f'{{{A}}}graphicData'), None)
        if gd is not None and 'chart' in (gd.get('uri') or ''):   # Word 原生图表（柱状图等）不是图片，导不出，单列提示
            charts.append((P, caption(i))); continue
        if ext is not None and min(int(ext.get('cx')), int(ext.get('cy'))) < MIN_EMU: skipped += 1; continue
        ids = [b.get(f'{{{R}}}embed') for b in draw.iter(f'{{{A}}}blip')] + [v.get(f'{{{R}}}id') for v in draw.iter(f'{{{V}}}imagedata')]
        for rid in filter(None, ids):
            if (P, rid) in seen: continue
            seen.add((P, rid))
            path = posixpath.normpath(posixpath.join('word', rels[rid]))
            k = sum(1 for x in rows if x[0] == P) + 1
            h = md5(zf.read(path)).hexdigest()
            if h in first:                       # 同一张图插了两次：不重复看，但要提示
                same.append((P, first[h])); continue
            first[h] = P
            try:
                im = Image.open(io.BytesIO(zf.read(path)))
                if im.mode in ('RGBA', 'LA', 'P') or 'transparency' in im.info:   # 透明底垫白，免得转 RGB 后变黑、黑字看不见
                    im = im.convert('RGBA'); bg = Image.new('RGBA', im.size, (255, 255, 255, 255))
                    bg.alpha_composite(im); im = bg
                im = im.convert('RGB')
                im.thumbnail((MAX_SIDE, MAX_SIDE))
                name = f'P{P}-{k}.png'
                im.save(os.path.join(out, '图片', name))
                rows.append((P, caption(i), name))
            except Exception:
                failed.append((P, caption(i), posixpath.basename(path)))

with open(os.path.join(out, 'figures.md'), 'w', encoding='utf-8') as f:
    f.write(f'# 本章图片\n\n稿件：`{os.path.basename(src)}`　导出 {len(rows)} 张，跳过小图 {skipped} 张，转不了 {len(failed)} 张。\n\n')
    f.write('| 段 | 图题 | 文件 |\n|---|---|---|\n' + ''.join(f'| P{P} | {c} | 图片/{n_} |\n' for P, c, n_ in rows))
    if same:
        f.write('\n## 同一张图重复插入（请核对）\n\n' + ''.join(f'- P{P} 与 P{a} 完全相同\n' for P, a in same))
    if charts:
        f.write('\n## Word 原生图表（导不出：看排版页面，数据在 word/charts/chartN.xml）\n\n' + ''.join(f'- P{P} {c}\n' for P, c in charts))
    if failed:
        f.write('\n## 转不了（要看排版页面）\n\n' + ''.join(f'- P{P} {c}：{n_}\n' for P, c, n_ in failed))
print(f'导出 {len(rows)} 张，跳过小图 {skipped} 张，重复 {len(same)} 张，转不了 {len(failed)} 张，原生图表 {len(charts)} 个 → {out}')

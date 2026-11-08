"""公式审校清单导出：所有交付都用本脚本生成，版式固定，与 references/样例-正斜体重审清单.pdf 一致。
用法: export_review.py 审校记录.json 输出目录 [--formats pdf|md|both]
记录格式见 references/report.md“审校记录 JSON”。本脚本只排版、做一致性检查，不判定对错；导出后仍须渲染目视核对。
原页截图和红框都由本脚本按每条的 boxes 从 source_pdf 生成（红框为必做项），不接受自备截图；
找不到目标、目标不唯一、字形条目框数少于 count 时不导出。
"""
import argparse, json, re, sys
from collections import Counter, OrderedDict
from pathlib import Path
from xml.sax.saxutils import escape
import pymupdf
from PIL import Image as PILImage, ImageDraw
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Image, Spacer, KeepTogether, PageBreak

SECTIONS = OrderedDict([('确认', '确认问题'), ('待复核', '待专业复核'), ('纠偏', '撤回与纠偏记录')])
KINDS = ('字形', '数学')
STYLES = ('upright', 'italic', 'bold-italic', 'bold')
SUP = {'normal': '', 'sub': 'sub', 'super': 'super'}
DPI = 200                       # 截图分辨率；导出器按约 110dpi 显示
MARK = re.compile(r'【(.+?)】')  # find 里用【】括出要框的字

sys.path.insert(0, str(Path(__file__).resolve().parent))
from symbol_inventory import decode   # 还原 MathType 私用区编码和 Unicode 数学字母，便于按文字查找

def page_lines(page):
    """每行返回 [(字, bbox)]，字已还原，空白保留但查找时忽略"""
    out = []
    for b in page.get_text('rawdict')['blocks']:
        for l in b.get('lines', []):
            chars = [(decode(c['c'], s['font'])[0], pymupdf.Rect(c['bbox'])) for s in l['spans'] for c in s['chars']]
            if chars: out.append(chars)
    return out

def find_targets(lines, find):
    """在本页逐行查找 find（忽略空白）；返回每处匹配的 [框选目标的 Rect…] 和该行文字"""
    parts = MARK.split(find)            # 偶数位是上下文，奇数位是要框的字
    plain = ''.join(parts).replace(' ', '')
    if not plain: return []
    spans, pos = [], 0
    for k, t in enumerate(parts):
        t = t.replace(' ', '')
        if k % 2: spans.append((pos, pos + len(t)))
        pos += len(t)
    if not spans: spans = [(0, len(plain))]   # 没有【】就框整段
    hits = []
    for chars in lines:
        keep = [c for c in chars if not c[0].isspace()]
        text = ''.join(c[0] for c in keep)
        i = text.find(plain)
        while i >= 0:
            rects = []
            for a, z in spans:
                r = pymupdf.Rect(keep[i + a][1])
                for c in keep[i + a + 1:i + z]: r |= c[1]
                rects.append(r)
            hits.append((rects, ''.join(c[0] for c in chars).strip()))
            i = text.find(plain, i + 1)
    return hits

def snap(rect, lines, text=None):
    """rect（多来自符号清单、扫描结果）收紧到其中各字的字形范围；给了 text 只取这些字（免得带上相邻行）；
    框里没有字（图片区域）就原样返回"""
    hit = [c[1] for chars in lines for c in chars
           if not c[0].isspace() and (not text or c[0] in text)
           and rect.contains(pymupdf.Point((c[1].x0 + c[1].x1) / 2, (c[1].y0 + c[1].y1) / 2))]
    if not hit: return rect
    r = pymupdf.Rect(hit[0])
    for h in hit[1:]: r |= h
    return r

def resolve_boxes(d, doc):
    """把每条 boxes 换成页面坐标，算出截图范围；返回错误列表"""
    pymupdf.TOOLS.set_small_glyph_heights(True)   # 字符框按字形高度，不按整行行高，红框才贴字
    errs = []
    for e in d['items']:
        tag = e['id']
        if not 1 <= e['page'] <= doc.page_count:
            errs.append(f'{tag}：第 {e["page"]} 页超出 source_pdf 页数 {doc.page_count}'); continue
        page = doc[e['page'] - 1]; lines = None; rects = []
        for b in e['boxes']:
            if 'rect' in b:
                r = b['rect']
                if not (isinstance(r, list) and len(r) == 4 and all(isinstance(v, (int, float)) for v in r)):
                    errs.append(f'{tag}：rect 必须是 [x0, y0, x1, y1]'); continue
                r = pymupdf.Rect(r)
                if r.is_empty or not r.intersects(page.rect): errs.append(f'{tag}：rect {b["rect"]} 不在第 {e["page"]} 页内'); continue
                lines = lines or page_lines(page)
                rects.append(snap(r, lines, b.get('text')))
            elif b.get('find'):
                lines = lines or page_lines(page)
                hits = find_targets(lines, b['find']); n = b.get('nth')
                if not hits:
                    errs.append(f'{tag}：第 {e["page"]} 页找不到“{b["find"]}”（分式、上下标常被拆行，可改用符号清单或扫描结果里的 rect）')
                elif n is None and len(hits) > 1:
                    errs.append(f'{tag}：“{b["find"]}”在第 {e["page"]} 页出现 {len(hits)} 次，请加长上下文或用 nth 指定第几处：'
                                + '；'.join(f'{k + 1}. {t[:30]}' for k, (_, t) in enumerate(hits)))
                elif n is not None and not (isinstance(n, int) and 1 <= n <= len(hits)):
                    errs.append(f'{tag}：nth={n} 超出范围（共 {len(hits)} 处）')
                else:
                    rects += hits[(n or 1) - 1][0]
            else:
                errs.append(f'{tag}：boxes 的每一项要给 find 或 rect')
        if not rects: continue        # 每个 box 都已报错
        if e['kind'] == '字形' and len(rects) < e['count'] and not e.get('elsewhere'):
            errs.append(f'{tag}：字形条目 count={e["count"]}，但只框了 {len(rects)} 处；每个错误位置都要框到，'
                        '不在本页截图里的位置写进 elsewhere')
        u = pymupdf.Rect(rects[0])
        for r in rects[1:]: u |= r
        if e.get('clip'):
            clip = pymupdf.Rect(e['clip'])
            if not clip.contains(u): errs.append(f'{tag}：clip 没有包住全部红框'); continue
        else:   # 默认取红框所在的整行（与红框上下重叠的各行）再留边
            clip = pymupdf.Rect(u)
            for chars in (lines or page_lines(page)):
                lb = pymupdf.Rect(chars[0][1])
                for c in chars[1:]: lb |= c[1]
                if min(lb.y1, u.y1) - max(lb.y0, u.y0) > 0.3 * min(lb.height, u.height): clip |= lb
            m = e.get('margin', 6)        # 左右留 6pt，上下只留 3pt，免得带出相邻行的半截字
            clip = (clip + (-m, -3, m, 3)) & page.rect
        e['_rects'], e['_clip'] = rects, clip
    return errs

def render_evidence(e, doc, out):
    """按 clip 截原页、画不遮字的红框；同时留一张无框图作工作证据"""
    page = doc[e['page'] - 1]; clip = e['_clip']
    pix = page.get_pixmap(clip=clip, dpi=DPI, alpha=False)
    im = PILImage.frombytes('RGB', (pix.width, pix.height), pix.samples)
    (out / '无框').mkdir(parents=True, exist_ok=True); im.save(out / '无框' / f'{e["id"]}.png')
    k, pad = DPI / 72, 1.5           # 红框外扩 1.5pt，不压字
    dr = ImageDraw.Draw(im)
    for r in e['_rects']:
        dr.rectangle([(r.x0 - pad - clip.x0) * k, (r.y0 - pad - clip.y0) * k, (r.x1 + pad - clip.x0) * k, (r.y1 + pad - clip.y0) * k],
                     outline=(220, 0, 0), width=3)
    p = out / f'{e["id"]}.png'; im.save(p); e['_image'] = p

def fonts(font_dir):
    """注册字体；每种字形按顺序找第一个有该字的字体（Times 缺 ∬ ∇ ⋯ 等时用 STIX Two 补）"""
    mac = sys.platform == 'darwin'
    table = [('CN', 'Songti.ttc', 6), ('TR', 'Times New Roman.ttf', 0), ('TI', 'Times New Roman Italic.ttf', 0),
             ('TBI', 'Times New Roman Bold Italic.ttf', 0), ('TB', 'Times New Roman Bold.ttf', 0),
             ('SX', 'STIXTwoText.ttf', 0), ('SXI', 'STIXTwoText-Italic.ttf', 0), ('AU', 'Arial Unicode.ttf', 0)] if mac else \
            [('CN', 'simsun.ttc', 0), ('TR', 'times.ttf', 0), ('TI', 'timesi.ttf', 0), ('TBI', 'timesbi.ttf', 0), ('TB', 'timesbd.ttf', 0)]
    have = set()
    for name, file, idx in table:
        p = font_dir / file
        if p.exists():
            pdfmetrics.registerFont(TTFont(name, str(p), **({'subfontIndex': idx} if file.endswith('.ttc') else {})))
            have.add(name)
    if not {'CN', 'TR', 'TI', 'TBI'} <= have: raise SystemExit('缺少宋体或 Times New Roman 字体，请用 --font-dir 指定')
    chain = {'upright': ['TR', 'SX', 'AU'], 'italic': ['TI', 'SXI'], 'bold-italic': ['TBI'], 'bold': ['TB', 'TR']}
    return {k: [f for f in v if f in have] for k, v in chain.items()}

def font_for(chain, style, ch):
    for f in chain[style]:
        if ord(ch) in pdfmetrics.getFont(f).face.charToGlyph: return f
    return None

def check(d, src):
    errs = []
    for k in ('title', 'review_type', 'date', 'scope', 'basis', 'scope_pages', 'coverage', 'source_pdf', 'items'):
        if k not in d: errs.append(f'缺少字段 {k}')
    if errs: return errs
    sp = Path(d['source_pdf']); sp = sp if sp.is_absolute() else src.parent / sp
    if not sp.is_file(): return [f'找不到 source_pdf：{d["source_pdf"]}']
    d['_pdf'] = sp
    cov = d['coverage']
    if type(cov.get('complete')) is not bool: errs.append('coverage.complete 必须是 true/false')
    if cov.get('complete') and cov.get('gaps'): errs.append('coverage.complete 为 true 时 gaps 必须为空')
    if not cov.get('complete') and not cov.get('gaps'): errs.append('覆盖未完成时必须在 coverage.gaps 写明缺口')
    for k in KINDS:
        if k not in cov.get('reviewed', {}): errs.append(f'coverage.reviewed 缺少“{k}”的已审范围说明')
    scope = set(d['scope_pages'])
    ids, last = set(), {}
    for e in d['items']:
        tag = e.get('id', '?')
        for k in ('id', 'section', 'kind', 'page', 'reading_order', 'category', 'count', 'locate', 'boxes', 'reason'):
            if k not in e or e[k] in ('', None): errs.append(f'{tag}：缺少 {k}')
        if errs and errs[-1].startswith(tag): continue
        if 'image' in e: errs.append(f'{tag}：不再接受自备截图 image，截图和红框由导出器按 boxes 生成')
        if not isinstance(e['boxes'], list) or not e['boxes'] or not all(isinstance(b, dict) for b in e['boxes']):
            errs.append(f'{tag}：boxes 必须是非空列表，每项为 {{find…}} 或 {{rect…}}；红框为必做项')
        if tag in ids: errs.append(f'{tag}：编号重复')
        ids.add(tag)
        if e['section'] not in SECTIONS: errs.append(f'{tag}：section 只能是 {"、".join(SECTIONS)}')
        if e['kind'] not in KINDS: errs.append(f'{tag}：kind 只能是 字形 或 数学')
        if e['page'] not in scope: errs.append(f'{tag}：第 {e["page"]} 页不在 scope_pages 内')
        key = (e['section'], e['page'])
        if e['reading_order'] <= last.get(key, 0): errs.append(f'{tag}：同一区同一页内 reading_order 必须递增')
        last[key] = e['reading_order']
        if type(e['count']) is not int or e['count'] < 1: errs.append(f'{tag}：count 必须是正整数（错误位置数）')
        n = sum(1 for k in ('correct_runs', 'correct_image', 'correct_text') if e.get(k))
        if n != 1: errs.append(f'{tag}：correct_runs、correct_image、correct_text 三者必须且只能给一个')
        for r in e.get('correct_runs') or []:
            if r.get('style') not in STYLES or r.get('position', 'normal') not in SUP: errs.append(f'{tag}：correct_runs 的 style/position 不支持')
        for k in ('correct_image',):
            if e.get(k):
                p = Path(e[k]); p = p if p.is_absolute() else src.parent / p
                if not p.is_file(): errs.append(f'{tag}：找不到图片 {e[k]}')
                e['_' + k] = p
    return errs

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('records', type=Path); ap.add_argument('output', type=Path)
    ap.add_argument('--formats', choices=['pdf', 'md', 'both'], default='pdf')
    ap.add_argument('--font-dir', type=Path, default=Path('/System/Library/Fonts/Supplemental') if sys.platform == 'darwin' else Path('C:/Windows/Fonts'))
    a = ap.parse_args()
    d = json.loads(a.records.read_text('utf8'))
    errs = check(d, a.records)
    if not errs:
        doc = pymupdf.open(d['_pdf'])
        errs = resolve_boxes(d, doc)
    if errs:
        print('记录有问题，没有导出：'); [print(' -', x) for x in errs]; sys.exit(1)
    a.output.mkdir(parents=True, exist_ok=True)
    if any((a.output / x).exists() for x in ('审校清单.pdf', '审校清单.md')): sys.exit('输出目录里已有清单，请换一个新目录')
    for e in d['items']: render_evidence(e, doc, a.output / '证据')
    items = sorted(d['items'], key=lambda e: (list(SECTIONS).index(e['section']), e['page'], e['reading_order']))
    conf = [e for e in items if e['section'] == '确认']
    cnt = Counter(e['kind'] for e in conf)
    cov = d['coverage']
    summary = [
        f'审查范围：{d["scope"]}。',
        f'执行技能：formula-proofread。字形已审：{cov["reviewed"]["字形"]}；数学已审：{cov["reviewed"]["数学"]}。',
        f'确认问题 {len(conf)} 条（字形 {cnt["字形"]} 条、数学 {cnt["数学"]} 条），涉及 {sum(e["count"] for e in conf)} 个位置；'
        f'待专业复核 {sum(1 for e in items if e["section"] == "待复核")} 条。条数按问题计，同一截图内多个目标按说明计数。',
        ('覆盖状态：已完成指定范围。' if cov['complete'] else '覆盖状态：<b>未完成</b>，不能据此声称全稿无遗漏。缺口：' + '；'.join(cov['gaps']) + '。'),
        f'判据：{d["basis"]}',
        '阅读方法：先按原稿页码定位，再核对截图和文字说明。“正确格式”只列本条修改的目标，原有的其余上下标、撇号、帽号和内容保留；截图中其余文字不属于本条修改范围。',
    ] + d.get('notes', []) + ['原稿 PDF 和 Word 均未修改。']

    def page_label(e):
        return f'原稿第{e["page"]}页' + (f'（书内第{e["printed_page"]}页）' if e.get('printed_page') else '')

    if a.formats in ('md', 'both'):
        assets = a.output / 'assets'; assets.mkdir(exist_ok=True)
        md = [f'# {d["title"]}', ''] + [re.sub(r'</?b>', '**', s) for s in summary]
        for sec, name in SECTIONS.items():
            rows = [e for e in items if e['section'] == sec]
            if not rows: continue
            md += ['', f'## {name}', '', '| 原稿页 | 编号·类别·处数 | 原文定位与截图 | 正确格式 | 理由 |', '|---|---|---|---|---|']
            for e in rows:
                img = f'assets/{e["id"]}.png'; PILImage.open(e['_image']).convert('RGB').save(a.output / img)
                if e.get('correct_runs'):
                    cf = ''.join(f'<{SUP[r.get("position","normal")]}>{r["text"]}</{SUP[r.get("position","normal")]}>' if r.get('position', 'normal') != 'normal' else r['text'] for r in e['correct_runs'])
                elif e.get('correct_image'):
                    ci = f'assets/{e["id"]}-correct.png'; PILImage.open(e['_correct_image']).convert('RGB').save(a.output / ci); cf = f'![正确格式]({ci})'
                else: cf = e['correct_text']
                cell = lambda s: str(s).replace('|', '｜').replace('\n', '<br>')
                md.append(f'| {cell(page_label(e))} | {e["id"]} {cell(e["category"])}（{e["kind"]}）〔{e["count"]}处〕 | {cell(e["locate"] + (f'；其余位置：{e["elsewhere"]}' if e.get('elsewhere') else ''))}<br>![原文]({img}) | {cell(cf)} | {cell(e["reason"])} |')
        if not cov['complete']:
            md += ['', '## 未核查与覆盖缺口', ''] + [f'- {g}' for g in cov['gaps']]
        (a.output / '审校清单.md').write_text('\n'.join(md) + '\n', 'utf8')

    if a.formats in ('pdf', 'both'):
        chain = fonts(a.font_dir)
        grey = HexColor('#555555')
        body = ParagraphStyle('body', fontName='CN', fontSize=10.5, leading=17, wordWrap='CJK', spaceAfter=6)
        title = ParagraphStyle('title', parent=body, fontSize=17, leading=26, spaceAfter=14)
        sect = ParagraphStyle('sect', parent=body, fontSize=15, leading=22, spaceBefore=4, spaceAfter=10)
        pageh = ParagraphStyle('pageh', parent=body, fontSize=16, leading=24, spaceAfter=10)
        head = ParagraphStyle('head', parent=body, fontSize=12, leading=18, spaceAfter=3)
        small = ParagraphStyle('small', parent=body, fontSize=8.5, leading=13, textColor=grey, spaceAfter=4)
        corr = ParagraphStyle('corr', parent=body, fontSize=12, leading=22, spaceAfter=3)

        def runs_markup(e):
            out = []
            for r in e['correct_runs']:
                pos = r.get('position', 'normal')
                for ch in r['text']:
                    f = 'CN' if '\u4e00' <= ch <= '\u9fff' else font_for(chain, r['style'], ch)
                    if not f: sys.exit(f'{e["id"]}：正确格式里的“{ch}”没有可用字体，请改用 correct_image')
                    t = f'<font name="{f}" size="{9 if pos != "normal" else 13}">{escape(ch)}</font>'
                    out.append(f'<{SUP[pos]}>{t}</{SUP[pos]}>' if pos != 'normal' else t)
            return ''.join(out)

        def img(p, maxh):
            w, h = PILImage.open(p).size
            s = min((A4[0] - 32 * mm) / w, maxh / h, 1.0 * 72 / 110)   # 截图按约 110dpi 显示，不放大糊图
            return Image(str(p), width=w * s, height=h * s, hAlign='LEFT')

        def footer(c, doc):
            c.setFont('CN', 8); c.setFillColor(grey)
            c.drawString(16 * mm, 10 * mm, f'{d["review_type"]} · {d["date"]}')
            c.drawRightString(A4[0] - 16 * mm, 10 * mm, f'报告第{doc.page}页')

        story = [Paragraph(escape(d['title']), title)] + [Paragraph(s if '<b>' in s else escape(s), body) for s in summary]
        for sec, name in SECTIONS.items():
            rows = [e for e in items if e['section'] == sec]
            if not rows: continue
            story += [PageBreak(), Paragraph(name, sect)]
            last = None
            for e in rows:
                if last is not None and last != e['page']: story.append(PageBreak())
                if last != e['page']: story.append(Paragraph(escape(page_label(e)), pageh))
                last = e['page']
                label = '正确格式：' if sec == '确认' else ('条件化建议：' if sec == '待复核' else '纠偏说明：')
                block = [Paragraph(f'{escape(e["id"])} {escape(e["category"])}（{e["kind"]}） 〔{e["count"]}处〕', head),
                         Paragraph('原文定位：' + escape(e['locate']) + (f'；其余位置：{escape(e["elsewhere"])}' if e.get('elsewhere') else ''), small),
                         img(e['_image'], 70 * mm), Spacer(1, 4)]
                if e.get('correct_runs'): block.append(Paragraph(f'<b>{label}</b>' + runs_markup(e), corr))
                elif e.get('correct_image'): block += [Paragraph(f'<b>{label}</b>', corr), img(e['_correct_image'], 45 * mm)]
                else: block.append(Paragraph(f'<b>{label}</b>' + escape(e['correct_text']), body))
                block += [Paragraph(escape(e['reason']), small), Spacer(1, 12)]
                story.append(KeepTogether(block))
        if not cov['complete']:
            story += [PageBreak(), Paragraph('未核查与覆盖缺口', sect)] + [Paragraph('· ' + escape(g), body) for g in cov['gaps']]
        SimpleDocTemplate(str(a.output / '审校清单.pdf'), pagesize=A4, leftMargin=16 * mm, rightMargin=16 * mm,
                          topMargin=16 * mm, bottomMargin=18 * mm, title=d['title']).build(story, onFirstPage=footer, onLaterPages=footer)
    print(f'已导出：确认 {len(conf)} 条、待复核 {sum(1 for e in items if e["section"] == "待复核")} 条，'
          f'红框 {sum(len(e["_rects"]) for e in items)} 个 → {a.output}。截图在“证据/”，无框原图在“证据/无框/”。请渲染 PDF 逐页目视核对。')

if __name__ == '__main__':
    main()

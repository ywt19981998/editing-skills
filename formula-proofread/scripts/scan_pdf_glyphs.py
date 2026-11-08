"""从排版 PDF 的字体信息里找正斜体候选，不依赖 Word 公式能否读出（MathType、图片转曲以外的公式都适用）。
只出候选，不判对错：每条都要结合语义和原页确认。

用法: scan_pdf_glyphs.py 书稿.pdf 输出目录 [--pages 1-50,60] [--crops]
输出: 输出目录/glyph-candidates.json（候选）、glyph-report.md（汇总与清单）、formula-pages.json（含公式字体的页，作覆盖台账的待审页清单）
      --crops 时为每条候选裁一张带框的局部图（输出目录/crops/）
查: ①微分号 d 为斜体（d 后紧跟变量）；②自然常数 e 为斜体且后接上标；③函数名（sin、ln、exp、lim、max 等）为斜体；
    ④数字为斜体；⑤π 为斜体。另统计 d 的正体、斜体数量，便于看全书是否混用。
局限：只认字体名带 Italic/Oblique 或字体标志为斜体的字；用剪切矩阵合成的斜体、转曲轮廓字、图片公式查不到，仍须逐页看图。
"""
import argparse, json, re
from collections import Counter, defaultdict
from pathlib import Path
import pymupdf
from symbol_inventory import decode   # 还原 Word 公式的 Unicode 数学字母（𝑑、𝑒）和 MathType 私用区编码

FUNCS = ('arcsin arccos arctan arccot sinh cosh tanh coth sin cos tan cot sec csc '
         'ln lg log exp lim max min sup inf det arg sgn grad div rot Re Im').split()
MATH_FONT = re.compile(r'Symbol|MT ?Extra|Euclid|Cambria ?Math|STIX|Math|Italic', re.I)
LETTER = re.compile(r'[A-Za-zα-ωΑ-Ωϑϕϵ]')

def italic(ch):
    if ch.get('msty'): return '斜' in ch['msty']
    return bool(re.search(r'Italic|Oblique|-It\b|Ital', ch['font'])) or bool(ch['flags'] & 2)

def parse_pages(spec, n):
    if not spec: return list(range(1, n + 1))
    out = []
    for part in spec.split(','):
        a, _, b = part.partition('-')
        out += list(range(int(a), int(b or a) + 1))
    return [p for p in out if 1 <= p <= n]

def lines_of(page):
    """每行返回逐字列表：c（字）、font、flags、size、bbox、origin"""
    raw = page.get_text('rawdict')
    for b in raw['blocks']:
        for ln in b.get('lines', []):
            chars = []
            for sp in ln['spans']:
                for c in sp['chars']:
                    ch, _, msty = decode(c['c'], sp['font'])
                    chars.append(dict(c=ch, msty=msty, font=sp['font'], flags=sp['flags'], size=sp['size'],
                                      bbox=list(c['bbox']), origin=list(c['origin'])))
            if chars: yield chars

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pdf', type=Path); ap.add_argument('out', type=Path)
    ap.add_argument('--pages'); ap.add_argument('--crops', action='store_true')
    a = ap.parse_args()
    doc = pymupdf.open(a.pdf)
    pages = parse_pages(a.pages, doc.page_count)
    a.out.mkdir(parents=True, exist_ok=True)
    cands, stats, fpages = [], Counter(), {}

    def add(pno, kind, chars, i, j, line, note):
        x0 = min(chars[k]['bbox'][0] for k in range(i, j)); y0 = min(chars[k]['bbox'][1] for k in range(i, j))
        x1 = max(chars[k]['bbox'][2] for k in range(i, j)); y1 = max(chars[k]['bbox'][3] for k in range(i, j))
        lo, hi = max(0, i - 12), min(len(chars), j + 12)
        cands.append(dict(page=pno, kind=kind, text=''.join(chars[k]['c'] for k in range(i, j)),
                          context=''.join(c['c'] for c in chars[lo:hi]), font=chars[i]['font'],
                          bbox=[round(v, 1) for v in (x0, y0, x1, y1)], note=note))

    for pno in pages:
        page = doc[pno - 1]
        nmath = 0
        page_lines = list(lines_of(page))
        allc = [c for ln in page_lines for c in ln]
        # ② 自然常数 e、⑥ 虚数单位 i：上标在 PDF 里常被拆成单独一行，按位置在整页里找 e 右上方的小字
        for chars in page_lines:
            s = ''.join(c['c'] for c in chars)
            for i, ch in enumerate(chars):
                if ch['c'] != 'e' or not italic(ch): continue
                if (i and LETTER.match(chars[i - 1]['c'])) or (i + 1 < len(chars) and LETTER.match(chars[i + 1]['c'])): continue
                x1, base, sz = ch['bbox'][2], ch['origin'][1], ch['size']
                sup = [c for c in allc if c['size'] < sz * 0.85 and base - sz * 0.9 < c['origin'][1] < base - sz * 0.2
                       and x1 - 1 <= c['bbox'][0] <= x1 + sz * 3]
                if not sup: continue
                add(pno, '常数e斜体', chars, i, i + 1, s, '若 e 是自然对数的底（e 的幂），国标应为正体；若是变量另判')
                sup.sort(key=lambda c: c['bbox'][0])
                st = ''.join(c['c'] for c in sup)
                for k, c in enumerate(sup):
                    if c['c'] == 'i' and italic(c) and not (k and LETTER.match(sup[k - 1]['c']) and sup[k - 1]['c'] not in 'ωπ'):
                        lo, hi = max(0, k - 6), k + 6
                        cands.append(dict(page=pno, kind='虚数单位i斜体', text='i', context=st[lo:hi], font=c['font'],
                                          bbox=[round(v, 1) for v in c['bbox']], note='e 的指数里的 i 若是虚数单位，国标应为正体；若是求和指标另判'))
        for chars in page_lines:
            s = ''.join(c['c'] for c in chars)
            nmath += sum(1 for c in chars if MATH_FONT.search(c['font']) and not c['c'].isspace())
            for i, ch in enumerate(chars):
                c = ch['c']; prev = chars[i - 1]['c'] if i else ' '; nxt = chars[i + 1] if i + 1 < len(chars) else None
                # ① 微分号 d：前面不是字母（不在单词里），后面紧跟一个字母
                # “dxdy”“dxdx”里第二个 d 前面紧挨着 x，按“d+字母”成对的规律放行
                after_pair = i >= 2 and chars[i - 2]['c'] == 'd' and LETTER.match(prev)
                if c == 'd' and nxt and LETTER.match(nxt['c']) and (not LETTER.match(prev) or after_pair):
                    after = chars[i + 1:i + 5]
                    if not italic(ch) and re.match(r'[a-z]{3,}', ''.join(x['c'] for x in after)):  # 正文英文单词（do、data…）
                        continue
                    stats['d·斜体' if italic(ch) else 'd·正体'] += 1
                    if italic(ch):
                        add(pno, '微分号d斜体', chars, i, min(i + 2, len(chars)), s, '若 d 是微分算子（dx、dy/dx、d(…)），国标应为正体；若是变量 d（距离、直径等）则斜体正确')
                # ⑤ π 斜体
                if c in 'πΠ' and italic(ch):
                    add(pno, 'π斜体', chars, i, i + 1, s, '圆周率 π 应为正体')
                # ④ 数字斜体
                if c.isdigit() and italic(ch) and not (prev.isdigit() and italic(chars[i - 1])):
                    j = i
                    while j < len(chars) and (chars[j]['c'].isdigit() or chars[j]['c'] == '.') and italic(chars[j]): j += 1
                    add(pno, '数字斜体', chars, i, j, s, '数字一律正体（注意 WPS 导出可能换字体造成假象，回看原页）')
            # ③ 函数名斜体：整个函数名的字母都是斜体，且前后不是字母
            for f in FUNCS:
                for m in re.finditer(re.escape(f), s):
                    i, j = m.start(), m.end()
                    if (i and LETTER.match(s[i - 1])) or (j < len(s) and re.match(r'[A-Za-z]', s[j]) and f not in ('ln', 'lg')):
                        continue
                    if j < len(s) and f in ('ln', 'lg') and re.match(r'[A-Za-z]', s[j]): continue
                    if all(italic(chars[k]) for k in range(i, j)):
                        add(pno, '函数名斜体', chars, i, j, s, f'函数名 {f} 应为正体（GB/T 3505 的评定长度 ln 等专用符号除外，先查定义）')
        if nmath: fpages[pno] = nmath

    # 去重（同页同位置同类只留一条）
    seen, uniq = set(), []
    for c in cands:
        key = (c['page'], c['kind'], tuple(c['bbox']))
        if key not in seen: seen.add(key); uniq.append(c)
    for k, c in enumerate(uniq, 1): c['id'] = f'G{k:04d}'

    if a.crops and uniq:
        (a.out / 'crops').mkdir(exist_ok=True)
        for c in uniq:
            page = doc[c['page'] - 1]
            x0, y0, x1, y1 = c['bbox']
            clip = pymupdf.Rect(max(0, x0 - 60), max(0, y0 - 14), min(page.rect.x1, x1 + 60), min(page.rect.y1, y1 + 14))
            pix = page.get_pixmap(clip=clip, dpi=220)   # 只渲染局部图，在图上画框，不改原 PDF
            from PIL import Image, ImageDraw
            im = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
            sx, sy = pix.width / clip.width, pix.height / clip.height
            ImageDraw.Draw(im).rectangle([(x0 - clip.x0 - 1.5) * sx, (y0 - clip.y0 - 1.5) * sy, (x1 - clip.x0 + 1.5) * sx, (y1 - clip.y0 + 1.5) * sy], outline=(220, 0, 0), width=3)
            c['crop'] = f'crops/{c["id"]}-p{c["page"]}.png'; im.save(a.out / c['crop'])

    (a.out / 'glyph-candidates.json').write_text(json.dumps(uniq, ensure_ascii=False, indent=1), 'utf8')
    (a.out / 'formula-pages.json').write_text(json.dumps(dict(pages_scanned=pages, formula_pages=sorted(fpages),
        note='含数学字体字符的页；只作待审页清单，不含图片公式、转曲公式所在页，仍须逐页看图'), ensure_ascii=False, indent=1), 'utf8')
    by = Counter(c['kind'] for c in uniq)
    L = [f'# PDF 正斜体候选\n', f'文件：`{a.pdf.name}`　扫描 {len(pages)} 页，其中含数学字体的页 {len(fpages)} 页。\n',
         '以下全部是**候选**，须逐条结合定义和原页判定；查不到合成斜体、转曲字和图片公式，扫描结果不能代替逐页看图。\n',
         f'微分号 d 统计：斜体 {stats["d·斜体"]} 处，正体 {stats["d·正体"]} 处' + ('（**全书混用**）' if stats['d·斜体'] and stats['d·正体'] else '') + '。\n',
         '| 类别 | 候选数 |\n|---|---|'] + [f'| {k} | {v} |' for k, v in by.most_common()]
    for kind in by:
        L.append(f'\n## {kind}\n\n| # | PDF页 | 原文 | 上下文 | 字体 |\n|---|---|---|---|---|')
        for c in uniq:
            if c['kind'] == kind:
                L.append(f"| {c['id']} | {c['page']} | {c['text']} | {c['context'].replace('|', '｜')} | {c['font']} |")
    (a.out / 'glyph-report.md').write_text('\n'.join(L) + '\n', 'utf8')
    print(f'扫描 {len(pages)} 页，含数学字体 {len(fpages)} 页；候选 {len(uniq)} 条：' + '，'.join(f'{k} {v}' for k, v in by.most_common())
          + f'；d 斜体 {stats["d·斜体"]}、正体 {stats["d·正体"]} → {a.out}')

if __name__ == '__main__':
    main()

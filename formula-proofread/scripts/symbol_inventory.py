"""符号清单：逐页列出 PDF 里的数学字母，归组、比对，并按“符号表”自动判定。读的是 PDF 文字层，不用 OCR。

三步用法:
  1. 列清单  symbol_inventory.py 书稿.pdf 输出目录 [--pages 1-50,60]
     输出 symbol-inventory.json/.md（分组、重点候选、全部出现位置）、symbol-table.json（符号表模板，附定义句）、
          image-regions.json（图片格式的图：要看图）、symbol-decisions.json（逐组判定模板）
  2. 填符号表（模型做）：读每个符号的定义句，填 meaning（含义）、expected（应有字形）、scope（适用页，默认全书）；
     同一字母不同含义写成多条，用 scope 或 position 区分。
  3. 自动比对  symbol_inventory.py --apply 输出目录/symbol-inventory.json 符号表.json [判定.json]
     按符号表逐处比对字形，写出 symbol-decisions.json（正确/错误/部分错误，错误处逐条列出），
     没有符号表条目的组标“待判”，交模型看整行公式判定。
  核对  symbol_inventory.py --check 输出目录/symbol-inventory.json 判定.json   # 每一组都有结论才算字形审完

字形写法（expected 与清单一致，只比“·”前面的部分）：斜体、正体、粗斜体、粗正体、花体、哥特体、空心体。
不列入：正常位置上的正体多字母词（DDoS、HGNN 等名称缩写）、参考文献（从“参考文献”标题到下一章）。上下标一律列入。
"""
import argparse, json, re, statistics, sys, unicodedata
from collections import Counter, defaultdict
from pathlib import Path
import pymupdf

SYMBOL_MAP = dict(zip('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ',
                      'αβχδεφγηιϕκλμνοπθρστυϖωξψζΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ'))
FUNCS = set('argmin argmax arg min max sin cos tan cot sec csc arcsin arccos arctan sinh cosh tanh ln lg log exp lim sup inf det tr rank sgn diag mod gcd'.split())
DESC = set('min max str att ref opt avg mean tot total rated eff in out on off real pred true test train sys loc glob init fin new old'.split())
SANS = re.compile(r'Arial|Helvetica|Calibri|Verdana|Tahoma|Segoe|SimHei|Hei|YaHei|PingFang|DengXian|Sans', re.I)
SPECIAL = re.compile(r'Euclid|MT ?Extra|Fraktur|Script|BlackBoard|MSBM|MSAM|Wingding', re.I)
CJK = re.compile(r'[一-鿿]')
LETTER = re.compile(r'[A-Za-zÀ-ÖØ-öø-ÿα-ωΑ-Ωϑϕϖς]')   # 含带重音的拉丁字母，避免把 GÉANT 这类词切断
MATH_STYLE = [('BOLD ITALIC', '粗斜体'), ('BOLD', '粗正体'), ('ITALIC', '斜体'), ('SCRIPT', '花体'), ('FRAKTUR', '哥特体'),
              ('DOUBLE-STRUCK', '空心体'), ('SANS-SERIF', '无衬线'), ('MONOSPACE', '等宽')]
DEF_WORDS = r'(表示|代表|为|是|指|记为|记作|定义为|denote)'

def family(font):
    """字体族名：去掉子集前缀和 PS、-BoldItalic、MT 等后缀（字体名可能被截断，如 BoldItal）"""
    f = font.split('+')[-1]
    f = re.sub(r'(PS)?[-,]?(BoldItal\w*|Bold|Ital\w*|Oblique|Regular|It|Bd)?(MT)?$', '', f)
    return f or font

def decode(ch, font):
    """还原 Unicode 数学字母和符号字体私用区编码；返回 (字, 是否经过还原, 码位自带的字形或 None)"""
    o = ord(ch)
    if ch == 'ℎ': return 'h', True, '斜体'
    if 0x1D400 <= o <= 0x1D7FF:
        m = re.match(r'MATHEMATICAL (.+?) (CAPITAL|SMALL) (.+)$', unicodedata.name(ch, ''))
        if m:
            sty = next((v for k, v in MATH_STYLE if m.group(1).startswith(k)), m.group(1))
            base = m.group(3)
            try:
                letter = unicodedata.lookup(f'GREEK {m.group(2)} LETTER {base}') if len(base) > 1 else (base if m.group(2) == 'CAPITAL' else base.lower())
            except KeyError:
                letter = ch
            return letter, True, sty
        return ch, False, None
    if 0xF020 <= o <= 0xF0FF:
        b = chr(o - 0xF000)
        return (SYMBOL_MAP.get(b, b) if re.search('Symbol', font, re.I) else b), True, None
    if re.search('Symbol', font, re.I) and ch in SYMBOL_MAP:
        return SYMBOL_MAP[ch], True, None
    return ch, False, None

def style_of(c):
    if c.get('msty'): return f"{c['msty']}·{family(c['font'])}"          # Unicode 数学字母：字形由码位决定
    fam = family(c['font'])
    for pat, lab in ((r'EuclidMathOne|Euclid ?Math ?One|Script|RSFS', '花体'), (r'Fraktur', '哥特体'), (r'EuclidMathTwo|BlackBoard|MSBM', '空心体')):
        if re.search(pat, fam, re.I): return f'{lab}·{fam}（特殊字体）'    # MathType 特殊字体的字形由字体本身决定
    it = bool(re.search(r'Italic|Oblique|-It\b', c['font'])) or bool(c['flags'] & 2)
    bd = bool(re.search(r'Bold|Bd\b', c['font'])) or bool(c['flags'] & 16)
    kind = '特殊字体' if SPECIAL.search(fam) else ('无衬线' if SANS.search(fam) else '')
    return f"{'粗' if bd else ''}{'斜体' if it else '正体'}·{fam}" + (f'（{kind}）' if kind else '')

def shape(style):
    return style.split('·')[0]

def parse_pages(spec, n):
    if not spec: return list(range(1, n + 1))
    out = []
    for part in str(spec).replace('p', '').split(','):
        if part.strip() in ('', '全书'): return list(range(1, n + 1))
        a, _, b = part.partition('-'); out += range(int(a), int(b or a) + 1)
    return [p for p in out if 1 <= p <= n]

# ---------------------------------------------------------------- 读页面
def page_lines(page):
    """每行：字列表＋外框、最大字号、基线。上下标常被 PDF 拆成单独一行，后面按位置挂回所属的行"""
    out = []
    for b in page.get_text('rawdict')['blocks']:
        for ln in b.get('lines', []):
            chars = []
            for sp in ln['spans']:
                for c in sp['chars']:
                    ch, dec, msty = decode(c['c'], sp['font'])
                    chars.append(dict(c=ch, dec=dec, msty=msty, font=sp['font'], flags=sp['flags'], size=sp['size'],
                                      oy=c['origin'][1], bbox=c['bbox']))
            if not chars: continue
            main = max(c['size'] for c in chars)
            base = statistics.median([c['oy'] for c in chars if c['size'] >= main * 0.9])
            x0 = min(c['bbox'][0] for c in chars); x1 = max(c['bbox'][2] for c in chars)
            y0 = min(c['bbox'][1] for c in chars); y1 = max(c['bbox'][3] for c in chars)
            out.append(dict(chars=chars, main=main, base=base, bbox=(x0, y0, x1, y1), text=''.join(c['c'] for c in chars)))
    for L in out:
        L['parent'] = None
        for P in out:   # 找比它大一号、紧挨在左边或上下的行：本行就是它的上标或下标
            if P is L or P['main'] < L['main'] * 1.2: continue
            near_x = P['bbox'][0] - 2 <= L['bbox'][0] <= P['bbox'][2] + P['main'] * 1.5
            near_y = P['bbox'][1] - L['main'] <= L['base'] <= P['bbox'][3] + L['main'] * 0.8
            if near_x and near_y and (L['parent'] is None or abs(P['base'] - L['base']) < abs(L['parent']['base'] - L['base'])):
                L['parent'] = P
    for L in out:
        P = L['parent']
        ref_main, ref_base = (P['main'], P['base']) if P else (L['main'], L['base'])
        L['ctx'] = '正文' if CJK.search(L['text'] + (P['text'] if P else '')) else '公式'
        for c in L['chars']:
            c['style'] = style_of(c)
            if P: c['pos'] = '上标' if c['oy'] < ref_base - ref_main * 0.15 else '下标'
            else: c['pos'] = '正常' if c['size'] >= ref_main * 0.8 else ('上标' if c['oy'] < ref_base - ref_main * 0.15 else '下标')
    return out

def scan(pdf, pages):
    doc = pymupdf.open(pdf)
    groups, skipped, images, defs, excluded = {}, Counter(), [], defaultdict(list), []
    in_refs = False
    for pno in range(1, doc.page_count + 1):
        page = doc[pno - 1]
        lines = page_lines(page)
        lines.sort(key=lambda L: (round(L['bbox'][1]), L['bbox'][0]))
        for L in lines:
            t = L['text'].strip()
            if re.fullmatch(r'参\s*考\s*文\s*献', t): in_refs = True; L['skip'] = True; continue
            if in_refs and re.match(r'第\s*\d+\s*章', t): in_refs = False
            L['skip'] = in_refs
        if pno not in pages: continue
        if any(L['skip'] for L in lines): excluded.append(pno)
        for info in page.get_image_info():
            x0, y0, x1, y1 = info['bbox']
            if (x1 - x0) * (y1 - y0) > 1500: images.append(dict(page=pno, bbox=[round(v, 1) for v in info['bbox']]))
        for L in lines:
            if L['skip']: continue
            chars, text, ctx = L['chars'], L['text'], L['ctx']
            i = 0
            while i < len(chars):
                if not LETTER.match(chars[i]['c']): i += 1; continue
                j = i + 1
                while j < len(chars) and LETTER.match(chars[j]['c']) and chars[j]['style'] == chars[i]['style'] and chars[j]['pos'] == chars[i]['pos']:
                    j += 1
                tok = ''.join(c['c'] for c in chars[i:j]); first = chars[i]
                plain_word = (len(tok) > 1 and tok not in FUNCS and first['pos'] == '正常' and shape(first['style']) in ('正体', '粗正体')   # 含算法框里的粗体关键词
                              and not first['dec'] and '特殊字体' not in first['style'])
                if len(tok) == 1 or tok in FUNCS or first['pos'] != '正常':
                    pieces = [(i, j, tok)]                         # 上下标整体一组（min、str、ij），不拆
                elif plain_word:
                    skipped[tok] += 1; pieces = []                 # 正常位置的普通正体词、缩写
                else:
                    pieces = [(k, k + 1, chars[k]['c']) for k in range(i, j)]   # 公式里连写的变量逐个拆开
                for a, z, tk in pieces:
                    c0 = chars[a]; cx = ctx
                    if ctx == '正文' and len(tk) == 1 and re.search(r'[（(]\s*$', text[:a]) and re.match(r'\s*[)）]', text[z:]):
                        cx = '正文·括注'                           # (a)、（b）一类标号
                    key = (tk, c0['style'], cx, c0['pos'])
                    g = groups.setdefault(key, dict(token=tk, style=c0['style'], context=cx, position=c0['pos'],
                                                    count=0, pages=[], examples=[], occurrences=[]))
                    g['count'] += 1
                    if pno not in g['pages']: g['pages'].append(pno)
                    host = L['parent']['text'] if L['parent'] else text      # 上下标的例句用所属那一行
                    around = (host[:60] + ' ‖ ' + text) if L['parent'] else text[max(0, a - 15): z + 15]
                    bbox = [round(v, 1) for v in (chars[a]['bbox'][0], chars[a]['bbox'][1], chars[z - 1]['bbox'][2], chars[z - 1]['bbox'][3])]
                    g['occurrences'].append(dict(page=pno, bbox=bbox, around=around))
                    if len(g['examples']) < 3 and all(e['page'] != pno for e in g['examples']):
                        g['examples'].append(dict(page=pno, around=around, bbox=bbox))
                    # 定义句：正文里“X 表示……”“其中 X 为……”
                    if ctx == '正文' and c0['pos'] == '正常' and len(defs[tk]) < 6:
                        after = text[z:z + 14]
                        if re.match(r'\s*[)）]?\s*' + DEF_WORDS, after) or re.search(r'(其中|式中|令|设|记)[，,：:]?\s*$', text[max(0, a - 6):a]):
                            s = text[max(0, a - 20): z + 40]
                            if all(s != d['sentence'] for d in defs[tk]): defs[tk].append(dict(page=pno, sentence=s, style=c0['style']))
                i = j
    return groups, skipped, images, defs, excluded

def flag(groups):
    by = defaultdict(lambda: {'正文': Counter(), '公式': Counter()})
    for g in groups.values():
        if len(g['token']) == 1 and g['position'] == '正常' and g['context'] in ('正文', '公式'):
            by[g['token']][g['context']][shape(g['style'])] += g['count']
    for g in groups.values():
        g['flags'] = []; t = g['token']; sh = shape(g['style'])
        if len(t) == 1 and g['position'] == '正常' and g['context'] in ('正文', '公式'):
            f = by[t]['公式']; tot = sum(f.values())
            if g['context'] == '正文' and f and sh not in f:
                m = f.most_common(1)[0][0]
                g['flags'].append(f'正文与公式字形不一致：公式中多为{m}（{f[m]}处），此处为{sh}')
            if g['context'] == '公式' and tot >= 5:
                m, n = f.most_common(1)[0]
                if sh != m and n / tot >= 0.6 and g['count'] / tot <= 0.4:
                    g['flags'].append(f'公式中同一字母字形不一致：多为{m}（{n}/{tot}处），此处为{sh}')
            if g['context'] == '正文' and sh in ('正体', '粗正体') and not re.match(r'[α-ωΑ-Ω]', t):
                g['flags'].append('正文中的单个字母为正体：若是变量应为斜体，若是矢量/矩阵/张量应为粗斜体；单位、标签、缩写可为正体')
        if g['position'] != '正常' and (t in DESC or len(t) >= 3) and '斜' in sh:
            g['flags'].append('说明性上下标（min、max、str、ref 等词）应为正体；若是变量连写（如 ij）则斜体正确')
        if re.fullmatch(r'[ΑΒΕΖΗΙΚΜΝΟΡΤΥΧ]', t):
            g['flags'].append('希腊大写字母，与拉丁字母同形（如 Ζ 与 Z）：若本意是拉丁字母，是录入错误，检索、排版都会出问题')
        if '特殊字体' in g['style']:
            g['flags'].append('MathType 特殊字体（花体/哥特体/空心体），部分阅读器缺字体会显示成别的样子，须在正确渲染的版面上确认，再按本书体例判断')
    return {t: {k: dict(v) for k, v in d.items()} for t, d in sorted(by.items())}

# ---------------------------------------------------------------- 符号表比对
def apply_table(inv, table, npages):
    """按符号表逐处比对：符号＋（可选）位置＋适用页 → 应有字形"""
    rules = defaultdict(list)
    for r in table:
        if not r.get('expected'): continue
        rules[r['symbol']].append(dict(r, _pages=set(parse_pages(r.get('scope') or '全书', npages))))
    out, pending = {}, []
    for g in inv['groups']:
        cand = [r for r in rules.get(g['token'], []) if r.get('position') in (None, '', g['position'])]
        if not cand:
            pending.append(g['id']); out[g['id']] = dict(decision='待判', reason='符号表里没有对应条目，交模型看整行公式判定'); continue
        bad, unknown, meaning = [], [], set()
        for o in g['occurrences']:
            r = next((r for r in cand if o['page'] in r['_pages']), None)
            if r is None: unknown.append(o); continue
            meaning.add(r.get('meaning', ''))
            if shape(g['style']) != r['expected']: bad.append(dict(o, expected=r['expected'], meaning=r.get('meaning', '')))
        if unknown:
            pending.append(g['id'])
            out[g['id']] = dict(decision='待判', reason=f'{len(unknown)} 处不在符号表任何条目的适用页内（如 p{unknown[0]["page"]}）'); continue
        sem = '；'.join(sorted(meaning))
        if not bad: out[g['id']] = dict(semantic=sem, decision='正确', reason=f'符合符号表：{sem}', auto=True)
        else:
            dec = '错误' if len(bad) == len(g['occurrences']) else '部分错误'
            out[g['id']] = dict(semantic=sem, decision=dec, auto=True, errors=bad,
                                reason=f'应为{bad[0]["expected"]}（{bad[0]["meaning"]}），此处为{shape(g["style"])}，共 {len(bad)} 处：' +
                                       '、'.join(f'p{b["page"]}' for b in bad[:20]) + ('……' if len(bad) > 20 else ''))
    return out, pending

# ---------------------------------------------------------------- 主程序
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('pdf', nargs='?'); ap.add_argument('out', nargs='?'); ap.add_argument('--pages')
    ap.add_argument('--check', nargs=2, metavar=('清单.json', '判定.json'))
    ap.add_argument('--apply', nargs='+', metavar='文件', help='清单.json 符号表.json [输出判定.json]')
    a = ap.parse_args()
    if a.check:
        inv = json.loads(Path(a.check[0]).read_text('utf8')); dec = json.loads(Path(a.check[1]).read_text('utf8'))
        miss = [g['id'] for g in inv['groups'] if dec.get(g['id'], {}).get('decision') not in ('正确', '错误', '部分错误', '待复核')]
        noreason = [k for k, v in dec.items() if v.get('decision') in ('错误', '部分错误', '待复核') and not v.get('reason')]
        print(f'共 {len(inv["groups"])} 组，已判定 {len(inv["groups"]) - len(miss)} 组；未判定 {len(miss)} 组' +
              (f'：{"、".join(miss[:40])}' + ('……' if len(miss) > 40 else '') if miss else ''))
        if noreason: print('判为错误/部分错误/待复核但没写理由：' + '、'.join(noreason))
        sys.exit(1 if miss or noreason else 0)
    if a.apply:
        invp = Path(a.apply[0]); inv = json.loads(invp.read_text('utf8'))
        table = json.loads(Path(a.apply[1]).read_text('utf8'))
        outp = Path(a.apply[2]) if len(a.apply) > 2 else invp.parent / 'symbol-decisions.json'
        old = json.loads(outp.read_text('utf8')) if outp.exists() else {}
        res, pending = apply_table(inv, table, max(inv['pages_scanned']))
        for k, v in old.items():            # 模型已手工判定的组保留，不被自动结果覆盖
            if v.get('decision') in ('正确', '错误', '部分错误', '待复核') and not v.get('auto'): res[k] = v; pending = [p for p in pending if p != k]
        outp.write_text(json.dumps(res, ensure_ascii=False, indent=1), 'utf8')
        c = Counter(v['decision'] for v in res.values())
        print(f'自动比对：' + '，'.join(f'{k} {v} 组' for k, v in c.most_common()) + f' → {outp}')
        if pending: print('待判组（交模型看整行公式）：' + '、'.join(pending[:60]) + ('……' if len(pending) > 60 else ''))
        return
    doc = pymupdf.open(a.pdf); pages = parse_pages(a.pages, doc.page_count)
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    groups, skipped, images, defs, excluded = scan(a.pdf, pages)
    letters = flag(groups)
    order = sorted(groups.values(), key=lambda g: (not g['flags'], g['token'].lower(), g['token'], g['context'], g['position'], -g['count']))
    for k, g in enumerate(order, 1): g['id'] = f'S{k:04d}'
    inv = dict(pdf=str(a.pdf), pages_scanned=pages, excluded_reference_pages=excluded, groups=order, letters=letters,
               skipped_words=dict(skipped.most_common(80)), skipped_total=sum(skipped.values()))
    (out / 'symbol-inventory.json').write_text(json.dumps(inv, ensure_ascii=False, indent=1), 'utf8')
    (out / 'image-regions.json').write_text(json.dumps(images, ensure_ascii=False, indent=1), 'utf8')
    (out / 'symbol-decisions.json').write_text(json.dumps({g['id']: dict(token=g['token'], style=g['style'], context=g['context'],
        position=g['position'], semantic='', decision='', reason='') for g in order}, ensure_ascii=False, indent=1), 'utf8')
    # 符号表模板：每个出现过的符号一条，附定义句和各字形的处数
    usage = defaultdict(Counter)
    for g in order: usage[(g['token'], g['position'] if len(g['token']) > 1 else '')][shape(g['style'])] += g['count']
    table = []
    for (tk, pos), cnt in sorted(usage.items(), key=lambda x: (x[0][0].lower(), x[0][0], x[0][1])):
        table.append(dict(symbol=tk, position=pos or None, meaning='', expected='', scope='全书',
                          styles_seen=dict(cnt), definitions=defs.get(tk, [])[:6]))
    (out / 'symbol-table.json').write_text(json.dumps(table, ensure_ascii=False, indent=1), 'utf8')
    flagged = [g for g in order if g['flags']]
    total = sum(g['count'] for g in order)
    pg = lambda ps: '、'.join(map(str, ps[:12])) + (f'……共{len(ps)}页' if len(ps) > 12 else '')
    L = [f'# 符号清单：{Path(a.pdf).name}\n',
         f'扫描 {len(pages)} 页，数学字母 {total} 处，归为 **{len(order)} 组**；重点候选 **{len(flagged)} 组**；符号表 {len(table)} 个符号（其中 {sum(1 for t in table if t["definitions"])} 个找到定义句）。\n',
         f'未列入：正常位置的正体多字母词 {sum(skipped.values())} 处（名称、缩写，常见的见 JSON 的 skipped_words）；参考文献所在页 {pg(excluded) if excluded else "无"}。\n',
         f'图片格式的图 {len(images)} 处（{pg(sorted(set(i["page"] for i in images)))} 页），文字层里没有，须看图，见 image-regions.json。\n' if images else '本范围没有图片格式的图。\n',
         '## 一、重点候选\n', '| 编号 | 符号 | 字形 | 位置 | 处数 | 页 | 提示 | 例句 |', '|---|---|---|---|---|---|---|---|']
    for g in flagged:
        ex = g['examples'][0]
        L.append(f"| {g['id']} | {g['token']} | {g['style']} | {g['context']}·{g['position']} | {g['count']} | {pg(g['pages'])} | {'；'.join(g['flags'])} | p{ex['page']}：{ex['around'].replace('|', '｜')} |")
    L += ['\n## 二、全部分组\n', '| 编号 | 符号 | 字形 | 位置 | 处数 | 页 | 例句 |', '|---|---|---|---|---|---|---|']
    for g in order:
        ex = g['examples'][0]
        L.append(f"| {g['id']} | {g['token']} | {g['style']} | {g['context']}·{g['position']} | {g['count']} | {pg(g['pages'])} | p{ex['page']}：{ex['around'].replace('|', '｜')} |")
    (out / 'symbol-inventory.md').write_text('\n'.join(L) + '\n', 'utf8')
    print(f'扫描 {len(pages)} 页：数学字母 {total} 处，{len(order)} 组，重点候选 {len(flagged)} 组；符号表 {len(table)} 个；'
          f'图片 {len(images)} 处；参考文献页 {len(excluded)} 页 → {out}')

if __name__ == '__main__':
    main()

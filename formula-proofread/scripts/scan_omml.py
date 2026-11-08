"""扫描 docx 中的 Word 公式（OMML）属性，全书统计字形与写法问题，不改原稿。

用法：~/.venvs/book-review/bin/python scan_omml.py 书稿.docx [输出.json] [--all]（--all 输出每一处位置，默认每类最多 8 例）

只读 docx 内部 XML，不需要安装 Word。结果是候选清单：
- 函数名/运算符（sin、ln、lim 等）未设正体：确定性高（OMML 默认斜体，只有函数结构 m:func 或正体属性才显示为正体）；
- π、e、d、i 的正斜体统计：给出正体/斜体数量，供判断是否全书统一，需结合语义；
- 撇号、增量符号、不等号、省略号、乘号的编码混用；
- MathType（OLE）公式与 Word 公式混用的位置。
MathType 和图片公式读不到内容，须看排版 PDF。
"""
import sys, json, re, zipfile, collections
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
def q(ns, t): return f'{{{ns}}}{t}'
def ln(el): return etree.QName(el).localname

FUNCS = ['arcsin', 'arccos', 'arctan', 'arccot', 'sinh', 'cosh', 'tanh', 'coth', 'sin', 'cos', 'tan', 'cot',
         'sec', 'csc', 'ln', 'lg', 'log', 'lim', 'max', 'min', 'sup', 'inf', 'exp', 'det', 'grad', 'div', 'rot', 'sgn', 'arg']
FUNC_RE = re.compile('(' + '|'.join(sorted(FUNCS, key=len, reverse=True)) + ')')
CHAR_GROUPS = {
    '撇号': {'′': '′（U+2032，规范）', "'": "'（英文单引号，排版时可能显示为撇号，需看PDF）", '’': '’（中文右单引号）', 'ˊ': 'ˊ（声调符）'},
    '增量符号': {'Δ': 'Δ（希腊字母，规范）', '∆': '∆（U+2206 增量符号）', '△': '△（三角形，误用）'},
    '小于等于': {'≤': '≤', '⩽': '⩽'},
    '大于等于': {'≥': '≥', '⩾': '⩾'},
    '省略号': {'⋯': '⋯（居中，规范）', '…': '…（底线）', '...': '...（三个句点）'},
    '乘号': {'×': '×', '*': '*（星号，除卷积外不应作乘号）', '·': '·（中点）', '⋅': '⋅（点乘）'},
}

def upright(r):
    """该 m:r 是否显示为正体：函数名结构内、m:sty=p、m:nor（按正文样式）"""
    if any(ln(a) == 'fName' for a in r.iterancestors()): return True
    rpr = r.find(q(M, 'rPr'))
    if rpr is not None:
        sty = rpr.find(q(M, 'sty'))
        if sty is not None and sty.get(q(M, 'val')) in ('p', 'b'): return True
        if rpr.find(q(M, 'nor')) is not None: return True
    return False

def bold(r):
    rpr = r.find(q(M, 'rPr')); sty = rpr.find(q(M, 'sty')) if rpr is not None else None
    return sty is not None and sty.get(q(M, 'val')) in ('b', 'bi')

def ptext(p):
    return ''.join(t.text or '' for t in p.iter(q(W, 't'), q(M, 't')))

def main(path, out=None, full=False):
    z = zipfile.ZipFile(path)
    body = etree.fromstring(z.read('word/document.xml')).find(q(W, 'body'))
    heading, n_para = '', 0
    stats = collections.defaultdict(collections.Counter)
    samples = collections.defaultdict(list)
    def add(cat, key, where, text, limit=8):
        stats[cat][key] += 1
        if full or len(samples[(cat, key)]) < limit: samples[(cat, key)].append({'位置': where, '内容': text[:60]})
    n_omml = n_ole = 0
    for p in body.iter(q(W, 'p')):
        n_para += 1
        txt = ptext(p).strip()
        if re.match(r'^(第\s*\d+\s*章|\d+(\.\d+)+)\s*\S', txt) and len(txt) < 40: heading = txt[:30]
        where = f'{heading} ｜ 段{n_para}'
        wpos = {'heading': heading, 'para': n_para}
        maths = [m for m in p.iter(q(M, 'oMath'))]
        n_omml += len(maths)
        oles = list(p.iter('{urn:schemas-microsoft-com:office:office}OLEObject'))
        if oles:
            n_ole += len(oles)
            add('公式编辑器', 'MathType/OLE 公式', where, txt)
            for _ in oles[1:]: stats['公式编辑器']['MathType/OLE 公式'] += 1
        for m in maths:
            for r in m.iter(q(M, 'r')):
                t = ''.join(x.text or '' for x in r.iter(q(M, 't')))
                if not t: continue
                up = upright(r)
                for fm in FUNC_RE.finditer(t):
                    add('函数名正斜体', '正体（正确）' if up else '斜体（应为正体）', where, t)
                for fm in re.finditer(r'(?<![A-Za-z])(Sin|Cos|Tan|Cot|Ln|Lg|Log|Lim|Max|Min|Exp)(?![a-z])', t):
                    add('函数名首字母大写', fm.group(), where, t)
                if up and not any(ln(a) == 'fName' for a in r.iterancestors()):
                    rest = FUNC_RE.sub('', t)
                    rest = re.sub(r'(?<![A-Za-z])(kg|km|cm|mm|mol|Pa|Hz|min|max|lim|sup|inf|dx|dy|dz|dt|[dπe])(?![A-Za-z])', '', rest)
                    if re.search(r'[A-Za-z]', rest) and not bold(r):
                        add('字母排成正体（核对是否为变量）', '正体字母', where, t)
                if not up:
                    for um in re.finditer(r'\d\s*(kg|km|cm|mm|m/s²|m/s|m|s|Pa|N|J|W|V|A|mol|K|h|Hz|℃)(?![A-Za-z])', t):
                        add('单位排成斜体', um.group(1), where, t)
                if 'π' in t: add('π 的字形', '正体' if up else '斜体', where, t)
                # e：作为指数底数（e^x）或单独的 e
                if t == 'e' or re.fullmatch(r'e', t.strip()):
                    parent = r.getparent()
                    if parent is not None and ln(parent) == 'e' and parent.getparent() is not None and ln(parent.getparent()) == 'sSup':
                        add('常数 e 的字形', '正体' if up else '斜体', where, t)
                # d：微分号，如 dx、dy、d(…)，出现在分式分子分母或 d 后紧跟变量
                if re.fullmatch(r'd[a-zA-Zα-ωθφψ]?', t) or re.match(r'^d[xyztuvsθ]\b', t):
                    add('微分号 d 的字形', '正体' if up else '斜体', where, t)
                if bold(r): add('粗体字母（向量、矩阵需粗斜体，核对是否该粗）', '粗体', where, t, limit=12)
                for grp, chars in CHAR_GROUPS.items():
                    for ch, label in chars.items():
                        if ch in t: add(grp, label, where, t, limit=5)
        # 正文（非公式）里的撇号、增量符号也统计，便于发现公式外的混用
        for r in p.findall(q(W, 'r')):
            t = ''.join(x.text or '' for x in r.iter(q(W, 't')))
            for grp in ('撇号', '增量符号'):
                for ch, label in CHAR_GROUPS[grp].items():
                    if ch in t and grp == '增量符号': add(grp + '（正文）', label, where, t, limit=3)
    res = {'文件': path, 'Word公式数': n_omml, 'MathType/OLE公式数': n_ole,
           '统计': {k: dict(v) for k, v in stats.items()},
           '例子': {f'{c}｜{k}': v for (c, k), v in samples.items()}}
    if out: json.dump(res, open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'Word 公式 {n_omml} 个，MathType/OLE 公式 {n_ole} 个')
    for c, v in res['统计'].items():
        print(f'【{c}】' + '；'.join(f'{k} {n}' for k, n in v.items()))
    print('（混用：同一组出现两种以上写法即需全书统一；函数名“斜体”项为确定问题，其余需结合语义判断）')

if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if a != '--all']
    main(args[0], args[1] if len(args) > 1 else None, full='--all' in sys.argv)

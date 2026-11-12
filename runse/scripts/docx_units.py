# 把 docx 段落拆成“可以改的文字”和“不能动的块”（公式、图、脚注号、超链接、带制表符的字等）。
# 准备、核对、写回三个脚本共用这一份，保证三处看到的段落文字一模一样；核对写回结果的脚本另写一套，不用这里的代码。
import re, zipfile, copy
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
O = 'urn:schemas-microsoft-com:office:office'
XSP = '{http://www.w3.org/XML/1998/namespace}space'

def q(t): return f'{{{W}}}{t}'
def ln(el): return etree.QName(el).localname if isinstance(el.tag, str) else ''
def nsp(el): return etree.QName(el).namespace if isinstance(el.tag, str) else ''

# 占位：〔式n〕公式，〔图n〕图片或文本框，〔注n〕脚注尾注号，〔符n〕特殊字体的符号；n 在每段内各自从 1 数起
PH = re.compile(r'〔(?:式|图|注|符)\d+〕')
CONTAINERS = ('hyperlink', 'smartTag', 'fldSimple', 'ins', 'sdt', 'sdtContent', 'customXml')
SLINE = re.compile(r'^(S\d+)(【[^】]*】)?\s?(.*)$')


def load(path):
    """读 docx，返回 (ZipFile, document.xml 的根)"""
    z = zipfile.ZipFile(path)
    return z, etree.fromstring(z.read('word/document.xml'))


def has_revisions(doc):
    return next(doc.iter(q('ins'), q('del'), q('moveFrom'), q('moveTo')), None) is not None


def plain(r):
    """只装一段文字的普通 run：可以拆开来改"""
    return [ln(c) for c in r] in (['t'], ['rPr', 't'])


def clean_para(p):
    """删掉 Word 留下的分页记号和拼写标记：它们不是内容，却会把一句话的文字隔成改不了的碎块"""
    for e in list(p.iter(q('lastRenderedPageBreak'), q('proofErr'))):
        par = e.getparent()
        if e.tail and e.tail.strip(): continue
        par.remove(e)


def _is_equation(obj):
    ole = obj.find(f'{{{O}}}OLEObject')
    prog = (ole.get('ProgID') or '') if ole is not None else ''
    return ole is None or prog.startswith('Equation') or 'MathType' in prog


def shown(r, cnt):
    """一个 run 在导出文字里的样子；cnt 记本段各类占位数到第几个"""
    out = []
    def ph(kind):
        cnt[kind] = cnt.get(kind, 0) + 1; out.append(f'〔{kind}{cnt[kind]}〕')
    for ch in r:
        ns, tag = nsp(ch), ln(ch)
        if ns == W:
            if tag == 't': out.append(ch.text or '')
            elif tag == 'tab': out.append('\t')
            elif tag == 'noBreakHyphen': out.append('-')
            elif tag == 'object': ph('式' if _is_equation(ch) else '图')
            elif tag in ('pict', 'drawing'): ph('图')
            elif tag in ('footnoteReference', 'endnoteReference'): ph('注')
            elif tag == 'sym': ph('符')
        elif tag == 'AlternateContent': ph('图')
    return ''.join(out)


def _inner(el, cnt):
    """容器（超链接等）里的文字，整体当一块"""
    out = ''
    for ch in el:
        ns, tag = nsp(ch), ln(ch)
        if ns == M and tag in ('oMath', 'oMathPara'):
            cnt['式'] = cnt.get('式', 0) + 1; out += f"〔式{cnt['式']}〕"
        elif ns == W and tag == 'r':
            if ch.find(q('fldChar')) is None: out += shown(ch, cnt)
        elif ns == W and tag in CONTAINERS: out += _inner(ch, cnt)
    return out


def build_units(p, infield=None):
    """按顺序返回段落里看得见的东西：
       {'k':'txt','run':r}                          可以改的文字
       {'k':'atom','first':el,'last':el,'disp':s}    不能动的块，disp 是它在导出文字里的样子
       infield：传一个 set 进来，落在域（交叉引用、文献序号等）显示文字里的 run 会记进去"""
    units = []; cnt = {}; depth = 0
    for ch in p:
        ns, tag = nsp(ch), ln(ch)
        if ns == M and tag in ('oMath', 'oMathPara'):
            cnt['式'] = cnt.get('式', 0) + 1
            units.append({'k': 'atom', 'first': ch, 'last': ch, 'disp': f"〔式{cnt['式']}〕"}); continue
        if ns != W: continue
        if tag == 'r':
            fc = ch.find(q('fldChar'))
            if fc is not None:                       # 域的起止标记本身看不见、不动；只记下现在是不是在域里
                ty = fc.get(q('fldCharType'))
                depth += 1 if ty == 'begin' else (-1 if ty == 'end' else 0); continue
            d = shown(ch, cnt)
            if not d: continue                       # 没有可见内容的 run（域代码、批注引用等）
            if plain(ch):
                units.append({'k': 'txt', 'run': ch})
                if depth and infield is not None: infield.add(ch)
            else:
                units.append({'k': 'atom', 'first': ch, 'last': ch, 'disp': d})
        elif tag in CONTAINERS:
            d = _inner(ch, cnt)
            if d: units.append({'k': 'atom', 'first': ch, 'last': ch, 'disp': d})
    if depth: raise ValueError('段里的域没有配对')
    return units


def unit_text(u):
    return (u['run'].find(q('t')).text or '') if u['k'] == 'txt' else u['disp']


def para_text(p, infield=None):
    return ''.join(unit_text(u) for u in build_units(p, infield))


def fixed_parts(p):
    """这一段里不能动的块（按顺序），和域里的显示文字（改了要提醒）"""
    p = copy.deepcopy(p); clean_para(p)          # 和写回脚本一样先去掉分页记号、拼写标记，免得把能改的文字算成不能动的
    inf = set(); units = build_units(p, inf)
    atoms = [u['disp'] for u in units if u['k'] == 'atom' and u['disp'].strip()]
    fields = [unit_text(u) for u in units if u['k'] == 'txt' and u['run'] in inf and unit_text(u).strip()]
    return atoms, fields


def body_rows(doc):
    """正文里一行一行的东西：('p', 文字, 段落元素) 和 ('t', 各格文字, 表格行元素)；空段不算"""
    rows = []
    for el in doc.find(q('body')):
        tag = ln(el)
        if tag == 'p':
            try: t = para_text(el)
            except ValueError: t = ''.join(x.text or '' for x in el.iter(q('t')))
            if t.strip(): rows.append(('p', t, el))
        elif tag == 'tbl':
            for tr in el.iter(q('tr')):
                cells = []
                for tc in tr.findall(q('tc')):
                    cells.append(''.join(para_text(x) for x in tc.iter(q('p'))).strip())
                rows.append(('t', ' | '.join(cells), tr))
    return rows


def read_slines(path):
    """读正文.txt，返回 [(段号, 标记, 文字)]"""
    out = []
    for line in open(path, encoding='utf-8'):
        m = SLINE.match(line.rstrip('\n'))
        if m: out.append((m.group(1), m.group(2) or '', m.group(3)))
    return out


def align(rows, slines):
    """把正文.txt 的段号对到 docx 的行上；对不上（稿子换过、正文不是从这份稿子导的）就报错。返回 {段号: 元素}"""
    def same(row, s):
        sid, tag, text = s
        if tag.startswith('【表格'): return row[0] == 't'
        return row[0] == 'p' and row[1].strip() == text
    cands = [a for a in range(len(rows) - len(slines) + 1) if all(same(rows[a + k], s) for k, s in enumerate(slines))]
    if len(cands) != 1:
        raise SystemExit(f'正文.txt 和这份 Word 对不上（能对上的位置有 {len(cands)} 个）。确认正文是从这份稿子导出的；稿子改过就重新跑准备脚本。')
    return {s[0]: rows[cands[0] + k][2] for k, s in enumerate(slines)}

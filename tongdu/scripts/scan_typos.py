# 错别字、重复字词、易混词、标点、量和单位、数字用法、英文的本地初筛：只出候选，交 AI 逐条判断
# 用法: scan_typos.py 书稿.docx 输出目录 [--from P233] [--to P288] [--words 易错词表.md] [--units space|nospace|off] [--with-punct]
#   --with-punct：连标点、括号引号配对、公式残句的候选也出。默认不出：句子和标点归润色技能（2026-11-12 编辑定）
#   --units：本书数值与单位之间的体例。space（国标，默认）报不空格的；nospace 报空了格的；off 不查
# 输出: 输出目录/typo-candidates.json（候选问题，带全文唯一 ctx）、typo-report.md（按类列表，给人看）
import sys, os, re, json, zipfile, logging
from collections import Counter, defaultdict
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
HERE = os.path.dirname(os.path.abspath(__file__))
HAN = r'一-鿿'

# ---------- 参数 ----------
args = sys.argv[1:]
def opt(name, default=None):
    if name in args:
        i = args.index(name); v = args[i + 1]; del args[i:i + 2]; return v
    return default
p_from, p_to = opt('--from'), opt('--to')
words_md = opt('--words', os.path.join(HERE, '..', 'references', '易错词表.md'))
units_style = opt('--units', 'space')
WITH_PUNCT = '--with-punct' in args
if WITH_PUNCT: args.remove('--with-punct')
src, outdir = args[0], args[1]
pnum = lambda s: int(s.lstrip('PpTt')) if s else None
p_from, p_to = pnum(p_from), pnum(p_to)

# ---------- 读取段落 ----------
# P 号与 dump_docx.py 一致：正文段落和表格的每一行各占一个号
doc = etree.fromstring(zipfile.ZipFile(src).read('word/document.xml'))
plabel, n = {}, 0
for el in doc.find(q('body')):
    if el.tag == q('p'):
        n += 1
        for x in el.iter(q('p')): plabel[x] = n
    elif el.tag == q('tbl'):
        for tr in el.iter(q('tr')):
            n += 1
            for x in tr.iter(q('p')): plabel[x] = n

M_OMATH = '{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath'
paras = []   # 每段：P 号、本段自有文字（不含嵌套文本框段落）、逐字是否上标、修订脚本 track 能看到的文字
for p in doc.iter(q('p')):
    text, sup = [], []
    for t in p.iter(q('t')):
        if next(t.iterancestors(q('p'))) is not p: continue
        r = next(t.iterancestors(q('r')), None)
        va = r.find(f'{q("rPr")}/{q("vertAlign")}') if r is not None else None
        s = t.text or ''
        text.append(s); sup += [va is not None and va.get(q('val')) == 'superscript'] * len(s)
    direct = ''.join(r.find(q('t')).text or '' for r in p.findall(q('r')) if len(r.findall(q('t'))) == 1)
    # 公式对象（MathType、Word 公式）在本段文字里的位置：记下它前面有几个字
    objs = []; k = 0
    for el in p.iter(q('t'), q('object'), M_OMATH):
        if next(el.iterancestors(q('p')), None) is not p: continue
        if el.tag == q('t'): k += len(el.text or '')
        elif not objs or objs[-1] != k: objs.append(k)
    paras.append(dict(P=plabel.get(p, 0), text=''.join(text), sup=sup, direct=direct, objs=objs))
ALL = [x['text'] for x in paras]

def in_range(P):
    return (p_from is None or P >= p_from) and (p_to is None or P <= p_to)

# ---------- 上下文定位（ctx 全文唯一，且修订脚本在 ctx 里找到的第一个 old 就是这一处） ----------
def make_ctx(pi, k, old):
    t = ALL[pi]
    for pad in (4, 6, 8, 12, 20, 40):
        for left in (pad, 0):
            a = max(0, k - left); s = t[a: k + len(old) + pad]
            if len(s.strip()) < 4: continue
            if sum(x.count(s) for x in ALL) == 1 and t.find(s) == a and t.find(old, a) == k:
                return s
    return t if sum(x.count(t) for x in ALL) == 1 else None

# ---------- 词表 ----------
def md_tables(path):
    """返回 {小节标题: [行(单元格列表)]}；单元格里的 \\| 还原为 |"""
    out, sec = defaultdict(list), ''
    for line in open(path, encoding='utf-8'):
        line = line.rstrip('\n')
        if line.startswith('## '): sec = line[3:].strip(); continue
        if not line.startswith('|') or re.match(r'^\|[\s\-|:]+\|$', line): continue
        cells = [c.strip().replace('\\|', '|') for c in re.split(r'(?<!\\)\|', line.strip()[1:-1])]
        out[sec].append(cells)
    return out
tabs = md_tables(words_md)
sure, confusable = [], []
for sec, rows in tabs.items():
    body = rows[1:]   # 去掉表头
    if '确定错词' in sec:
        for c in body:
            if c[0]: sure.append(dict(wrong=c[0], right=c[1], excl=[x for x in re.split('[、，,]', c[2]) if x] if len(c) > 2 else [], note=c[3] if len(c) > 3 else ''))
    elif '易混词' in sec:
        for c in body:
            if c[0]: confusable.append(dict(pat=re.compile(c[0]), excl=re.compile(c[1]) if len(c) > 1 and c[1] else None, hint=c[2] if len(c) > 2 else ''))

# ---------- 候选 ----------
cands = []
def add(pi, k, old, typ, rule, hint, new=None):
    if not WITH_PUNCT and typ in ('标点', '配对', '病句'): return      # 句子和标点归润色，通读不判断这几类候选
    ctx = make_ctx(pi, k, old)
    t = ALL[pi]
    cands.append(dict(type=typ, rule=rule, para=f"P{paras[pi]['P']}", old=old, new=new, hint=hint, ctx=ctx,
                      edit_ok=bool(ctx) and ctx in paras[pi]['direct'],
                      around=t[max(0, k - 15): k + len(old) + 15], off=min(15, k), pos=k))

import jieba
jieba.setLogLevel(logging.WARNING)
AA_OK = set('常常 渐渐 往往 慢慢 刚刚 仅仅 稍稍 人人 处处 个个 天天 年年 层层 步步 点点 种种 多多 谢谢 看看 试试 想想 紧紧 '
            '深深 高高 明明 偏偏 统统 白白 默默 纷纷 频频 重重 略略 轻轻 久久 匆匆 隐隐 微微 累累 等等 一一 每每 时时 事事 '
            '条条 件件 句句 字字 片片 次次 代代 大大 小小 好好 早早 远远 长长 满满 静静 悄悄 缓缓 徐徐 稳稳 牢牢 区区 '
            '屡屡 万万 环环 节节 样样 项项 恰恰 寥寥 历历 赫赫 孜孜 源源 滚滚 茫茫 勃勃 跃跃 井井 楚楚 洋洋 绵绵 漫漫 '
            '悠悠 翩翩 姗姗 袅袅 熊熊 蒸蒸 津津 栩栩 彬彬 绰绰 依依 闪闪 冉冉 莘莘 芸芸 碌碌 浩浩 荡荡 哈哈 呵呵 嘻嘻'.split())
unit_space = [0, 0]   # 本范围内数值与单位之间：有空格、无空格各几处
PUNCT_END = '。！？：；…）)”’》」』】—~～.!?:;'

for pi, para in enumerate(paras):
    P, t = para['P'], para['text']
    if not in_range(P) or not t.strip(): continue
    is_heading = len(t) < 50 and not re.search('[。；，]', t)

    # 1. 重复字词
    toks, pos = [], 0
    for w in jieba.cut(t):
        toks.append((pos, w)); pos += len(w)
    starts = {s: w for s, w in toks}
    tok_at = {}
    for s, w in toks:
        for j in range(s, s + len(w)): tok_at[j] = (s, w)
    for m in re.finditer(rf'([{HAN}])\1', t):
        k = m.start()
        if m.group(0) in AA_OK or t[k:k + 4] in ('的的确确',): continue
        s1, w1 = tok_at[k]; s2, w2 = tok_at[k + 1]
        if s1 == s2 and len(w1) >= 2: continue                       # “常常”“谢谢”等被切成一个词
        if s1 != s2 and len(w1) >= 2 and len(w2) >= 2: continue      # “为了|了解”“具有|有效”：两个词首尾相接
        if k > 0 and t[k - 1] == m.group(1): continue                 # 三连字另行处理
        sure_ = len(w1) == 1 and len(w2) == 1
        add(pi, k, m.group(0), '重复字词', 'AA', '同一字连写，疑多字' + ('' if sure_ else '（一侧是词的一部分，多半是正常搭配，请看语境）'),
            m.group(1) if sure_ else None)
    for m in re.finditer(rf'([{HAN}]{{2,4}})\1', t):
        add(pi, m.start(), m.group(0), '重复字词', 'ABAB', '同一词连写，疑多字', m.group(1))
    for m in re.finditer(r'\b([A-Za-z]{2,})\s+\1\b', t, re.I):
        add(pi, m.start(), m.group(0), '重复字词', 'word-word', '英文单词重复', m.group(1))
    # 5～30 字的片段原样连写两次（术语或单位重复粘贴，如“社会贡献弃风弃光电量社会贡献弃风弃光电量”“W/m²W/m²”）
    for m in re.finditer(r'(\S{3,30}?)\1', t):
        g = m.group(1)
        if len(g) < 5 and not re.search(r'[A-Za-z/²³]', g): continue          # 短的中文重复由 AA/ABAB 规则处理
        if re.fullmatch(r'[\d.\s…—\-_=*·]+', g): continue                     # 数字串、省略号、分隔线
        add(pi, m.start(), m.group(0), '重复字词', '片段重复', '同一片段连写两次，疑重复粘贴', g)
    # 不可见字符：零宽字符多为复制粘贴残留，应删；私用区字符多是 Symbol 等符号字体的字（如减号），要看版面确认显示正常
    for m in re.finditer(r'[\u200b-\u200f\u2060\ufeff]', t):
        k = m.start()
        add(pi, max(0, k - 1), t[max(0, k - 1): k + 2], '不可见字符', '零宽字符', f'含零宽字符 U+{ord(m.group(0)):04X}（看不见，会干扰检索和排版），应删除',
            t[max(0, k - 1): k] + t[k + 1: k + 2])
    for m in re.finditer(r'[\ue000-\uf8ff]', t):
        k = m.start()
        add(pi, max(0, k - 1), t[max(0, k - 1): k + 2], '不可见字符', '符号字体字符', f'私用区字符 U+{ord(m.group(0)):04X}，多为 Symbol 等符号字体里的符号（如减号），提取文字时显示不出来；请看版面确认是什么字，必要时换成标准字符')

    # 2. 确定错词
    for e in sure:
        k = t.find(e['wrong'])
        while k >= 0:
            if not any(x in t[max(0, k - len(x)): k + len(e['wrong']) + len(x)] and e['wrong'] in x
                       and t.find(x, max(0, k - len(x))) <= k for x in e['excl']):
                add(pi, k, e['wrong'], '易错词', '词表', ('词表：' + e['note']) if e['note'] else '词表确定错词', e['right'])
            k = t.find(e['wrong'], k + 1)

    # 3. 易混词
    for e in confusable:
        for m in e['pat'].finditer(t):
            # 排除：匹配文字连同其前后各 6 个字里能找到排除式，就不算
            if e['excl'] and e['excl'].search(t[max(0, m.start() - 6): m.end() + 6]): continue
            add(pi, m.start(), m.group(0), '易混词', e['pat'].pattern[:20], e['hint'])

    # 公式残句：“其中”后紧跟公式就断句（“其中vᵢ；”），缺“为……”“表示……”的说明，要补全或删去
    for k in paras[pi]['objs']:
        if t[max(0, k - 2):k] in ('其中', '式中') and t[k:k + 1] in '；。，':
            add(pi, max(0, k - 2), t[max(0, k - 2):k + 1], '病句', '公式残句', '“其中”后面只有公式就断句，缺少说明（如“其中，vᵢ为……”）；请补全说明或删去这半句')
    # 4. 标点
    for m in re.finditer(rf'(?<=[{HAN}])[,;:!?] ?|[,;:!?] ?(?=[{HAN}])', t):
        cn = {',': '，', ';': '；', ':': '：', '!': '！', '?': '？'}[m.group(0)[0]]
        add(pi, m.start(), m.group(0), '标点', '半角标点', '中文句中用了半角标点', cn)
    for m in re.finditer(rf'(?<=[{HAN}）”])\.(?![\dA-Za-z])', t):
        add(pi, m.start(), '.', '标点', '半角句号', '中文句中用了英文句号', '。')
    for m in re.finditer(r'([，。、；：！？])\1|[，、；：。][，、；：。]', t):
        add(pi, m.start(), m.group(0), '标点', '重复标点', '标点重复或连用', m.group(0)[0] if m.group(0)[0] == m.group(0)[-1] else None)
    for m in re.finditer(rf'\(([^()]*[{HAN}][^()]*)\)', t):
        add(pi, m.start(), m.group(0), '标点', '半角括号', '中文内容用了半角括号', '（' + m.group(1) + '）')
    for m in re.finditer(rf'(?<![A-Za-z])([A-Za-z]|sin|cos|tan|log|ln|exp|max|min|softmax)（([^（）{HAN}，]{{1,30}})）', t):
        add(pi, m.start(), m.group(0), '标点', '公式括号', '函数、算式里的括号一般用半角“()”（如 O(N log N)），请按本书公式体例确认',
            f'{m.group(1)}({m.group(2)})')
    for m in re.finditer(rf'(?<=[{HAN}])"|"(?=[{HAN}])', t):
        add(pi, m.start(), '"', '标点', '直引号', '中文里用了英文直引号，应为“”')
    for m in re.finditer(rf'(?<!…)…(?!…)|(?<=[{HAN}])\.\.\.', t):
        add(pi, m.start(), m.group(0), '标点', '省略号', '省略号应为“……”（两个“…”）', '……')
    # 单个“—”前后都是正文文字才报；行首的“—”多半另一半在公式对象里（“式中”说明行），版面上是完整的“——”
    for m in re.finditer(rf'(?<=[{HAN}”）》A-Za-z])—(?=[{HAN}“（《A-Za-z])', t):
        add(pi, m.start(), '—', '标点', '破折号', '破折号应占两个字位，作“——”', '——')
    for m in re.finditer(r'([。！？；])(\[\d+(?:[-–,，～]\d+)*\])', t):
        add(pi, m.start(), m.group(0), '标点', '文献序号位置', '文献序号应放在句末标点之前', m.group(2) + m.group(1))
    # 三项及以上顿号列举、末项用“以及”时，“以及”前加逗号（两项“A以及B”不加）
    for m in re.finditer(r'以及', t):
        k = m.start()
        if k == 0 or t[k - 1] in '，,；;：:': continue
        seg = re.split(r'[，。；：！？,;:]', t[:k])[-1]
        if '、' in seg and len(seg) <= 60:
            add(pi, k - 1, t[k - 1:k + 2], '标点', '以及前逗号', '三项及以上顿号列举、末项用“以及”时，“以及”前加逗号', t[k - 1] + '，以及')
    # 冒号后带序号的分项列举（（1）……，（2）……），项与项之间用分号（GB/T 15834 分号用于分项列举）
    SEQ = r'(?:（\d{1,2}）|\(\d{1,2}\)|\d{1,2}\)|\d{1,2}）|[①②③④⑤⑥⑦⑧⑨⑩]|[二三四五六七八九十]是)'
    for c in re.finditer(r'：', t):
        sent = re.split(r'[。！？]', t[c.end():])[0]
        if not re.match(r'\s*' + SEQ.replace('[二三四五六七八九十]是', '一是'), sent): continue
        for m in re.finditer(r'，(?=\s*' + SEQ + ')', sent):
            k = c.end() + m.start()
            add(pi, k, t[k:k + 3], '标点', '分项列举用分号', '冒号后带序号的分项列举，项与项之间用分号', '；' + t[k + 1:k + 3])
    for m in re.finditer(r' +(?=[，。、；：！？）])', t):
        add(pi, m.start(), m.group(0) + t[m.end()], '标点', '标点前空格', '中文标点前有空格', t[m.end()])
    if not is_heading:
        for m in re.finditer(rf'(?<=[{HAN}]) +(?=[{HAN}])', t):
            add(pi, m.start() - 1, t[m.start() - 1: m.end() + 1], '标点', '汉字间空格', '汉字之间有空格', t[m.start() - 1] + t[m.end()])
    m = re.match(r'^\s*(\d+)、', t)
    if m: add(pi, m.start(1), m.group(0).strip(), '标点', '序号顿号', '阿拉伯数字序号后用“.”不用“、”（GB/T 15834）', m.group(1) + '.')
    if not is_heading and len(t.strip()) > 20 and t.rstrip()[-1] not in PUNCT_END and not re.search(r'（\d+-\d+）\s*$', t):
        k = len(t.rstrip()) - 1
        add(pi, max(0, k - 3), t[max(0, k - 3): k + 1], '标点', '段末无标点', '段落结尾没有标点：可能漏标点，也可能是段落被误断（与下一段连读判断）')

    # 5. 括号、引号配对
    PAIR = {'（': '）', '(': ')', '【': '】', '《': '》', '“': '”', '‘': '’', '[': ']', '「': '」'}
    CLOSE = {v: k_ for k_, v in PAIR.items()}
    stack = []
    for k, ch in enumerate(t):
        if ch in PAIR and not (ch == '‘' and False):
            stack.append((k, ch))
        elif ch in CLOSE:
            if stack and PAIR[stack[-1][1]] == ch: stack.pop()
            elif stack and {stack[-1][1], ch} in ({'（', ')'}, {'(', '）'}):
                k0, o = stack.pop()
                add(pi, k0, t[k0:k + 1], '配对', '括号全半角混用', f'左括号“{o}”与右括号“{ch}”全半角不一致')
            else:
                if re.match(r'^\s*[A-Za-z0-9]{1,2}$', t[:k]): continue          # “a）”“1）”式序号
                if ch == '’' and k > 0 and re.match('[A-Za-z]', t[k - 1]): continue   # 英文撇号
                add(pi, k, ch, '配对', '多余右括号/引号', f'“{ch}”前面没有对应的“{CLOSE[ch]}”')
    for k0, o in stack:
        add(pi, k0, o, '配对', '缺右括号/引号', f'“{o}”在本段没有对应的“{PAIR[o]}”（跨段引文除外）')

    # 6. 全半角
    for m in re.finditer(r'[０-９Ａ-Ｚａ-ｚ％＋＝＜＞]+', t):
        half = ''.join(chr(ord(c) - 0xFEE0) for c in m.group(0))
        add(pi, m.start(), m.group(0), '全半角', '全角字母数字', '数字、字母、运算符应为半角', half)
    # 英文里的标点用半角：英文单词之间的弯撇号、全角逗号冒号分号（Word 用中文字体会显示成全角）
    for m in re.finditer(r"(?<=[A-Za-z])[’‘](?=[A-Za-z])", t):
        add(pi, m.start(), m.group(0), '全半角', '英文撇号', "英文里的撇号用半角 '", "'")
    for m in re.finditer(r'(?<=[A-Za-z])([，：；])(?= ?[A-Za-z]{2})', t):
        add(pi, m.start(), m.group(0), '全半角', '英文标点', '英文里的标点用半角', {'，': ',', '：': ':', '；': ';'}[m.group(1)])
    for m in re.finditer(r'(?<=\d)\s*~\s*(?=\d)', t):
        add(pi, m.start(), m.group(0), '全半角', '范围号', '数值范围号中文书稿一般用“～”', '～')

    # 7. 量和单位
    UNIT = r'(km|cm|mm|nm|μm|m|kg|g|h|min|ms|s|kW·h|kWh|MWh|GWh|TWh|W|kW|MW|GW|TW|Hz|kHz|MHz|GHz|V|kV|mA|A|hPa|kPa|MPa|Pa|K|J|kJ|MJ|N|L|dB)'
    for m in re.finditer(rf'(?<![A-Za-z\d.])(\d+(?:\.\d+)?[亿万]?)[ \u00a0]+{UNIT}(?![A-Za-z·])', t):   # 含“3.9亿 kW”
        unit_space[0] += 1
        if units_style == 'nospace':
            add(pi, m.start(), m.group(0), '量和单位', '数值与单位间有空格', '本书体例数值与单位符号之间不空格', m.group(1) + m.group(2))
    for m in re.finditer(rf'(?<![A-Za-z\d.])(\d+(?:\.\d+)?){UNIT}(?![A-Za-z·])', t):
        unit_space[1] += 1
        if units_style == 'space':
            add(pi, m.start(), m.group(0), '量和单位', '数值与单位间无空格',
                '国标写法数值与单位符号之间留一个空格（如“1 h”“30 km”）；本书若统一不空格，用 --units nospace 重跑', m.group(1) + ' ' + m.group(2))
    BADCASE = {'S': 's', 'H': 'h', 'MIN': 'min', 'Min': 'min', 'KM': 'km', 'Km': 'km', 'KG': 'kg', 'Kg': 'kg', 'KW': 'kW', 'Kw': 'kW',
               'KWH': 'kW·h', 'kwh': 'kW·h', 'HZ': 'Hz', 'hz': 'Hz', 'KHZ': 'kHz', 'MHZ': 'MHz', 'GHZ': 'GHz', 'mw': 'MW', 'Mw': 'MW', 'gw': 'GW', 'Gw': 'GW',
               'MS': 'ms', 'KV': 'kV', 'pa': 'Pa', 'HPA': 'hPa', 'hpa': 'hPa', 'KPA': 'kPa', 'Kpa': 'kPa', 'Mpa': 'MPa', 'MPA': 'MPa', 'l': 'L'}
    for m in re.finditer(r'(?<![A-Za-z\d.])(\d+(?:\.\d+)?)\s?(' + '|'.join(sorted(BADCASE, key=len, reverse=True)) + r')(?![A-Za-z\d])', t):
        add(pi, m.start(2), m.group(2), '量和单位', '单位符号大小写', f'单位符号大小写有误，应为“{BADCASE[m.group(2)]}”', BADCASE[m.group(2)])
    for m in re.finditer(r'(?<=\d)\s*[xX]\s*(?=\d)', t):
        add(pi, m.start(), m.group(0), '量和单位', '乘号', '乘号应为“×”', '×')
    for m in re.finditer(r'(?<![A-Za-z])(km|cm|mm|m|s)([23])(?![\d])', t):
        k = m.start(2)
        if k < len(para['sup']) and not para['sup'][k]:
            add(pi, m.start(), m.group(0), '量和单位', '指数未上标', f'单位指数“{m.group(2)}”应为上标')
    for m in re.finditer(r'(?<![\d.,])\d{5,}(?![\d.])', t):
        if not re.match(r'(19|20)\d{2}$', m.group(0)):
            add(pi, m.start(), m.group(0), '量和单位', '多位数分节', '五位及以上整数一般三位分节（如“45 000”），与全书体例核对')

# 7b. 数字用法（GB/T 15835—2011）：只出候选。选汉字还是阿拉伯数字标准不强求，全书按“同类别同形式”统一
RANGE_SEP = '[~～—–]'
CN_DIGIT = '一二三四五六七八九十'
COUNT_WORD = '个种类项'
num_forms = Counter()
for pi, para in enumerate(paras):
    P, t = para['P'], para['text']
    if not in_range(P): continue
    for m in re.finditer(rf'(?<![\d.%])(\d+(?:\.\d+)?)({RANGE_SEP})(\d+(?:\.\d+)?)%', t):
        add(pi, m.start(), m.group(0), '数字用法', '范围省略百分号', '范围两端都要写“%”（GB/T 15835—2011 5.1.3：不写为15~30%）', f'{m.group(1)}%{m.group(2)}{m.group(3)}%')
    for m in re.finditer(rf'(?<![\d.万亿])(\d+(?:\.\d+)?)({RANGE_SEP})(\d+(?:\.\d+)?)(万|亿)(元|千瓦|吨|人)?', t):
        u = m.group(4) + (m.group(5) or '')
        add(pi, m.start(), m.group(0), '数字用法', '范围省略万亿', '范围前一个数值不能省略“万”“亿”（5.1.3：不写为9~16亿）', f'{m.group(1)}{u}{m.group(2)}{m.group(3)}{u}')
    for m in re.finditer(rf'(?<![\d.])\d+(?:\.\d+)?{RANGE_SEP}\d+(?:\.\d+)?×10', t):
        add(pi, m.start(), m.group(0), '数字用法', '范围省略幂', '范围两端都要写“×10ⁿ”（5.1.3：不写为4.3~5.7×10⁶）；指数是上标，请看原文后改')
    for m in re.finditer(r'(?<![\d.])\d+[百千](?![克米瓦伏帕赫卡焦牛安字分])', t):
        add(pi, m.start(), m.group(0), '数字用法', '数字与百千混写', '一般数值不混用阿拉伯数字和汉字（5.3：4000不写为“4千”），只有“万”“亿”可以')
    for m in re.finditer(rf'(?<![第{CN_DIGIT}])([{CN_DIGIT}])、([{CN_DIGIT}])(?=[个天年月日周米岁次倍种类人项条位])', t):
        add(pi, m.start(), m.group(0), '数字用法', '概数加顿号', '两数连用表概数时不用顿号（5.2.1）；若确是列举两项，则不改', m.group(1) + m.group(2))
    for m in re.finditer(r'(?<![\d.第图表式节章步层级类组号])([1-9])、([1-9])(?=[个天年月周岁次倍种人项])', t):
        add(pi, m.start(), m.group(0), '数字用法', '阿拉伯数字表概数', '数字连用表概数应用汉字（4.2.2：三四个月）；若是列举两个数值，则不改')
    for m in re.finditer(r'\d+几|几\d+', t):
        add(pi, m.start(), m.group(0), '数字用法', '含几的概数', '含“几”的概数应用汉字（4.2.2：二十几、几千）')
    for m in re.finditer(r'(?<![\d.纪])(\d{2})年代', t):
        if '世纪' in t[max(0, m.start() - 12): m.start()]: continue       # “20世纪80年代末到90年代初”“20世纪80—90年代”
        add(pi, m.start(), m.group(0), '数字用法', '年份简写', '年代写全，如“20世纪90年代”（5.1.4：四位年份不简写为两位）')
    for m in re.finditer(r'(?<![\d.])(\d{2})年(?=\d{1,2}月)', t):
        add(pi, m.start(), m.group(0), '数字用法', '年份简写', '四位年份不简写为两位（5.1.4：“1990年”不写为“90年”）')
    for m in re.finditer(r'[○〇零一二三四五六七八九]{4}年(?=\d)', t):
        add(pi, m.start(), m.group(0), '数字用法', '年月日混用', '同一日期不混用汉字和阿拉伯数字（4.3 示例3）')
    for m in re.finditer(r'[一二]零[○〇零一二三四五六七八九]{2}年', t):
        add(pi, m.start(), m.group(0), '数字用法', '编号用○', '年份是编号，“0”写作“○”（5.2.5）', m.group(0).replace('零', '○'))
    for m in re.finditer(r'(?<![\dA-Za-z.）)\]])\.\d+', t):
        add(pi, m.start(), m.group(0), '数字用法', '纯小数缺0', '纯小数要写小数点前的“0”（5.1.2）', '0' + m.group(0))
    for m in re.finditer(r'\d。\d{1,3}(?![\d年])', t):
        add(pi, m.start(), m.group(0), '数字用法', '小数点', '小数点用“.”，不用句号（5.1.2）', m.group(0).replace('。', '.'))
    # 同类别同形式：同一段里同一量词既有“3个”又有“三个”的报候选，全章数量只作统计，给定体例用
    for w in COUNT_WORD:
        ar = list(re.finditer(rf'(?<![\d.第])(?:[3-9]|10)(?={w})', t))
        cn = list(re.finditer(rf'(?<![第{CN_DIGIT}])[三四五六七八九十](?={w})', t))
        num_forms['阿拉伯'] += len(ar); num_forms['汉字'] += len(cn)
        if ar and cn:
            add(pi, ar[0].start(), ar[0].group(0) + w, '数字用法', '同段数字形式不一', f'同一段里“{ar[0].group(0)}{w}”和“{cn[0].group(0)}{w}”并用，同类别应同形式（4.3）')
    num_forms['千分撇'] += len(re.findall(r'(?<![\d.])\d{1,3}(?:,\d{3})+(?!\d)', t))
    num_forms['五位以上不分节'] += len(re.findall(r'(?<![\d.,])\d{5,}(?![\d.,])', t))

# 7c. 英文：本书高频词的近形低频词（疑拼错）、同一词内部大小写不同（GitHub/Github）。专有名词要按名词表看
EN_WORD = re.compile(r"(?<![A-Za-z])[A-Za-z][A-Za-z]{3,}(?![A-Za-z])")
def edit1(a, b):
    """编辑距离为 1，或相邻两字母对调"""
    if a == b or abs(len(a) - len(b)) > 1: return False
    if len(a) == len(b):
        diff = [i for i in range(len(a)) if a[i] != b[i]]
        return len(diff) == 1 or (len(diff) == 2 and diff[1] == diff[0] + 1 and a[diff[0]] == b[diff[1]] and a[diff[1]] == b[diff[0]])
    if len(a) > len(b): a, b = b, a
    i = 0
    while i < len(a) and a[i] == b[i]: i += 1
    return a[i:] == b[i + 1:]
def inflected(a, b):
    a, b = sorted((a, b), key=len)
    return any(b == a + x for x in ('s', 'd', 'r', 'es', 'ed', 'er'))
en_vocab, en_forms = Counter(), defaultdict(Counter)
for x in ALL:
    for m in EN_WORD.finditer(x):
        en_vocab[m.group(0).lower()] += 1; en_forms[m.group(0).lower()][m.group(0)] += 1
en_first = {}
for pi, para in enumerate(paras):
    if not in_range(para['P']): continue
    for m in EN_WORD.finditer(para['text']):
        en_first.setdefault(m.group(0), (pi, m.start()))
common = [w for w, c in en_vocab.items() if c >= 5]
for form, (pi, k) in en_first.items():
    w = form.lower()
    if len(w) >= 6 and en_vocab[w] <= 2:
        near = [v for v in common if edit1(w, v) and en_vocab[v] >= 3 * en_vocab[w] and not inflected(w, v)]
        if near:
            v = max(near, key=en_vocab.get)
            add(pi, k, form, '英文', '疑似拼写', f'全书“{v}”出现 {en_vocab[v]} 次，这个写法只有 {en_vocab[w]} 次，疑拼错')
    variants = {f[0].lower() + f[1:] for f in en_forms[w]}          # 句首大写不算差别
    if len(variants) > 1 and form != en_forms[w].most_common(1)[0][0]:
        add(pi, k, form, '英文', '大小写不一', '同一词有多种大小写：' + '、'.join(f'{f}（{n}）' for f, n in en_forms[w].most_common()) + '，请按名词表核对')

# 8. 总述里的数量（“分为三个阶段”“包括以下四个方面”）与后文实际列举的项数
CN_NUM = {'两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}
MARK = re.compile(r'^\s*(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|（[1-9]）|\([1-9]\)|[1-9][）)]|第[一二三四五六七八九十]|首先|其次|再次|然后|最后)')
CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩'
for pi, para in enumerate(paras):
    P, t = para['P'], para['text']
    if not in_range(P): continue
    for m in re.finditer(r'(?:分为|分成|包括(?:以下)?|包含(?:以下)?|有以下|归纳为|概括为)([两二三四五六七八九十2-9])(?:个)?(阶段|步|步骤|类|种|个方面|方面|部分|环节|点|条)', t):
        want = CN_NUM.get(m.group(1)) or int(m.group(1))
        # 先看本段冒号之后是否就地列举（“首先……；其次……；最后……”或①②③）
        tail = t[m.end():]
        inline = re.findall(r'(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|（[1-9]）|首先|其次|再次|最后|第[一二三四五六七八九十]|一是|二是|三是|四是|五是)', tail)
        got = len(inline)
        if got < 2:   # 本段没列，往后看连续的编号段落
            got, j = 0, pi + 1
            while j < len(paras) and j <= pi + 30:
                tx = paras[j]['text']
                if not tx.strip(): j += 1; continue
                if re.match(r'^\s*\d+\.\d+', tx) and len(tx) < 50: break           # 遇到下一节标题
                mk = MARK.match(tx)
                if mk:
                    n_in = len(re.findall(r'(首先|其次|再次|然后|最后|第[一二三四五六七八九十]，)', tx))
                    got += n_in if n_in >= 2 else 1     # “首先……；其次……；最后……”写在同一段
                elif got and len(tx) > 0 and not re.match(r'^\s*[\[（(]', tx) and got >= want: break
                j += 1
            # 编号是①②③时，以最大编号为准（中间可能夹着说明段）
            marks = [MARK.match(paras[x]['text']).group(1) for x in range(pi + 1, j) if MARK.match(paras[x]['text'])]
            circ = [CIRCLED.index(x) + 1 for x in marks if x in CIRCLED]
            if circ: got = max(circ)
        if got and got != want:
            add(pi, m.start(), m.group(0), '列举数量', '数量与列举项数', f'总述说“{m.group(0)}”，后文实际列出 {got} 项，请核对（编号也可能缺漏）')

# 9. 式下符号说明的版式：全书要一致（引出语“式中，”、接排或分行、“为/是/表示/——”），按体例表核对
FORMULA = re.compile(r'^\s*([（(]\s*\d+[-–.]\d+\s*[）)])?\s*$')        # 只剩公式编号或空（公式对象）的段
SYM_LINE = re.compile(r'^\s*(\[图片\]|[A-Za-zα-ωΑ-Ω]|—).{0,80}$')    # 以符号（或公式对象后的“—”）开头的短行（分行写的说明）
ex_groups = []
for pi, para in enumerate(paras):
    P, t = para['P'], para['text']
    if not in_range(P): continue
    m = re.match(r'^\s*(式中|其中)([，,：:]?)', t)
    if not m: continue
    prev = next((paras[j]['text'] for j in range(pi - 1, max(-1, pi - 3), -1) if paras[j]['text'].strip() or j == pi - 1), '')
    if not FORMULA.match(prev) and '\t' not in prev: continue          # 前一段不是公式，不是式下说明（如“其中，加权平均法……”）
    grp, j = [pi], pi + 1
    while j < len(paras) and SYM_LINE.match(paras[j]['text']) and ('—' in paras[j]['text'] or re.search(r'为|是|表示', paras[j]['text'])):
        grp.append(j); j += 1
    txt = '\n'.join(paras[x]['text'] for x in grp)
    feats = []
    if m.group(0).strip() != '式中，': feats.append(f'引出语作“{m.group(0).strip() or m.group(1)}”')
    if len(grp) > 1: feats.append(f'分{len(grp)}行')
    if '—' in txt: feats.append('用破折号')
    if re.search(r'[A-Za-zα-ωΑ-Ω′\'\]）)\d]\s*(是|表示)', txt): feats.append('动词用“是/表示”')
    ex_groups.append((P, feats))
    if feats:
        add(pi, m.start(1), m.group(0).strip() or m.group(1), '体例', '式下说明版式',
            '式下符号说明：' + '、'.join(feats) + '。按本书体例核对（如“式中，X为……；Y为……。”接排一段），全书须一致')

# 标题、图表题清单：这些行字少、常被略读，要单独逐字看
titles = []
for para in paras:
    P, t = para['P'], para['text'].strip()
    if not in_range(P) or not t: continue
    if (len(t) < 50 and not re.search('[。；，]', t) and re.match(r'^(第\s*\d+\s*章|\d+(\.\d+){0,3}\s|[图表]\s*\d)', t)):
        titles.append((P, t))

# 去重：同一段同一位置同一规则只留一条；已被确定错词覆盖的易混词不再列
seen, uniq = set(), []
sure_spans = [(c['para'], c['pos'], c['pos'] + len(c['old'])) for c in cands if c['type'] == '易错词']
for c in cands:
    key = (c['para'], c['around'], c['rule'], c['old'])
    if key in seen: continue
    if c['type'] == '易混词' and any(P == c['para'] and a < c['pos'] + len(c['old']) and c['pos'] < b for P, a, b in sure_spans): continue
    seen.add(key); uniq.append(c)
for i, c in enumerate(uniq, 1): c['id'] = f'T{i:03d}'

os.makedirs(outdir, exist_ok=True)
json.dump(uniq, open(os.path.join(outdir, 'typo-candidates.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump([dict(P=f'P{P}', text=t) for P, t in titles], open(os.path.join(outdir, 'titles.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
by = Counter(c['type'] for c in uniq)
with open(os.path.join(outdir, 'typo-report.md'), 'w', encoding='utf-8') as f:
    f.write(f'# 错别字与标点初筛候选\n\n稿件：`{os.path.basename(src)}`　范围：{p_from or "开头"}–{p_to or "结尾"}　词表：`{os.path.basename(words_md)}`\n\n')
    f.write('以下是程序候选，**须逐条看语境判断**，不能直接全部写入修订。\n\n')
    f.write(f'数值与单位符号之间：有空格 {unit_space[0]} 处，无空格 {unit_space[1]} 处（判断本书体例用）。\n\n')
    f.write(f'数字形式（3～10 接“个种类项”）：阿拉伯 {num_forms["阿拉伯"]} 处，汉字 {num_forms["汉字"]} 处；多位数：千分撇 {num_forms["千分撇"]} 处，五位以上不分节 {num_forms["五位以上不分节"]} 处（判断本书体例用）。\n\n')
    f.write(f'式下符号说明 {len(ex_groups)} 组：' + ('；'.join(f'P{P} ' + ('、'.join(fs) or '接排“为”') for P, fs in ex_groups) or '无') + '（全书须一致，判断本书体例用）。\n\n')
    f.write('| 类别 | 数量 |\n|---|---|\n' + ''.join(f'| {k} | {v} |\n' for k, v in by.items()) + '\n')
    f.write('## 标题与图表题（逐字看错别字、漏字）\n\n' + ''.join(f'- P{P}：{t}\n' for P, t in titles) + '\n')
    for typ in by:
        f.write(f'## {typ}\n\n| # | 段 | 原文 | 建议 | 提示 |\n|---|---|---|---|---|\n')
        for c in uniq:
            if c['type'] != typ: continue
            a, o = c['around'], c['off']
            ar = (a[:o] + '**' + a[o:o + len(c['old'])] + '**' + a[o + len(c['old']):]).replace('|', '\\|')
            f.write(f"| {c['id']} | {c['para']} | {ar} | {c['new'] or ''} | {c['hint']}{'' if c['ctx'] else '（未找到唯一 ctx）'} |\n")
        f.write('\n')
bad = sum(1 for c in uniq if not c['ctx'])
print(f'候选 {len(uniq)} 条：' + '，'.join(f'{k} {v}' for k, v in by.items()) + (f'；{bad} 条未找到唯一 ctx' if bad else '')
      + f'；数值与单位有空格 {unit_space[0]} 处、无空格 {unit_space[1]} 处 → {outdir}')

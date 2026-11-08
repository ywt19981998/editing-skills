# 在 docx 副本上写入真实修订（w:del/w:ins）与批注；用法: revise_docx.py 原稿 输出 规则.json
# 规则：edits（段内改字）、comments（批注）、inserts（插入新段落：after_ctx 定位锚点段，paragraphs 为 [{like_ctx 样板段, segments 文字分段}]）、scope_from（只在此文字所在段及之后定位）
import sys, zipfile, json, copy, datetime, re
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return f'{{{W}}}{t}'
src, dst, spec = sys.argv[1], sys.argv[2], json.load(open(sys.argv[3], encoding='utf-8'))
# 修订人随运行方自动区分：Codex 目录下运行显示 Codex，否则显示 Claude；可用环境变量 REVISE_AUTHOR 覆盖
import os
AUTHOR = os.environ.get('REVISE_AUTHOR') or ('Codex' if '/.codex/' in os.path.abspath(__file__) else 'Claude')
DATE = datetime.datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ')
zin = zipfile.ZipFile(src)
doc = etree.fromstring(zin.read('word/document.xml'))
rid = [9000]

# 规则可选键 scope_from：只在含该文字的段落及其后的段落里定位（如只改第6章时写章题），防止误改前面章节
_scope = spec.get('scope_from')
_start_el = None
if _scope:
    _start_el = next((p for p in doc.iter(q('p')) if _scope in ''.join(t.text or '' for t in p.iter(q('t')))), None)
    if _start_el is None: raise SystemExit(f'未找到范围起点: {_scope}')
def paras():
    """按文档顺序逐段返回（每次现取，插入的新段落也在内）；有 scope_from 时从起点段开始"""
    started = _start_el is None
    for p in doc.iter(q('p')):
        if not started and p is _start_el: started = True
        if started: yield p
def nid():
    rid[0] += 1; return str(rid[0])

def plain(r):
    """只含 rPr 和单个 w:t 的普通文字 run"""
    kids = [etree.QName(c).localname for c in r]
    return r.getparent().tag == q('p') and kids in (['t'], ['rPr', 't'])

def merge_runs(p):
    """合并段内相邻、格式完全相同的普通 run（Word 常因拼写检查等把同格式文字切碎）"""
    prev = None
    for r in list(p.findall(q('r'))):
        if prev is not None and plain(prev) and plain(r) and r.getprevious() is prev:
            a, b = prev.find(q('rPr')), r.find(q('rPr'))
            sa = etree.tostring(a) if a is not None else b''
            sb = etree.tostring(b) if b is not None else b''
            if sa == sb:
                prev.find(q('t')).text = (prev.find(q('t')).text or '') + (r.find(q('t')).text or '')
                prev.find(q('t')).set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                p.remove(r); continue
        prev = r

def isolate(ctx, target):
    """在含 ctx 的段落里，把 target（须落在单个 run 内）切成独立 run 并返回"""
    for p in paras():
        ptxt = ''.join(t.text or '' for t in p.iter(q('t')))
        if ctx not in ptxt: continue
        merge_runs(p)
        for r in p.iter(q('r')):
            ts = r.findall(q('t'))
            if len(ts) != 1 or target not in (ts[0].text or ''): continue
            text = ts[0].text
            # 用 ctx 在 run 内定位 target，避免同段重复
            base = text.find(ctx) if ctx in text else 0
            i = text.find(target, base)
            if i < 0: i = text.find(target)   # 目标在 ctx 之前时从头找
            if i < 0: continue
            parts = [text[:i], text[i:i+len(target)], text[i+len(target):]]
            parent, pos = r.getparent(), r.getparent().index(r)
            new = []
            for s in parts:
                if not s: new.append(None); continue
                nr = copy.deepcopy(r); t = nr.find(q('t')); t.text = s
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); new.append(nr)
            parent.remove(r)
            for k, nr in enumerate([x for x in new if x is not None]):
                parent.insert(pos + k, nr)
            return new[1]
    raise SystemExit(f'未找到: {ctx} / {target}')

def text_runs(p):
    """段落的直接子 run 中带单个 w:t 的文字 run（已删除/已插入的修订不计入）"""
    return [r for r in p.findall(q('r')) if len(r.findall(q('t'))) == 1]

def split_at(p, pos):
    """确保在段落文字的第 pos 个字符处是 run 边界"""
    acc = 0
    for r in text_runs(p):
        t = r.find(q('t')); n = len(t.text or '')
        if acc < pos < acc + n:
            k = pos - acc
            r2 = copy.deepcopy(r)
            t.text, r2.find(q('t')).text = t.text[:k], t.text[k:]
            for x in (t, r2.find(q('t'))): x.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.addnext(r2); return
        acc += n

def track(ctx, old, new, exact=False, author=None):
    au = author or AUTHOR
    """支持跨 run 的最小差异修订：去掉公共前后缀后，删除中段并插入新文字"""
    a = 0
    while a < min(len(old), len(new)) and old[a] == new[a]: a += 1
    b = 0
    while b < min(len(old), len(new)) - a and old[-1-b] == new[-1-b]: b += 1
    n_mid = new[a:len(new)-b]
    for p in paras():
        full = ''.join(r.find(q('t')).text or '' for r in text_runs(p))
        ci = full.find(ctx)
        if ci < 0 or (exact and full != ctx): continue   # exact：整段（单元格）文字须与 ctx 完全相同
        i = full.find(old, ci)
        if i < 0 or i > ci + len(ctx): i = full.find(old)
        if i < 0: continue
        s0, e0 = i + a, i + len(old) - b
        merge_runs(p); split_at(p, s0); split_at(p, e0)
        acc, dels, anchor, prev = 0, [], None, None
        for r in text_runs(p):
            n = len(r.find(q('t')).text or '')
            if n and s0 <= acc and acc + n <= e0: dels.append(r)
            if n and anchor is None and acc == s0: anchor = r      # 插入点之后的第一个 run
            if n and acc + n == s0: prev = r                        # 插入点之前的 run
            acc += n
        base = copy.deepcopy(dels[0] if dels else (prev if prev is not None else anchor))
        for r in dels:
            d = etree.Element(q('del')); d.set(q('id'), nid()); d.set(q('author'), au); d.set(q('date'), DATE)
            r.addprevious(d); d.append(r); r.find(q('t')).tag = q('delText')
        if n_mid:
            ins = etree.Element(q('ins')); ins.set(q('id'), nid()); ins.set(q('author'), au); ins.set(q('date'), DATE)
            base.find(q('t')).text = n_mid
            base.find(q('t')).set('{http://www.w3.org/XML/1998/namespace}space', 'preserve'); ins.append(base)
            if dels: dels[-1].getparent().addnext(ins)
            elif prev is not None: prev.addnext(ins)
            else: anchor.addprevious(ins)
        return
    raise SystemExit(f'未找到: {ctx} / {old}')

def span_runs(ctx, target):
    """target 跨多个格式不同的 run 时，返回覆盖它的首尾 run（批注范围略宽于目标）"""
    for p in paras():
        ptxt = ''.join(t.text or '' for t in p.iter(q('t')))
        if ctx not in ptxt: continue
        # 含修订插入的文字（w:ins 里的 run），不含已删除的文字
        runs = [r for r in p.iter(q('r')) if r.find(q('t')) is not None
                and not any(etree.QName(a).localname == 'del' for a in r.iterancestors())]
        full = ''.join(r.find(q('t')).text or '' for r in runs)
        i = full.find(target, max(full.find(ctx), 0))
        if i < 0: i = full.find(target)
        if i < 0: continue
        pos, first, last = 0, None, None
        for r in runs:
            n = len(r.find(q('t')).text or '')
            if first is None and pos + n > i: first = r
            if pos < i + len(target) <= pos + n: last = r; break
            pos += n
        if first is not None and last is not None:
            def top(r):   # 取段落的直接子元素（插入的 run 取外层的 w:ins）
                while r.getparent() is not p: r = r.getparent()
                return r
            if first.getparent() is last.getparent(): return first, last
            return top(first), top(last)
    raise SystemExit(f'未找到: {ctx} / {target}')

old_cx = etree.fromstring(zin.read('word/comments.xml')) if 'word/comments.xml' in zin.namelist() else None
cid_base = (max([int(c.get(q('id'))) for c in old_cx.findall(q('comment'))] or [-1]) + 1) if old_cx is not None else 0
comments = []
def comment(ctx, target, text):
    try:
        first = last = isolate(ctx, target)
    except SystemExit:
        first, last = span_runs(ctx, target)
    cid = str(cid_base + len(comments))
    parent = first.getparent()
    s = etree.Element(q('commentRangeStart')); s.set(q('id'), cid)
    e = etree.Element(q('commentRangeEnd')); e.set(q('id'), cid)
    ref = etree.Element(q('r')); cr = etree.SubElement(ref, q('commentReference')); cr.set(q('id'), cid)
    parent.insert(parent.index(first), s)
    parent.insert(parent.index(last) + 1, e); parent.insert(parent.index(e) + 1, ref)
    comments.append((cid, text))

def find_para(ctx):
    for p in paras():
        if ctx in ''.join(t.text or '' for t in p.iter(q('t'))): return p
    raise SystemExit(f'未找到段落: {ctx}')

def insert_paras(after_ctx, paras, author=None):
    """在 after_ctx 段之后以修订方式插入新段落；每段仿照 like_ctx 段的段落格式和文字格式。
    segments 与样板段的文字 run 一一对应（如节标题“3.3”“  ”“本章小结”），多出的段用最后一个 run 的格式。
    样板段里的图片（如节号底图）不复制；锚点段带分节符时，分节符移到最后一个新段上。"""
    au = author or AUTHOR
    anchor = find_para(after_ctx)
    last = anchor
    for spec_p in paras:
        like = find_para(spec_p['like_ctx'])
        np_ = etree.Element(q('p'))
        ppr = copy.deepcopy(like.find(q('pPr'))) if like.find(q('pPr')) is not None else etree.Element(q('pPr'))
        for x in ppr.findall(q('sectPr')): ppr.remove(x)
        prpr = ppr.find(q('rPr'))
        if prpr is None: prpr = etree.SubElement(ppr, q('rPr'))
        for x in list(prpr):
            if etree.QName(x).localname in ('ins', 'del'): prpr.remove(x)
        mark = etree.Element(q('ins')); mark.set(q('id'), nid()); mark.set(q('author'), au); mark.set(q('date'), DATE)
        prpr.insert(0, mark)                                   # 段落标记本身也标为插入
        np_.append(ppr)
        tmpl = [r for r in like.findall(q('r')) if r.find(q('t')) is not None] or [etree.Element(q('r'))]
        ins = etree.SubElement(np_, q('ins')); ins.set(q('id'), nid()); ins.set(q('author'), au); ins.set(q('date'), DATE)
        # 样板段开头的图片 run（如节号底图，节号白字压在上面）一并复制，图片编号改为不重复
        for r in like.findall(q('r')):
            if r.find(q('t')) is not None: break
            if r.find(q('drawing')) is not None:
                dr = copy.deepcopy(r)
                for el in dr.iter():
                    if etree.QName(el).localname == 'docPr': el.set('id', str(60000 + int(nid())))
                ins.append(dr)
        for k, seg in enumerate(spec_p['segments']):
            r = etree.SubElement(ins, q('r'))
            rpr = tmpl[min(k, len(tmpl) - 1)].find(q('rPr'))
            if rpr is not None: r.append(copy.deepcopy(rpr))
            t = etree.SubElement(r, q('t')); t.text = seg
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        last.addnext(np_); last = np_
    sect = anchor.find(q('pPr') + '/' + q('sectPr')) if anchor.find(q('pPr')) is not None else None
    if sect is not None and last is not anchor:
        anchor.find(q('pPr')).remove(sect); last.find(q('pPr')).append(sect)

# 批注先按原文定位；定位不到的（目标在插入的段落或修订插入的文字里）等改字之后再定位一次
_later = []
for c in spec.get('comments', []):
    try: comment(c['ctx'], c['target'], c['text'])
    except SystemExit: _later.append(c)
for s in spec.get('inserts', []): insert_paras(s['after_ctx'], s['paragraphs'], AUTHOR + '（建议）' if s.get('decision') else None)
for e in spec.get('edits', []): track(e['ctx'], e['old'], e['new'], e.get('exact', False), AUTHOR + '（建议）' if e.get('decision') else None)
for c in _later: comment(c['ctx'], c['target'], c['text'])

# 写 comments.xml 及关系、内容类型
cx = old_cx if old_cx is not None else etree.Element(q('comments'), nsmap={'w': W})
for cid, text in comments:
    c = etree.SubElement(cx, q('comment')); c.set(q('id'), cid); c.set(q('author'), AUTHOR); c.set(q('date'), DATE); c.set(q('initials'), 'C')
    for line in text.split('\n'):
        p = etree.SubElement(c, q('p')); r = etree.SubElement(p, q('r')); t = etree.SubElement(r, q('t')); t.text = line
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
rels = zin.read('word/_rels/document.xml.rels').decode('utf-8')
ct = zin.read('[Content_Types].xml').decode('utf-8')
if comments and 'comments.xml' not in rels:
    rels = rels.replace('</Relationships>', '<Relationship Id="rIdClaudeComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>')
    ct = ct.replace('</Types>', '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>')
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    if item.filename == 'word/document.xml': data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    elif item.filename == 'word/comments.xml' and comments: data = etree.tostring(cx, xml_declaration=True, encoding='UTF-8', standalone=True)
    elif item.filename == 'word/_rels/document.xml.rels': data = rels.encode('utf-8')
    elif item.filename == '[Content_Types].xml': data = ct.encode('utf-8')
    else: data = zin.read(item.filename)
    zout.writestr(item, data)
if comments and old_cx is None: zout.writestr('word/comments.xml', etree.tostring(cx, xml_declaration=True, encoding='UTF-8', standalone=True))
zout.close()
print(f'修订 {len(spec.get("edits", []))} 处，插入段落 {sum(len(x["paragraphs"]) for x in spec.get("inserts", []))} 段，批注 {len(comments)} 条 → {dst}')

# 这一章审稿用哪份稿当底稿：润色过的章用“润色后底稿”，没润色的用作者原稿。准备、机械规则、合并三个脚本共用。
import os

def chapter_src(cfg, book, n):
    """返回 (底稿, 作者原稿, 润色目录)。润色目录为 None 表示这一章没有润色过。
    润色产物按润色技能 runse 的约定放在 <书目录>/润色/第NN章/（可用配置“润色目录”改），
    里面有 第NN章-润色后底稿.docx、正文.txt、终稿.json 三样才算润色过。"""
    NN = f'{int(n):02d}'
    fill = lambda s: s.replace('{n}', str(int(n))).replace('{NN}', NN)
    ov = (cfg.get('原稿覆盖') or {}).get(NN) or (cfg.get('原稿覆盖') or {}).get(str(int(n)))
    author = fill(cfg['原稿'])
    rd = os.path.join(book, fill(cfg.get('润色目录', '润色/第{NN}章')))
    base = os.path.join(rd, f'第{NN}章-润色后底稿.docx')
    if all(os.path.exists(p) for p in (base, os.path.join(rd, '正文.txt'), os.path.join(rd, '终稿.json'))):
        if ov and os.path.abspath(ov) != os.path.abspath(base): author = ov      # 覆盖里写的是另一份作者原稿
        return base, author, rd
    return (ov or author), (ov or author), None

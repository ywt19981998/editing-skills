#!/usr/bin/env python3
"""把旧格式的总名词表瘦身成“只有名词”的新格式（一次性迁移用，原文件不动）。

用法: slim_glossary.py <旧 名词表.md> -o <新 名词表.md>

保留: 标题和稿本说明、“本书名词类别清单”、“一、采用名称”的 # / 所指对象 / 采用名称 / 适用范围 / 状态 五列、
      “二、非同一对象，不得合并”。
去掉: 各章范围、依据原文及位置、已见变体及位置、目录核对附记、待用户确认（待定的事只记在 待定决策.md）。
瘦身后人工过一遍：采用名称栏里夹带的过程说明（谁哪天指出、第几章改了几处）可以再删。
"""
import sys, re


def main():
    args = sys.argv[1:]
    if len(args) != 3 or args[1] != '-o':
        print(__doc__); sys.exit(2)
    lines = open(args[0], encoding='utf-8').read().splitlines()
    out, section, kept, n = [], None, 0, 0
    for line in lines:
        if line.startswith('## '):
            section = line
            if '类别清单' in line:
                out += ['', '## 〇、本书名词类别清单（防整类漏收的检查单，不是收录范围）']
            elif '采用名称' in line:
                out += ['', '## 一、采用名称', '', '| # | 所指对象 | 采用名称 | 适用范围 | 状态 |', '|---|---|---|---|---|']
            elif '非同一对象' in line:
                out += ['', '## 二、非同一对象，不得合并']
            continue
        if section is None:                       # 文件头
            if line.strip() and not line.startswith('最近更新'):
                out.append(line)
            continue
        if '类别清单' in section or '非同一对象' in section:
            if line.strip() or (out and out[-1].strip()):
                out.append(line)
        elif '采用名称' in section and re.match(r'\|\s*T\d+', line):
            c = [x.strip() for x in line.strip().strip('|').split('|')]
            n += 1
            if len(c) == 7:
                out.append(f'| {c[0]} | {c[1]} | {c[2]} | {c[4]} | {c[6]} |'); kept += 1
            else:
                out.append(line)                  # 列数不对的原样留下，人工处理
    text = '\n'.join(out).rstrip() + '\n'
    open(args[2], 'w', encoding='utf-8').write(text)
    print(f'条目 {n} 条（按五列改写 {kept} 条）；{sum(len(l) for l in lines)} 字 → {len(text)} 字 → {args[2]}')


if __name__ == '__main__':
    main()

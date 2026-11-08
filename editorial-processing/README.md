# 编辑加工辅助脚本

独立于一致性审查、公式审校两个技能的辅助脚本，不修改那两个技能。都用共用环境 `~/.venvs/book-review/bin/python` 运行。

## 编号与引用检查

已移到 `~/.claude/skills/tongdu/scripts/scan_numbering.py 书稿.docx 输出目录`

查图号、表号、公式号的跳号、重号、顺序颠倒、章号不符，图表未引用、先见图后见文，引用的编号不存在，图表编号与首次引用顺序不一致，编号写法不统一。结果是候选问题（`numbering-issues.json` 带 ctx，可转成 revise_docx 规则），须逐条判断后再写入修订。尚未用真实书稿校准。

## 决策单与按决策重新生成

适用于一致性审查的“有倾向的待确认”（`待定决策.md` + 修订规则里带 `"decision": "D-xx"` 的 edit）。

1. **出决策单**：`scripts/decision_sheet.py 待定决策.md 决策单.md 第02章-修订规则.json 第03章-修订规则.json ...`
   按决策列出 A（建议，已写入修订）涉及几处、举例，B 是什么，状态。用户回复“D-12 选B”“其余按建议”。
2. **写决策选择**：`决策选择.json`
   ```json
   {"D-16": {"choice": "A"},
    "D-12": {"choice": "B", "edits": [{"chapter": "第02章", "ctx": "...", "old": "经纬网格", "new": "经纬度网格"}]}}
   ```
   B 方案的修订由 Claude 按 `待定决策.md` 的“另一个选项”现写，用 `scripts/find_term.py 底稿.docx 词1 词2 [--exclude ...]` 找出全部位置和唯一 ctx（P 号与 dump_docx.py 一致）。
3. **重新生成**：`scripts/apply_decisions.py 底稿.docx 第03章-修订规则.json 决策选择.json 输出.docx [--chapter 第03章]`
   - 选 A：该决策的修订去掉 decision 标记，修订人由“Claude（建议）”变为“Claude”；删去【建议·D-xx】批注。
   - 选 B：撤回该决策的修订和【建议·D-xx】批注，写入 B 方案中属于本章的修订。
   - 未拍板的保持原样。先预检（ctx 找不到、修改重叠都会停下，不出文件），合成规则存为 `输出.rules.json`，再调用一致性技能的 `revise_docx.py`。
   - 底稿必须是当初生成该章修订稿所用的底稿（如第3章：`原稿对照修正/第3章-v3-已接受还原.docx`）。不传决策时，重新生成的稿与原修订稿逐字一致（第3章已验证）。
4. 拍板后更新 `待定决策.md` 的状态列。

**限制**：重新生成会覆盖用户在旧修订稿上的手工处理，所以要先拍板再通读。通读中途改选、在已有修订痕迹的稿子上补改，还没有做。

## 通读（已移出）

通读初筛已做成独立技能 `~/.claude/skills/tongdu/`（SKILL.md、scan_typos.py、scan_numbering.py、check_rules.py、make_rules.py、易错词表、通读提示词、体例模板都在那里）。本目录只保留一致性流程的辅助工具：决策单、按决策重新生成、定位工具。

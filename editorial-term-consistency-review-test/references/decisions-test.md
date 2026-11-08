# Decisions 对比试验（test 版专有）

目的：同一批“只差一个字”的变体候选，分别交 OpenAI Decisions（`gpt-6-luna`，只答是否并给概率）和大模型（Claude 子代理）判断“是否同一对象”，比较结论和用时，决定以后能否让 Decisions 先筛。

**试验不影响审查本身**：本章审查子代理照正式流程做完、主代理收完摘要后再做试验；试验结果不写进修订规则、名词表或待定决策，只出对比报告。

脚本：`~/.venvs/book-review/bin/python ~/.claude/skills/editorial-term-consistency-review-test/scripts/decisions_compare.py`（下文记作 `DC`），试验目录默认 `<工作目录>/Decisions试验/第NN章/`。

## 步骤

1. **出候选**：
   `~/.venvs/book-review/bin/python ~/.claude/skills/editorial-term-consistency-review-test/scripts/variant_pairs.py source/第NN章-docx段落.txt [图表文字.txt] <试验目录>/变体候选.md`
2. **准备材料**：`DC prepare source/第NN章-docx段落.txt <试验目录>/变体候选.md <试验目录>`。两边看到的是同一份 `candidates.md`（每对两端各最多 2 处原句和所在小节），保证比较公平。
3. **发送前确认**：先跑 `DC decisions <试验目录>`（不加 `--confirm-live` 不发送），把候选对数、请求次数、材料字数告诉用户，说明会发给 OpenAI 并按 API 计费；用户确认后才加 `--confirm-live` 发送。密钥从登录钥匙串 `bianjia-decisions` 读，不经手、不打印、不落盘；钥匙串里没有时请用户本人在终端运行 `~/.codex/skills/bianjia-decisions/scripts/bianjia_decisions.py --setup-keychain`，不要把密钥贴进聊天。
4. **大模型判断**：
   - `DC stamp <试验目录> llm-start`
   - 派一个 `general-purpose` 子代理（干净上下文，`run_in_background: false`），提示词用下文模板；
   - 子代理返回后 `DC stamp <试验目录> llm-end`。
5. **对比**：`DC compare <试验目录>`，写出 `对比报告.md`：一致率、结论分布表、分歧逐条、两边用时。
6. **向用户报告**：一致率；**漏放风险**条数（大模型判同一对象、Decisions 判不同对象——若以后让 Decisions 直接放行“不同对象”，这些会漏改）；两边用时；Decisions 的“拿不准”有多少（太多说明材料不够，参见公式审校 203 组全判“待复核”的教训）。分歧较多时，抽几条回原文看谁对，写进报告末尾。

## 用时口径

- Decisions：脚本记录的请求往返时间（不含准备材料）。
- 大模型：从派子代理到返回的整段时间（含读文件、思考、写文件）。子代理返回的 token 用量，主代理补进报告。
- 两者都不含本章精读建名词表的时间。这一项两条路线都省不掉。

## 大模型判断子代理提示词模板

```
你负责判断一批书稿“写法配对”是否指同一对象，只做判断，不改任何书稿或其他文件。全程用简体中文。

1. 完整读取 <试验目录>/candidates.md。每对是脚本找出的“只差一个字”的两种写法，附两端原句和所在小节。
2. 只根据这份材料判断（不读书稿其他部分，与另一个判断模型看到的材料相同），每对给一个结论：
   - 同一对象：同一术语、方法、模型或名称的两种写法，应统一成一种；
   - 不同对象：不同概念、上下位概念、函数与数值、有意对比的两个对象，或只是普通叙述用词、不是名称；
   - 拿不准：材料不足以判断。
   判断原则同一致性审查：完整名称逐字比较；先辨对象再定是否统一；保留有意义的限定词；普通叙述的合理改写不算术语漂移。
3. 把结论写进 <试验目录>/llm-judgments.json（已有模板，键是编号），每项填 "label"（只能是 同一对象／不同对象／拿不准）和 "reason"（一句话）。每一对都要填。
4. 返回：三种结论各几对、最拿不准的 3 对编号。不超过 8 行。
```

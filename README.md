# 编辑加工技能备份

这是 Claude Code 的个人技能目录（`~/.claude/skills`）的备份，私有仓库。

## 内容

| 技能 | 作用 |
|---|---|
| `editorial-term-consistency-review` | 书稿一致性审查（正式版）：逐章精读建名词表，查名称、数量、结构、数据、图文的前后不一致。2026-11-08 改版；改版前的旧版可从 git 提交 `e593250` 取回（该提交里的 `legacy/` 目录） |
| `editorial-term-consistency-review-test` | 一致性审查 test 版（Decisions 对比试验），平时不用 |
| `runse` | 润色：拿到稿子后的第一遍。以段为单位把句子改通顺、标点用对，公式、数值、引注不动；脚本核对加独立复查后写成 Word 修订稿，并生成润色后底稿给后两遍用（2026-11-12 建） |
| `tongdu` | 通读：在润色后的稿子上做。错别字、数字和单位、英文拼写、图表编号直接改成修订，知识性和数据疑点只批注；句子和标点 2026-11-12 起移到润色 |
| `formula-proofread` | 公式审校：字形（正斜体、黑白体）与数学正确性 |
| `editorial-processing` | 编辑加工辅助脚本 |

## 恢复

换电脑或误删后，一条命令恢复（目录不存在时）：

```bash
git clone https://github.com/ywt19981998/editing-skills.git ~/.claude/skills
```

目录已存在时，先把旧目录移走再克隆，或在目录内 `git pull`。

## 需要注意

- 脚本需要 `lxml`：系统默认的 `python3` 没有，用 `/usr/bin/python3` 或 `~/.venvs/book-review/bin/python`。
- 改完技能后检查每个 `SKILL.md` 的首行是 `---`。
- 每次修改提交一次，再 `git push` 同步到这里。
- 密钥不在仓库里：OpenAI、Anthropic 的密钥放在 macOS 钥匙串（服务名 `bianjia-openai-decisions`、`bianjia-anthropic-api`）。
- Codex 那份副本在 `~/.codex/skills/`，不在本仓库。

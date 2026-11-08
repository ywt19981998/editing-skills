# 2026-11-07 旧版备份（2026-11-08 改版前）

2026-11-08 改版内容：判断栏三选一、名词类别清单、脚本候选前置（term_candidates.py、variant_pairs.py）、门禁 check_glossary.py、机械替换脱离子代理、两张专表、压缩复核、建议修订门槛。

**切回旧版**（三个文件覆盖回去即可，scripts/ 里新增的三个脚本不影响旧流程）：

```bash
cd ~/.claude/skills/editorial-term-consistency-review && cp legacy/2026-11-07旧版/SKILL-旧版.md SKILL.md && cp legacy/2026-11-07旧版/templates-旧版.md references/templates.md && cp legacy/2026-11-07旧版/chapter-subagent-prompt-旧版.md references/chapter-subagent-prompt.md
```

技能目录（~/.claude/skills）同日纳入 git，之后每次改版 `git log` 可查、`git checkout <commit> -- editorial-term-consistency-review/` 可回。

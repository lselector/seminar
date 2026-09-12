# Skill names use hyphens, never underscores

Skill names — both the directory under `.claude/skills/<name>/` and the `name:` field in that skill's `SKILL.md` frontmatter — must use **hyphens** for word separation, never underscores. 

So `/mycron-apply`, not `/mycron_apply`. 

This rule applies **only to skill names**, not to underlying Python tool files. Python modules can't have hyphens, so a skill `/mycron-apply` legitimately maps to `tools/mycron_apply.py`.
Don't try to rename Python files to use hyphens.

If you spot an existing skill that violates this rule, flag it for renaming rather than silently leaving it.

# Claude Code Instructions

@AGENTS.md

## Claude Code specific rules

- Treat `AGENTS.md`, current code, current branch, and Git history as the source of truth.
- Do not depend on restoring a previous Claude session to continue work.
- Read only files relevant to the current task before broad repository scans.
- Use `.claude/agents/` subagents only when isolated context clearly helps; do not create or call one for routine edits.
- Use `.claude/skills/` skills only for genuinely repeated procedures; do not invoke a skill merely because it exists.
- Prefer local files, Git, and terminal/compiler work before external MCP calls.
- Do not enumerate or probe all configured MCP servers. Use only the MCP required by the current task.
- Do not create persistent progress diaries. Update `AGENTS.md` Current State instead, and rely on Git history for old work.
- Never store API keys, tokens, passwords, or other secrets in repository files.
- Keep commits meaningful rather than creating a commit for every tiny edit.
- Preserve the copyright/trademark/font rules in `AGENTS.md` and `docs/LEGAL_ASSET_POLICY.md`.

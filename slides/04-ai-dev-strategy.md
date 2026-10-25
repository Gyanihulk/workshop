---
slide: AI development strategy (2 slides, optional 3rd)
position: Part 3 (Build a real project), after "How to pick a project you'll actually finish"
status: draft, not built
rule: no token-saving figure on a slide unless you measured it live or it is cited on the slide
---

# Research notes (2026-10-25)

Mostly third-party sources. Verify anything marked CHECK against the vendor's own docs before presenting.

## Graphify (answers the open question in HANDOFF.md)
- Open-source tool and Claude Code skill that builds a queryable knowledge graph of a codebase, so the agent
  queries structure instead of re-reading files. Ships a skill plus a PreToolUse hook that points Claude to
  GRAPH_REPORT.md before Grep/Glob.
- Reported savings: ~1.7k tokens per query vs ~123k reading raw files (71.5x) on a 52-file mixed corpus; other
  posts say 6.8x (code review) and up to 49x (daily tasks). These are the maintainer's / blogs' numbers.
- Caveats: small projects may not repay the overhead; a competing tool (code-review-graph) reports 100-350 tokens
  for targeted queries vs ~1,500 for graphify BFS. => Measure on your own repo (live demo idea below).
- Sources: mindstudio.ai graphify article, dev.to "Graphify + code-review-graph", flowtivity.ai, graaph.org.

## Claude Code token habits (all from third-party guides, CHECK against Anthropic docs)
- Clear context between unrelated tasks (/clear); compact manually before auto-compaction (/compact).
  Thresholds quoted by guides conflict (40% left vs 93-95% full): do not put a number on a slide.
- Keep CLAUDE.md short: it loads every session. Path-scoped rules load only when matching files are edited.
- Skills: only name + description are loaded at start; the body loads when relevant.
- Subagents keep noisy work out of the main thread, but cost more on small tasks.
- Plan mode before non-trivial work. Filter tool output. Deny/ignore big folders (permissions.deny).
- Watch usage with /context and /usage (or a status line).
- Sources: analyticsvidhya.com (23 tips), kdnuggets.com (7 ways), composio.dev (9 ways), computingforgeeks.com.

## Codex (OpenAI)
- Runs as a VS Code extension (also forks like Cursor/Windsurf), a CLI, a desktop app and a cloud agent in ChatGPT;
  same account/config/AGENTS.md across surfaces. Cloud tasks run in isolated sandboxes preloaded with the repo.
- AGENTS.md is the repo-level instruction file (Claude's equivalent is CLAUDE.md).
- CHECK: which ChatGPT plans include Codex (sources disagree).

## Google / GCP
- Gemini CLI and Gemini Code Assist IDE extensions stop serving individual and Google AI Pro/Ultra users on
  2026-06-18 (Google's own FAQ); Standard/Enterprise keep access; individuals are pointed to Antigravity CLI.
- Gemini Developer API still has a free tier via Google AI Studio (CHECK limits).
- Google Cloud: $300 welcome credit for 90 days, plus an Always Free tier of 20+ products (official docs).
  If the trial is not upgraded the billing account closes and resources stop.
- CHECK: Cloud Run / Cloud Functions / Firestore free limits, any 2026 student offer (none confirmed).

# Draft slide A: Pick your AI coding setup
Four cards: Claude Code (skills, CLAUDE.md, subagents) | Codex (VS Code extension, AGENTS.md, cloud tasks) |
Google (AI Studio free API tier, GCP trial credits for deploying) | Graphify (graph of your repo as a skill).
Each card: what it is, pick it when, one catch.

# Draft slide B: Get more work out of every token
1. Clear between tasks. 2. Keep CLAUDE.md / AGENTS.md short. 3. Put repeat instructions in a skill.
4. Plan before you build. 5. Give noisy work to a subagent. 6. Give the agent a map, then measure.
Callout: "Measure it yourself: same question with and without the map, compare what the agent read."

# Optional slide C: Ship it on the cloud
build -> container -> deploy -> live URL. Free credits first, upgrade only when you have users.

# Live demo (audience-friendly, no unverified numbers)
Ask the same question about one repo twice, once without and once with a graph; show /context both times.

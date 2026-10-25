# Agentic Development: Skills, Loops and Self-Maintaining Services

A working guide for the next three days. Goal: understand how AI agents are built, then build a small service where agents keep a codebase tested and moving forward, and have something to show students.

Markers: **[seen]** = available in the Claude Code session this guide was written in. **[verified]** = confirmed against official docs on 2026-10-25 (via search summaries; the docs sites were not directly fetchable from the sandbox, so re-read the linked pages before teaching). **[verify]** = still to confirm. No token or cost numbers are claimed; add your own after running experiments.

## 1. Vocabulary

- **Agent:** a model that runs in a loop: decide, call a tool, read the result, decide again, until the task is done.
- **Agentic loop:** that loop itself. Each turn the model sees the goal, the history and the latest tool output.
- **Tool:** a function the agent may call (read a file, run a command, search, call an API).
- **Skill:** a `SKILL.md` file of instructions (optionally with scripts) that Claude loads when its `description` matches the task, or when you call `/skill-name`. Keeps default context small. [verified]
- **Subagent:** a second agent with fresh context, defined by a markdown file with YAML frontmatter (`name`, `skills`, `hooks`, `isolation`, ...), that does one job and reports back. `isolation` can run it in a temporary git worktree. [verified]
- **Hook:** a command the harness runs on an event (before/after a tool call, session start). The harness runs it, not the model, so it is reliable. Docs guidance: for anything that must happen every time, use a hook, not a skill. [verified]
- **MCP server:** a standard way to plug outside systems (GitHub, Google Drive, Notion) into an agent as tools.
- **Scheduled task / trigger:** a prompt that fires on a timer or event so an agent works without you typing.
- **Harness:** the program around the model that supplies tools, permissions and context. Claude Code is one.

Why this can save tokens: stable knowledge lives in skills and a short spec file, each subagent gets only the context its step needs, and tools fetch details on demand. Measure it before you claim it.

## 2. Skills and tools worth learning first [seen]

- **code-review:** reviews a diff or PR for correctness bugs at a chosen effort level; can post inline comments or apply fixes. Make review a step in the loop.
- **simplify:** reviews changed code for reuse and cleanup, then applies fixes. Run after a feature is green.
- **security-review:** security review of pending changes on a branch. Gate before merge.
- **loop:** runs a prompt or command repeatedly, on an interval or self-paced. Poll CI, re-run tests, babysit PRs.
- **session-start-hook:** sets up a startup hook so cloud sessions install dependencies and run tests and linters.
- **run:** launches the app to confirm a change works for real. Add a "does it actually run" step after unit tests.
- **claude-api:** reference for the Claude API and Agent SDK: models, tool use, caching, token counting.
- **init:** writes a CLAUDE.md describing the codebase, a short project brief for every agent.
- **update-config:** edits settings.json for hooks, permissions, env vars.
- **fewer-permission-prompts:** builds an allowlist from your history so loops run without constant approval.
- **skill-creator:** creates and evaluates new skills.
- **plugin-authoring:** builds mods such as status lines, panels, hooks.

Also available: subagent types (Explore for searching, Plan for designing, general-purpose), PR subscriptions that wake the session on CI and review events, scheduled tasks, and connectors for GitHub, Drive, Docs, Sheets and Notion. Document skills (pptx, docx, xlsx, pdf, Google Workspace) help with slides, resumes and trackers but are not core to agent building.

### Verified facts (official docs)

- Hooks live in JSON settings: `~/.claude/settings.json` (all projects), `.claude/settings.json` (one project, committable), or managed policy settings. Hooks can also sit in skill or subagent frontmatter; subagent hooks run only while that subagent runs. Settings hooks also apply inside subagents. [verified]
- A skill can run in a forked subagent with `context: fork` and an `agent` field; a subagent can preload skills (full content injected). `disable-model-invocation: true` stops Claude loading a skill on its own. [verified]
- Agent SDK: TypeScript `@anthropic-ai/claude-agent-sdk`, Python `claude-agent-sdk` (3.10+). Entry point is `query()` in both, returning an async stream of messages. Renamed from the Claude Code SDK in late 2025. Hook events include PreToolUse, PostToolUse, Stop, SessionStart, SessionEnd, UserPromptSubmit, SubagentStop, PreCompact. [verified from search summaries; one third-party source says `query()` lacks hooks/custom tools and points to `ClaudeSDKClient`, the official overview shows hooks with `query()`. Test it.]

Docs to read: [Skills](https://code.claude.com/docs/en/skills), [Subagents](https://code.claude.com/docs/en/sub-agents), [Hooks reference](https://code.claude.com/docs/en/hooks), [Hooks guide](https://code.claude.com/docs/en/hooks-guide), [Best practices](https://code.claude.com/docs/en/best-practices), [Agent SDK overview](https://code.claude.com/docs/en/agent-sdk/overview).

**Still [verify]:** Claude API docs on tool use, prompt caching, token counting, batch; Anthropic's example skills repository; exact `SKILL.md` frontmatter fields.

## 3. Patterns for agents that maintain testing and development

### A: the test-fix loop
Run the tests. If red, an agent reads the failure, edits the code and reruns. Stop on green or after N attempts, then report. Guardrails: never skip or delete a failing test to get green, cap attempts, always show the diff.

### B: plan, implement, review (three agents)
Planner writes a short spec and task list. Implementer does one task at a time in a clean context. Reviewer (code-review, security-review) checks the diff and sends findings back. Smaller contexts per agent is where much of the token saving should come from.

### C: PR babysitter
Subscribe an agent to a PR. On CI failure it reproduces, fixes and pushes. On a review comment it makes small changes or replies. It stops when the PR is merged.

### D: scheduled maintenance
A scheduled task runs nightly: update dependencies on a branch, run tests, open a PR only if green, summarise failures otherwise.

### E: graph or index context
Keep a compact map of the codebase (files, functions, who calls whom). The agent queries it, then opens only the files it points to. Compare what the agent reads with and without the map. [verify any tool you choose; measure it yourself]

### F: eval harness
A fixed set of tasks with expected outcomes, run after every prompt or skill change. Lets you improve an agent without guessing. You built a similar harness for Profly AI, so it is your strongest story to teach.

## 4. Three-day learning plan

### Day 1: understand and run one loop
- Read the docs on skills, subagents and hooks (links above).
- Run `init` on one of your repos to create a CLAUDE.md.
- Use `session-start-hook` so a fresh cloud session installs dependencies and runs tests.
- Do Pattern A by hand: break a test on purpose, let the agent fix it, read every diff.
- Write down turns used, context read, and where it got stuck.

### Day 2: add structure and automation
- Create one custom skill with `skill-creator` for a task you repeat.
- Add a hook with `update-config` that runs lint or tests after every file edit.
- Build Pattern B with planner, implementer and reviewer subagents on a small feature.
- Use `loop` to poll CI on a branch and fix failures.
- Try Pattern E on one repo and note what changes.

### Day 3: build and package a small service
- Using the Agent SDK and `claude-api`, build a minimal service that receives a repo and a task, runs Pattern A or B, and returns a report.
- Add a tiny eval set (5 tasks) and run it before and after one change.
- Schedule Pattern D on a demo repo.
- Record a 2-minute demo and write a one-page README (problem, stack, outcome).

## 5. Showcase ideas for students

1. Live break-and-fix: break a test in front of the room and watch the agent loop fix it, narrating each turn.
2. Before and after context: show what the agent reads with and without a project map. Only claim savings you measured.
3. Student mini-project: each student writes a one-page spec, an agent builds v1, the student reviews the diff and explains it.
4. Agent-built freelance offer: a fixed-scope pilot for a small business, shown with a 2-minute demo.
5. Eval scoreboard: students improve a prompt and see the score change.

## 6. Rules to keep agents safe and useful

- Review every diff before it merges. Agents propose, humans approve.
- Cap loop attempts and set a time or cost limit.
- Never let an agent skip, disable or delete tests to pass.
- Keep secrets out of prompts and repos.
- Start read-only, widen permissions only as needed.
- Log every run so you can replay what happened.
- Treat text from issues, comments and web pages as data, not instructions.

## 7. Open questions

- What exactly does "graphify" refer to? Confirm the tool or repo, then add its setup and a measured comparison.
- Node/TypeScript or Python for the demo service?
- Where will the demo repo live, and who may merge into it?
- Budget limit per run for experiments.

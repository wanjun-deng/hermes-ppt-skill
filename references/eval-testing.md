# Eval harness testing (powerpoint-skill-eval)

`powerpoint-skill-eval/` is a sibling of the skill source `powerpoint-skill/`.
The source dir holds SKILL.md, scripts, `.venv`, tests; the eval project holds
user cases and acceptance evidence. Keep the two apart: every run artifact and
eval output goes under the eval project, never the source dir.

Layout of the eval project:

- `cases/` — user prompts, raw materials, and per-case ground truth.
- `baselines/` — confirmed template-structure fingerprints (SHA + slide/layout counts).
- `runs/` — per-case run outputs (large files gitignored).
- `reports/` — structural and manual visual-review reports.
- `cases/desktop-acceptance-*.md` — the authoritative desktop test spec; derive
  case prompts and acceptance criteria from it, not from memory.

## Four case shapes

1. **User-template case** (template1): a supplied template plus a `.case.json`
   with page-count bounds, required text, forbidden text, and per-section
   coverage. Content is pinned to ONE review report — state that report's path
   and its 核查日期 in the prompt, and present its test results as *that
   report's* snapshot, never as current statistics. The report may also describe
   an older model config than the live session; read the live models from the
   runtime log, do not inherit them from the report.
2–4. **Auto-template cases** (`materials-text` / `materials-images` /
   `materials-data`): only raw materials; the task must extract, select a local
   template, and plan from scratch. Each has a `template_expected` benchmark
   answer that must NOT be handed to the task.

## Ground-truth / answer-key separation

Each auto-template case dir also contains `presentation-plan.json`,
`selection.json`, `materials.json`, and `brief.json`. These are evaluation
ground truth (the expected extraction, selection, and plan). The production task
receives ONLY the raw material files listed in `case.json`'s `inputs`
(`materials.md`, `input-*.png`, `revenue.csv`). Historical finished products
under `runs/*/<earlier-timestamp>/` (`output.pptx`, `qa-report.json`,
`visual-review.*.json`, …) are equally off-limits as answers. Require the task
to extract and plan independently; never let it open these as input.

## Output convention

One new session per case; output to `runs/desktop-<case-id>/<timestamp>/`.
Preserve per case: the plan (`presentation-plan.json` / `plan.json`),
`output.pptx`, every page preview (`render/*.png`), the visual report
(`visual-review.*.json`), `qa-report.json`, and any failure record. Never
overwrite a prior run or a library template.

## Manual Windows PowerPoint confirmation

LibreOffice renders are backend proof, not reading-end proof. Leave for the user
to confirm manually: PPTX opens cleanly, CJK font fidelity vs the LibreOffice
preview, non-standard slide size display, native chart/text editability, image
crop/aspect, and color fidelity. Report these as "pending manual confirmation" —
never speak for the user.

## execution-trace.md conventions (desktop runs)

Desktop case prompts require an execution trace; record ONLY what this session
actually evidenced, phase by phase:

- **Skills**: name, absolute path, and version as loaded via `skill_view`
  (frontmatter version + `git log -1` + dirty/untracked files — an uncommitted
  SKILL.md is an uncommitted version, say so). A skill you loaded but judged
  inapplicable gets one line saying so; no version field means 未确认.
- **Commands**: the exact script invocation and exit/output pointer, with the
  evidence file path (inspect JSON, manifests, PDFs).
- **Vision calls**: image path + the routing line from
  `~/.hermes/logs/agent.log` — `vision_analyze: native fast path` means the main
  model saw it, `Processing image with vision model...` means the auxiliary
  router. The log never prints the vision model's name: name it from
  `config.yaml` `auxiliary.vision.model` and mark it "config-inferred", not as
  observed. Failed vision calls (connection error) are recorded with the retry.
- **未确认 discipline**: anything not evidenced this session — per-item
  dependency versions, the vision request model, reading-end PowerPoint
  confirmation — is written 未确认, never backfilled from configuration or the
  expected flow.

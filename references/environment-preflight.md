# Environment preflight

Read-only probes to confirm the powerpoint-dev environment end-to-end, or to
isolate a failure. Run scripts from the skill directory with the project
`.venv/bin/python` (not bare `python`) so their `pptx_*` imports resolve.

Give every audit its own timestamped directory —
`runs/desktop-preflight/<YYYYMMDD-HHMMSS>/` under the **eval project** (sibling
`powerpoint-skill-eval/`), never the skill source dir `powerpoint-skill/` — and
write all probe output inside it (`environment/`, `logs/`, `environment.json`,
`library-check.json`, `preflight-report.md`, `preflight-report.json`). Output
written to a fixed or reused run dir cannot be distinguished from stale evidence.

```bash
cd /path/to/powerpoint-skill

# 1. Dependency + render probe (writes probe.pptx + slide PNG into the audit dir)
.venv/bin/python scripts/pptx_doctor.py --probe-dir RUN/environment --output RUN/environment.json
# ok:true requires pptx/lxml/PIL/jsonschema importable AND a real render passed.

# 2. Config entrypoint drift (dry-run only; never --apply unless asked)
.venv/bin/python scripts/configure_hermes.py
# config_changed:false + broken_link:null = no drift; disabled must contain
# powerpoint and pptx-template-generation but NOT powerpoint-dev.

# 3. Template library integrity (hashes, shape refs, preview/probe manifests)
.venv/bin/python ../powerpoint-template-library/scripts/validate_library.py --output RUN/library-check.json

# 4. Chinese font availability
fc-list :lang=zh family | sort -u   # need at least one CJK face (Noto Sans SC, WenQuanYi Zen Hei, ...)
```

## Verify each dependency, not the aggregate

`ok:true` is one boolean over seven imports plus a render, so it hides which item
is weak. Probe every item separately and report it individually: import `pptx`,
`lxml`, `PIL`, `jsonschema`, `docx`, `pypdf`, `openpyxl` and print each version,
and resolve `soffice`, `pdftoppm`, `pdftocairo`, `fc-match` to absolute paths.
A passing aggregate with one missing reader still fails on the first DOCX/PDF/XLSX
input.

`pptx-py` and the project `.venv/bin/python` are different interpreters with
different sight lines. `pptx-py` is a shim to the Hermes venv and carries only
the core `pptx/lxml/PIL/jsonschema`; `docx`/`pypdf`/`openpyxl` live in the project
`.venv`. Read material files with the project `.venv`, never through `pptx-py`.

## Render proof vs vision

A successful render is proven by `render_probe.status == "passed"` plus a valid
PNG whose bytes you sample. Sample a known coordinate for the expected RGB (e.g.
a filled shape the probe inserted) and count non-white pixels with PIL — this
proves content was drawn. It is NOT proof anyone looked at the image.

The doctor probe page is not a layout sample: it deliberately drops a filled
rectangle over its own subtitle, so the probe PNG always shows text hidden by a
shape. Do not report that overlap as an environment defect, and do not reuse the
probe layout in a real deck. Read it as two facts only — a square of the expected
RGB rendered at the expected place, and the vision step can see the collision.

## Vision check

Call `vision_analyze` on the rendered PNG. If it returns
`[Unsupported Image]`, the image did not enter the model's visual context: the
active model lacks vision, or no auxiliary vision model is configured. Report
the check as failed and do not substitute script knowledge for observation. Fix
it via Hermes config (`auxiliary.vision.provider`/`model`, or a vision-capable
main model), not via the PPT scripts.

Record the evidence for the vision row: tool name, image path, image SHA256,
whether pixels entered the context (native multimodal return vs
`[Unsupported Image]`), the model that saw them, and what it actually described.

`vision_tools` log lines do not print the model name. To name "the model that
saw them", read `~/.hermes/logs/agent.log` (runtime evidence, not `config.yaml`):
the `agent.turn_context` line names the live main model; `run_agent: OpenAI
client created (agent_init, ...) model=<m>` lines at session build name the
auxiliary clients (the vision one matches `auxiliary.vision.model`); and
`vision_analyze: native fast path` means the main model saw it natively, whereas
`Processing image with vision model...` means the auxiliary vision router path.
On a suspected collision or clipping, call again with a `region` crop and keep
both observations — the crop is what separates a real overlap from a guess made
by reading a small render.

## Chinese font names

`fc-list :lang=zh family | sort -u` proves a CJK face is installed; it does not
prove the family name you plan to write resolves. `fc-match "Noto Sans CJK SC"`
falls back to a Latin-only `Noto Sans` on a machine that ships the `Noto Sans SC`
family, which silently substitutes a font without CJK glyphs. Run `fc-match` on
each family the plan will name and write only names that match exactly.

## Report shape

Separate the facts that are easy to conflate:

- Two questions about a skill: its `SKILL.md` exists on disk, and this session
  actually loaded it through the skill tool (`skill_view` returned the body,
  `readiness_status`). Record them as separate rows — a directory listing cannot
  prove loading.
- Two model facts: what the session is actually running, and `model.default` in
  `~/.hermes/config.yaml`. Read both; claim a match only then. The running model
  is proven by the session's `agent.turn_context` line (and its `API call` model
  fields) in `~/.hermes/logs/agent.log` — never by `config.yaml` alone.
- Three Windows facts: the backend file exists, the Desktop displayed it, the
  user opened it in PowerPoint. Only the first is provable from the backend; mark
  the rest pending manual confirmation and never speak for the user.
- Config drift as dry-run output: run `configure_hermes.py` with no `--apply` and
  quote `config_changed`, `broken_link` and `disabled` from its JSON.

Each row carries the item, pass/fail/unconfirmed, the concrete command output or
hash behind it, and the problem plus suggested repair. Evidence must come from
this session — a prior report's conclusion is not this session's proof.

Check the skill's git baseline too: a frontmatter version newer than the HEAD
commit means the version under test is uncommitted. Name the HEAD commit and the
dirty files before formal testing rather than reporting the version as a release.

## Windows access (WSL)

`wslpath -w <path>` yields the `\\wsl.localhost\<Distro>\...` UNC path the user
opens from Windows Explorer. Backend files existing is a separate fact from the
user successfully opening them — mark the latter as pending manual confirmation.

---
name: powerpoint-dev
description: Create and edit PPTX presentations from a user template, or select a suitable local template from supplied materials and requirements; inspect slides visually and verify the finished deck.
license: MIT
metadata:
  version: 1.5.0
  author: Nous Research and local development
  platforms: [linux, macos, windows]
  hermes:
    tags: [pptx, powerpoint, presentations, slides, office, templates, vision]
    category: productivity
    related_skills: [docx, xlsx, pdf]
---

Use this as the single entrypoint for local PowerPoint development. Resolve script
paths relative to this skill's directory returned by `skill_view`; terminal cwd
need not be the skill directory. Use `pptx-py`, or the project `.venv/bin/python`
when its optional material readers are needed. Run `scripts/pptx_doctor.py` first
in a new environment and probe each dependency individually: its aggregate `ok`
hides which item is weak, so report per item with the output behind it. Commands
emit JSON and nonzero exit status on errors.

When the user requests a finished deck, continue through generation and review.
Show a plan and pause only when the user asks to approve it, or a missing decision
prevents meeting a hard requirement. Historical tasks requiring confirmation do
not establish a permanent approval rule. Resolve ordinary preferences explicitly
and keep working. Explain unsupported requirements before changing their meaning.

For small existing-deck edits, use [the CLI reference](references/cli-reference.md).
For a finished deck, read [the production workflow](references/production-workflow.md).
It covers both user templates and automatic local-template matching. For visual
inspection and report fields, read [visual QA](references/visual-qa.md). To
verify or diagnose the local environment before a run, see
[environment preflight](references/environment-preflight.md). For the eval
harness and desktop acceptance cases (`powerpoint-skill-eval/`), read
[the eval testing reference](references/eval-testing.md).

The shared workflow is:

1. Record requirements and sources. Extract facts with `pptx_materials.py`;
   retain file hashes, pages, paragraphs and cell positions. Do not invent
   missing values or interpret missing formula caches as zero.
2. If a template was supplied, inspect and render it. Use its own masters and
   layouts. Otherwise use `pptx_templates.py` to filter the local catalog by
   executable object capabilities, then compare eligible page previews and
   content capacity. Aesthetic similarity alone is insufficient.
3. Author `presentation-plan.json`: template SHA, source slide refs, explicit
   shape bindings, source IDs, page contents and required coverage. Compile
   with `pptx_workflow.py compile`, then generate with `build`. The compiler
   clears unmapped sample text except protected attribution and explicit
   preserved shapes. Review graphics and inherited master content too.
4. Render every final page. Call Hermes `vision_analyze` on template previews,
   relevant material images, and every final slide. Enlarge suspicious regions;
   compare with the untouched template to identify existing defects.
5. Record a visual report using final render hashes. Run `pptx_workflow.py qa`.
   Revise the semantic plan and regenerate when issues appear. Check numbers,
   units, example residue and native editability as well as visual appearance.
6. Deliver PPTX, previews and the QA result. Missing rendering, failed vision,
   unreviewed slides or stale hashes mean a draft with incomplete verification.

Reuse `powerpoint.plan/v1` and strict inventory refs. Do not create ad-hoc slide
manipulation code when the executor supports the operation. Chart slides cannot
be duplicated; tables must match existing dimensions. Matching must account for
chart type, axis units, usable picture slots and page reuse. Text replacement
preserves the first styled paragraph/run, not arbitrary mixed formatting.
Capacity estimates are warnings; rendered fit determines acceptance.

The library defaults to sibling `powerpoint-template-library/`; commands accept
`--library` for another location. Keep provider credit and source records. Never
treat a template's sample figures as user facts. Keep each run's brief, materials,
plan, execution result, PPTX, render manifest, visual review and QA together.

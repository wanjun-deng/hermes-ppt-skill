Use these commands from the skill directory, or pass absolute script paths.
`pptx-py` supports the core tools. Optional Word/PDF/Excel readers are installed
in this checkout's `.venv/bin/python`; `pyproject.toml` records their version ranges.

```bash
.venv/bin/python scripts/pptx_doctor.py --probe-dir RUN/environment
.venv/bin/python scripts/pptx_materials.py material.docx figures.csv photo.png --output RUN/materials.json
.venv/bin/python scripts/pptx_templates.py RUN/brief.json --output RUN/selection.json
```

For a supplied template, skip library selection: run `pptx_read.py inspect` and
`pptx_render.py`, then combine the strict inventory with the previews to choose
source pages and slots. For library selection, `brief.json` contains a `pages`
array of content needs, for example:

```json
{
  "schema": "powerpoint.brief/v1",
  "preferences": ["经营分析", "数据优先"],
  "slide_size": "16:9",
  "pages": [
    {"id": "cover", "role": "cover"},
    {"id": "revenue", "chart": "COLUMN_CLUSTERED"}
  ]
}
```

Roles use the actual profile vocabulary. Where a semantic requirement has no
exact role match, inspect possible pages rather than inventing a profile role.
Other page constraints are `photos` (count) and `table` (`[rows, columns]`).
Filtering checks executable combinations and chart-page non-reuse. The shortlist
still needs Agent review of meaning, capacity, image ratios and previews. No
candidate means a concrete gap to explain, not permission to lower requirements.

Author `presentation-plan.json` against
`schemas/powerpoint-presentation-v1.schema.json`. Its template has `path`, exact
`sha256` and `slide_count`. Each page has a unique `id`, original `source_slide`
(`slide_id` and `index`), and explicit `text`, `images`, `tables` or `charts`
operations using shape refs from inventory. Copy known page-slot refs from the
profile; do not select shapes by guessed names.

Selection output includes each chosen page's compact `slots`, `slide_count`,
template SHA and attribution. Use these returned values; read the full profile
only when you need alternatives or more design context. A semantic text item
requires `shape` and `value`; do not add a `role` field to that item.

This is the complete shape of a minimal one-page plan. Replace the template SHA,
slide and shape refs with actual inventory or selection output before running:

```json
{
  "schema": "powerpoint.presentation/v1",
  "template": {
    "path": "/absolute/path/template.pptx",
    "sha256": "0000000000000000000000000000000000000000000000000000000000000000",
    "slide_count": 7
  },
  "materials": {
    "path": "/absolute/path/run/materials.json",
    "sha256": "0000000000000000000000000000000000000000000000000000000000000000"
  },
  "pages": [{
    "id": "cover",
    "source_slide": {"slide_id": 256, "index": 0},
    "text": [{"shape": {"shape_id": 5, "path": [0]}, "value": "合成数据汇报"}]
  }],
  "requirements": [{"id": "title", "slide_indexes": [0], "text": ["合成数据汇报"]}],
  "forbidden_text": ["[Company / project name]"]
}
```

`materials.path` points to the JSON manifest created by `pptx_materials.py`, not
the original CSV. Omit `source_ids` for a purely editorial heading; when supplied,
they must exactly match record IDs such as `file1:2` in that manifest. Use
`source_ids` on factual values and independently verify their correspondence.

Attach `materials: {path, sha256}` when material sources are used. Operations may
have `source_ids` copied from `materials.json`. The compiler rejects changed
sources, unknown IDs and unresolved extraction failures. A valid source ID is a
traceability link, not proof of semantic truth: verify each number and unit
against its cited record. Required wording should be listed in `requirements`
with `id`, `text` (array), and optionally `slide_indexes`. `forbidden_text` lists
sample strings and old labels that must disappear.

Unmapped slide text and notes are cleared by default. Explicit `preserve` refs
and recognized provider credit remain. Tables/charts must be filled or explicitly
preserved. Inspect retained images and inherited master/layout objects because
they can contain sample labels too. For Showeet templates keep visible credit,
for example map a footer slot to `Designed by Showeet.com`.

```bash
.venv/bin/python scripts/pptx_workflow.py compile RUN/presentation-plan.json --output RUN/plan.json
.venv/bin/python scripts/pptx_workflow.py build RUN/presentation-plan.json --outdir RUN/final
```

The compiler duplicates reusable pages, retains chart pages at most once,
removes unused originals and reorders by stable aliases. It calls the existing
executor's validation. `build` saves the low-level plan, page map, execution
result, inventory and render manifest. It initially reports `draft` because
generation and rendering alone do not complete visual review.

Text operations accept `font_size_pt` and `alignment` (`left`, `center`, `right`)
for measured repairs. Prefer reducing redundant prose or a more suitable page
before shrinking text. Check Chinese fonts, original character spacing and line
spacing; changes to these can affect the render even when geometry is preserved.
For demonstrated font-substitution or spacing defects, `font_family` applies to
Latin and East Asian runs, `character_spacing_pt` controls tracking, and
`auto_size: "none"` with `word_wrap: true` keeps the intended fixed box.

### Sizing replacements before you write them

- **Group-scaled child boxes are narrower than their inventory bounds.** Text
  boxes nested in groups carry `chOff`/`chExt` child coordinates; the effective
  width is `bounds × (group ext cx / chExt cx)` — routinely ~0.75×. Compute this
  before replacing short template text (a big "95%", a stage title) with longer
  strings, or the text wraps in LibreOffice.
- **LibreOffice wraps overflowing text, it does not clip it.** A single glyph
  pushed past the effective box lands on a second line ON TOP of the content
  below — a stray blue "1" from "19/11" can be misread by vision as a decorative
  bar or a missing character. Prevent with `font_size_pt` sized to the effective
  box (estimate: original width × glyph-count ratio), and confirm the render
  rather than trusting the box math.
- **Tag/label graphics are sized to the ORIGINAL text, not the text-box bounds.**
  Replacing a 4-glyph label with 10+ glyphs wraps it out of the colored tag even
  when the inventory's box width looks roomy. Keep label replacements within the
  original glyph count, or shorten until the tag renders it whole.
- **Display-size cover titles:** estimate the replacement width from the
  original (chars × per-char width at the original pt) and set `font_size_pt` so
  the estimate fits the box; for two lines, join with `\n` (the executor renders
  it as separate lines) and size so each line fits.
- **`forbidden_text` is substring-matched against visible text.** Use the full
  distinctive placeholder phrase ("演示文稿是一种") rather than a common word
  ("演示文稿") that legitimate content may contain (e.g. "跨演示文稿"). If QA
  flags your own wording, tighten the pattern — do not rewrite content to dodge
  an over-broad pattern you control.

Image operations accept `fit: preserve|cover|contain`. `cover` crops without
stretching, `contain` centers the whole image within the original box; both
preserve the template's picture effects. True empty picture placeholders can
receive images. Non-picture decorative shapes cannot.

For a category chart whose metric differs from the template, set `number_format`
explicitly and use `display_units: "none"` to remove inherited axis multipliers.
For example `number_format: "0\"万元\""` is suitable when input values themselves
are in 万元. The tool does not convert units. Bubble/XY data and duplicated chart
pages remain unsupported.

After inspecting every image, save the review described in
[visual QA](visual-qa.md), then run:

```bash
.venv/bin/python scripts/pptx_workflow.py qa RUN/final --review RUN/final/visual-review.json
```

Revise the semantic plan, regenerate from the original template, and review the
new final hashes. After three unsuccessful correction rounds, report the draft
and remaining issues. A user-requested change of scope can reset the plan; a
minor repair within the authorized requirement does not require another approval.

Material readers preserve body block order in DOCX, page locations in text PDFs,
CSV rows, Excel cell coordinates and formula-cache warnings. They do not OCR
scanned PDF pages or calculate Excel formulas. DOCX headers, footnotes, tracked
revisions and text boxes require additional inspection if relevant. See the
primary [DOCX API](https://python-docx.readthedocs.io/en/latest/api/document.html),
[PDF extraction notes](https://pypdf.readthedocs.io/en/stable/user/extract-text.html),
and [Excel reader API](https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.excel.html).

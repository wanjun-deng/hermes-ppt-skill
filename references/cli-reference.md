# Powerpoint Skill

Create, inspect, and edit PowerPoint (.pptx) presentations using the
python-pptx library. Six CLI scripts plus shared inventory, text, and
atomic-I/O modules cover deck creation, shape-centric inspection, edits, template-driven
brand decks, and slide rendering — all offline, no PowerPoint installation
required.

## When to Use

- The user asks to build a slide deck, report presentation, or pitch deck.
- You need to extract text, notes, tables, chart data, or images from a
  .pptx someone shared.
- You need to update an existing deck: replace text, refresh or patch
  chart data, swap a logo, duplicate/remove/reorder slides, set
  backgrounds, footers, hyperlinks, or speaker notes.
- You must produce an on-brand deck from a company .pptx template.
- Do NOT use this for .ppt (legacy binary) files — convert them first with
  `soffice --convert-to pptx old.ppt` if LibreOffice is available.

## Prerequisites

- Python 3.10+ with dependencies declared in `pyproject.toml`
  (`python-pptx` 1.0.2; `pytest` and `jsonschema` for development).
- Optional: LibreOffice (`soffice`) plus poppler (`pdftoppm` or
  `pdftocairo`) for rendering slides to PNGs and for PDF export.
  `pptx_render.py` detects both with `shutil.which` and degrades
  gracefully (reports `{"rendered": false, "missing": [...]}`, exit 0)
  when absent — all create/read/edit operations work without them.
- Check availability via `terminal`:
  `pptx-py -c "import pptx; print(pptx.__version__)"` and `which soffice pdftoppm`.

## How to Run

All CLI scripts live in `scripts/`, take `--help`, print JSON, and exit nonzero on failure. Use the installed `pptx-py` shim or this project’s `.venv/bin/python`; the project environment includes optional material readers. Resolve script paths from the skill directory returned by `skill_view`, rather than assuming the terminal cwd.

```bash
pptx-py scripts/pptx_create.py deck.json out.pptx
pptx-py scripts/pptx_read.py inspect deck.pptx
pptx-py scripts/pptx_read.py export-images deck.pptx --outdir ./img
pptx-py scripts/pptx_edit.py deck.pptx --replace-text "Old Corp" "New Corp"
pptx-py scripts/pptx_edit.py deck.pptx --chart-data update.json
pptx-py scripts/pptx_edit.py deck.pptx --duplicate-slide 2
pptx-py scripts/pptx_edit.py deck.pptx --remove-slide 3 --move-slide 2 0
pptx-py scripts/pptx_from_template.py brand.pptx out.pptx --values vals.json
pptx-py scripts/pptx_from_complete_template.py template.pptx plan.json out.pptx
pptx-py scripts/pptx_render.py deck.pptx --outdir ./render  # slide PNGs
```

Author JSON specs with `write_file`; inspect script output and generated
JSON with `read_file`.

## Quick Reference

| Task | Command |
|---|---|
| New deck from spec | `pptx_create.py spec.json out.pptx` |
| 16:9 vs 4:3 | `"slide_size": "16:9"` or `"4:3"` in the spec |
| Shape inventory as JSON | `pptx_read.py inspect deck.pptx` |
| Export images with locators | `pptx_read.py export-images deck.pptx --outdir DIR` |
| Replace text | `pptx_edit.py deck.pptx --replace-text OLD NEW` |
| Replace chart data | `pptx_edit.py deck.pptx --chart-data spec.json` |
| Patch one series | same flag, spec with `"ops"` (see below) |
| Swap picture | `pptx_edit.py deck.pptx --swap-image N NAME new.png` |
| Duplicate slide | `pptx_edit.py deck.pptx --duplicate-slide N` |
| Remove slide | `pptx_edit.py deck.pptx --remove-slide N` |
| Reorder slide | `pptx_edit.py deck.pptx --move-slide FROM TO` |
| Slide background | `pptx_edit.py deck.pptx --set-background N RRGGBB` |
| Hyperlink runs | `pptx_edit.py deck.pptx --hyperlink N TEXT URL` |
| Slide number on | `pptx_edit.py deck.pptx --enable-slide-number N` |
| Footer text | `pptx_edit.py deck.pptx --set-footer N TEXT` |
| Set notes | `pptx_edit.py deck.pptx --set-notes N TEXT` |
| Append notes | `pptx_edit.py deck.pptx --append-notes N TEXT` |
| Fill template | `pptx_from_template.py tpl.pptx out.pptx --values v.json` |
| Complex template plan | `pptx_from_complete_template.py tpl.pptx plan.json out.pptx` |
| Render slide PNGs | `pptx_render.py deck.pptx --outdir DIR` |

## Procedure

### 1. Create a deck

Write a `powerpoint.create/v1` JSON spec (see `schemas/` and
`pptx_create.py --help`), then run `pptx_create.py`. Unknown fields and
unknown layout/shape/chart types are rejected rather than silently replaced
with defaults. Per slide you can set: `layout` (title,
title_content, section, two_content, title_only, blank), `title`,
`subtitle`, `bullets` (strings, or dicts with `level` 0-4, `size` pt,
`bold`, `italic`, `font`, `color` hex, `link` URL for a hyperlink),
`background` (solid hex), `footer` (text; enables the layout's footer
placeholder), `slide_number` (true; enables the layout's slide-number
placeholder), `images` (path + left/top/width/height in inches), `tables`
(`rows` as list-of-lists), `shapes` (rectangle, rounded_rectangle, oval,
diamond, right_arrow, chevron, with `fill` hex + optional `text`),
`charts` (bar, bar_h, line, pie with `categories` + `series`), and
`notes` (speaker notes). On the `blank` layout, an explicit
`title` creates a normal title textbox so the value is not silently dropped.

### 2. Read a deck

`pptx_read.py inspect deck.pptx` returns the canonical
`powerpoint.inventory/v2` document. It includes the template SHA-256,
slide/layout/master inventories, speaker notes, and a flat pre-order list
of every slide shape (including group containers and children). Each slide
has a strict `{slide_id,index}` `recommended_ref`; each editable shape has
its ID, name, hierarchy path, group path, placeholder metadata, z-order,
transform, fill/line style, local and slide-absolute bounds, capabilities,
paragraph/run detail, and its own `recommended_ref`. Rotated or flipped
geometry that is not represented by the axis-aligned bounds is marked
`accuracy: "approximate"` and emitted in top-level `warnings`.

Use `pptx_read.py export-images deck.pptx --outdir DIR` to export
pictures. Its `powerpoint.image-export/v1` manifest maps each output file
back to the source slide and shape locator, so visual analysis results can
be applied to the correct picture.

Filled picture placeholders also have `capabilities.picture: true`, even
though their API `type` is `PLACEHOLDER`. They can be exported and swapped
through the same image operations. Empty picture placeholders are not filled pictures. Plan image operations can now insert into true picture placeholders; the small edit CLI still targets existing pictures.

### 3. Edit a deck

`pptx_edit.py` combines operations in one pass; use `--output` to keep the
original. Text replacement scans slide shapes, table cells, and notes.
Image swap retargets the picture's relationship id so position and size
are preserved. Slide removal drops the relationship and the `<p:sldId>`
entry; reorder moves the `<p:sldId>` element within `<p:sldIdLst>`
(python-pptx has no public API for either — the script does the XML-level
work). `--duplicate-slide N` appends an independent deep copy of slide N:
shape XML plus image/media/hyperlink relationships are cloned and rIds
remapped, so editing the copy never touches the original. Chart slides
are refused (see Pitfalls). `--set-notes`/`--append-notes` edit speaker
notes; `--set-background`, `--hyperlink`, `--enable-slide-number`, and
`--set-footer` handle deck polish.

Chart updates take a JSON spec via `--chart-data`. Full replace:
`{"slide": 0, "chart": 0, "categories": [...], "series": {...}}`. For
surgical edits, pass `"ops"` instead — a list of
`{"op": "update_series", "name": ..., "values": [...]}`,
`add_series`, `remove_series`, `rename_category` (`from`/`to` or
`index`), and `set_title`. python-pptx can only swap a chart's entire
dataset (`replace_data`), so ops are implemented as read-existing →
modify → replace; the per-part UX is a wrapper, and any chart data not
expressible as categories + numeric series will be normalized by the
round-trip.

### 4. Build from a template

`pptx_from_template.py` replaces `{{token}}` values recursively across
normal and grouped shapes, tables, and notes. Add `--strict` to reject
unused values or unresolved tokens. Appended slides use a
`powerpoint.template-add/v1` spec and a layout `{index,name}` locator copied
from inventory, so duplicate or renamed layouts cannot be selected silently.

For complex templates, use a `powerpoint.plan/v1` document. The root must
contain the exact template SHA-256 and slide count. Original slides use the
inventory's `{slide_id,index}` reference; `duplicate` and `append` may declare
an `as` alias, and later operations use `{"alias":"..."}`. This preserves the
target across deletes and reorders. Shape targets are strict inventory
`recommended_ref` objects. Runtime role inference is disabled unless the
plan explicitly sets `allow_inferred_roles: true`.

```bash
pptx-py scripts/pptx_from_complete_template.py template.pptx plan.json out.pptx
```

```json
{
  "schema": "powerpoint.plan/v1",
  "template": {
    "sha256": "COPY_FROM_INVENTORY",
    "slide_count": 19
  },
  "operations": [
    {
      "action": "duplicate",
      "slide": {"slide_id": 256, "index": 0},
      "as": "cover_copy"
    },
    {
      "action": "fill",
      "slide": {"alias": "cover_copy"},
      "text": [{
        "shape": {
          "shape_id": 23,
          "path": [2, 1],
          "name": "TextBox 16"
        },
        "value": "Annual review"
      }]
    }
  ]
}
```

Operations execute sequentially. Schema validation happens first, original
slide/layout references are bound against the untouched template, and each
operation is revalidated against the current runtime deck. Table dimensions
must match. Chart slides cannot be duplicated. Successful execution emits
`powerpoint.plan-result/v1`; failures emit `powerpoint.error/v1` with an
`operation_index` when applicable.

### 5. Visual verification

`pptx_render.py deck.pptx --outdir ./render` converts the deck to PDF
with `soffice --headless` and renders in an isolated temporary directory.
Each conversion uses its own temporary LibreOffice user profile so it does
not reuse an interactive instance or write to the user's normal profile.
It verifies that the PNG count equals the slide count and emits a
`powerpoint.render/v1` manifest containing slide index, output path,
SHA-256, width, and height. Review only files listed in this manifest.
When either optional tool is missing the script exits 0 with
`{"rendered": false, "missing": [...]}` and guidance; use inventory for
structural verification, but do not treat it as a visual substitute.

## Converting to PDF

If LibreOffice is installed, export the finished deck to PDF directly:

```bash
soffice --headless --convert-to pdf --outdir ./out deck.pptx
```

The output lands at `./out/deck.pdf`. Fonts not installed on the host are
substituted, so render-verify (Procedure step 5) before shipping the PDF.
There is no offline pure-Python .pptx→PDF path; if `soffice` is absent,
say so rather than approximating.

## Pitfalls

- **Filled text must fit the original box**: `fill` keeps the template's run
  properties, and its box geometry — Canva-style decks size each text box
  (often with trailing letter-spacing) to the exact original string. A longer
  value wraps, collides with the decorative `‖`/label shapes next to it, or
  overflows a side panel and gets clipped at render time. Compare the intended
  value's length with the original text (`pptx_read.py inspect`) and stay at
  or below it, then render-verify. The same boxes also clip the *template's own*
  Chinese copy, so a clipped line is not automatically a generation bug.
- **Run splitting**: PowerPoint fragments paragraph text into runs at
  spell-check and edit boundaries. `--replace-text` first merges adjacent
  runs whose formatting is identical, so matches split across such runs
  are replaced with formatting fully preserved. Only when a match spans
  *genuinely differently-formatted* runs is the paragraph rewritten with
  the first run's formatting — verify those slides after replacement.
- **Chart slides cannot be duplicated**: each chart relationship embeds a
  separate XLSX workbook part; cloning that graph reliably is not
  supported, so `--duplicate-slide` refuses chart slides cleanly instead
  of corrupting the deck. Rebuild the chart on a new slide instead.
  External-hyperlink and image/media rels are carried over; layout and
  notes rels are recreated fresh.
- **Chart ops are a wrapper**: python-pptx replaces the whole dataset;
  `"ops"` round-trips existing plot data through `replace_data`, and
  changing chart *type* is not possible.
- **Plan slide references are strict**: copy `{slide_id,index}` from the
  exact inventory and use aliases for newly created slides. Do not write raw
  mutable source indexes into `powerpoint.plan/v1`.
- **Role inference is opt-in**: explicit `targets` are safer. If inference is
  enabled, render and inspect the chosen result because heuristics can be
  ambiguous.
- **The ad-hoc edit CLI has a fixed operation order**:
  `pptx_edit.py` is intended for small maintenance edits. When several
  slide-structural operations depend on one another, use
  `powerpoint.plan/v1` (strict slide refs and aliases), or chain separate
  edit invocations through explicit output files; do not infer execution
  order from the position of CLI flags.
- **Reordering is XML-level**: python-pptx has no supported reorder API.
  `--move-slide` manipulates `<p:sldIdLst>` directly; safe for ordinary
  decks but re-read the deck afterward to confirm.
- **Copying slides between decks is unsupported** — duplication works
  only within one deck, where layouts and masters are shared.
- Footer/slide-number enablement copies the placeholder from the slide's
  layout; on layouts without those placeholders, `--set-footer` fails
  with a clear message (add a textbox instead).
- Hyperlinks apply to whole runs; `--hyperlink` links every run
  containing the given text on that slide.
- The default python-pptx template is 4:3; the create script sets 16:9
  unless the spec says otherwise. Custom templates keep their own size.
- Layout indexes vary by template. For brand templates, list layout names
  first: `pptx_read.py inspect template.pptx` (`layouts`).
- A blank layout has no title placeholder. When `title` is explicit,
  `pptx_create.py` adds a normal title textbox; ad-hoc python-pptx code
  must make that choice explicitly as well.
- Create/edit/template outputs are saved to a temporary package, reopened,
  ZIP-checked, and atomically published. A failed save does not replace an
  existing valid target.
- Always pass `encoding="utf-8"` when writing spec files; tokens like
  `{{city}}` may be filled with non-ASCII values.

## Verification

1. After any create/edit, run `pptx_read.py inspect OUT.pptx` and check
   slide count, shape locators, texts, tables, notes, and chart values.
2. Run `pptx_read.py export-images OUT.pptx --outdir DIR`; verify the
   manifest and file hashes to confirm pictures are embedded and mapped.
3. Render every slide with `pptx_render.py deck.pptx --outdir ./render`
   and review each PNG with `vision_analyze` — this catches overlapping
   shapes, truncated text, and color problems the outline cannot. If rendering or vision is unavailable, record visual verification as incomplete.
4. Validate all JSON documents against the schemas in `schemas/`.
5. Run the bundled contract tests:
   `PYTHONDONTWRITEBYTECODE=1 pptx-py -m pytest tests/ -q -p no:cacheprovider`.

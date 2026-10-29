# Scoped edits, geometry, and fonts

Use `pptx_constrained_edit.py` for edits that must preserve unrelated content.
Template production continues through `pptx_workflow.py`; its sample-text
clearing behavior is intentionally unsuitable for precise edits to an original.

## Contract and execution

Read the source inventory first. Copy its SHA, slide count, slide
`recommended_ref` and leaf shape `recommended_ref` into an edit contract:

```json
{
  "schema": "powerpoint.edit/v1",
  "source": {"sha256": "COPY_FROM_INVENTORY", "slide_count": 2},
  "mode": "precise",
  "operations": [{
    "action": "replace_text",
    "slide": {"slide_id": 256, "index": 0},
    "shape": {"shape_id": 2},
    "old": "150",
    "new": "180",
    "expected_count": 1
  }]
}
```

```bash
.venv/bin/python scripts/pptx_constrained_edit.py source.pptx edit.json output.pptx --report edit-proof.json
```

Use a separate output path. The source SHA is checked before editing and again
before publication. Unknown fields and mismatched locators fail. Changes are
validated in the temporary, reopened PPTX before atomic publication; failure
preserves any existing output. The proof binds the contract and package changes
to the source and output hashes. It is a scope proof, not a visual pass.

| Mode | Allowed operations | Frozen properties |
| --- | --- | --- |
| `precise` | `replace_text`, `set_font`, `set_text_style`, `set_geometry` on explicitly referenced leaves | Everything outside each operation's named object and attributes |
| `restyle` | `set_font`, `set_text_style`, `set_geometry` | All text, data, media and unrelated properties |
| `enrich` | `add_text`, `append_slide` | Existing objects, slide order, masters, layouts and package parts |

Only supported operations are executable in this first scope contract. Chart-data
mutation, picture replacement, removal and reordering remain available through
the existing CLIs, without this new scope-proof guarantee. Do not silently
substitute an unproved edit when proof was requested.

`replace_text` defaults to exactly one occurrence within the selected text
object (or its table cells). Set `expected_count` explicitly for repeated terms.
It keeps original runs and hyperlinks; replacement text takes the first matched
character's formatting. Unmatched suffixes retain their original style. Fields,
paragraphs and explicit line breaks are boundaries. Whole-frame replacement
used in template production still retains only the first paragraph/run style.

`set_font` accepts any subset of `latin`, `east_asian`, `complex_script`.
A font-only fix should specify only that slot. `set_text_style` accepts
`font_size_pt`, `bold`, `color` (RRGGBB) and `alignment`
(left/center/right). `set_geometry.bounds` contains `left/top/width/height`
in inches. Target a leaf in a group, not the group itself. Native chart frames
have no ordinary text frame; chart-specific edits use the existing chart API.

`add_text` contains a strict `slide` ref, `bounds`, `text`, and optional
`style`. The style accepts the text-style fields plus `latin/east_asian`.
`append_slide` takes `layout: {index, name}` from the source inventory and
a nonempty `text` array of these boxes (without `slide`).
No page is rewritten or reordered. Additions must still be factually supported;
the script proves preservation, not the truth of new claims.

## Geometry advisories

```bash
.venv/bin/python scripts/pptx_geometry.py output.pptx --baseline source.pptx --output geometry-report.json
```

The report identifies off-canvas boxes, near-total occlusion by a later explicit
opaque rectangle, and geometry requiring visual inspection. It records strict
shape refs and coordinates, including group scaling. Rotated/flipped objects
have approximate bounds and are not classified as exact geometry defects.

All findings are advisory by default. Decorative bleed and intentional layering
may be valid. Baseline matches include geometry and related object identities;
an original problem is distinguished from a new one. A template build supplies
the source-to-output page map so duplicated pages compare to the correct source.
`--fail-on-new` is an explicit opt-in for controlled test fixtures.

Standard `build` saves `geometry-report.json` before rendering. Standard
`qa` regenerates it against the final artifact and, when available and unchanged,
the original template. Geometry counts appear as an advisory entry in the
existing QA content checks; they never substitute for visual review.

These checks do not calculate glyph layout, clipping, actual contrast, or
inherited master artwork. Render every page and use the existing hash-bound
visual review. A scope-verified edit remains visually unreviewed until that pass.

## Font assignment and resolution

```bash
.venv/bin/python scripts/pptx_fonts.py "WenQuanYi Zen Hei" --text "中文营收120万元"
.venv/bin/python scripts/pptx_doctor.py --font "WenQuanYi Zen Hei"
.venv/bin/python scripts/pptx_create.py deck.json output.pptx --font-latin Arial --font-cjk "WenQuanYi Zen Hei"
```

No font override is applied to an existing template unless requested. Creation
flags are explicit overrides for generated text, table cells, chart defaults
and chart title runs. Per-bullet `font` and template-fill `font_family` use
the same public helper for Latin/East Asian slots. The helper preserves OOXML
child order and existing hyperlinks.

The read-only probe reports the requested and resolved families, actual file,
substitution and coverage of the supplied characters. A missing fontconfig
returns `unconfirmed`; it is not a pass. Probe every required family and actual
representative text, not just a generic installed-CJK check. A resolution/coverage
pass does not prove Windows PowerPoint uses the same fonts. No fonts are installed
automatically.

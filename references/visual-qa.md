Use Hermes `vision_analyze` for images. A vision-capable main model receives
pixels directly; otherwise Hermes can route to a configured auxiliary model.
Check the actual session's model and available tools, not only a name containing
“vision”. A successful render is not evidence that anyone looked at the image.
If `vision_analyze` returns `[Unsupported Image]`, the pixels did not enter the
model's visual context — the model cannot see the image. Treat the vision check
as failed and name the actual model as the blocker; never describe slide content
from the generator script, the text you set, or inventory read-back in place of
looking. Pixel sampling (PIL: assert the expected RGB at a known coordinate,
count non-white pixels) proves the render drew content, and is render evidence
only — not a substitute for visual inspection.

Look at template previews while assigning page/shape roles, at material images
while selecting pictures, and at every final page before delivery. For suspected
clipping or alignment, inspect a larger region and compare the original template.
Keep decorative overlap separate from unintended collisions. Native tables and
charts are established by inventory and read-back, not inferred from pixels.

The final review uses `powerpoint.visual-review/v1`; its schema is in `schemas/`.
Set `artifact_sha256` to the final PPTX hash and `reviewer` to the actual reviewer
or model. Each slide needs `slide_index`, exact rendered `image_sha256`, concrete
`observations`, and all six checks:

| Check | What to examine |
| --- | --- |
| `text_fit` | Clipped characters, overflow, tiny text and broken line wraps |
| `overlap` | Unintended collisions with text, rules, decoration or photos |
| `images` | Stretching, cropping, effects, relevance and legibility |
| `contrast` | Text/background contrast and essential detail |
| `template_consistency` | Original visual language, type hierarchy and alignment |
| `sample_residue` | Old dates, names, example facts, prompts and labels |

Values are `pass`, `fail` or `not_applicable`, based on actual inspection.
Findings include `severity` (`info`, `warning`, `blocking`) and visible `evidence`;
add a suggested repair, normalized region, or reliable `shape_ref` when useful.
Do not fabricate a shape match from approximate rotated/group bounds.

`pptx_workflow.py qa` requires final PPTX and image hashes, all pages reviewed,
all checks present, no failed check and no blocking finding. It also checks the
visible slide text against requirements and forbidden sample phrases. Notes
cannot satisfy requirements. Structural text presence still requires visual
confirmation that it is readable. Missing render/model calls, incomplete reviews,
stale hashes or parser errors retain a draft status.

If a model raises a false positive, preserve its original report and document a
second inspection with the concrete evidence and reviewer. Never edit a model's
failed result into a pass without checking the page. Revisions invalidate old
review hashes. At delivery, give the verified PPTX and current preview paths;
identify any remaining warning and whether final-reader compatibility was tested.

### Settling a "clipped / missing text" claim before editing

Vision models repeatedly flag text as "cut at the edge" when a colored band runs
to the slide edge or a crop boundary sits near the text — confirm with evidence
in this order, cheapest first:

1. **Slide XML** (zipfile read of `output.pptx` → `ppt/slides/slideN.xml`):
   proves what was WRITTEN. If the full string is there, it is a render-level
   issue, not a plan error.
2. **PDF text layer**: `soffice --headless --convert-to pdf` then
   `pdftotext -bbox` — the words and their coordinates show whether every glyph
   rendered inside the container.
3. **Pixel scan of the PNG**: find the rightmost light pixel in each text row
   band (PIL, threshold above the background's RGB sum) and compare with the
   container's edge in inches.
4. **Template comparison**: measure the SAME slot in the untouched template
   render. If the template's own placeholder text extends further right than
   yours, the layout is template-faithful — record the claim as a false positive
   with these measurements, never silently "fix" what is already correct.

### Revision rounds

After each rebuild, diff the new per-slide render SHA against the previous
render manifest: pixel-identical pages keep their prior review (same
`image_sha256`), only changed pages need re-inspection. But update
`artifact_sha256` whenever the PPTX bytes change — plan metadata edits change
the file hash even when every render stays identical, and QA rejects stale
hashes.

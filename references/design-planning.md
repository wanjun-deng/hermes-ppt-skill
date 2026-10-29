# Optional page intent and design plan

Use this sidecar when template selection or content density needs explicit
reasoning. It records what each page should communicate before binding content
to template objects. It does not replace the existing execution plan or
automatically recolor/reformat a template.

```json
{
  "schema": "powerpoint.design/v1",
  "objective": "Explain quarterly performance and the next action",
  "audience": "Internal management",
  "branding": "template",
  "pages": [
    {
      "id": "cover",
      "purpose": "Identify the reporting period and topic",
      "focus": "Quarterly performance",
      "kind": "cover",
      "density": "low"
    },
    {
      "id": "revenue",
      "purpose": "Compare the two quarters using the supplied data",
      "focus": "Revenue increased",
      "kind": "chart",
      "density": "normal",
      "must_keep": ["120", "150", "万元"]
    }
  ]
}
```

Page IDs must match `presentation-plan.json` exactly, with no duplicates.
The execution plan determines slide order. `must_keep` adds visible-text
requirements scoped to the final output page and participates in build/QA;
it does not prove numerical accuracy or source attribution. The content still
needs the original materials and native chart data checks.

```bash
.venv/bin/python scripts/pptx_workflow.py compile presentation-plan.json --design design-plan.json --output plan.json
.venv/bin/python scripts/pptx_workflow.py build presentation-plan.json --design design-plan.json --outdir RUN/final
```

The build saves the sidecar and its interpretation in `compilation.json`.
High density produces an advisory readability warning. Colors and fonts, when
provided, record user/brand intent for the author to follow and verify; they
are not automatic styling commands. Honor the supplied template and user colors.
There are no mandatory palettes, alternating backgrounds, page badges, crop
directions, or universal line-count rules. Prefer a different suitable template
page, concise wording or an allowed split when content does not fit.

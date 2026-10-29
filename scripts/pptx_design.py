"""Optional presentation intent and design constraints, separate from execution."""
import json
from pathlib import Path
from jsonschema import Draft202012Validator


def review_design(design, pages):
    schema = Path(__file__).resolve().parents[1] / "schemas/powerpoint-design-v1.schema.json"
    json.dumps(design, allow_nan=False)
    Draft202012Validator(json.loads(schema.read_text(encoding="utf-8"))).validate(design)
    ids = [page["id"] for page in design["pages"]]
    if len(ids) != len(set(ids)) or set(ids) != {page["id"] for page in pages}:
        raise ValueError("design page IDs must match presentation page IDs exactly")
    by_id = {page["id"]: page for page in design["pages"]}
    requirements, warnings = [], []
    for index, page in enumerate(pages):
        item = by_id[page["id"]]
        if item.get("must_keep"):
            requirements.append({"id": "design:" + page["id"], "text": item["must_keep"], "slide_indexes": [index]})
        if item.get("density") == "high":
            warnings.append({"page_id": page["id"], "code": "design_density_review",
                             "message": "Check readability at the actual slide size; this is advisory."})
    return {"schema": "powerpoint.design-review/v1", "objective": design["objective"],
            "branding": design["branding"], "pages": [by_id[p["id"]] for p in pages],
            "requirements": requirements, "warnings": warnings}

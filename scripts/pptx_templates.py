#!/usr/bin/env python3
"""Filter a local library by executable page requirements before semantic ranking."""
import argparse
import json
from pathlib import Path
import sys
from jsonschema import Draft202012Validator

from pptx_common import write_error
from pptx_inventory import sha256_file

DEFAULT_LIBRARY = Path(__file__).resolve().parents[2] / "powerpoint-template-library"


def select_templates(brief, library=DEFAULT_LIBRARY):
    schema = Path(__file__).resolve().parents[1] / "schemas/powerpoint-brief-v1.schema.json"
    Draft202012Validator(json.loads(schema.read_text())).validate(brief)
    library = Path(library).resolve()
    catalog = json.loads((library / "catalog.json").read_text())
    if len({p["id"] for p in brief["pages"]}) != len(brief["pages"]):
        raise ValueError("brief page ids must be unique")
    candidates, rejected = [], []
    for entry in catalog["templates"]:
        template = library / entry["template"]
        profile = json.loads((library / entry["profile"]).read_text())
        if sha256_file(template) != entry["template_sha256"] or profile["template_sha256"] != entry["template_sha256"]:
            rejected.append({"id": entry["id"], "reasons": ["stale_template_profile"]})
            continue
        pages = [p for p in profile["pages"] if p["auto_selection_eligible"]]
        options_by_need, reasons = [], []
        for need in brief.get("pages", []):
            options = []
            for page in pages:
                slots = page["slots"]
                if need.get("role") and need["role"] != page["role"]:
                    continue
                if need.get("chart"):
                    charts = [s for s in slots if s["kind"] == "chart" and s.get("data_update_supported") and (need["chart"] == "any" or s["chart_type"] == need["chart"])]
                    if not charts:
                        continue
                if need.get("table") and not any(s["kind"] == "table" and [s["rows"], s["columns"]] == need["table"] for s in slots):
                    continue
                if need.get("photos", 0) > sum(s["kind"] == "picture" and s.get("role_hint") == "photo" and s.get("replacement_supported") for s in slots):
                    continue
                options.append(page)
            options.sort(key=lambda p: any(s["kind"] == "chart" for s in p["slots"]))
            options_by_need.append(options)
            if not options:
                reasons.append("no_executable_page_for:" + need["id"])
        order = sorted(range(len(options_by_need)), key=lambda i: len(options_by_need[i]))
        assigned = {}
        def assign(position, used):
            if position == len(order):
                return True
            index = order[position]
            reusable = [p for p in options_by_need[index] if not any(s["kind"] == "chart" for s in p["slots"])]
            for page in reusable[:1] or options_by_need[index]:
                has_chart = any(s["kind"] == "chart" for s in page["slots"])
                key = page["slide"]["slide_id"]
                if has_chart and key in used:
                    continue
                assigned[index] = page
                if assign(position + 1, used | {key} if has_chart else used):
                    return True
            return False
        if not reasons and not assign(0, set()):
            reasons.append("page_combination_requires_unsupported_chart_duplication")
        width, height = (float(x) for x in brief.get("slide_size", "16:9").split(":"))
        size = profile["presentation"]["slide_size"]
        if height <= 0 or abs(size["width_inches"] / size["height_inches"] - width / height) > .01:
            reasons.append("library_does_not_match_requested_slide_size")
        if reasons:
            rejected.append({"id": entry["id"], "reasons": reasons})
            continue
        matches = []
        for index, need in enumerate(brief.get("pages", [])):
            chosen = assigned[index]
            matches.append({"page_id": need["id"], "slide": chosen["slide"], "role": chosen["role"],
                            "preview": str((library / entry["profile"]).parent / chosen["preview"]),
                            "slots": [{k: v for k, v in slot.items() if k in {"shape", "kind", "role_hint", "sample_text", "rows", "columns", "chart_type", "capacity_advice", "replacement_supported", "role_hint_status"}} for slot in chosen["slots"]],
                            "alternative_slides": [p["slide"] for p in options_by_need[index] if p != chosen]})
        preferences = brief.get("preferences", [])
        tags = " ".join(entry["scenarios"] + entry["style"])
        score = sum(str(tag).casefold() in tags.casefold() for tag in preferences)
        candidates.append({"id": entry["id"], "name": entry["name"], "template": str(template),
                           "sha256": entry["template_sha256"], "profile": str(library / entry["profile"]),
                           "slide_count": profile["presentation"]["slide_count"], "attribution": profile["attribution"],
                           "score": score, "reason": entry["selection_reason"], "limitations": entry["limitations"],
                           "page_matches": matches, "semantic_and_visual_review_required": True})
    candidates.sort(key=lambda row: (-row["score"], row["id"]))
    return {"schema": "powerpoint.template-selection/v1", "matched": bool(candidates),
            "candidates": candidates[:3], "rejected": rejected}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("brief")
    parser.add_argument("--library", default=str(DEFAULT_LIBRARY))
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    try:
        brief = json.loads(Path(args.brief).read_text())
        result = select_templates(brief, args.library)
        Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["matched"] else 1
    except Exception as exc:
        write_error(exc, stream=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

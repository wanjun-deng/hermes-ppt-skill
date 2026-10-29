"""Content and hash-bound visual acceptance shared by workflow and evaluation."""
from pathlib import Path
import json
import re

from jsonschema import Draft202012Validator

from pptx_inventory import sha256_file

VISUAL_CHECKS = ("text_fit", "overlap", "images", "contrast", "template_consistency", "sample_residue")


def slide_text(slide, include_notes=False):
    values = [slide.get("notes") or ""] if include_notes else []
    for shape in slide["shapes"]:
        content = shape["content"]
        if "text" in content:
            values.append(content["text"])
        if content.get("table"):
            values.extend(str(cell) for row in content["table"]["cells"] for cell in row)
        if content.get("chart"):
            chart = content["chart"]
            values.append(chart.get("title") or "")
            values.extend(str(value) for value in chart["categories"])
    return "\n".join(values)


def collect_text(inventory, include_notes=False):
    return "\n".join(slide_text(s, include_notes) for s in inventory["slides"])


def content_checks(inventory, requirements=(), forbidden=()):
    checks = []
    for item in requirements:
        indices = item.get("slide_indexes", list(range(len(inventory["slides"]))))
        if any(type(i) is not int or i < 0 or i >= len(inventory["slides"]) for i in indices):
            raise ValueError("requirement has an invalid slide index")
        text = "\n".join(slide_text(inventory["slides"][i]) for i in indices).casefold()
        missing = [value for value in item["text"] if str(value).casefold() not in text]
        checks.append({"id": item["id"], "passed": not missing, "slide_indexes": indices, "missing": missing})
    text = collect_text(inventory)
    residues = [value for value in forbidden if value.casefold() in text.casefold()]
    tokens = sorted(set(re.findall(r"{{[^{}]+}}", collect_text(inventory, True))))
    checks.append({"id": "sample_residue", "passed": not residues, "found": residues})
    checks.append({"id": "unresolved_tokens", "passed": not tokens, "found": tokens})
    return checks


def verify_render(manifest, artifact, slide_count):
    if not manifest or manifest.get("rendered") is not True:
        raise ValueError("all final slides must be rendered")
    if manifest["source"]["sha256"] != sha256_file(artifact):
        raise ValueError("render manifest belongs to a different PPTX")
    if manifest["source"]["slide_count"] != slide_count or len(manifest["slides"]) != slide_count:
        raise ValueError("rendered page count mismatch")
    for index, row in enumerate(manifest["slides"]):
        if row["slide_index"] != index or sha256_file(row["output"]) != row["sha256"]:
            raise ValueError("render image order or hash mismatch")
    return True


def verify_visual(review, manifest, artifact):
    if not review or review.get("schema") != "powerpoint.visual-review/v1":
        raise ValueError("visual review is missing")
    schema = Path(__file__).resolve().parents[1] / "schemas/powerpoint-visual-review-v1.schema.json"
    errors = list(Draft202012Validator(json.loads(schema.read_text())).iter_errors(review))
    if errors:
        raise ValueError("invalid visual review: " + errors[0].message)
    if review.get("artifact_sha256") != sha256_file(artifact):
        raise ValueError("visual review belongs to a different PPTX")
    if not str(review.get("reviewer", "")).strip():
        raise ValueError("visual reviewer is required")
    pages = review.get("slides", [])
    if len(pages) != len(manifest["slides"]):
        raise ValueError("not all final slides have been reviewed")
    findings = []
    for index, (page, rendered) in enumerate(zip(pages, manifest["slides"])):
        if page.get("slide_index") != index or page.get("image_sha256") != rendered["sha256"]:
            raise ValueError("visual review page order or image hash mismatch")
        if not str(page.get("observations", "")).strip():
            raise ValueError("each slide needs visual observations")
        checks = page.get("checks", {})
        if set(checks) != set(VISUAL_CHECKS) or any(v not in {"pass", "fail", "not_applicable"} for v in checks.values()):
            raise ValueError("visual review checks are incomplete")
        if "fail" in checks.values():
            findings.append({"slide_index": index, "code": "visual_check_failed"})
        for finding in page.get("findings", []):
            if finding.get("severity") not in {"info", "warning", "blocking"} or not finding.get("evidence"):
                raise ValueError("visual finding needs severity and evidence")
            if finding.get("severity") == "blocking":
                findings.append(dict(finding, slide_index=index))
    return findings


def quality_report(inventory, artifact, requirements=(), forbidden=(), manifest=None, review=None, geometry=None):
    checks = content_checks(inventory, requirements, forbidden)
    if geometry is not None:
        if geometry["artifact_sha256"] != sha256_file(artifact):
            raise ValueError("geometry report belongs to a different PPTX")
        checks.append({"id": "geometry_advisory", "passed": True,
                       "finding_count": len(geometry["findings"]),
                       "new_finding_count": geometry["new_finding_count"],
                       "visual_review_required": True})
    errors, findings = [], []
    rendered = visual = False
    try:
        rendered = verify_render(manifest, artifact, len(inventory["slides"]))
        findings = verify_visual(review, manifest, artifact)
        visual = not findings
    except (ValueError, KeyError, OSError, TypeError) as exc:
        errors.append(str(exc))
    passed = all(row["passed"] for row in checks) and rendered and visual
    return {"schema": "powerpoint.qa/v1", "artifact": str(Path(artifact).resolve()),
            "artifact_sha256": sha256_file(artifact), "status": "passed" if passed else "needs_revision" if findings else "draft",
            "passed": passed, "content_checks": checks, "render_verified": rendered,
            "visual_verified": visual, "visual_findings": findings, "errors": errors}

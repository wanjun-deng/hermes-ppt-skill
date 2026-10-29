#!/usr/bin/env python3
"""Compile source-aware page plans, build decks, and enforce final acceptance."""
import argparse
import json
from pathlib import Path
import re
import shutil
import sys

from jsonschema import Draft202012Validator
from pptx import Presentation

from pptx_common import PptxSkillError, write_error
from pptx_from_complete_template import compile_plan, execute_plan
from pptx_inventory import inspect_presentation, resolve_shape_ref, sha256_file
from pptx_quality import quality_report
from pptx_render import render

ROOT = Path(__file__).resolve().parents[1]


def read_json(path):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"{path}: expected a JSON document; materials.path must point to the extracted materials.json manifest, not a raw CSV/DOCX/PDF") from exc


def save_json(path, payload):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


def resolve_path(base, value):
    path = Path(value).expanduser()
    return path.resolve() if path.is_absolute() else (Path(base) / path).resolve()


def compile_presentation(spec, base):
    Draft202012Validator(read_json(ROOT / "schemas/powerpoint-presentation-v1.schema.json")).validate(spec)
    template = resolve_path(base, spec["template"]["path"])
    inventory = inspect_presentation(str(template))
    if inventory["template"]["sha256"] != spec["template"]["sha256"] or inventory["presentation"]["slide_count"] != spec["template"]["slide_count"]:
        raise ValueError("presentation plan does not match the template version")
    records = {}
    if spec.get("materials"):
        material_path = resolve_path(base, spec["materials"]["path"])
        if sha256_file(material_path) != spec["materials"]["sha256"]:
            raise ValueError("material manifest changed")
        material = read_json(material_path)
        if material.get("ready") is not True:
            raise ValueError("material extraction has unresolved blocking issues")
        for file in material["files"]:
            if sha256_file(file["path"]) != file["sha256"]:
                raise ValueError("source material changed: " + file["path"])
        records = {row["id"]: row for row in material["records"]}
    prs = Presentation(template)
    operations, retained, targets, page_map, warnings = [], set(), [], [], []
    ids = set()
    for index, page in enumerate(spec["pages"]):
        if page["id"] in ids:
            raise ValueError("duplicate page id")
        ids.add(page["id"])
        ref = page["source_slide"]
        source = inventory["slides"][ref["index"]]
        if source["recommended_ref"] != ref:
            raise ValueError("source slide locator mismatch")
        slide = prs.slides[ref["index"]]
        preserve = {int(resolve_shape_ref(slide, shape).shape_id) for shape in page.get("preserve", [])}
        payload = {"notes": page.get("notes", "")}
        bound_ids = set()
        for kind in ("text", "images", "tables", "charts"):
            payload[kind] = []
            for item in page.get(kind, []):
                bound = resolve_shape_ref(slide, item["shape"])
                if bound.shape_id in bound_ids:
                    raise ValueError("a shape is assigned more than once on page " + page["id"])
                bound_ids.add(bound.shape_id)
                source_ids = item.get("source_ids", [])
                if any(source_id not in records for source_id in source_ids):
                    raise ValueError(f"unknown material source id on page {page['id']}; valid IDs: {sorted(records)}")
                operation = {k: v for k, v in item.items() if k != "source_ids"}
                if kind == "images":
                    operation["path"] = str(resolve_path(base, operation["path"]))
                payload[kind].append(operation)
        for shape in source["shapes"]:
            sid = shape["shape_id"]
            if sid in bound_ids or sid in preserve:
                continue
            if shape["capabilities"]["table"] or shape["capabilities"]["chart"]:
                raise ValueError("unmapped table/chart must be filled or explicitly preserved")
            if shape["capabilities"]["text"] and shape["content"].get("text", "").strip():
                text = shape["content"]["text"]
                if re.fullmatch(r"\s*(?:(?:Designed by\s+)?(?:www\.)?Showeet\.com|Deckary|Copyright[^\n]{0,80})\s*", text, re.I):
                    continue
                payload["text"].append({"shape": shape["recommended_ref"], "value": "", "required": False})
        has_chart = any(s["capabilities"]["chart"] for s in source["shapes"])
        if has_chart:
            if ref["slide_id"] in retained:
                raise ValueError("chart page cannot be reused: chart duplication is unsupported")
            retained.add(ref["slide_id"])
            target = ref
            operations.append(dict(payload, action="fill", slide=ref))
        else:
            alias = "page_" + str(index + 1)
            target = {"alias": alias}
            operations.append(dict(payload, action="duplicate", slide=ref, **{"as": alias}))
        targets.append(target)
        page_map.append({"id": page["id"], "source_slide": ref, "target": target, "output_index": index})
        for item in page.get("text", []):
            old = next(s for s in source["shapes"] if s["shape_id"] == resolve_shape_ref(slide, item["shape"]).shape_id)
            if len(str(item["value"])) > max(8, len(old["content"].get("text", "")) * 1.5):
                warnings.append({"page_id": page["id"], "code": "text_capacity_risk", "shape": item["shape"]})
    for slide in inventory["slides"]:
        if slide["slide_id"] not in retained:
            operations.append({"action": "remove", "slide": slide["recommended_ref"]})
    for index, target in enumerate(targets):
        operations.append({"action": "reorder", "slide": target, "to": index})
    plan = {"schema": "powerpoint.plan/v1", "template": {k: spec["template"][k] for k in ("sha256", "slide_count")}, "operations": operations}
    compile_plan(prs, plan, str(template))
    return template, plan, {"schema": "powerpoint.compilation/v1", "pages": page_map, "warnings": warnings}


def build(spec_path, outdir, should_render=True):
    spec_path, outdir = Path(spec_path).resolve(), Path(outdir).resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    spec = read_json(spec_path)
    template, plan, compilation = compile_presentation(spec, spec_path.parent)
    previous = outdir / "output.pptx"
    if previous.exists():
        revision = outdir / "revisions" / sha256_file(previous)
        revision.mkdir(parents=True, exist_ok=True)
        for name in ("output.pptx", "plan.json", "plan-result.json", "qa-report.json", "run.json"):
            if (outdir / name).is_file():
                shutil.copyfile(outdir / name, revision / name)
        for path in outdir.glob("visual-review*.json"):
            shutil.copyfile(path, revision / path.name)
        if (outdir / "render-manifest.json").exists():
            old_render = read_json(outdir / "render-manifest.json")
            for row in old_render.get("slides", []):
                archived = revision / "render" / Path(row["output"]).name
                archived.parent.mkdir(exist_ok=True)
                shutil.copyfile(row["output"], archived)
                row["output"] = str(archived)
            save_json(revision / "render-manifest.json", old_render)
    save_json(outdir / "run.json", {"schema": "powerpoint.run/v1", "status": "building", "visual_review_pending": True})
    save_json(outdir / "plan.json", plan)
    save_json(outdir / "compilation.json", compilation)
    save_json(outdir / "presentation-plan.json", spec)
    artifact = outdir / "output.pptx"
    result = execute_plan(str(template), plan, str(artifact))
    save_json(outdir / "plan-result.json", result)
    inventory = inspect_presentation(str(artifact))
    save_json(outdir / "inventory.json", inventory)
    manifest = None
    if should_render:
        try:
            manifest = render(str(artifact), str(outdir / "render"), "slide", 120)
        except Exception as exc:
            save_json(outdir / "run.json", {"schema": "powerpoint.run/v1", "status": "render_failed", "artifact_sha256": sha256_file(artifact), "error": str(exc), "visual_review_pending": True})
            raise
        save_json(outdir / "render-manifest.json", manifest)
    qa = quality_report(inventory, artifact, spec.get("requirements", []), spec.get("forbidden_text", []), manifest)
    save_json(outdir / "qa-report.json", qa)
    save_json(outdir / "run.json", {"schema": "powerpoint.run/v1", "status": qa["status"], "skill_version": "1.5.0", "template_sha256": sha256_file(template), "artifact_sha256": sha256_file(artifact), "visual_review_pending": True})
    return {"ok": True, "artifact": str(artifact), "status": qa["status"], "visual_review_pending": True}


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    compiler = commands.add_parser("compile")
    compiler.add_argument("spec")
    compiler.add_argument("--output", required=True)
    builder = commands.add_parser("build")
    builder.add_argument("spec")
    builder.add_argument("--outdir", required=True)
    builder.add_argument("--no-render", action="store_true")
    qa = commands.add_parser("qa")
    qa.add_argument("run_dir")
    qa.add_argument("--review", required=True)
    args = parser.parse_args(argv)
    try:
        if args.command == "compile":
            spec = Path(args.spec).resolve()
            _, plan, compilation = compile_presentation(read_json(spec), spec.parent)
            save_json(args.output, plan)
            result = compilation
        elif args.command == "build":
            result = build(args.spec, args.outdir, not args.no_render)
        else:
            folder = Path(args.run_dir).resolve()
            spec = read_json(folder / "presentation-plan.json")
            manifest = read_json(folder / "render-manifest.json") if (folder / "render-manifest.json").exists() else None
            result = quality_report(inspect_presentation(str(folder / "output.pptx")), folder / "output.pptx", spec.get("requirements", []), spec.get("forbidden_text", []), manifest, read_json(args.review))
            save_json(folder / "qa-report.json", result)
            state = read_json(folder / "run.json")
            state.update(status=result["status"], visual_review_pending=not result["visual_verified"])
            save_json(folder / "run.json", state)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 1 if args.command == "qa" and not result["passed"] else 0
    except Exception as exc:
        write_error(exc, stream=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

"""End-to-end tests for the powerpoint skill helper scripts.

Runs entirely offline. Exercises create, read, template-fill (with
non-ASCII values), and edit (text + chart data + remove/move slide).
"""
import json
import os
import subprocess
import sys

import pytest

SKILL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(SKILL, "scripts")


def run(script, *args):
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    proc = subprocess.run(
        [sys.executable, os.path.join(SCRIPTS, script), *args],
        capture_output=True, text=True, encoding="utf-8", env=env)
    assert proc.returncode == 0, f"{script} failed: {proc.stderr}"
    return json.loads(proc.stdout)


def write_json(path, obj):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False)


def assert_schema(filename, payload):
    from pathlib import Path
    from jsonschema import Draft202012Validator

    schema_path = Path(SKILL) / "schemas" / filename
    with open(schema_path, encoding="utf-8") as handle:
        schema = json.load(handle)
    Draft202012Validator(schema).validate(payload)


def inspect_deck(path):
    return run("pptx_read.py", "inspect", str(path))


def slide_texts(slide):
    return [
        shape["content"]["text"] for shape in slide["shapes"]
        if "text" in shape["content"]
    ]


def slide_tables(slide):
    return [
        shape["content"]["table"]["cells"] for shape in slide["shapes"]
        if "table" in shape["content"]
    ]


def slide_images(slide):
    return [
        shape["content"]["image"] for shape in slide["shapes"]
        if "image" in shape["content"]
    ]


def slide_charts(slide):
    return [
        shape["content"]["chart"] for shape in slide["shapes"]
        if "chart" in shape["content"]
    ]


def first_shape_ref(slide, capability):
    return next(
        shape["recommended_ref"] for shape in slide["shapes"]
        if shape["capabilities"][capability]
    )


def slide_ref(inventory, index):
    return dict(inventory["slides"][index]["recommended_ref"])


def layout_ref(inventory, name):
    return dict(next(
        layout["recommended_ref"] for layout in inventory["layouts"]
        if layout["name"] == name
    ))


def plan_v1(inventory, operations, *, strict=False, values=None):
    plan = {
        "schema": "powerpoint.plan/v1",
        "template": {
            "sha256": inventory["template"]["sha256"],
            "slide_count": inventory["presentation"]["slide_count"],
        },
        "operations": operations,
    }
    if strict:
        plan["strict"] = True
    if values is not None:
        plan["values"] = values
    return plan


def test_json_schemas_are_valid():
    from pathlib import Path
    from jsonschema import Draft202012Validator

    schema_dir = Path(SKILL) / "schemas"
    schemas = sorted(schema_dir.glob("*.schema.json"))
    assert {path.name for path in schemas} >= {
        "powerpoint-create-result-v1.schema.json",
        "powerpoint-create-v1.schema.json",
        "powerpoint-edit-result-v1.schema.json",
        "powerpoint-error-v1.schema.json",
        "powerpoint-image-export-v1.schema.json",
        "powerpoint-inventory-v2.schema.json",
        "powerpoint-plan-result-v1.schema.json",
        "powerpoint-plan-v1.schema.json",
        "powerpoint-render-v1.schema.json",
        "powerpoint-template-add-v1.schema.json",
        "powerpoint-template-result-v1.schema.json",
    }
    for schema_path in schemas:
        with open(schema_path, encoding="utf-8") as handle:
            Draft202012Validator.check_schema(json.load(handle))



@pytest.fixture(scope="module")
def workdir(tmp_path_factory):
    return tmp_path_factory.mktemp("pptx")


@pytest.fixture(scope="module")
def sample_png(workdir):
    # 1x1 red PNG, hardcoded bytes -> no Pillow dependency.
    import base64
    data = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4"
        "z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==")
    path = workdir / "red.png"
    path.write_bytes(data)
    return str(path)


@pytest.fixture(scope="module")
def deck(workdir, sample_png):
    """Create a deck exercising every create feature."""
    spec = {
        "schema": "powerpoint.create/v1",
        "slide_size": "16:9",
        "slides": [
            {"layout": "title", "title": "Annual Review",
             "subtitle": "Fiscal year 2026",
             "notes": "Welcome the audience."},
            {"layout": "title_content", "title": "Agenda",
             "bullets": [
                 "Overview",
                 {"text": "Details", "level": 1, "bold": True,
                  "size": 20, "color": "CC0000", "font": "Arial"},
                 {"text": "Deep dive", "level": 2, "italic": True},
             ],
             "notes": "Keep this under two minutes."},
            {"layout": "blank", "title": None,
             "images": [{"path": sample_png, "left": 0.5, "top": 0.5,
                         "width": 2, "height": 2}],
             "tables": [{"left": 3, "top": 1, "width": 6, "height": 2,
                         "rows": [["Region", "Sales"],
                                  ["North", "120"], ["South", "80"]]}],
             "shapes": [{"type": "rounded_rectangle", "left": 10,
                         "top": 1, "width": 2.5, "height": 1,
                         "fill": "4472C4", "text": "Callout",
                         "text_color": "FFFFFF"}]},
            {"layout": "title_only", "title": "Charts",
             "charts": [
                 {"type": "bar", "left": 0.5, "top": 1.5, "width": 4,
                  "height": 3, "title": "Sales",
                  "categories": ["Q1", "Q2"],
                  "series": {"North": [10, 20], "South": [7, 13]}},
                 {"type": "line", "left": 4.7, "top": 1.5, "width": 4,
                  "height": 3, "categories": ["Jan", "Feb", "Mar"],
                  "series": {"Trend": [1, 3, 2]}},
                 {"type": "pie", "left": 8.9, "top": 1.5, "width": 4,
                  "height": 3, "categories": ["A", "B", "C"],
                  "series": {"Share": [50, 30, 20]}}]},
        ],
    }
    assert_schema("powerpoint-create-v1.schema.json", spec)
    spec_path = workdir / "deck.json"
    write_json(spec_path, spec)
    out = workdir / "deck.pptx"
    result = run("pptx_create.py", str(spec_path), str(out))
    assert_schema("powerpoint-create-result-v1.schema.json", result)
    assert result["ok"] and result["slides"] == 4
    return str(out)


def test_create_and_outline(deck):
    outline = inspect_deck(deck)
    assert outline["presentation"]["slide_count"] == 4
    assert outline["presentation"]["slide_size"]["width_inches"] == pytest.approx(13.333, abs=0.01)
    s0, s1, s2, s3 = outline["slides"]
    assert "Annual Review" in slide_texts(s0)
    assert "Fiscal year 2026" in slide_texts(s0)
    assert s0["notes"] == "Welcome the audience."
    # bullets present on slide 1
    assert any("Deep dive" in t for t in slide_texts(s1))
    # table content
    assert slide_tables(s2)[0][0] == ["Region", "Sales"]
    assert slide_tables(s2)[0][2] == ["South", "80"]
    # image inventory + shape text
    assert len(slide_images(s2)) == 1 and slide_images(s2)[0]["extension"] == "png"
    assert "Callout" in slide_texts(s2)
    # three charts with correct data
    assert len(slide_charts(s3)) == 3
    bar = slide_charts(s3)[0]
    assert bar["categories"] == ["Q1", "Q2"]
    north = next(s for s in bar["series"] if s["name"] == "North")
    assert north["values"] == [10.0, 20.0]


def test_inventory_schema_locators_and_layouts(deck):
    from pptx import Presentation

    sys.path.insert(0, SCRIPTS)
    from pptx_inventory import resolve_shape_ref

    inventory = inspect_deck(deck)
    assert_schema("powerpoint-inventory-v2.schema.json", inventory)
    assert inventory["schema"] == "powerpoint.inventory/v2"
    assert len(inventory["template"]["sha256"]) == 64
    assert inventory["layouts"][0]["placeholders"]
    assert "shapes" in inventory["layouts"][0]
    assert inventory["masters"]

    prs = Presentation(deck)
    for slide_record, slide in zip(inventory["slides"], prs.slides):
        for shape_record in slide_record["shapes"]:
            resolved = resolve_shape_ref(
                slide, shape_record["recommended_ref"])
            assert resolved.shape_id == shape_record["shape_id"]
            assert resolved.name == shape_record["name"]

    title = next(
        shape for shape in inventory["slides"][0]["shapes"]
        if shape["placeholder"] and shape["placeholder"]["idx"] == 0)
    assert title["recommended_ref"] == {"placeholder": 0}
    assert set(title["bounds"]) == {
        "accuracy", "local_emu", "slide_emu", "slide_inches"}


def test_inventory_rejects_mismatched_locator(deck, workdir):
    inventory = inspect_deck(deck)
    ref = dict(first_shape_ref(inventory["slides"][0], "text"))
    ref["name"] = "Wrong shape name"
    plan_path = workdir / "mismatched_locator.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 0),
        "text": [{"shape": ref, "value": "Must not be written"}],
    }]))
    proc = run_raw(
        "pptx_from_complete_template.py", deck, str(plan_path),
        str(workdir / "mismatched_locator.pptx"))
    assert proc.returncode != 0
    error = json.loads(proc.stderr)
    assert error["schema"] == "powerpoint.error/v1"
    assert "name mismatch" in error["error"]["message"]


def test_inventory_warns_about_duplicate_shape_names(workdir):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    first = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(2), Inches(1))
    second = slide.shapes.add_textbox(
        Inches(4), Inches(1), Inches(2), Inches(1))
    first.name = "Repeated name"
    second.name = "Repeated name"
    template = workdir / "duplicate-shape-names.pptx"
    prs.save(str(template))

    inventory = inspect_deck(template)
    warning = next(
        item for item in inventory["slides"][0]["warnings"]
        if item["code"] == "duplicate_shape_name"
    )
    assert warning["name"] == "Repeated name"
    assert warning["shape_ids"] == [first.shape_id, second.shape_id]
    assert any(item["code"] == "duplicate_shape_name"
               and item["scope"] == "slide"
               for item in inventory["warnings"])



def test_inventory_marks_rotated_bounds_as_approximate(workdir):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_textbox(
        Inches(1), Inches(1), Inches(3), Inches(1))
    shape.text = "Rotated"
    shape.rotation = 30
    template = workdir / "rotated-shape.pptx"
    prs.save(str(template))

    inventory = inspect_deck(template)
    record = inventory["slides"][0]["shapes"][0]
    assert record["transform"]["rotation_degrees"] == 30
    assert record["bounds"]["accuracy"] == "approximate"
    warning = next(
        item for item in inventory["warnings"]
        if item["code"] == "approximate_bounds"
    )
    assert warning["scope"] == "slide"
    assert warning["shape_id"] == shape.shape_id


def test_reader_returns_json_error_for_invalid_file(workdir):
    invalid = workdir / "not-a-presentation.pptx"
    invalid.write_text("not a zip package", encoding="utf-8")
    proc = run_raw("pptx_read.py", "inspect", str(invalid))
    assert proc.returncode != 0
    error = json.loads(proc.stderr)
    assert error["schema"] == "powerpoint.error/v1"
    assert error["ok"] is False


def test_inventory_group_geometry_and_plan_roundtrip(workdir):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    group = slide.shapes.add_group_shape()
    child = group.shapes.add_textbox(
        Inches(1), Inches(2), Inches(2), Inches(1))
    child.text = "Nested text"
    group.left = Inches(3)
    group.top = Inches(1)
    group.width = Inches(4)
    group.height = Inches(2)
    template = workdir / "grouped-template.pptx"
    prs.save(str(template))

    inventory = inspect_deck(template)
    shapes = inventory["slides"][0]["shapes"]
    group_record = next(shape for shape in shapes
                        if shape["capabilities"]["group"])
    child_record = next(shape for shape in shapes
                        if shape["content"].get("text") == "Nested text")
    assert group_record["path"] == [0]
    assert child_record["path"] == [0, 0]
    assert child_record["group_path"][0]["shape_id"] == (
        group_record["shape_id"])
    assert child_record["bounds"]["slide_emu"] == (
        group_record["bounds"]["slide_emu"])

    plan_path = workdir / "grouped-plan.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "fill",
        "slide": slide_ref(inventory, 0),
        "text": [{
            "shape": child_record["recommended_ref"],
            "value": "Updated nested text",
        }],
    }]))
    output = workdir / "grouped-output.pptx"
    run("pptx_from_complete_template.py", str(template), str(plan_path),
        str(output))
    updated = inspect_deck(output)
    assert "Updated nested text" in slide_texts(updated["slides"][0])
    assert "Nested text" not in slide_texts(updated["slides"][0])



def test_complete_template_replaces_token_inside_group(workdir):
    from pptx import Presentation
    from pptx.util import Inches

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    group = slide.shapes.add_group_shape()
    group.shapes.add_textbox(
        Inches(1), Inches(1), Inches(3), Inches(1)).text = "{{nested}}"
    template = workdir / "group-token-template.pptx"
    prs.save(str(template))
    inventory = inspect_deck(template)
    plan_path = workdir / "group-token-plan.json"
    write_json(plan_path, plan_v1(
        inventory, [], strict=True, values={"nested": "Replaced in group"}))
    output = workdir / "group-token-output.pptx"
    result = run("pptx_from_complete_template.py", str(template),
                 str(plan_path), str(output))
    assert result["operations"]["filled_text"] == 1
    rendered = inspect_deck(output)
    assert "Replaced in group" in slide_texts(rendered["slides"][0])


def test_complete_template_rejects_wrong_template_hash(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "wrong-template-hash.json"
    plan = plan_v1(inventory, [])
    plan["template"]["sha256"] = "0" * 64
    write_json(plan_path, plan)
    proc = run_raw(
        "pptx_from_complete_template.py", deck, str(plan_path),
        str(workdir / "wrong-template-output.pptx"))
    assert proc.returncode != 0
    error = json.loads(proc.stderr)
    assert error["error"]["code"] == "template_mismatch"
    assert "sha256 mismatch" in error["error"]["message"]



def test_locator_rejects_unknown_fields(deck):
    from pptx import Presentation

    sys.path.insert(0, SCRIPTS)
    from pptx_inventory import InventoryError, resolve_shape_ref

    prs = Presentation(deck)
    with pytest.raises(InventoryError, match="unknown fields"):
        resolve_shape_ref(prs.slides[0], {"placeholder": 0, "typo": 1})


def test_atomic_save_does_not_replace_valid_target_on_failure(workdir):
    sys.path.insert(0, SCRIPTS)
    from pptx_common import PptxSkillError, atomic_save_presentation

    target = workdir / "existing-output.pptx"
    target.write_bytes(b"existing-valid-user-data")

    class InvalidPresentation:
        def save(self, path):
            with open(path, "wb") as handle:
                handle.write(b"not a pptx")

    with pytest.raises(PptxSkillError, match="invalid PPTX"):
        atomic_save_presentation(InvalidPresentation(), target)
    assert target.read_bytes() == b"existing-valid-user-data"
    assert not list(workdir.glob(".existing-output.pptx.*.tmp.pptx"))


def test_create_rejects_unknown_layout_with_json_error(workdir):
    spec_path = workdir / "invalid-layout.json"
    write_json(spec_path, {
        "schema": "powerpoint.create/v1",
        "slides": [{"layout": "typo-layout"}],
    })
    proc = run_raw("pptx_create.py", str(spec_path),
                   str(workdir / "invalid-layout.pptx"))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert_schema("powerpoint-error-v1.schema.json", error)
    assert error["error"]["code"] == "invalid_layout_ref"
    assert not (workdir / "invalid-layout.pptx").exists()


def test_create_rejects_invalid_bullet_field_types(workdir):
    spec_path = workdir / "invalid-bullet-type.json"
    write_json(spec_path, {
        "schema": "powerpoint.create/v1",
        "slides": [{
            "layout": "title_content",
            "bullets": [{"text": "Bad", "bold": "yes"}],
        }],
    })
    proc = run_raw("pptx_create.py", str(spec_path),
                   str(workdir / "invalid-bullet-type.pptx"))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert error["error"]["code"] == "invalid_type"


def test_create_blank_layout_keeps_explicit_title(workdir):
    spec_path = workdir / "blank-title.json"
    write_json(spec_path, {
        "schema": "powerpoint.create/v1",
        "slides": [{"layout": "blank", "title": "Visible title"}],
    })
    output = workdir / "blank-title.pptx"
    run("pptx_create.py", str(spec_path), str(output))
    inventory = inspect_deck(str(output))
    assert "Visible title" in slide_texts(inventory["slides"][0])


def test_create_43_size(workdir):
    spec_path = workdir / "small.json"
    write_json(spec_path, {"schema": "powerpoint.create/v1",
                           "slide_size": "4:3",
                           "slides": [{"layout": "title", "title": "T"}]})
    out = workdir / "small.pptx"
    run("pptx_create.py", str(spec_path), str(out))
    outline = inspect_deck(str(out))
    assert [outline["presentation"]["slide_size"]["width_inches"], outline["presentation"]["slide_size"]["height_inches"]] == [10.0, 7.5]


def test_notes_mode(deck):
    result = inspect_deck(deck)
    assert result["slides"][1]["notes"] == "Keep this under two minutes."


def test_image_export(deck, workdir):
    out_dir = workdir / "exported"
    result = run("pptx_read.py", "export-images", deck, "--outdir", str(out_dir))
    assert_schema("powerpoint-image-export-v1.schema.json", result)
    assert result["schema"] == "powerpoint.image-export/v1"
    assert len(result["exports"]) == 1
    inventory_ref = first_shape_ref(inspect_deck(deck)["slides"][2], "picture")
    assert result["exports"][0]["shape_ref"] == inventory_ref
    exported = result["exports"][0]["output"]
    assert os.path.getsize(exported) > 0
    with open(exported, "rb") as fh:
        assert fh.read(8) == b"\x89PNG\r\n\x1a\n"


def test_template_fill_non_ascii(workdir):
    """Template token fill with non-ASCII values, forced under LC_ALL=C."""
    tpl_spec = workdir / "tpl.json"
    write_json(tpl_spec, {"schema": "powerpoint.create/v1", "slides": [
        {"layout": "title", "title": "{{city}} report",
         "subtitle": "Prepared by {{author}}",
         "notes": "Deck for {{city}}."},
        {"layout": "title_content", "title": "Data",
         "tables": [{"rows": [["Site", "{{city}}"]]}]},
    ]})
    template = workdir / "template.pptx"
    run("pptx_create.py", str(tpl_spec), str(template))

    values = workdir / "values.json"
    write_json(values, {"city": "Z\u00fcrich \u2014 \u2018Bericht\u2019",
                        "author": "Beispiel GmbH"})
    out = workdir / "filled.pptx"
    result = run("pptx_from_template.py", str(template), str(out),
                 "--values", str(values), "--strict")
    assert_schema("powerpoint-template-result-v1.schema.json", result)
    assert result["schema"] == "powerpoint.template-result/v1"
    assert result["tokens_filled"] == 4
    assert result["unused_values"] == []
    assert result["remaining_tokens"] == []

    outline = inspect_deck(str(out))
    expected = "Z\u00fcrich \u2014 \u2018Bericht\u2019"
    assert f"{expected} report" in slide_texts(outline["slides"][0])
    assert outline["slides"][0]["notes"] == f"Deck for {expected}."
    assert slide_tables(outline["slides"][1])[0][0][1] == expected


def test_template_add_slides(deck, workdir):
    inventory = inspect_deck(deck)
    add_spec = workdir / "add.json"
    add_payload = {
        "schema": "powerpoint.template-add/v1",
        "slides": [{
            "layout": layout_ref(inventory, "Title and Content"),
            "title": "Appended", "bullets": ["from layout"],
            "notes": "appended slide",
        }],
    }
    assert_schema("powerpoint-template-add-v1.schema.json", add_payload)
    write_json(add_spec, add_payload)
    out = workdir / "appended.pptx"
    result = run("pptx_from_template.py", deck, str(out),
                 "--add-slides", str(add_spec))
    assert_schema("powerpoint-template-result-v1.schema.json", result)
    assert result["schema"] == "powerpoint.template-result/v1"
    outline = inspect_deck(str(out))
    assert outline["presentation"]["slide_count"] == 5
    assert "Appended" in slide_texts(outline["slides"][4])


def test_edit_replace_text(deck, workdir):
    edited = workdir / "edited.pptx"
    result = run("pptx_edit.py", deck, "--output", str(edited),
                 "--replace-text", "Annual Review", "Semi-Annual Review",
                 "--replace-text", "South", "West")
    assert_schema("powerpoint-edit-result-v1.schema.json", result)
    assert result["replacements"] == 2  # title + table cell
    outline = inspect_deck(str(edited))
    assert "Semi-Annual Review" in slide_texts(outline["slides"][0])
    assert slide_tables(outline["slides"][2])[0][2][0] == "West"
    # formatting survives a run-level replace: red bold bullet untouched
    assert any("Deep dive" in t for t in slide_texts(outline["slides"][1]))


def test_edit_chart_data(deck, workdir):
    spec = workdir / "chart_update.json"
    write_json(spec, {"slide": 3, "chart": 0,
                      "categories": ["Q3", "Q4"],
                      "series": {"North": [30, 40], "South": [21, 34]}})
    edited = workdir / "chart_edited.pptx"
    run("pptx_edit.py", deck, "--output", str(edited),
        "--chart-data", str(spec))
    outline = inspect_deck(str(edited))
    bar = slide_charts(outline["slides"][3])[0]
    assert bar["categories"] == ["Q3", "Q4"]
    north = next(s for s in bar["series"] if s["name"] == "North")
    assert north["values"] == [30.0, 40.0]


def test_edit_remove_and_move_slide(deck, workdir):
    edited = workdir / "reordered.pptx"
    run("pptx_edit.py", deck, "--output", str(edited),
        "--remove-slide", "2", "--move-slide", "2", "0")
    outline = inspect_deck(str(edited))
    assert outline["presentation"]["slide_count"] == 3
    # charts slide (was index 3, then 2 after removal) moved to front
    assert len(slide_charts(outline["slides"][0])) == 3
    assert "Annual Review" in slide_texts(outline["slides"][1])


def test_edit_swap_image(deck, workdir):
    import base64
    blue = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNg"
        "YPj/HwADAgH/p5UronAAAAAASUVORK5CYII=")
    blue_path = workdir / "blue.png"
    blue_path.write_bytes(blue)
    outline = inspect_deck(deck)
    # find picture shape name via python-pptx directly
    sys.path.insert(0, SCRIPTS)
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(deck)
    name = next(s.name for s in prs.slides[2].shapes
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
    edited = workdir / "swapped.pptx"
    run("pptx_edit.py", deck, "--output", str(edited),
        "--swap-image", "2", name, str(blue_path))
    out_dir = workdir / "swapped_images"
    result = run("pptx_read.py", "export-images", str(edited), "--outdir", str(out_dir))
    with open(result["exports"][0]["output"], "rb") as fh:
        assert fh.read() == blue


@pytest.mark.parametrize("mode", ["plan", "edit"])
def test_filled_picture_placeholder_inventory_export_and_swap(workdir, sample_png, mode):
    import base64
    from pathlib import Path
    from pptx import Presentation

    template = workdir / f"picture-placeholder-{mode}.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[8])
    picture = slide.placeholders[1].insert_picture(sample_png)
    original_bounds = tuple(getattr(picture, name) for name in ("left", "top", "width", "height"))
    original_crop = (picture.crop_left, picture.crop_top, picture.crop_right, picture.crop_bottom)
    prs.save(template)

    inventory = inspect_deck(template)
    assert_schema("powerpoint-inventory-v2.schema.json", inventory)
    image = next(shape for shape in inventory["slides"][0]["shapes"]
                 if shape["capabilities"]["picture"])
    assert image["type"] == "PLACEHOLDER"
    assert image["placeholder"]["type"] == "PICTURE"
    exported = run("pptx_read.py", "export-images", str(template), "--outdir",
                   str(workdir / f"placeholder-original-{mode}"))
    assert_schema("powerpoint-image-export-v1.schema.json", exported)
    with open(exported["exports"][0]["output"], "rb") as handle:
        assert handle.read() == Path(sample_png).read_bytes()

    replacement = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNg"
        "YPj/HwADAgH/p5UronAAAAAASUVORK5CYII=")
    blue_path = workdir / f"placeholder-blue-{mode}.png"
    blue_path.write_bytes(replacement)
    output = workdir / f"placeholder-replaced-{mode}.pptx"
    if mode == "plan":
        plan_path = workdir / "placeholder-image-plan.json"
        write_json(plan_path, plan_v1(inventory, [{
            "action": "fill", "slide": slide_ref(inventory, 0),
            "images": [{"shape": image["recommended_ref"], "path": str(blue_path)}],
        }]))
        run("pptx_from_complete_template.py", str(template), str(plan_path), str(output))
    else:
        run("pptx_edit.py", str(template), "--output", str(output),
            "--swap-image", "0", image["name"], str(blue_path))
    swapped = Presentation(output).slides[0].placeholders[1]
    assert swapped.is_placeholder
    assert swapped.image.blob == replacement
    assert tuple(getattr(swapped, name) for name in ("left", "top", "width", "height")) == original_bounds
    assert (swapped.crop_left, swapped.crop_top, swapped.crop_right, swapped.crop_bottom) == original_crop
    exported = run("pptx_read.py", "export-images", str(output), "--outdir",
                   str(workdir / f"placeholder-swapped-{mode}"))
    with open(exported["exports"][0]["output"], "rb") as handle:
        assert handle.read() == replacement


def test_help_flags():
    for script in ("pptx_create.py", "pptx_read.py", "pptx_edit.py",
                   "pptx_from_template.py"):
        proc = subprocess.run(
            [sys.executable, os.path.join(SCRIPTS, script), "--help"],
            capture_output=True, text=True, encoding="utf-8")
        assert proc.returncode == 0 and "usage" in proc.stdout.lower()


# ---------------------------------------------------------------------------
# parity features: render, run-merge replace, chart ops, duplication,
# polish (background/hyperlink/slide-number/footer), notes editing
# ---------------------------------------------------------------------------

def run_raw(script, *args):
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    return subprocess.run(
        [sys.executable, os.path.join(SCRIPTS, script), *args],
        capture_output=True, text=True, encoding="utf-8", env=env)


def test_render_all_slides(workdir):
    import shutil
    spec_path = workdir / "render_spec.json"
    write_json(spec_path, {"schema": "powerpoint.create/v1", "slides": [
        {"layout": "title", "title": "One"},
        {"layout": "title_content", "title": "Two", "bullets": ["b"]},
        {"layout": "blank"}]})
    deck3 = workdir / "render_me.pptx"
    run("pptx_create.py", str(spec_path), str(deck3))
    out_dir = workdir / "render_out"
    out_dir.mkdir(exist_ok=True)
    stale = out_dir / "slide-999.png"
    stale.write_bytes(b"stale")
    result = run("pptx_render.py", str(deck3), "--outdir", str(out_dir))
    assert_schema("powerpoint-render-v1.schema.json", result)
    have_tools = shutil.which("soffice") and (
        shutil.which("pdftoppm") or shutil.which("pdftocairo"))
    if have_tools:
        assert result["schema"] == "powerpoint.render/v1"
        assert result["rendered"] is True
        assert len(result["slides"]) == 3
        assert str(stale) not in [item["output"] for item in result["slides"]]
        for item in result["slides"]:
            assert item["width_px"] > 0 and item["height_px"] > 0
            assert len(item["sha256"]) == 64
            with open(item["output"], "rb") as fh:
                assert fh.read(8) == b"\x89PNG\r\n\x1a\n"
    else:
        assert result["rendered"] is False
        assert result["missing"]
        assert "guidance" in result


def test_replace_across_identically_formatted_runs(workdir):
    """A match split mid-word into equal-format runs keeps formatting."""
    import copy as cp
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1), Inches(1),
                                   Inches(6), Inches(1))
    para = box.text_frame.paragraphs[0]
    run_obj = para.add_run()
    run_obj.text = "Say TotalWord now"
    run_obj.font.bold = True
    run_obj.font.size = Pt(20)
    run_obj.font.color.rgb = RGBColor.from_string("CC0000")
    # simulate PowerPoint's spell-check split: same rPr, seam mid-match
    second = cp.deepcopy(run_obj._r)
    run_obj._r.addnext(second)
    run_obj.text = "Say Total"
    para.runs[1].text = "Word now"
    path = workdir / "split_runs.pptx"
    prs.save(str(path))

    edited = workdir / "split_runs_edited.pptx"
    result = run("pptx_edit.py", str(path), "--output", str(edited),
                 "--replace-text", "TotalWord", "MergedWord")
    assert result["replacements"] == 1

    prs2 = Presentation(str(edited))
    para2 = prs2.slides[0].shapes[0].text_frame.paragraphs[0]
    assert "".join(r.text for r in para2.runs) == "Say MergedWord now"
    for r in para2.runs:
        assert r.font.bold is True
        assert r.font.size == Pt(20)
        assert str(r.font.color.rgb) == "CC0000"


def test_chart_surgical_ops(deck, workdir):
    from pptx import Presentation
    spec = workdir / "chart_ops.json"
    write_json(spec, {"slide": 3, "chart": 0, "ops": [
        {"op": "update_series", "name": "North", "values": [99, 88]},
        {"op": "add_series", "name": "East", "values": [1, 2]},
        {"op": "remove_series", "name": "South"},
        {"op": "rename_category", "from": "Q1", "to": "Q1 FY26"},
        {"op": "set_title", "title": "Updated Sales"}]})
    edited = workdir / "chart_ops.pptx"
    run("pptx_edit.py", deck, "--output", str(edited),
        "--chart-data", str(spec))
    outline = inspect_deck(str(edited))
    bar = slide_charts(outline["slides"][3])[0]
    assert bar["categories"] == ["Q1 FY26", "Q2"]
    names = [s["name"] for s in bar["series"]]
    assert "South" not in names and "East" in names
    north = next(s for s in bar["series"] if s["name"] == "North")
    assert north["values"] == [99.0, 88.0]
    east = next(s for s in bar["series"] if s["name"] == "East")
    assert east["values"] == [1.0, 2.0]
    chart = next(s.chart for s in Presentation(str(edited)).slides[3].shapes
                 if s.has_chart)
    assert chart.chart_title.text_frame.text == "Updated Sales"


def test_chart_op_unknown_series_fails(deck, workdir):
    spec = workdir / "chart_bad.json"
    write_json(spec, {"slide": 3, "chart": 0, "ops": [
        {"op": "update_series", "name": "Nowhere", "values": [1, 2]}]})
    proc = run_raw("pptx_edit.py", deck, "--output",
                   str(workdir / "never.pptx"), "--chart-data", str(spec))
    assert proc.returncode != 0
    assert "Nowhere" in proc.stderr


def test_duplicate_image_slide_is_independent(deck, workdir):
    import base64
    dup = workdir / "dup.pptx"
    result = run("pptx_edit.py", deck, "--output", str(dup),
                 "--duplicate-slide", "2")
    assert result["duplicated_to"] == 4
    outline = inspect_deck(str(dup))  # re-opens cleanly
    assert outline["presentation"]["slide_count"] == 5
    s4 = outline["slides"][4]
    assert len(slide_images(s4)) == 1
    assert slide_tables(s4)[0][0] == ["Region", "Sales"]
    assert "Callout" in slide_texts(s4)

    # independence: swap the image on the COPY, original stays red
    blue = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGNg"
        "YPj/HwADAgH/p5UronAAAAAASUVORK5CYII=")
    blue_path = workdir / "blue2.png"
    blue_path.write_bytes(blue)
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    prs = Presentation(str(dup))
    name = next(s.name for s in prs.slides[4].shapes
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
    run("pptx_edit.py", str(dup), "--swap-image", "4", name,
        str(blue_path))
    prs = Presentation(str(dup))
    orig = next(s for s in prs.slides[2].shapes
                if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
    copy_pic = next(s for s in prs.slides[4].shapes
                    if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
    assert copy_pic.image.blob == blue
    assert orig.image.blob != blue


def test_duplicate_chart_slide_refused(deck, workdir):
    proc = run_raw("pptx_edit.py", deck, "--output",
                   str(workdir / "never2.pptx"), "--duplicate-slide", "3")
    assert proc.returncode != 0
    assert "chart" in proc.stderr.lower()


def test_create_polish_features(workdir):
    from pptx import Presentation
    spec_path = workdir / "polish.json"
    write_json(spec_path, {"schema": "powerpoint.create/v1", "slides": [
        {"layout": "title_content", "title": "Polished",
         "background": "112233", "footer": "Confidential draft",
         "slide_number": True,
         "bullets": [{"text": "Visit example",
                      "link": "https://example.com/info"}]}]})
    out = workdir / "polish.pptx"
    run("pptx_create.py", str(spec_path), str(out))
    prs = Presentation(str(out))
    slide = prs.slides[0]
    assert str(slide.background.fill.fore_color.rgb) == "112233"
    ph_idx = [ph.placeholder_format.idx for ph in slide.placeholders]
    assert 12 in ph_idx  # slide number enabled
    footer = next(ph for ph in slide.placeholders
                  if ph.placeholder_format.idx == 11)
    assert footer.text_frame.text == "Confidential draft"
    link_runs = [r for sh in slide.shapes if sh.has_text_frame
                 for p in sh.text_frame.paragraphs for r in p.runs
                 if r.hyperlink.address]
    assert link_runs[0].hyperlink.address == "https://example.com/info"


def test_edit_polish_features(deck, workdir):
    from pptx import Presentation
    edited = workdir / "polish_edit.pptx"
    result = run("pptx_edit.py", deck, "--output", str(edited),
                 "--set-background", "0", "004400",
                 "--hyperlink", "1", "Overview", "https://example.com/x",
                 "--enable-slide-number", "1",
                 "--set-footer", "1", "Footer via edit")
    assert result["background_set"] and result["footer_set"]
    assert result["hyperlinked_runs"] == 1
    assert result["slide_number_enabled"] is True
    prs = Presentation(str(edited))
    assert str(prs.slides[0].background.fill.fore_color.rgb) == "004400"
    s1 = prs.slides[1]
    assert any(ph.placeholder_format.idx == 12 for ph in s1.placeholders)
    footer = next(ph for ph in s1.placeholders
                  if ph.placeholder_format.idx == 11)
    assert footer.text_frame.text == "Footer via edit"
    links = [r.hyperlink.address for sh in s1.shapes if sh.has_text_frame
             for p in sh.text_frame.paragraphs for r in p.runs
             if r.hyperlink.address]
    assert links == ["https://example.com/x"]


def test_set_and_append_notes(deck, workdir):
    edited = workdir / "notes_edit.pptx"
    result = run("pptx_edit.py", deck, "--output", str(edited),
                 "--set-notes", "0", "Fresh notes")
    assert result["notes_set"]
    run("pptx_edit.py", str(edited), "--append-notes", "0", "Second line")
    notes = inspect_deck(str(edited))
    assert notes["slides"][0]["notes"] == "Fresh notes\nSecond line"


def test_complete_template_generation_plan(deck, workdir):
    inventory = inspect_deck(deck)
    table_ref = first_shape_ref(inventory["slides"][2], "table")
    chart_ref = first_shape_ref(inventory["slides"][3], "chart")

    plan = workdir / "generation_plan.json"
    plan_payload = plan_v1(inventory, [
        {"action": "fill", "slide": slide_ref(inventory, 0),
         "text": [{"shape": {"placeholder": 0},
                   "value": "Complete template report"}]},
        {"action": "duplicate", "slide": slide_ref(inventory, 2),
         "as": "table_copy",
         "tables": [{"shape": table_ref,
                     "rows": [["Region", "Sales"],
                              ["North", "140"], ["South", "90"]]}]},
        {"action": "fill", "slide": slide_ref(inventory, 3),
         "charts": [{"shape": chart_ref,
                     "categories": ["Q3", "Q4"],
                     "series": {"North": [30, 40],
                                "South": [21, 34]}}]},
        {"action": "append",
         "layout": layout_ref(inventory, "Title and Content"),
         "as": "summary", "title": "Summary", "bullets": ["Complete"]},
    ])
    assert_schema("powerpoint-plan-v1.schema.json", plan_payload)
    write_json(plan, plan_payload)
    output = workdir / "complete.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan),
                 str(output))
    assert_schema("powerpoint-plan-result-v1.schema.json", result)
    assert result["schema"] == "powerpoint.plan-result/v1"
    assert result["slides_before"] == 4
    assert result["slides_after"] == 6
    assert result["operations"]["duplicated_slides"] == 1
    assert result["operations"]["updated_tables"] == 1
    assert result["operations"]["updated_charts"] == 1
    assert result["aliases"]["summary"]["index"] == 5
    outline = inspect_deck(str(output))
    assert "Complete template report" in slide_texts(outline["slides"][0])
    assert slide_tables(outline["slides"][4])[0][1] == ["North", "140"]
    assert slide_texts(outline["slides"][5])[0] == "Summary"


def test_complete_template_role_binding(deck, workdir):
    inventory = inspect_deck(deck)
    title_ref = first_shape_ref(inventory["slides"][0], "text")
    plan = workdir / "role_generation_plan.json"
    write_json(plan, plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 0),
        "targets": {"title": title_ref},
        "text": [{"role": "title", "required": True,
                  "value": "Role-bound title"}],
    }], strict=True))
    output = workdir / "role-bound.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan),
                 str(output))
    assert result["operations"]["filled_text"] == 1
    outline = inspect_deck(str(output))
    assert "Role-bound title" in slide_texts(outline["slides"][0])



def test_complete_template_rejects_unbound_role_by_default(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "unbound-role.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 0),
        "text": [{"role": "title", "value": "Should not run"}],
    }]))
    proc = run_raw("pptx_from_complete_template.py", deck, str(plan_path),
                   str(workdir / "unbound-role.pptx"))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert error["error"]["code"] == "unbound_role"


def test_complete_template_infers_common_role(deck, workdir):
    inventory = inspect_deck(deck)
    plan = workdir / "inferred_role_plan.json"
    inferred_plan = plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 0),
        "text": [{"role": "title", "required": True,
                  "value": "Inferred title"}],
    }], strict=True)
    inferred_plan["allow_inferred_roles"] = True
    write_json(plan, inferred_plan)
    output = workdir / "inferred-role.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan),
                 str(output))
    assert result["operations"]["filled_text"] == 1
    outline = inspect_deck(str(output))
    assert "Inferred title" in slide_texts(outline["slides"][0])


def test_complete_template_bullets_require_body_placeholder(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "bullets-on-title-layout.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "append",
        "layout": layout_ref(inventory, "Title Slide"),
        "bullets": ["Must not overwrite subtitle"],
    }]))
    output = workdir / "bullets-on-title-layout.pptx"
    proc = run_raw("pptx_from_complete_template.py", deck, str(plan_path),
                   str(output))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert error["error"]["operation_index"] == 0
    assert "body placeholder" in error["error"]["message"]
    assert not output.exists()


def test_complete_template_rejects_table_dimension_mismatch(deck, workdir):
    inventory = inspect_deck(deck)
    table_ref = first_shape_ref(inventory["slides"][2], "table")
    plan = workdir / "bad_generation_plan.json"
    write_json(plan, plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 2),
        "tables": [{"shape": table_ref, "rows": [["only one row"]]}],
    }]))
    proc = run_raw("pptx_from_complete_template.py", deck, str(plan),
                   str(workdir / "bad_complete.pptx"))
    assert proc.returncode != 0
    error = json.loads(proc.stderr)
    assert "dimensions do not match" in error["error"]["message"]


def test_complete_template_duplicate_then_fill_alias(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "duplicate-alias.json"
    write_json(plan_path, plan_v1(inventory, [
        {"action": "duplicate", "slide": slide_ref(inventory, 0),
         "as": "cover_copy"},
        {"action": "fill", "slide": {"alias": "cover_copy"},
         "text": [{"shape": {"placeholder": 0},
                   "value": "Alias-targeted copy"}]},
    ]))
    output = workdir / "duplicate-alias.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan_path),
                 str(output))
    alias_index = result["aliases"]["cover_copy"]["index"]
    rendered = inspect_deck(output)
    assert "Alias-targeted copy" in slide_texts(rendered["slides"][alias_index])



def test_complete_template_append_then_fill_alias(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "append-alias.json"
    write_json(plan_path, plan_v1(inventory, [
        {"action": "append",
         "layout": layout_ref(inventory, "Title and Content"),
         "as": "new_summary"},
        {"action": "fill", "slide": {"alias": "new_summary"},
         "text": [{"shape": {"placeholder": 0},
                   "value": "Filled after append"}]},
    ]))
    output = workdir / "append-alias.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan_path),
                 str(output))
    alias_index = result["aliases"]["new_summary"]["index"]
    rendered = inspect_deck(output)
    assert "Filled after append" in slide_texts(rendered["slides"][alias_index])


def test_complete_template_original_ref_survives_remove_and_reorder(deck, workdir):
    inventory = inspect_deck(deck)
    target_id = inventory["slides"][3]["slide_id"]
    plan_path = workdir / "stable-slide-ref.json"
    write_json(plan_path, plan_v1(inventory, [
        {"action": "remove", "slide": slide_ref(inventory, 1)},
        {"action": "reorder", "slide": slide_ref(inventory, 3), "to": 0},
    ]))
    output = workdir / "stable-slide-ref.pptx"
    result = run("pptx_from_complete_template.py", deck, str(plan_path),
                 str(output))
    assert result["operations"]["removed_slides"] == 1
    assert result["operations"]["reordered_slides"] == 1
    rendered = inspect_deck(output)
    assert rendered["slides"][0]["slide_id"] == target_id


def test_complete_template_rejects_non_object_payload_item(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "bad-payload-item.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "fill", "slide": slide_ref(inventory, 0),
        "text": ["bad"],
    }]))
    proc = run_raw("pptx_from_complete_template.py", deck, str(plan_path),
                   str(workdir / "bad-payload-item.pptx"))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert error["error"]["code"] == "invalid_type"
    assert "must be an object" in error["error"]["message"]


def test_complete_template_rejects_out_of_range_reorder(deck, workdir):
    inventory = inspect_deck(deck)
    plan_path = workdir / "bad-reorder.json"
    write_json(plan_path, plan_v1(inventory, [{
        "action": "reorder", "slide": slide_ref(inventory, 0), "to": 99,
    }]))
    proc = run_raw("pptx_from_complete_template.py", deck, str(plan_path),
                   str(workdir / "bad-reorder.pptx"))
    assert proc.returncode == 2
    error = json.loads(proc.stderr)
    assert error["error"]["code"] == "invalid_slide_ref"
    assert error["error"]["operation_index"] == 0


def test_render_help_flag():
    proc = run_raw("pptx_render.py", "--help")
    assert proc.returncode == 0 and "usage" in proc.stdout.lower()

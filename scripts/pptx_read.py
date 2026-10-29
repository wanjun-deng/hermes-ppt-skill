#!/usr/bin/env python3
"""Inspect a PowerPoint file or export pictures with stable shape locators.

Commands:
  inspect INPUT
      Emit the canonical powerpoint.inventory/v2 document.

  export-images INPUT --outdir DIR
      Export every embedded picture and emit a manifest that maps each file
      back to the source slide and shape locator.
"""
import argparse
import json
import sys

from pptx.exc import PackageNotFoundError

from pptx_inventory import InventoryError, export_images, inspect_presentation


def write_json(payload, pretty=False, stream=None):
    print(json.dumps(
        payload,
        ensure_ascii=False,
        indent=2 if pretty else None,
        separators=None if pretty else (",", ":"),
    ), file=stream or sys.stdout)


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(
        description="Inspect a .pptx or export pictures with shape locators.",
        epilog=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)

    inspect_cmd = commands.add_parser(
        "inspect", help="emit powerpoint.inventory/v2 JSON")
    inspect_cmd.add_argument("pptx", help="path to the .pptx file")
    inspect_cmd.add_argument(
        "--pretty", action="store_true", help="pretty-print JSON")

    images_cmd = commands.add_parser(
        "export-images", help="export pictures and a locator manifest")
    images_cmd.add_argument("pptx", help="path to the .pptx file")
    images_cmd.add_argument("--outdir", required=True,
                            help="directory for exported pictures")
    images_cmd.add_argument(
        "--pretty", action="store_true", help="pretty-print JSON")

    args = parser.parse_args(argv)
    try:
        if args.command == "inspect":
            result = inspect_presentation(args.pptx)
        else:
            result = export_images(args.pptx, args.outdir)
    except InventoryError as exc:
        write_json({
            "schema": "powerpoint.error/v1",
            "ok": False,
            "error": {"code": "unsafe_inventory", "message": str(exc)},
        }, stream=sys.stderr)
        return 2
    except PackageNotFoundError as exc:
        write_json({
            "schema": "powerpoint.error/v1",
            "ok": False,
            "error": {"code": "invalid_presentation", "message": str(exc)},
        }, stream=sys.stderr)
        return 2
    except OSError as exc:
        write_json({
            "schema": "powerpoint.error/v1",
            "ok": False,
            "error": {"code": "io_error", "message": str(exc)},
        }, stream=sys.stderr)
        return 2
    except (AttributeError, KeyError, TypeError, ValueError) as exc:
        write_json({
            "schema": "powerpoint.error/v1",
            "ok": False,
            "error": {"code": "inspection_failed", "message": str(exc)},
        }, stream=sys.stderr)
        return 2

    write_json(result, pretty=args.pretty)
    return 0


if __name__ == "__main__":
    sys.exit(main())

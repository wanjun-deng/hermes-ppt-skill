#!/usr/bin/env python3
"""Render every slide of a .pptx to validated per-slide PNG images.

LibreOffice converts the deck to PDF and Poppler converts that PDF to PNG.
The command emits a ``powerpoint.render/v1`` manifest. Missing optional tools
are reported with ``rendered: false`` and exit status 0.
"""
import argparse
import glob
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation

from pptx_common import PptxSkillError, write_error, write_json


class RenderError(PptxSkillError):
    code = "render_failed"


def _sha256_file(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _png_size(path):
    with open(path, "rb") as handle:
        header = handle.read(24)
    if len(header) < 24 or header[:8] != b"\x89PNG\r\n\x1a\n":
        raise RenderError(f"renderer produced an invalid PNG: {path}")
    return int.from_bytes(header[16:20], "big"), int.from_bytes(
        header[20:24], "big")


def _natural_page_key(path):
    match = re.search(r"-(\d+)\.png$", os.path.basename(path))
    if not match:
        raise RenderError(f"unexpected rendered filename: {path}")
    return int(match.group(1))


def find_tools():
    """Return (soffice, splitter, missing) using shutil.which."""
    soffice = shutil.which("soffice")
    splitter = shutil.which("pdftoppm") or shutil.which("pdftocairo")
    missing = []
    if not soffice:
        missing.append("soffice")
    if not splitter:
        missing.append("pdftoppm (or pdftocairo)")
    return soffice, splitter, missing


def render(pptx_path, out_dir, prefix, dpi):
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    if not os.path.isfile(pptx_path):
        raise RenderError(f"input presentation not found: {pptx_path}",
                          code="input_not_found")
    if not isinstance(dpi, int) or isinstance(dpi, bool) or dpi <= 0:
        raise RenderError("dpi must be a positive integer", code="invalid_dpi")
    if (not prefix or os.path.basename(prefix) != prefix
            or prefix in {".", ".."}):
        raise RenderError("prefix must be a plain filename prefix",
                          code="invalid_prefix")

    try:
        expected_pages = len(Presentation(pptx_path).slides)
    except Exception as exc:
        raise RenderError(f"cannot open input presentation: {exc}",
                          code="invalid_presentation") from exc
    source = {
        "path": pptx_path,
        "sha256": _sha256_file(pptx_path),
        "slide_count": expected_pages,
    }
    soffice, splitter, missing = find_tools()
    if missing:
        return {
            "schema": "powerpoint.render/v1",
            "ok": True,
            "rendered": False,
            "source": source,
            "missing": missing,
            "guidance": "Install LibreOffice and poppler-utils; use "
                        "pptx_read.py inspect for structural verification.",
        }

    os.makedirs(out_dir, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".pptx-render-", dir=out_dir) as tmp:
        pdf_dir = os.path.join(tmp, "pdf")
        png_dir = os.path.join(tmp, "png")
        os.makedirs(pdf_dir)
        os.makedirs(png_dir)
        profile_uri = Path(tmp, "libreoffice-profile").as_uri()
        conversion_env = dict(os.environ, DCONF_BACKEND="memory")
        try:
            converted = subprocess.run(
                [soffice, f"-env:UserInstallation={profile_uri}",
                 "--headless", "--convert-to", "pdf",
                 "--outdir", pdf_dir, pptx_path],
                capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=300, env=conversion_env)
        except subprocess.TimeoutExpired as exc:
            raise RenderError("LibreOffice conversion timed out",
                              code="render_timeout") from exc
        pdfs = glob.glob(os.path.join(pdf_dir, "*.pdf"))
        if converted.returncode != 0 or len(pdfs) != 1:
            message = converted.stderr.strip() or converted.stdout.strip()
            raise RenderError(
                f"LibreOffice PDF conversion failed (exit {converted.returncode}): "
                f"{message or 'no PDF produced'}")

        temporary_prefix = os.path.join(png_dir, "page")
        try:
            split = subprocess.run(
                [splitter, "-png", "-r", str(dpi), pdfs[0], temporary_prefix],
                capture_output=True, text=True, encoding="utf-8",
                errors="replace", timeout=300)
        except subprocess.TimeoutExpired as exc:
            raise RenderError("Poppler rendering timed out",
                              code="render_timeout") from exc
        if split.returncode != 0:
            raise RenderError(
                f"{os.path.basename(splitter)} failed: {split.stderr.strip()}")

        generated = sorted(
            glob.glob(temporary_prefix + "-*.png"), key=_natural_page_key)
        if len(generated) != expected_pages:
            raise RenderError(
                f"rendered page count mismatch: expected {expected_pages}, "
                f"got {len(generated)}", code="page_count_mismatch")

        slides = []
        for slide_index, temporary_png in enumerate(generated):
            output = os.path.join(
                out_dir, f"{prefix}-{slide_index + 1:03d}.png")
            width, height = _png_size(temporary_png)
            os.replace(temporary_png, output)
            slides.append({
                "slide_index": slide_index,
                "output": output,
                "sha256": _sha256_file(output),
                "width_px": width,
                "height_px": height,
            })

    return {
        "schema": "powerpoint.render/v1",
        "ok": True,
        "rendered": True,
        "source": source,
        "dpi": dpi,
        "slides": slides,
    }


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(
        description="Render each slide of a .pptx to a validated PNG.",
        epilog=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("pptx", help="path to the .pptx file")
    parser.add_argument("--outdir", default="render",
                        help="directory for PNGs (default: ./render)")
    parser.add_argument("--prefix", default="slide",
                        help="PNG filename prefix (default: slide)")
    parser.add_argument("--dpi", type=int, default=100,
                        help="render resolution (default: 100)")
    args = parser.parse_args(argv)
    try:
        result = render(args.pptx, args.outdir, args.prefix, args.dpi)
    except PptxSkillError as exc:
        write_error(exc, stream=sys.stderr)
        return 2
    except (OSError, ValueError, subprocess.SubprocessError) as exc:
        write_error(exc, stream=sys.stderr)
        return 2
    write_json(result, stream=sys.stdout, pretty=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())

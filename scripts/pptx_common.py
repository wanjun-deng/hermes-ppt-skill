#!/usr/bin/env python3
"""Shared CLI output, error, and atomic PowerPoint save helpers."""

import json
import os
import tempfile
import zipfile

from pptx import Presentation


ERROR_SCHEMA = "powerpoint.error/v1"


class PptxSkillError(Exception):
    """A user-correctable PowerPoint skill error."""

    code = "pptx_error"

    def __init__(self, message, *, code=None, details=None):
        super().__init__(message)
        self.code = code or self.code
        self.details = details

    def as_error(self):
        result = {"code": self.code, "message": str(self)}
        if self.details is not None:
            result["details"] = self.details
        return result


def write_json(payload, *, stream, pretty=False):
    print(json.dumps(
        payload,
        ensure_ascii=False,
        indent=2 if pretty else None,
        separators=None if pretty else (",", ":"),
    ), file=stream)


def write_error(error, *, stream):
    if isinstance(error, PptxSkillError):
        error = error.as_error()
    elif isinstance(error, Exception):
        error = {"code": "execution_failed", "message": str(error)}
    write_json({"schema": ERROR_SCHEMA, "ok": False, "error": error},
               stream=stream)


def validate_pptx_file(path):
    """Verify ZIP integrity and that python-pptx can reopen the package."""
    try:
        with zipfile.ZipFile(path) as package:
            bad_member = package.testzip()
    except (OSError, zipfile.BadZipFile) as exc:
        raise PptxSkillError(
            f"invalid PPTX package: {exc}", code="invalid_output") from exc
    if bad_member is not None:
        raise PptxSkillError(
            f"invalid PPTX package member: {bad_member}", code="invalid_output")
    try:
        Presentation(path)
    except Exception as exc:
        raise PptxSkillError(
            f"generated PPTX cannot be reopened: {exc}",
            code="invalid_output") from exc


def atomic_save_presentation(prs, output, *, validator=None):
    """Save, validate, and atomically replace ``output`` on one filesystem."""
    output = os.path.abspath(os.fspath(output))
    parent = os.path.dirname(output)
    if not os.path.isdir(parent):
        raise PptxSkillError(
            f"output directory does not exist: {parent}", code="io_error")
    descriptor, temporary = tempfile.mkstemp(
        prefix=f".{os.path.basename(output)}.", suffix=".tmp.pptx", dir=parent)
    os.close(descriptor)
    try:
        prs.save(temporary)
        validate_pptx_file(temporary)
        if validator is not None:
            validator(temporary)
        if os.path.exists(output):
            os.chmod(temporary, os.stat(output).st_mode)
        os.replace(temporary, output)
    except Exception:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass
        raise
    return output

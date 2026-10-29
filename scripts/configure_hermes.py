#!/usr/bin/env python3
"""Consolidate the authorized local PPT skill entrypoint, with a config backup."""
import argparse
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import tempfile

import yaml


def configure(home, workspace, apply=False):
    home, workspace = Path(home).resolve(), Path(workspace).resolve()
    path = home / "config.yaml"
    original = path.read_text(encoding="utf-8")
    config = yaml.safe_load(original)
    skills = dict(config.get("skills") or {})
    disabled = skills.get("disabled") or []
    disabled = [disabled] if isinstance(disabled, str) else list(disabled)
    disabled = [x for x in disabled if x != "powerpoint-dev"]
    for name in ("powerpoint", "pptx-template-generation"):
        if name not in disabled:
            disabled.append(name)
    external = skills.get("external_dirs") or []
    external = [external] if isinstance(external, str) else list(external)
    if str(workspace) not in external:
        external.append(str(workspace))
    skills.update(disabled=disabled, external_dirs=external)
    replacement = yaml.safe_dump({"skills": skills}, allow_unicode=True, sort_keys=False)
    for key, value in yaml.compose(original).value:
        if key.value == "skills":
            updated = original[:key.start_mark.index] + replacement + original[value.end_mark.index:]
            break
    else:
        updated = original.rstrip() + "\n" + replacement
    if yaml.safe_load(updated) != dict(config, skills=skills):
        raise ValueError("configuration round-trip changed unrelated settings")
    if config.get("skills") == skills:
        updated = original
    link = home / "skills/productivity/powerpoint-dev"
    broken = link.is_symlink() and not link.exists()
    report = {"applied": apply, "config_changed": updated != original,
              "disabled": disabled, "external_dirs": external,
              "broken_link": str(link) if broken else None, "backup": None}
    if apply and (updated != original or broken):
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        backup = home / "backups" / ("powerpoint-dev-" + stamp)
        backup.mkdir(parents=True, mode=0o700)
        (backup / "config.yaml").write_text(original, encoding="utf-8")
        os.chmod(backup / "config.yaml", 0o600)
        if updated != original:
            descriptor, temporary = tempfile.mkstemp(dir=home, prefix=".ppt-config-")
            try:
                with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                    handle.write(updated)
                os.chmod(temporary, path.stat().st_mode & 0o777)
                os.replace(temporary, path)
            finally:
                if os.path.exists(temporary):
                    os.unlink(temporary)
        if broken:
            (backup / "broken-link-target.txt").write_text(os.readlink(link), encoding="utf-8")
            link.rename(backup / "powerpoint-dev.broken-link")
        report["backup"] = str(backup)
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--home", default=os.environ.get("HERMES_HOME", str(Path.home() / ".hermes")))
    parser.add_argument("--workspace", default=str(Path(__file__).resolve().parents[2]))
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()
    print(json.dumps(configure(args.home, args.workspace, args.apply), ensure_ascii=False, indent=2))

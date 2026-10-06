#!/usr/bin/env python3
"""Check exported VBA source, Git storage and CHANGELOG structure without Office."""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path
from typing import Any

from _gatelib import git_bytes, parse_report_args, run_gate, tracked_files

VBA_SUFFIXES = {".bas", ".cls", ".frm"}
# The VBE exports in the Windows code page; SACCR targets Western-European hosts.
VBA_ENCODING = "cp1252"
VB_NAME = re.compile(r'^Attribute VB_Name = "([^"]+)"\s*$', re.M)
OPTION_EXPLICIT = re.compile(r"^[ \t]*Option[ \t]+Explicit[ \t]*(?:'.*)?$", re.M | re.I)
FRX_REFERENCE = re.compile(r'"([^"\r\n]+\.frx)":([0-9A-Fa-f]+)')
SEMVER = r"(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)\.(?:0|[1-9]\d*)"
VERSION_HEADING = re.compile(r"^## \[([^\]]+)\](.*)$", re.M)
RELEASE_SUFFIX = re.compile(r" - (\d{4}-\d{2}-\d{2})")


def index_eol(root: Path) -> dict[str, tuple[str, str]]:
    """Map each tracked path to its (index EOL, attribute) pair from ``git ls-files --eol``."""
    completed = git_bytes(root, "ls-files", "--eol", "-z")
    if completed.returncode:
        raise RuntimeError(completed.stderr.decode("utf-8", errors="replace").strip())
    result: dict[str, tuple[str, str]] = {}
    for record in completed.stdout.split(b"\0"):
        if not record:
            continue
        info, _, path = record.decode("utf-8", errors="surrogateescape").partition("\t")
        fields = info.split()
        result[path] = (fields[0].removeprefix("i/"), " ".join(fields[2:]).removeprefix("attr/"))
    return result


def check_storage(root: Path) -> list[str]:
    """Text must be stored with LF in Git; exported VBA must also check out as CRLF."""
    findings = []
    for path, (eol, attr) in sorted(index_eol(root).items()):
        if eol in {"crlf", "mixed"}:
            findings.append(f"{path}: stored with {eol.upper()} in Git; renormalize to LF")
        # Git reports lone CR line endings as non-text; that is a defect only for declared text.
        elif eol == "-text" and attr.split()[:1] == ["text"]:
            findings.append(f"{path}: declared text but Git classifies the blob as non-text (lone CR?)")
        if Path(path).suffix.lower() in VBA_SUFFIXES and "eol=crlf" not in attr:
            findings.append(f"{path}: .gitattributes must check VBA source out as CRLF")
    return findings


def check_component(root: Path, path: str, tracked: set[str], names: dict[str, str]) -> list[str]:
    try:
        text = (root / path).read_bytes().decode(VBA_ENCODING)
    except UnicodeDecodeError as error:
        return [f"{path}: not decodable as {VBA_ENCODING}: {error}"]
    # Windows checkouts are CRLF; normalize so line anchors behave on every host.
    text = text.replace("\r\n", "\n")
    findings = []
    matches = VB_NAME.findall(text)
    if matches != [Path(path).stem]:
        findings.append(f"{path}: VB_Name must match the filename exactly")
    elif matches[0].casefold() in names:
        findings.append(f"{path}: duplicate component name (also {names[matches[0].casefold()]})")
    else:
        names[matches[0].casefold()] = path
    if not OPTION_EXPLICIT.search(text):
        findings.append(f"{path}: missing Option Explicit")
    if path.lower().endswith(".frm"):
        companions = FRX_REFERENCE.findall(text)
        if not companions:
            findings.append(f"{path}: no form resource reference")
        for filename, offset in companions:
            companion = (Path(path).parent / filename).as_posix()
            if Path(filename).name != filename or companion not in tracked:
                findings.append(f"{path}: missing or unsafe form resource {filename}")
            elif (root / companion).stat().st_size <= int(offset, 16):
                findings.append(f"{path}: resource offset is outside {filename}")
    return findings


def valid_date(text: str) -> bool:
    try:
        date.fromisoformat(text)
    except ValueError:
        return False
    return True


def check_changelog(root: Path) -> list[str]:
    path = root / "CHANGELOG.md"
    if not path.is_file():
        return ["CHANGELOG.md is missing"]
    text = path.read_text(encoding="utf-8")
    headings = VERSION_HEADING.findall(text)
    findings = []
    if not headings or headings[0] != ("Unreleased", ""):
        findings.append("CHANGELOG.md: the first version heading must be '## [Unreleased]'")
    seen = set()
    for label, suffix in headings:
        if label in seen:
            findings.append(f"CHANGELOG.md: duplicate heading [{label}]")
        seen.add(label)
        if label == "Unreleased":
            continue
        dated = RELEASE_SUFFIX.fullmatch(suffix)
        if not re.fullmatch(SEMVER, label) or not dated:
            findings.append(f"CHANGELOG.md: release heading must be '## [X.Y.Z] - YYYY-MM-DD': [{label}]{suffix}")
        elif not valid_date(dated.group(1)):
            findings.append(f"CHANGELOG.md: [{label}] date {dated.group(1)} is not a calendar date")
        if not re.search(rf"^\[{re.escape(label)}\]: \S+$", text, re.M):
            findings.append(f"CHANGELOG.md: missing link reference for [{label}]")
    if "Unreleased" in seen and not re.search(r"^\[Unreleased\]: \S+$", text, re.M):
        findings.append("CHANGELOG.md: missing link reference for [Unreleased]")
    return findings


def run_check(root: Path) -> dict[str, Any]:
    tracked = tracked_files(root)
    findings = check_storage(root)
    names: dict[str, str] = {}
    components = sorted(p for p in tracked if Path(p).suffix.lower() in VBA_SUFFIXES)
    for path in components:
        findings.extend(check_component(root, path, tracked, names))
    findings.extend(check_changelog(root))
    return {"schema_version": 1, "status": "fail" if findings else "pass",
            "components": len(components), "findings": findings}


def markdown_report(report: dict[str, Any]) -> str:
    lines = [f"Source integrity: {report['status'].upper()} ({report['components']} VBA component(s))"]
    lines.extend(report["findings"])
    return "\n".join(lines) + "\n"


def main() -> int:
    options = parse_report_args(sys.argv[1:], description=__doc__)
    return run_gate(options, build=lambda: run_check(options.root.resolve()),
                    markdown=markdown_report, errors=(OSError, RuntimeError, ValueError))


if __name__ == "__main__":
    raise SystemExit(main())

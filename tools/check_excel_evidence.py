#!/usr/bin/env python3
"""Validate a manual Excel evidence bundle against one exact candidate commit.

The validator never runs Excel and cannot authenticate what an operator did. It
checks that the record is complete and consistent, that the declared source
inventory is the exact Git content of the candidate, and that the retained logs
match their digests and the harness report. See docs/EXCEL_EVIDENCE.md.
Adapted from EXCEL-VBA-PROJECT-TEMPLATE b903fe44: manual execution only, and
the source inventory is derived from the repository layout.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Any

from _gatelib import git_bytes, run_gate

DESCRIPTION = "Validate a manual Excel evidence bundle against one exact candidate commit."
POLICY = ".github/excel-evidence-policy.json"
STAGES = ("import", "compile", "regression", "cleanup")
OUTCOMES = {"import": "IMPORT_FAILED", "compile": "COMPILE_FAILED",
            "regression": "TEST_FAILED", "cleanup": "CLEANUP_FAILED"}
# Imported components: production and tests, never examples (docs/REPOSITORY_STRUCTURE.md).
SOURCE_ROOTS = ("src", "tests")
VBA_SUFFIXES = (".bas", ".cls", ".frm")
SHA = re.compile(r"[0-9a-f]{40}")
DIGEST = re.compile(r"[0-9a-fA-F]{64}")
SCOPE = ("Checks a retained manual record and its source and log bindings; "
         "does not run Excel or authenticate the operator.")
RECORD_KEYS = ("schema_version repository candidate_sha execution availability_reason "
               "started_at finished_at operator environment sources stages harness")
ENVIRONMENT_KEYS = ("excel_version excel_build office_bitness os os_architecture runtime "
                    "references macro_policy vba_project_access trust_changes")


def require(condition: object, message: str) -> None:
    if not condition:
        raise ValueError(message)


def object_keys(value: Any, keys: str, label: str) -> dict[str, Any]:
    """Return ``value`` when it is an object with exactly ``keys``, else fail."""
    if not (isinstance(value, dict) and set(value) == set(keys.split())):
        raise ValueError(f"{label}: invalid object fields")
    return value


def nonempty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def relative(value: Any) -> bool:
    return (nonempty(value) and "\\" not in value and "\0" not in value
            and not PurePosixPath(value).is_absolute() and ".." not in value.split("/")
            and PurePosixPath(value).as_posix() == value and value != ".")


def decode(raw: bytes) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key: {key}")
            result[key] = value
        return result
    return json.loads(raw.decode("utf-8-sig"), object_pairs_hook=unique)


def string_list(value: Any, label: str) -> list[str]:
    require(isinstance(value, list) and all(nonempty(item) for item in value),
            f"invalid {label}")
    require(len(value) == len(set(value)), f"duplicate {label}")
    return list(value)


def require_commit(root: Path, sha: str) -> None:
    require(bool(SHA.fullmatch(sha)), "candidate must be a full 40-character lowercase SHA")
    found = git_bytes(root, "cat-file", "-e", f"{sha}^{{commit}}")
    require(found.returncode == 0, f"candidate commit {sha} is not in this repository")


def committed(root: Path, sha: str, path: str) -> bytes:
    require(relative(path), f"unsafe path: {path}")
    result = git_bytes(root, "show", f"{sha}:{path}")
    require(result.returncode == 0, f"candidate does not contain {path}")
    return result.stdout


def source_inventory(root: Path, sha: str) -> list[dict[str, str]]:
    """Return the imported components at ``sha`` with SHA-256 of their exact Git bytes."""
    listing = git_bytes(root, "ls-tree", "-r", "-z", "--name-only", sha, "--", *SOURCE_ROOTS)
    require(listing.returncode == 0, "cannot list candidate source")
    tracked = {item.decode("utf-8") for item in listing.stdout.split(b"\0") if item}
    paths = {path for path in tracked if path.lower().endswith(VBA_SUFFIXES)}
    # A form's binary resource is imported with it, so it is part of the identity.
    for path in tuple(paths):
        if path.lower().endswith(".frm") and path[:-4] + ".frx" in tracked:
            paths.add(path[:-4] + ".frx")
    require(bool(paths), "candidate has no importable VBA source")
    return [{"path": path, "sha256": hashlib.sha256(committed(root, sha, path)).hexdigest()}
            for path in sorted(paths)]


def load_policy(root: Path, sha: str) -> dict[str, Any]:
    policy = object_keys(decode(committed(root, sha, POLICY)),
                         "schema_version repository entry_point cases assertions "
                         "expected_error_cases required_references", "evidence policy")
    require(type(policy["schema_version"]) is int and policy["schema_version"] == 1,
            "unsupported evidence policy schema")
    require(nonempty(policy["repository"]) and nonempty(policy["entry_point"]),
            "policy needs a repository and an entry point")
    require(type(policy["assertions"]) is int and policy["assertions"] > 0,
            "policy needs a positive assertion count")
    cases = string_list(policy["cases"], "policy cases")
    errors = string_list(policy["expected_error_cases"], "policy expected-error cases")
    string_list(policy["required_references"], "policy references")
    require(bool(cases) and set(errors) <= set(cases),
            "expected-error cases must belong to the declared suite")
    return policy


def validate_identity(value: Any, policy: dict[str, Any], sha: str) -> dict[str, Any]:
    record = object_keys(value, RECORD_KEYS, "evidence record")
    require(type(record["schema_version"]) is int and record["schema_version"] == 1,
            "unsupported evidence record schema")
    require(record["repository"] == policy["repository"], "record repository differs from policy")
    require(record["candidate_sha"] == sha, "record candidate_sha differs from the candidate")
    require(record["execution"] in ("manual", "unavailable"), "execution must be manual or unavailable")
    times = []
    for key in ("started_at", "finished_at"):
        require(nonempty(record[key]), f"{key} is required")
        stamp = datetime.fromisoformat(record[key].replace("Z", "+00:00"))
        require(stamp.utcoffset() is not None, f"{key} needs a timezone")
        times.append(stamp)
    require(times[1] >= times[0], "finished_at precedes started_at")
    return record


def validate_environment(record: dict[str, Any], policy: dict[str, Any]) -> None:
    require(nonempty(record["operator"]), "operator and workstation are required")
    environment = object_keys(record["environment"], ENVIRONMENT_KEYS, "environment")
    require(environment["trust_changes"] is False, "a run must not change macro trust settings")
    require(all(nonempty(value) for key, value in environment.items()
                if key not in ("references", "trust_changes")),
            "environment fields must be nonempty")
    require(environment["office_bitness"] in ("32-bit", "64-bit"), "invalid Office bitness")
    require(environment["os"].casefold().startswith("windows"), "Excel for Windows is required")
    references = string_list(environment["references"], "references")
    require(sorted(references) == sorted(policy["required_references"]),
            "references differ from the required default set")


def retained_log(directory: Path, value: Any) -> str:
    log = object_keys(value, "path sha256", "log reference")
    require(relative(log["path"]), "unsafe log path")
    require(isinstance(log["sha256"], str) and DIGEST.fullmatch(log["sha256"]),
            "log digest must be 64 hexadecimal characters")
    path: Path = directory / log["path"]
    require(not any(item.is_symlink() for item in (path, *path.parents)), "symlinked log path")
    require(path.resolve().is_relative_to(directory.resolve()) and path.is_file(),
            f"missing log: {log['path']}")
    raw = path.read_bytes()
    require(hashlib.sha256(raw).hexdigest() == log["sha256"].lower(),
            f"log digest mismatch: {log['path']}")
    return raw.decode("utf-8-sig").replace("\r\n", "\n")


def validate_stages(record: dict[str, Any], directory: Path) -> tuple[list[str], dict[str, str]]:
    stages = object_keys(record["stages"], " ".join(STAGES), "stages")
    outcomes: list[str] = []
    logs: dict[str, str] = {}
    for name in STAGES:
        stage = object_keys(stages[name], "status detail log", f"{name} stage")
        status = stage["status"]
        require(status in ("PASS", "FAIL", "NOT_RUN", "TIMEOUT"), f"invalid {name} status")
        require(nonempty(stage["detail"]), f"{name} needs a detail")
        if status == "NOT_RUN":
            require(stage["log"] is None, f"NOT_RUN {name} cannot cite a log")
            outcomes.append("INCOMPLETE")
            continue
        logs[name] = retained_log(directory, stage["log"])
        if status != "PASS":
            outcomes.append("EXECUTION_TIMEOUT" if status == "TIMEOUT" else OUTCOMES[name])
    status = {name: stages[name]["status"] for name in STAGES}
    require(status["import"] == "PASS" or status["compile"] == "NOT_RUN",
            "compile cannot run after a failed import")
    require(status["compile"] == "PASS" or status["regression"] == "NOT_RUN",
            "regression cannot run after a failed compile")
    require("import" not in logs or record["candidate_sha"] in logs["import"],
            "import log must name the candidate SHA")
    return list(dict.fromkeys(outcomes)), logs


def single(log: str, key: str) -> str:
    values: list[str] = re.findall(rf"^{re.escape(key)}=(.*)$", log, re.MULTILINE)
    require(len(values) == 1, f"harness log needs exactly one {key}= line")
    return values[0].strip()


def lines(log: str, key: str, partial: bool) -> list[str]:
    """Return the ``key=`` values; a full report has exactly one, a partial one at most one."""
    values: list[str] = [value.strip() for value in
                         re.findall(rf"^{re.escape(key)}=(.*)$", log, re.MULTILINE)]
    require(len(values) <= 1 if partial else len(values) == 1,
            f"harness log needs {'at most' if partial else 'exactly'} one {key}= line")
    return values


def report_cases(record: dict[str, Any], policy: dict[str, Any], log: str,
                 partial: bool = False) -> list[str]:
    """Check the report header and return its CASE lines, which must follow the policy order.

    A partial report, kept from a timed-out run, is checked only on the lines it printed.
    """
    runs = re.findall(r"^SACCR TESTS[ \t]*$", log, re.MULTILINE)
    require(len(runs) <= 1 if partial else len(runs) == 1,
            "harness log must contain exactly one run")
    require(all(mode == "NORMAL" for mode in lines(log, "MODE", partial)),
            "harness log is not a normal-mode run")
    environment = record["environment"]
    for summary in lines(log, "ENVIRONMENT", partial):
        reported: dict[str, str] = {}
        for item in summary.split("; "):
            key, _, value = item.partition("=")
            reported[key] = value
        for key, field in (("version", "excel_version"), ("office", "office_bitness"),
                           ("runtime", "runtime")):
            require(reported.get(key) == environment[field],
                    f"harness ENVIRONMENT {key} differs from the record's {field}")
    cases = [case.strip() for case in re.findall(r"^CASE=(.*)$", log, re.MULTILINE)]
    require(cases == policy["cases"][:len(cases)], "harness CASE lines differ from the policy")
    require(len(cases) == record["harness"]["cases"], "harness CASE lines differ from the record")
    return cases


def validate_report(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    """Bind a passing or failed regression to one complete normal-mode harness report."""
    harness = record["harness"]
    report_cases(record, policy, log)
    for field in ("cases", "assertions", "failures"):
        require(single(log, field.upper()) == str(harness[field]),
                f"harness {field.upper()} differs from the record")
    cleanup = single(log, "CLEANUP").split(";")[0]
    require(cleanup in ("PASS", "FAIL"), "harness CLEANUP must be PASS or FAIL")
    # The harness reports COMPLETE only when both expected counts were reached.
    complete = (harness["cases"] == len(policy["cases"])
                and harness["assertions"] == policy["assertions"])
    require(harness["completeness"] == ("COMPLETE" if complete else "INCOMPLETE"),
            "harness completeness contradicts the counts")
    passed = complete and harness["failures"] == 0 and cleanup == "PASS"
    require(single(log, "RESULT") == (
        f"{'PASS' if passed else 'FAIL'}; completeness={harness['completeness']}; "
        f"cases={harness['cases']}; assertions={harness['assertions']}; "
        f"failures={harness['failures']}; cleanup={cleanup}"),
        "harness RESULT contradicts the record")
    regression = record["stages"]["regression"]["status"]
    require(regression == "FAIL" or (complete and harness["failures"] == 0),
            "regression PASS contradicts the harness results")
    require(regression == "PASS" or not passed, "regression FAIL contradicts a passing report")
    require(cleanup == "PASS" or record["stages"]["cleanup"]["status"] == "FAIL",
            "a harness cleanup failure must be recorded as cleanup FAIL")


def validate_partial_report(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    """A timed-out run keeps what was printed before it stopped; it never reached RESULT."""
    require(record["harness"]["completeness"] == "INCOMPLETE", "a timed-out run is INCOMPLETE")
    require(not re.search(r"^RESULT=", log, re.MULTILINE),
            "a log with a RESULT line finished; record it as PASS or FAIL, not TIMEOUT")
    report_cases(record, policy, log, partial=True)


def validate_harness(record: dict[str, Any], policy: dict[str, Any], log: str) -> None:
    harness = object_keys(record["harness"], "entry_point cases assertions failures "
                          "completeness expected_errors", "harness")
    require(harness["entry_point"] == policy["entry_point"], "wrong harness entry point")
    for field in ("cases", "assertions", "failures"):
        require(type(harness[field]) is int and harness[field] >= 0, f"invalid harness {field}")
    require(harness["completeness"] in ("COMPLETE", "INCOMPLETE"), "invalid completeness")
    require(isinstance(harness["expected_errors"], list), "expected_errors must be an array")
    observed = []
    for error in harness["expected_errors"]:
        error = object_keys(error, "case status detail", "expected error")
        require(error["status"] in ("PASS", "FAIL", "NOT_RUN") and nonempty(error["detail"]),
                "invalid expected-error result")
        observed.append(error["case"])
    require(observed == policy["expected_error_cases"], "expected-error results differ from policy")
    status = record["stages"]["regression"]["status"]
    if status == "TIMEOUT":
        validate_partial_report(record, policy, log)
        return
    require(status == "FAIL" or all(error["status"] == "PASS" for error in harness["expected_errors"]),
            "regression PASS needs every expected-error result to pass")
    validate_report(record, policy, log)


def validate_unavailable(record: dict[str, Any]) -> list[str]:
    require(nonempty(record["availability_reason"]), "unavailable needs a reason")
    require(all(record[key] is None for key in ("operator", "environment", "harness"))
            and record["sources"] == [] and record["stages"] == {},
            "unavailable cannot carry execution claims")
    return ["UNAVAILABLE"]


def validate_manual(root: Path, sha: str, policy: dict[str, Any], record: dict[str, Any],
                    directory: Path) -> list[str]:
    require(record["availability_reason"] is None, "a manual run has no unavailability reason")
    validate_environment(record, policy)
    require(record["sources"] == source_inventory(root, sha),
            "source inventory differs from the candidate's exact source")
    outcomes, logs = validate_stages(record, directory)
    if record["stages"]["regression"]["status"] == "NOT_RUN":
        require(record["harness"] is None, "a regression that did not run has no harness results")
    else:
        validate_harness(record, policy, logs["regression"])
    return outcomes


def evaluate(root: Path, sha: str, evidence: Path) -> dict[str, Any]:
    execution = "unresolved"
    outcomes: list[str] = []
    findings: list[str] = []
    try:
        require_commit(root, sha)
        policy = load_policy(root, sha)
        record = validate_identity(decode(evidence.read_bytes()), policy, sha)
        execution = record["execution"]
        if execution == "unavailable":
            outcomes = validate_unavailable(record)
        else:
            outcomes = validate_manual(root, sha, policy, record, evidence.parent)
    except (ValueError, TypeError, KeyError, AttributeError, OSError) as error:
        outcomes.append("EVIDENCE_INVALID")
        findings.append(str(error))
    return {"status": "fail" if outcomes else "pass", "candidate_sha": sha,
            "execution": execution, "outcomes": outcomes or ["PASS"],
            "findings": findings, "scope": SCOPE}


def markdown(report: dict[str, Any]) -> str:
    lines = [f"# Excel evidence: {report['status'].upper()}", "",
             f"Candidate: `{report['candidate_sha']}`; execution: `{report['execution']}`", "",
             "Outcomes: " + ", ".join(report["outcomes"]), ""]
    lines += [f"- {item}" for item in report["findings"]]
    return "\n".join(lines + ["", report["scope"], ""])


def main(arguments: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=DESCRIPTION)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--candidate-sha", required=True)
    parser.add_argument("--evidence", type=Path, help="host.json inside the retained bundle")
    parser.add_argument("--inventory", action="store_true",
                        help="print the candidate's source inventory for the record and exit")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--summary", type=Path)
    options = parser.parse_args(arguments)
    if options.inventory:
        try:
            require_commit(options.root, options.candidate_sha)
            print(json.dumps(source_inventory(options.root, options.candidate_sha), indent=2))
            return 0
        except (ValueError, OSError) as error:
            print(f"ERROR: {error}", file=sys.stderr)
            return 2
    if options.evidence is None:
        parser.error("--evidence is required unless --inventory is given")
    bundle = options.evidence.resolve().parent
    for path in (options.output, options.summary):
        if path is not None and path.resolve().is_relative_to(bundle):
            parser.error("reports must be written outside the evidence bundle")
    return run_gate(options, build=lambda: evaluate(options.root, options.candidate_sha,
                                                    options.evidence),
                    markdown=markdown, errors=(OSError,))


if __name__ == "__main__":
    raise SystemExit(main())

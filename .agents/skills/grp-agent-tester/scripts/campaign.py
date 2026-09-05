#!/usr/bin/env python3
"""Create and run reproducible Agent 1 GRP test campaigns."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import random
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
ROOT = SKILL_DIR.parents[2]
sys.path.insert(0, str(ROOT / "tools"))
from schedule_ir import ScheduleProject, validate_schedule_ir  # noqa: E402

SAFE_CASE_ID = re.compile(r"^[a-z0-9][a-z0-9._-]{0,79}$")
FACADES = ("НВФ", "СФТК", "модульный")
ENCLOSURES = ("трубошпунт", "СВГ", "БСС", "отвал")
GLAZING = (
    {"пвх": True, "витражи": False, "витражи_на_всю_высоту": False},
    {"пвх": True, "витражи": True, "витражи_на_всю_высоту": False},
    {"пвх": False, "витражи": True, "витражи_на_всю_высоту": True},
)
PILE_SETS = (
    (("БНС", 0, 1), ("забивные", 0, 1)),
    (("БНС", 320, 2), ("забивные", 0, 1)),
    (("БНС", 0, 1), ("забивные", 480, 2)),
    (("БНС", 240, 1), ("забивные", 360, 2)),
)
CONDITIONS = ("ППТ", "снос_застройки", "вынос_сетей", "зона_метро", "ВРИ", "ПЗЗ", "ОКН", "СЗЗ")


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest().upper()


def make_spec(index: int, rng: random.Random) -> tuple[dict, dict]:
    corpus_count = (1, 2, 3)[index % 3]
    underground = (0, 1, 2)[(index // 3) % 3]
    facade = FACADES[(index // 2) % len(FACADES)]
    parking = index % 4 in (1, 2)
    pile_profile = (index // 4) % len(PILE_SETS)
    finish_share = (0.0, 0.5, 1.0)[(index // 5) % 3]
    base_floors = (9, 18, 30, 45, 55, 70)[(index * 5 + index // 2) % 6]
    corpuses = []
    glazing_labels = []
    for number in range(corpus_count):
        glazing_index = (index + number) % len(GLAZING)
        glazing = dict(GLAZING[glazing_index])
        glazing_labels.append(("pvc", "mixed", "full_stained")[glazing_index])
        corpuses.append({
            "код": f"К{number + 1}",
            "этажей_надземных": base_floors + number * (7 + index % 5),
            "этажей_подземных": underground,
            "секций": 1 + (index + number) % 4,
            "сложный_конструктив": bool((index + number) % 2),
            "остекление": glazing,
        })
    mask = rng.randrange(1 << len(CONDITIONS))
    conditions = {name: bool(mask & (1 << position)) for position, name in enumerate(CONDITIONS)}
    start = date(2026, 1, 15) + timedelta(days=index * 11)
    piles = PILE_SETS[pile_profile]
    spec = {
        "название": f"Campaign case {index + 1:03d}",
        "старт_проекта": start.strftime("%d.%m.%Y"),
        "корпуса": corpuses,
        "паркинг": {"есть": parking, "код": "П1", "этажей_подземных": 1 + index % 2},
        "стилобат": {"есть": index % 6 == 0, "этажей_надземных": 1 if index % 6 == 0 else 0},
        "фасад": {"тип": facade},
        "нулевой_цикл": {
            "ограждение_котлована": ENCLOSURES[index % len(ENCLOSURES)],
            "сваи": [{"тип": kind, "количество": count, "установок": rigs} for kind, count, rigs in piles],
        },
        "отделка": {"доля_квартир_с_чистовой": finish_share},
        "индивидуальные_условия": conditions,
    }
    dimensions = {
        "corpus_count": corpus_count,
        "floors": [item["этажей_надземных"] for item in corpuses],
        "underground_floors": underground,
        "facade": facade,
        "glazing": glazing_labels,
        "parking": parking,
        "enclosure": spec["нулевой_цикл"]["ограждение_котлована"],
        "pile_profile": pile_profile,
        "finish_share": finish_share,
        "active_conditions": [name for name, enabled in conditions.items() if enabled],
        "support": "exploratory" if corpus_count >= 3 else "confirmed",
    }
    return spec, dimensions


def make_invalid_specs() -> list[tuple[str, dict]]:
    base, _ = make_spec(0, random.Random(0))
    mutations = (
        ("missing-corpuses", lambda value: value.pop("корпуса")),
        ("missing-underground", lambda value: [item.pop("этажей_подземных") for item in value["корпуса"]]),
        ("missing-facade", lambda value: value.pop("фасад")),
        ("missing-glazing", lambda value: [item.pop("остекление") for item in value["корпуса"]]),
        ("missing-piles", lambda value: value["нулевой_цикл"].pop("сваи")),
    )
    result = []
    for label, mutate in mutations:
        candidate = json.loads(json.dumps(base, ensure_ascii=False))
        mutate(candidate)
        candidate["название"] = f"Negative {label}"
        result.append((label, candidate))
    return result


def create_campaign(path: Path, count: int, seed: int, include_invalid: bool) -> int:
    if path.exists() or path.is_symlink():
        raise SystemExit(f"Campaign path already exists: {path}")
    data_dir = (ROOT / "data").resolve()
    try:
        path.resolve().relative_to(data_dir)
    except ValueError:
        pass
    else:
        raise SystemExit("Campaign path must not be inside data/")
    inputs = path / "inputs"
    inputs.mkdir(parents=True)
    rng = random.Random(seed)
    cases = []
    for index in range(count):
        spec, dimensions = make_spec(index, rng)
        case_id = f"valid-{index + 1:03d}"
        target = inputs / f"{case_id}.json"
        target.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        cases.append({"id": case_id, "input": str(target.relative_to(path)), "expected": "accept", "dimensions": dimensions})
    if include_invalid:
        for label, spec in make_invalid_specs():
            case_id = f"invalid-{label}"
            target = inputs / f"{case_id}.json"
            target.write_text(json.dumps(spec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            cases.append({"id": case_id, "input": str(target.relative_to(path)), "expected": "reject", "dimensions": {"negative": label}})
    manifest = {"schema_version": 1, "seed": seed, "valid_count": count, "include_invalid": include_invalid, "cases": cases}
    (path / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Created {len(cases)} cases in {path}")
    return 0


def run_command(command: list[str], log_path: Path, timeout: int) -> tuple[int, float, bool]:
    started = time.monotonic()
    child_env = os.environ.copy()
    child_env["PYTHONUTF8"] = "1"
    try:
        completed = subprocess.run(command, cwd=ROOT, env=child_env, text=True, encoding="utf-8",
                                   errors="replace", capture_output=True, timeout=timeout)
        code, timed_out = completed.returncode, False
        output = completed.stdout + ("\n[stderr]\n" + completed.stderr if completed.stderr else "")
    except subprocess.TimeoutExpired as exc:
        code, timed_out = 124, True
        output = (exc.stdout or "") + "\n[TIMEOUT]\n" + (exc.stderr or "")
    log_path.write_text(output, encoding="utf-8")
    return code, round(time.monotonic() - started, 3), timed_out


def failure_signature(text: str, code: int) -> str:
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if any("Traceback" in line for line in lines):
        for line in reversed(lines):
            if "Error" in line or "Exception" in line:
                return line[:240]
        return "uncontrolled traceback"
    for item in lines:
        if item.startswith("!") or "ОШИБКА" in item or "НАРУШЕН" in item:
            return item[:240]
    return f"exit-{code}"


def contained_path(root: Path, value: str) -> Path:
    candidate = (root / value).resolve()
    try:
        candidate.relative_to(root.resolve())
    except ValueError as exc:
        raise SystemExit(f"Manifest path escapes campaign directory: {value}") from exc
    return candidate


def validate_ir(path: Path, log_path: Path) -> tuple[int, list[str]]:
    try:
        schedule = ScheduleProject.from_json(path.read_text(encoding="utf-8"))
        issues = validate_schedule_ir(schedule)
        messages = [f"{item.code}: {item.message}" for item in issues]
        log_path.write_text("\n".join(messages) + ("\n" if messages else "OK\n"), encoding="utf-8")
        return (1 if issues else 0), messages
    except Exception as exc:  # malformed output is a test failure, not a campaign crash
        message = f"IR validation crashed: {type(exc).__name__}: {exc}"
        log_path.write_text(message + "\n", encoding="utf-8")
        return 2, [message]


def previous_results(path: Path) -> dict[str, dict]:
    candidates = sorted((path / "runs").glob("*/results.json")) if (path / "runs").exists() else []
    if not candidates:
        return {}
    data = json.loads(candidates[-1].read_text(encoding="utf-8"))
    return {item["id"]: item for item in data.get("results", [])}


def coverage(results: list[dict]) -> dict[str, list]:
    accepted = [item["dimensions"] for item in results if item["expected"] == "accept"]
    keys = ("corpus_count", "underground_floors", "facade", "parking", "enclosure",
            "pile_profile", "finish_share", "support")
    return {key: sorted({item.get(key) for item in accepted}, key=str) for key in keys}


def run_campaign(path: Path, timeout: int) -> int:
    manifest_path = path / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    prior = previous_results(path)
    prepared_cases = []
    seen_ids = set()
    for case in manifest.get("cases", []):
        case_id = case.get("id", "")
        if not SAFE_CASE_ID.fullmatch(case_id) or case_id in seen_ids:
            raise SystemExit(f"Unsafe or duplicate case id: {case_id}")
        seen_ids.add(case_id)
        input_path = contained_path(path, str(case.get("input", "")))
        if not input_path.is_file():
            raise SystemExit(f"Campaign input does not exist or is not a file: {input_path}")
        prepared_cases.append((case, input_path))
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    run_dir = path / "runs" / run_id
    run_dir.mkdir(parents=True)
    results = []
    for case, input_path in prepared_cases:
        case_id = case["id"]
        case_dir = run_dir / case_id
        case_dir.mkdir()
        xlsx = case_dir / f"{case_id}.xlsx"
        ir = case_dir / f"{case_id}.ir.json"
        build_code, build_seconds, timed_out = run_command(
            [sys.executable, str(ROOT / "tools" / "build_grp.py"), str(input_path), str(xlsx), "--ir", str(ir)],
            case_dir / "build.log", timeout,
        )
        excel_code = excel_seconds = ir_code = None
        ir_issues = []
        if build_code == 0 and xlsx.exists():
            excel_code, excel_seconds, _ = run_command(
                [sys.executable, str(ROOT / "tools" / "validate_grp.py"), str(xlsx)],
                case_dir / "excel-validation.log", timeout,
            )
        if build_code == 0 and ir.exists():
            ir_code, ir_issues = validate_ir(ir, case_dir / "ir-validation.log")
        log_text = (case_dir / "build.log").read_text(encoding="utf-8")
        controlled_rejection = build_code != 0 and "Traceback" not in log_text and not timed_out
        passed = ((case["expected"] == "accept" and build_code == 0 and excel_code == 0 and ir_code == 0)
                  or (case["expected"] == "reject" and controlled_rejection))
        ir_hash = digest(ir) if ir.exists() else None
        prior_ir_hash = prior.get(case_id, {}).get("ir_sha256")
        results.append({
            "id": case_id, "expected": case["expected"], "passed": passed,
            "build_exit": build_code, "excel_validation_exit": excel_code,
            "ir_validation_exit": ir_code, "ir_issues": ir_issues, "timed_out": timed_out,
            "build_seconds": build_seconds, "excel_validation_seconds": excel_seconds,
            "input_sha256": digest(input_path), "xlsx_sha256": digest(xlsx) if xlsx.exists() else None,
            "ir_sha256": ir_hash, "previous_ir_sha256": prior_ir_hash,
            "ir_changed_since_previous": prior_ir_hash is not None and prior_ir_hash != ir_hash,
            "failure_signature": (None if passed else
                                  (f"unexpectedly accepted: {case['dimensions'].get('negative')}"
                                   if case["expected"] == "reject" and build_code == 0 else
                                   (ir_issues[0] if ir_issues else failure_signature(log_text, build_code)))),
            "dimensions": case.get("dimensions", {}),
        })
        print(f"{'PASS' if passed else 'FAIL'} {case_id}: build={build_code}, excel={excel_code}, ir={ir_code}")
    context_path = ROOT / "instructions" / "context-manifest.json"
    context = json.loads(context_path.read_text(encoding="utf-8"))
    summary = {
        "run_id": run_id, "created_at": datetime.now().astimezone().isoformat(),
        "manifest_sha256": digest(manifest_path), "campaign_script_sha256": digest(Path(__file__)),
        "context_manifest_sha256": digest(context_path), "context_profile": context.get("profile"),
        "agent_policy_version": context.get("agent_policy_version"), "python": sys.version,
        "total": len(results),
        "passed": sum(item["passed"] for item in results),
        "failed": sum(not item["passed"] for item in results),
        "ir_changes_since_previous": sum(item["ir_changed_since_previous"] for item in results),
        "coverage": coverage(results), "results": results,
    }
    (run_dir / "results.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    failures = Counter(item["failure_signature"] for item in results if not item["passed"])
    lines = ["# Campaign report", "", f"Run: `{run_id}`", "", f"- Total: {summary['total']}",
             f"- Passed: {summary['passed']}", f"- Failed: {summary['failed']}",
             f"- IR changes since previous run: {summary['ir_changes_since_previous']}",
             f"- Manifest SHA256: `{summary['manifest_sha256']}`", f"- Context: `{summary['context_profile']}`",
             "", "## Coverage", ""]
    lines.extend(f"- {key}: `{', '.join(map(str, values))}`" for key, values in summary["coverage"].items())
    lines.extend(["", "## Failure signatures", ""])
    lines.extend(f"- {count} × `{name}`" for name, count in failures.most_common())
    lines.extend(["", "## Failed cases", ""])
    lines.extend(f"- `{item['id']}` — {item['failure_signature']}" for item in results if not item["passed"])
    (run_dir / "report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Report: {run_dir / 'report.md'}")
    return 1 if summary["failed"] else 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    create = commands.add_parser("create", help="create a new deterministic campaign")
    create.add_argument("path", type=Path)
    create.add_argument("--count", type=int, default=36)
    create.add_argument("--seed", type=int, default=20260904)
    create.add_argument("--include-invalid", action="store_true")
    run = commands.add_parser("run", help="run an existing campaign")
    run.add_argument("path", type=Path)
    run.add_argument("--timeout", type=int, default=300)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    path = args.path.resolve()
    if args.command == "create":
        if args.count < 1:
            raise SystemExit("--count must be positive")
        return create_campaign(path, args.count, args.seed, args.include_invalid)
    return run_campaign(path, args.timeout)


if __name__ == "__main__":
    raise SystemExit(main())

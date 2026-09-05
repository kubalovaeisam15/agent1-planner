# -*- coding: utf-8 -*-
"""Чтение и машинная проверка MPP через Microsoft Project COM."""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Literal
from uuid import uuid4
from xml.etree import ElementTree as ET

from grp_model import parse_links
from schedule_ir import ScheduleProject
from shared_sections import shared_section_errors

NS = "http://schemas.microsoft.com/project"
NAMESPACES = {"p": NS}
BRIDGE = Path(__file__).with_name("project_read_bridge.ps1")
LINK_TYPE = {0: "FF", 1: "FS", 2: "SF", 3: "SS"}
TENTHS_OF_MINUTE_PER_DAY = 24 * 60 * 10
HARD_CONSTRAINTS = {2, 3, 5, 7}
LINK_TYPE_FROM_RU = {"ОН": "FS", "НН": "SS", "ОО": "FF", "НО": "SF"}


@dataclass(frozen=True)
class MPPLink:
    predecessor_uid: int
    type: str
    lag_days: float


@dataclass
class MPPTask:
    uid: int
    task_id: int
    name: str
    outline_level: int
    summary: bool
    milestone: bool
    start: date | None
    finish: date | None
    duration_minutes: float | None
    percent_complete: int | None
    critical: bool
    total_slack_minutes: float | None
    constraint_type: int
    constraint_date: date | None
    deadline: date | None
    duration_text: str = ""
    predecessors: list[MPPLink] = field(default_factory=list)


@dataclass
class MPPSnapshot:
    name: str
    start: date | None
    finish: date | None
    tasks: list[MPPTask]
    calculation_mode: int | None = None


@dataclass(frozen=True)
class MPPIssue:
    severity: Literal["error", "warning", "info"]
    code: str
    message: str
    task_id: int | None = None
    task_name: str | None = None


def _text(node: ET.Element, name: str, default: str = "") -> str:
    return node.findtext(f"p:{name}", default=default, namespaces=NAMESPACES)


def _bool(node: ET.Element, name: str) -> bool:
    return _text(node, name, "0") == "1"


def _date(value: str) -> date | None:
    return date.fromisoformat(value[:10]) if value else None


def _number(value: str) -> float | None:
    return float(value) if value not in ("", None) else None


_DURATION = re.compile(
    r"^P(?:(?P<days>[0-9.]+)D)?(?:T(?:(?P<hours>[0-9.]+)H)?"
    r"(?:(?P<minutes>[0-9.]+)M)?(?:(?P<seconds>[0-9.]+)S)?)?$"
)


def duration_minutes(value: str) -> float | None:
    if not value:
        return None
    match = _DURATION.fullmatch(value)
    if not match:
        raise ValueError(f"Неизвестный формат длительности MSPDI: {value}")
    parts = {key: float(number or 0) for key, number in match.groupdict().items()}
    return (parts["days"] * 1440 + parts["hours"] * 60
            + parts["minutes"] + parts["seconds"] / 60)


def parse_mspdi(path: Path) -> MPPSnapshot:
    """Читает XML-снимок, созданный самим Microsoft Project."""
    root = ET.parse(path).getroot()
    tasks: list[MPPTask] = []
    for node in root.findall("p:Tasks/p:Task", NAMESPACES):
        task_id = int(_text(node, "ID", "0"))
        if task_id == 0:  # встроенная суммарная задача проекта
            continue
        links: list[MPPLink] = []
        for link in node.findall("p:PredecessorLink", NAMESPACES):
            raw_type = int(_text(link, "Type", "1"))
            links.append(MPPLink(
                predecessor_uid=int(_text(link, "PredecessorUID")),
                type=LINK_TYPE.get(raw_type, f"UNKNOWN-{raw_type}"),
                lag_days=float(_text(link, "LinkLag", "0")) /
                TENTHS_OF_MINUTE_PER_DAY,
            ))
        tasks.append(MPPTask(
            uid=int(_text(node, "UID")),
            task_id=task_id,
            name=_text(node, "Name"),
            outline_level=int(_text(node, "OutlineLevel", "0")),
            summary=_bool(node, "Summary"),
            milestone=_bool(node, "Milestone"),
            start=_date(_text(node, "Start")),
            finish=_date(_text(node, "Finish")),
            duration_minutes=duration_minutes(_text(node, "Duration")),
            percent_complete=(int(_text(node, "PercentComplete"))
                              if _text(node, "PercentComplete") else None),
            critical=_bool(node, "Critical"),
            total_slack_minutes=_number(_text(node, "TotalSlack")),
            constraint_type=int(_text(node, "ConstraintType", "0")),
            constraint_date=_date(_text(node, "ConstraintDate")),
            deadline=_date(_text(node, "Deadline")),
            duration_text="",
            predecessors=links,
        ))
    return MPPSnapshot(
        name=_text(root, "Name"),
        start=_date(_text(root, "StartDate")),
        finish=_date(_text(root, "FinishDate")),
        tasks=sorted(tasks, key=lambda task: task.task_id),
        calculation_mode=None,
    )


def parse_com_snapshot(path: Path) -> MPPSnapshot:
    """Читает компактный JSON, полученный из MPP через COM."""
    raw = json.loads(path.read_text(encoding="utf-8-sig"))
    uid_by_id = {int(item["id"]): int(item["uid"]) for item in raw["tasks"]}
    tasks: list[MPPTask] = []
    for item in raw["tasks"]:
        links = []
        predecessor_text = item.get("predecessors", "")
        # Русский Project помечает календарный лаг как «адней»
        # (астрономических дней), тогда как ГРП использует «дней».
        normalized = re.sub(r"а(?=д(?:н|ень|ня|ней))", "", predecessor_text)
        parsed_links = parse_links(normalized)
        for predecessor_id, kind, lag in parsed_links:
            predecessor_uid = uid_by_id.get(predecessor_id, -predecessor_id)
            links.append(MPPLink(
                predecessor_uid, LINK_TYPE_FROM_RU[kind], float(lag)))
        raw_count = sum(bool(part.strip()) for part in predecessor_text.split(";"))
        if len(parsed_links) != raw_count:
            links.append(MPPLink(0, "UNPARSED", 0))
        tasks.append(MPPTask(
            uid=int(item["uid"]),
            task_id=int(item["id"]),
            name=str(item["name"]),
            outline_level=int(item["outline_level"]),
            summary=bool(item["summary"]),
            milestone=bool(item["milestone"]),
            start=_date(item.get("start") or ""),
            finish=_date(item.get("finish") or ""),
            duration_minutes=_number(item.get("duration_minutes")),
            percent_complete=(int(item["percent_complete"])
                              if item.get("percent_complete") is not None else None),
            critical=bool(item["critical"]),
            total_slack_minutes=_number(item.get("total_slack_minutes")),
            constraint_type=int(item["constraint_type"]),
            constraint_date=_date(item.get("constraint_date") or ""),
            deadline=_date(item.get("deadline") or ""),
            duration_text=str(item.get("duration_text") or ""),
            predecessors=links,
        ))
    return MPPSnapshot(
        name=str(raw["name"]),
        start=_date(raw.get("start") or ""),
        finish=_date(raw.get("finish") or ""),
        tasks=sorted(tasks, key=lambda task: task.task_id),
        calculation_mode=(int(raw["calculation_mode"])
                          if raw.get("calculation_mode") is not None else None),
    )


def read_mpp(mpp_path: Path, *, timeout_seconds: int = 300) -> MPPSnapshot:
    """Без изменения исходного MPP получает компактный снимок через COM."""
    mpp_path = mpp_path.resolve()
    if not mpp_path.exists():
        raise FileNotFoundError(f"MPP не найден: {mpp_path}")
    if mpp_path.suffix.lower() != ".mpp":
        raise ValueError("Входной файл должен иметь расширение .mpp")
    if not BRIDGE.exists():
        raise FileNotFoundError(f"Не найден COM-мост: {BRIDGE}")
    # Временный снимок держим рядом с MPP без отдельного TemporaryDirectory:
    # в локальной песочнице Windows каталог с mode=0700 может стать недоступен
    # даже создавшему его процессу. Уникальный файл удаляется в finally.
    snapshot_path = mpp_path.parent / f".agent1-mpp-read-{uuid4().hex}.json"
    try:
        command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(BRIDGE), "-InputMpp", str(mpp_path),
            "-OutputJson", str(snapshot_path),
        ]
        completed = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout_seconds,
            check=False,
        )
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise RuntimeError(f"Microsoft Project COM завершился с ошибкой: {detail}")
        if not snapshot_path.exists():
            raise RuntimeError("Microsoft Project не создал JSON-снимок")
        return parse_com_snapshot(snapshot_path)
    finally:
        snapshot_path.unlink(missing_ok=True)


def _ancestor_map(tasks: list[MPPTask]) -> dict[int, tuple[int, ...]]:
    """Возвращает UID предков от корня к непосредственному родителю."""
    ancestors: dict[int, tuple[int, ...]] = {}
    path: list[int] = []
    for task in tasks:
        while len(path) >= task.outline_level:
            path.pop()
        ancestors[task.uid] = tuple(path)
        path.append(task.uid)
    return ancestors


def validate_snapshot(snapshot: MPPSnapshot, *, require_all_sections: bool = True) -> list[MPPIssue]:
    """Проверяет структуру и качество сети независимо от исходного IR."""
    issues: list[MPPIssue] = []
    issues.extend(MPPIssue("error", "MPP-SHARED-SECTION", message) for message in
                  shared_section_errors(((t.name, t.outline_level) for t in snapshot.tasks),
                                        require_all=require_all_sections))
    if snapshot.calculation_mode is not None and snapshot.calculation_mode != -1:
        issues.append(MPPIssue(
            "error", "MPP-CALCULATION-MODE",
            "Автоматический перерасчёт Microsoft Project отключён"))
    ids = [task.task_id for task in snapshot.tasks]
    uids = {task.uid for task in snapshot.tasks}
    uid_to_task = {task.uid: task for task in snapshot.tasks}
    successors: dict[int, set[int]] = {uid: set() for uid in uids}
    ancestors = _ancestor_map(snapshot.tasks)

    if len(ids) != len(set(ids)):
        issues.append(MPPIssue("error", "MPP-DUPLICATE-ID",
                               "Идентификаторы задач не уникальны"))
    previous_level = 0
    for index, task in enumerate(snapshot.tasks):
        label = {"task_id": task.task_id, "task_name": task.name}
        if (task.outline_level < 1 or (index == 0 and task.outline_level != 1)
                or (index and task.outline_level > previous_level + 1)):
            issues.append(MPPIssue("error", "MPP-WBS-LEVEL",
                                   "Некорректный скачок уровня WBS", **label))
        previous_level = task.outline_level
        if task.start is None or task.finish is None:
            issues.append(MPPIssue("error", "MPP-DATES-MISSING",
                                   "Не заполнены даты задачи", **label))
        elif task.finish < task.start:
            issues.append(MPPIssue("error", "MPP-DATES",
                                   "Окончание раньше начала", **label))
        if task.milestone and (task.duration_minutes or 0) != 0:
            issues.append(MPPIssue("error", "MPP-MILESTONE-DURATION",
                                   "У вехи ненулевая длительность", **label))
        if re.search(r"а(?=д(?:н|ень|ня|ней))", task.duration_text, re.IGNORECASE):
            issues.append(MPPIssue(
                "error", "MPP-ELAPSED-DURATION",
                f"Использована астрономическая длительность: {task.duration_text}",
                **label))
        if task.total_slack_minutes is not None and task.total_slack_minutes < 0:
            issues.append(MPPIssue("error", "MPP-NEGATIVE-SLACK",
                                   "Отрицательный общий резерв", **label))
        if task.constraint_type in HARD_CONSTRAINTS:
            issues.append(MPPIssue("warning", "MPP-HARD-CONSTRAINT",
                                   f"Жёсткое ограничение типа {task.constraint_type}", **label))
        elif task.constraint_type not in (0, 1, 4):
            issues.append(MPPIssue("warning", "MPP-CONSTRAINT",
                                   f"Ограничение типа {task.constraint_type}", **label))
        seen_links: set[tuple[int, str, float]] = set()
        for link in task.predecessors:
            key = (link.predecessor_uid, link.type, link.lag_days)
            if key in seen_links:
                issues.append(MPPIssue("error", "MPP-LINK-DUPLICATE",
                                       "Дублирующая связь", **label))
            seen_links.add(key)
            if link.predecessor_uid not in uids:
                issues.append(MPPIssue("error", "MPP-LINK-MISSING",
                                       "Предшественник отсутствует", **label))
            elif link.predecessor_uid == task.uid:
                issues.append(MPPIssue("error", "MPP-LINK-SELF",
                                       "Задача связана сама с собой", **label))
            else:
                successors[link.predecessor_uid].add(task.uid)

    leaf_tasks = [task for task in snapshot.tasks if not task.summary]
    summary_start_coverage = 0
    summary_finish_coverage = 0
    reporting_milestones = 0
    for task in leaf_tasks:
        label = {"task_id": task.task_id, "task_name": task.name}
        ancestor_tasks = [uid_to_task[uid] for uid in ancestors[task.uid]]
        ancestor_predecessors = {
            link.predecessor_uid for ancestor in ancestor_tasks
            for link in ancestor.predecessors
        }
        ancestor_successors = {
            successor for ancestor in ancestor_tasks
            for successor in successors[ancestor.uid]
        }
        effective_predecessors = (
            {link.predecessor_uid for link in task.predecessors} | ancestor_predecessors
        )
        effective_successors = successors[task.uid] | ancestor_successors
        reporting = task.milestone and any(
            ancestor.name == "Контрольные вехи" for ancestor in ancestor_tasks)

        if not task.predecessors and effective_predecessors:
            summary_start_coverage += 1
        if not successors[task.uid] and effective_successors:
            summary_finish_coverage += 1
        if not effective_predecessors:
            issues.append(MPPIssue("warning", "MPP-OPEN-START",
                                   "У задачи нет предшественника", **label))
        if not effective_successors and reporting:
            reporting_milestones += 1
        elif not effective_successors:
            issues.append(MPPIssue("warning", "MPP-OPEN-FINISH",
                                   "У задачи нет последователя", **label))

    if summary_start_coverage or summary_finish_coverage:
        issues.append(MPPIssue(
            "info", "MPP-SUMMARY-LINK-COVERAGE",
            "Связями суммарных строк покрыто задач без прямых связей: "
            f"начало — {summary_start_coverage}, окончание — {summary_finish_coverage} "
            "(DEC-26)"))
    if reporting_milestones:
        issues.append(MPPIssue(
            "info", "MPP-REPORTING-MILESTONES",
            f"Контрольных вех без последователей: {reporting_milestones}; "
            "это выходной контракт интеграции, а не управляющие задачи "
            "(typGRP.md §5)"))

    critical = {task.uid for task in snapshot.tasks if task.critical}
    critical_leaves = {task.uid for task in leaf_tasks if task.critical}
    if not critical_leaves:
        issues.append(MPPIssue("warning", "MPP-CRITICAL-EMPTY",
                               "Microsoft Project не определил критические задачи"))
    else:
        for uid in critical_leaves:
            task = uid_to_task[uid]
            ancestor_tasks = [uid_to_task[item] for item in ancestors[uid]]
            effective_preds = {link.predecessor_uid for link in task.predecessors}
            effective_succs = set(successors[uid])
            for ancestor in ancestor_tasks:
                effective_preds.update(
                    link.predecessor_uid for link in ancestor.predecessors)
                effective_succs.update(successors[ancestor.uid])
            if effective_preds and not effective_preds & critical:
                issues.append(MPPIssue(
                    "warning", "MPP-CRITICAL-GAP-IN",
                    "Критическая задача не имеет критического предшественника",
                    task.task_id, task.name))
            if effective_succs and not effective_succs & critical:
                issues.append(MPPIssue(
                    "warning", "MPP-CRITICAL-GAP-OUT",
                    "Критическая задача не имеет критического последователя",
                    task.task_id, task.name))
    return issues


def compare_with_ir(snapshot: MPPSnapshot, schedule: ScheduleProject) -> list[MPPIssue]:
    """Сверяет MPP с исходным Schedule IR по задачам, датам и связям."""
    issues: list[MPPIssue] = []
    critical_mismatches = 0
    deadline_count = 0
    if len(snapshot.tasks) != len(schedule.tasks):
        issues.append(MPPIssue(
            "error", "MPP-IR-TASK-COUNT",
            f"В MPP {len(snapshot.tasks)} задач, в IR {len(schedule.tasks)}"))
    project_by_position = {index: task for index, task in enumerate(
        snapshot.tasks, start=1)}
    uid_by_position = {index: task.uid for index, task in project_by_position.items()}
    ir_position = {task.task_id: index for index, task in enumerate(
        schedule.tasks, start=1)}

    for position, ir_task in enumerate(schedule.tasks, start=1):
        task = project_by_position.get(position)
        if task is None:
            issues.append(MPPIssue("error", "MPP-IR-TASK-MISSING",
                                   "Задача IR отсутствует в MPP", position, ir_task.name))
            continue
        label = {"task_id": task.task_id, "task_name": task.name}
        if task.name != ir_task.name:
            issues.append(MPPIssue("error", "MPP-IR-NAME",
                                   f"Имя не совпадает с IR: {ir_task.name}", **label))
        if task.outline_level != ir_task.outline_level:
            issues.append(MPPIssue("error", "MPP-IR-WBS",
                                   f"Уровень {task.outline_level}, в IR {ir_task.outline_level}",
                                   **label))
        expected_summary = ir_task.task_type == "summary"
        expected_milestone = ir_task.task_type == "milestone"
        type_matches = (
            (expected_summary and task.summary)
            or (expected_milestone and not task.summary and task.milestone)
            or (ir_task.task_type == "task" and not task.summary and not task.milestone)
        )
        # Project marks a zero-span summary as both Summary and Milestone.
        if not type_matches:
            issues.append(MPPIssue("error", "MPP-IR-TYPE",
                                   "Тип задачи не совпадает с IR", **label))
        if ir_task.duration_days is not None:
            expected_minutes = ir_task.duration_days * 1440
            if task.duration_minutes is None or abs(
                    task.duration_minutes - expected_minutes) > 0.01:
                issues.append(MPPIssue(
                    "error", "MPP-IR-DURATION",
                    f"Длительность {task.duration_minutes}, в IR {expected_minutes} мин",
                    **label))
        if task.start != ir_task.start or task.finish != ir_task.finish:
            issues.append(MPPIssue(
                "error", "MPP-IR-DATES",
                f"Даты {task.start}…{task.finish}, в IR {ir_task.start}…{ir_task.finish}",
                **label))
        if task.deadline is not None:
            deadline_count += 1
        expected_constraint = (
            4 if ir_task.constraint_type == "start_no_earlier_than"
            else (4 if (ir_task.task_type != "summary" and not ir_task.predecessors
                        and ir_task.start and ir_task.start > schedule.project_start) else 0)
        )
        if task.constraint_type != expected_constraint:
            issues.append(MPPIssue(
                "error", "MPP-IR-CONSTRAINT-TYPE",
                f"Тип ограничения {task.constraint_type}, в IR ожидается "
                f"{expected_constraint}", **label))
        if (expected_constraint == 4 and ir_task.constraint_date is not None
                and task.constraint_date != ir_task.constraint_date):
            issues.append(MPPIssue(
                "error", "MPP-IR-CONSTRAINT-DATE",
                f"Дата ограничения {task.constraint_date}, в IR "
                f"{ir_task.constraint_date}", **label))
        # Критичность суммарных строк Project вычисляет самостоятельно и она
        # не входит в контракт IR; сравниваются только задачи и вехи.
        if ir_task.task_type != "summary" and task.critical != ir_task.critical:
            critical_mismatches += 1

        expected_links = {
            (uid_by_position[ir_position[link.predecessor_id]],
             link.type, float(link.lag_days))
            for link in ir_task.predecessors
            if link.predecessor_id in ir_position
        }
        actual_links = {(link.predecessor_uid, link.type, link.lag_days)
                        for link in task.predecessors}
        if actual_links != expected_links:
            issues.append(MPPIssue("error", "MPP-IR-LINKS",
                                   "Набор связей не совпадает с IR", **label))
    if deadline_count:
        issues.append(MPPIssue(
            "error", "MPP-IR-DEADLINES",
            f"В MPP заданы дедлайны у {deadline_count} задач. Резерв DEC-31 "
            "не импортируется в Project и не должен влиять на критический путь"))
    if critical_mismatches:
        issues.append(MPPIssue(
            "warning", "MPP-IR-CRITICAL",
            "Microsoft Project пересчитал признак критичности иначе у "
            f"{critical_mismatches} задач"))
    return issues


def report_dict(snapshot: MPPSnapshot, issues: list[MPPIssue]) -> dict:
    errors = sum(issue.severity == "error" for issue in issues)
    warnings = sum(issue.severity == "warning" for issue in issues)
    infos = sum(issue.severity == "info" for issue in issues)
    counts: dict[str, int] = {}
    for issue in issues:
        counts[issue.code] = counts.get(issue.code, 0) + 1
    return {
        "project": {"name": snapshot.name, "start": str(snapshot.start),
                    "finish": str(snapshot.finish), "task_count": len(snapshot.tasks),
                    "calculation_mode": snapshot.calculation_mode,
                    "automatic_calculation": snapshot.calculation_mode == -1},
        "result": {"errors": errors, "warnings": warnings, "infos": infos,
                   "by_code": counts},
        "issues": [asdict(issue) for issue in issues],
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Проверить расписание Microsoft Project")
    parser.add_argument("mpp", type=Path, help="проверяемый .mpp")
    parser.add_argument("--ir", type=Path, help="исходный Schedule IR для точной сверки")
    parser.add_argument("--json-report", type=Path, help="записать полный отчёт JSON")
    parser.add_argument("--timeout", type=int, default=300, help="таймаут COM, секунд")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        snapshot = read_mpp(args.mpp, timeout_seconds=args.timeout)
        issues = validate_snapshot(snapshot)
        if args.ir:
            schedule = ScheduleProject.from_json(args.ir.read_text(encoding="utf-8"))
            issues.extend(compare_with_ir(snapshot, schedule))
        report = report_dict(snapshot, issues)
        if args.json_report:
            args.json_report.parent.mkdir(parents=True, exist_ok=True)
            args.json_report.write_text(
                json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"Ошибка проверки MPP: {exc}", file=sys.stderr)
        return 2

    result = report["result"]
    print(f"MPP: {snapshot.name} · задач: {len(snapshot.tasks)} · "
          f"{snapshot.start}…{snapshot.finish}")
    print(f"Ошибок: {result['errors']} · предупреждений: {result['warnings']} · "
          f"информационных: {result['infos']}")
    for code, count in sorted(result["by_code"].items()):
        print(f"  {code}: {count}")
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    raise SystemExit(main())

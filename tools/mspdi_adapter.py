# -*- coding: utf-8 -*-
"""Запись Schedule IR в Microsoft Project MPP через MSPDI XML и COM.

Python строит стандартный Project XML без внешних зависимостей. Локальный
PowerShell-мост открывает XML в установленном Microsoft Project и сохраняет
новый MPP. Существующие MPP по умолчанию не перезаписываются.
"""
from __future__ import annotations

import argparse
import json
import subprocess
import sys
from datetime import date
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any
from xml.etree import ElementTree as ET

from schedule_ir import ScheduleProject, validate_schedule_ir

ROOT = Path(__file__).resolve().parents[1]
BRIDGE = Path(__file__).with_name("project_com_bridge.ps1")
NS = "http://schemas.microsoft.com/project"
ET.register_namespace("", NS)

LINK_TYPE = {"FF": 0, "FS": 1, "SF": 2, "SS": 3}
ELAPSED_DAY = 8
TENTHS_OF_MINUTE_PER_DAY = 24 * 60 * 10


def _tag(name: str) -> str:
    return f"{{{NS}}}{name}"


def _add(parent: ET.Element, name: str, value: Any) -> ET.Element:
    child = ET.SubElement(parent, _tag(name))
    child.text = str(value)
    return child


def _dt(value: date) -> str:
    return f"{value.isoformat()}T00:00:00"


def _elapsed_duration(days: int) -> str:
    # Project 2021 обнуляет большие elapsed-длительности вида PT6000H.
    # ISO-форма P250D импортируется как 250 календарных (elapsed) дней.
    return f"P{days}D"


def _outline_numbers(schedule: ScheduleProject) -> dict[str, str]:
    """Строит 1, 1.1, 1.1.1 по порядку WBS, не полагаясь на task_id."""
    counters: dict[int, int] = {}
    result: dict[str, str] = {}
    for task in schedule.tasks:
        level = task.outline_level
        counters[level] = counters.get(level, 0) + 1
        for stale in [n for n in counters if n > level]:
            del counters[stale]
        result[task.task_id] = ".".join(str(counters[n]) for n in range(1, level + 1))
    return result


def schedule_to_mspdi(schedule: ScheduleProject) -> bytes:
    """Сериализует валидный Schedule IR в Microsoft Project XML (MSPDI)."""
    issues = validate_schedule_ir(schedule)
    if issues:
        detail = "; ".join(f"{i.code}{'[' + i.task_id + ']' if i.task_id else ''}"
                           for i in issues[:10])
        raise ValueError(f"Schedule IR не прошёл валидацию: {detail}")

    finish = max((task.finish for task in schedule.tasks if task.finish),
                 default=schedule.project_start)
    root = ET.Element(_tag("Project"))
    _add(root, "SaveVersion", 14)
    _add(root, "UID", schedule.schedule_id)
    _add(root, "Name", schedule.name)
    _add(root, "Title", schedule.name)
    _add(root, "ScheduleFromStart", 1)
    _add(root, "StartDate", _dt(schedule.project_start))
    _add(root, "FinishDate", _dt(finish))
    _add(root, "FYStartDate", 1)
    _add(root, "CriticalSlackLimit", 0)
    _add(root, "CalendarUID", 1)
    _add(root, "DefaultStartTime", "00:00:00")
    _add(root, "DefaultFinishTime", "00:00:00")
    _add(root, "MinutesPerDay", 1440)
    _add(root, "MinutesPerWeek", 10080)
    _add(root, "DaysPerMonth", 30)
    _add(root, "DefaultTaskType", 1)  # fixed duration
    _add(root, "DurationFormat", ELAPSED_DAY)
    _add(root, "WorkFormat", 2)
    _add(root, "HonorConstraints", 1)
    _add(root, "NewTasksEffortDriven", 0)
    _add(root, "NewTasksEstimated", 0)

    calendars = ET.SubElement(root, _tag("Calendars"))
    calendar = ET.SubElement(calendars, _tag("Calendar"))
    _add(calendar, "UID", 1)
    _add(calendar, "Name", "Calendar 7d")
    _add(calendar, "IsBaseCalendar", 1)
    _add(calendar, "BaseCalendarUID", -1)
    weekdays = ET.SubElement(calendar, _tag("WeekDays"))
    for day_type in range(1, 8):
        weekday = ET.SubElement(weekdays, _tag("WeekDay"))
        _add(weekday, "DayType", day_type)
        _add(weekday, "DayWorking", 1)
        working_times = ET.SubElement(weekday, _tag("WorkingTimes"))
        working_time = ET.SubElement(working_times, _tag("WorkingTime"))
        _add(working_time, "FromTime", "00:00:00")
        _add(working_time, "ToTime", "23:59:00")

    uid_by_id = {task.task_id: index for index, task in enumerate(schedule.tasks, start=1)}
    outline = _outline_numbers(schedule)
    tasks_node = ET.SubElement(root, _tag("Tasks"))

    project_task = ET.SubElement(tasks_node, _tag("Task"))
    _add(project_task, "UID", 0)
    _add(project_task, "ID", 0)
    _add(project_task, "Name", schedule.name)
    _add(project_task, "Type", 1)
    _add(project_task, "IsNull", 0)
    _add(project_task, "WBS", 0)
    _add(project_task, "OutlineNumber", 0)
    _add(project_task, "OutlineLevel", 0)
    _add(project_task, "Start", _dt(schedule.project_start))
    _add(project_task, "Finish", _dt(finish))
    _add(project_task, "Duration", _elapsed_duration((finish - schedule.project_start).days))
    _add(project_task, "DurationFormat", ELAPSED_DAY)
    _add(project_task, "Summary", 1)

    for index, task in enumerate(schedule.tasks, start=1):
        node = ET.SubElement(tasks_node, _tag("Task"))
        _add(node, "UID", uid_by_id[task.task_id])
        _add(node, "ID", index)
        _add(node, "Name", task.name)
        _add(node, "Type", 1)
        _add(node, "IsNull", 0)
        _add(node, "WBS", outline[task.task_id])
        _add(node, "OutlineNumber", outline[task.task_id])
        _add(node, "OutlineLevel", task.outline_level)
        _add(node, "Priority", 500)
        if task.start:
            _add(node, "Start", _dt(task.start))
        if task.finish:
            _add(node, "Finish", _dt(task.finish))
        if task.duration_days is not None:
            _add(node, "Duration", _elapsed_duration(task.duration_days))
            _add(node, "DurationFormat", ELAPSED_DAY)
        _add(node, "Work", "PT0H0M0S")
        _add(node, "EffortDriven", 0)
        _add(node, "Recurring", 0)
        _add(node, "Estimated", 0)
        _add(node, "Milestone", int(task.task_type == "milestone"))
        _add(node, "Summary", int(task.task_type == "summary"))
        _add(node, "Critical", int(task.critical))
        if task.percent_complete is not None:
            _add(node, "PercentComplete", task.percent_complete)
        if task.task_type != "summary" and task.duration_days is not None:
            # Project may derive an imported task's duration from the remaining
            # duration instead of Duration when no actual progress exists.
            _add(node, "ActualDuration", "PT0H0M0S")
            _add(node, "RegularWork", "PT0H0M0S")
            _add(node, "RemainingDuration", _elapsed_duration(task.duration_days))

        # Независимая задача с более поздним стартом — календарный якорь ГРП.
        # SNET сохраняет дату, не превращая все связанные задачи в ограничения.
        if (task.task_type != "summary" and not task.predecessors and task.start
                and task.start > schedule.project_start):
            _add(node, "ConstraintType", 4)
            _add(node, "CalendarUID", 1)
            _add(node, "ConstraintDate", _dt(task.start))
        else:
            _add(node, "ConstraintType", 0)
            _add(node, "CalendarUID", 1)
        if task.reserve_finish:
            _add(node, "Deadline", _dt(task.reserve_finish))
        notes = task.notes
        trace = [f"Schedule IR task_id: {task.task_id}"]
        if task.source_key:
            trace.append(f"source_key: {task.source_key}")
        if task.phase:
            trace.append(f"phase: {task.phase}")
        _add(node, "Notes", "\n".join(trace + ([notes] if notes else [])))

        for link in task.predecessors:
            pred = ET.SubElement(node, _tag("PredecessorLink"))
            _add(pred, "PredecessorUID", uid_by_id[link.predecessor_id])
            _add(pred, "Type", LINK_TYPE[link.type])
            _add(pred, "CrossProject", 0)
            _add(pred, "LinkLag", link.lag_days * TENTHS_OF_MINUTE_PER_DAY)
            _add(pred, "LagFormat", ELAPSED_DAY)

    ET.indent(root, space="  ")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def export_mpp(ir_path: Path, mpp_path: Path, *, timeout_seconds: int = 300) -> dict[str, Any]:
    """Создаёт новый MPP. Существующий файл намеренно не перезаписывает."""
    ir_path = ir_path.resolve()
    mpp_path = mpp_path.resolve()
    if mpp_path.suffix.lower() != ".mpp":
        raise ValueError("Выходной файл должен иметь расширение .mpp")
    if mpp_path.exists():
        raise FileExistsError(f"MPP уже существует: {mpp_path}")
    if not BRIDGE.exists():
        raise FileNotFoundError(f"Не найден COM-мост: {BRIDGE}")
    mpp_path.parent.mkdir(parents=True, exist_ok=True)

    schedule = ScheduleProject.from_json(ir_path.read_text(encoding="utf-8"))
    expected_finish = max(
        (task.finish for task in schedule.tasks if task.finish),
        default=schedule.project_start,
    )
    xml = schedule_to_mspdi(schedule)
    with TemporaryDirectory(prefix="agent1-mpp-", dir=mpp_path.parent) as temp_dir:
        temp = Path(temp_dir)
        xml_path = temp / "schedule.xml"
        durations_path = temp / "durations.json"
        report_path = temp / "report.json"
        xml_path.write_bytes(xml)
        duration_overrides = [
            {"id": index, "duration_minutes": task.duration_days * 24 * 60}
            for index, task in enumerate(schedule.tasks, start=1)
            if task.task_type != "summary" and task.duration_days is not None
        ]
        durations_path.write_text(
            json.dumps(duration_overrides, ensure_ascii=False), encoding="utf-8")
        sample_names = {
            "К1. Пуск тепла корпус",
            "К2. Пуск тепла корпус",
            "Получено РВЭ",
            "К1. По договору Отопление (контур для пуска тепла)",
        }
        sample_ids = ",".join(str(index) for index, task in enumerate(
            schedule.tasks, start=1) if task.name in sample_names)
        command = [
            "powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-File", str(BRIDGE),
            "-InputXml", str(xml_path),
            "-OutputMpp", str(mpp_path),
            "-ReportJson", str(report_path),
            "-DurationsJson", str(durations_path),
            "-SampleTaskIds", sample_ids,
        ]
        completed = subprocess.run(command, capture_output=True, text=True,
                                   timeout=timeout_seconds, check=False)
        if completed.returncode != 0:
            detail = (completed.stderr or completed.stdout).strip()
            raise RuntimeError(f"Microsoft Project COM завершился с ошибкой: {detail}")
        if not mpp_path.exists():
            raise RuntimeError("Microsoft Project не создал выходной MPP")
        if not report_path.exists():
            raise RuntimeError("COM-мост не создал отчёт проверки")
        report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    if report.get("task_count") != len(schedule.tasks):
        raise RuntimeError(
            f"MPP содержит {report.get('task_count')} задач вместо {len(schedule.tasks)}")
    if report.get("duration_override_count") != len(duration_overrides):
        raise RuntimeError(
            "MPP принял не все длительности: "
            f"{report.get('duration_override_count')} из {len(duration_overrides)}")
    if report.get("duration_mismatch_count"):
        raise RuntimeError(
            f"MPP изменил длительности {report['duration_mismatch_count']} задач")
    if report.get("project_start") != schedule.project_start.isoformat():
        raise RuntimeError(
            f"Начало MPP {report.get('project_start')} не совпадает с IR "
            f"{schedule.project_start.isoformat()}")
    if report.get("project_finish") != expected_finish.isoformat():
        raise RuntimeError(
            f"Окончание MPP {report.get('project_finish')} не совпадает с IR "
            f"{expected_finish.isoformat()}")
    return report


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Создать новый MPP из Schedule IR")
    parser.add_argument("ir", type=Path, help="Schedule IR v1.0 JSON")
    parser.add_argument("mpp", type=Path, help="новый выходной .mpp")
    parser.add_argument("--timeout", type=int, default=300, help="таймаут COM, секунд")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        report = export_mpp(args.ir, args.mpp, timeout_seconds=args.timeout)
    except (OSError, ValueError, RuntimeError, subprocess.SubprocessError) as exc:
        print(f"Ошибка экспорта MPP: {exc}", file=sys.stderr)
        return 1
    print(f"Создан MPP: {args.mpp.resolve()}")
    print(f"Задач: {report['task_count']} · начало: {report.get('project_start', '')} · "
          f"окончание: {report.get('project_finish', '')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

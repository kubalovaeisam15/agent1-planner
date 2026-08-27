# -*- coding: utf-8 -*-
"""Версионированная нейтральная модель расписания (Schedule IR).

IR отделяет planning logic от форматов выдачи. Генератор строит расписание один
раз, после чего Excel-, MPP- и MCP-адаптеры работают с одним контрактом.
Технические поля английские; пользовательские названия и пояснения сохраняются
на русском без преобразования.
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any
from uuid import NAMESPACE_URL, uuid5

from grp_model import dparse, parse_links

SCHEMA_VERSION = "1.0"
LINK_TYPES = {"FS", "SS", "FF", "SF"}
LINK_TYPE_FROM_RU = {"ОН": "FS", "НН": "SS", "ОО": "FF", "НО": "SF"}
TASK_TYPES = {"summary", "task", "milestone"}


@dataclass(frozen=True)
class ScheduleLink:
    predecessor_id: str
    type: str = "FS"
    lag_days: int = 0


@dataclass
class ScheduleTask:
    task_id: str
    name: str
    outline_level: int
    task_type: str
    parent_id: str | None = None
    source_key: str | None = None
    phase: str | None = None
    duration_days: int | None = None
    start: date | None = None
    finish: date | None = None
    reserve_finish: date | None = None
    percent_complete: int | None = None
    critical: bool = False
    calendar_id: str = "calendar-7d"
    notes: str = ""
    predecessors: list[ScheduleLink] = field(default_factory=list)


@dataclass(frozen=True)
class ScheduleCalendar:
    calendar_id: str = "calendar-7d"
    name: str = "Календарные дни, 7 дней в неделю"
    working_weekdays: tuple[int, ...] = (1, 2, 3, 4, 5, 6, 7)
    duration_unit: str = "calendar_day"


@dataclass
class ScheduleProject:
    schedule_id: str
    name: str
    project_start: date
    tasks: list[ScheduleTask]
    schema_version: str = SCHEMA_VERSION
    calendars: list[ScheduleCalendar] = field(default_factory=lambda: [ScheduleCalendar()])
    metadata: dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return _encode(asdict(self))

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=indent)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "ScheduleProject":
        calendars = [ScheduleCalendar(
            calendar_id=c["calendar_id"],
            name=c["name"],
            working_weekdays=tuple(c["working_weekdays"]),
            duration_unit=c["duration_unit"],
        ) for c in value.get("calendars", [])]
        tasks = []
        for item in value.get("tasks", []):
            task = dict(item)
            task["start"] = _date_or_none(task.get("start"))
            task["finish"] = _date_or_none(task.get("finish"))
            task["reserve_finish"] = _date_or_none(task.get("reserve_finish"))
            task["predecessors"] = [ScheduleLink(**link)
                                      for link in task.get("predecessors", [])]
            tasks.append(ScheduleTask(**task))
        return cls(
            schedule_id=value["schedule_id"],
            name=value["name"],
            project_start=_date_or_none(value["project_start"]),
            tasks=tasks,
            schema_version=value.get("schema_version", ""),
            calendars=calendars or [ScheduleCalendar()],
            metadata=dict(value.get("metadata", {})),
        )

    @classmethod
    def from_json(cls, value: str) -> "ScheduleProject":
        return cls.from_dict(json.loads(value))


@dataclass(frozen=True)
class IRIssue:
    code: str
    message: str
    task_id: str | None = None


def _encode(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, tuple):
        return [_encode(v) for v in value]
    if isinstance(value, list):
        return [_encode(v) for v in value]
    if isinstance(value, dict):
        return {k: _encode(v) for k, v in value.items()}
    return value


def _date_or_none(value: str | date | None) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return date.fromisoformat(value)


def _grp_date(value: Any) -> date | None:
    return dparse(str(value)) if value not in (None, "") else None


def _duration(value: Any) -> int | None:
    if value in (None, ""):
        return None
    return int(str(value).split()[0])


def schedule_from_grp(project: dict[str, Any], rows: list[dict[str, Any]]) -> ScheduleProject:
    """Преобразует готовые строки ГРП в нейтральный Schedule IR v1.0."""
    name = str(project["название"])
    project_start = dparse(str(project["старт_проекта"]))
    explicit_id = project.get("project_id")
    schedule_id = str(explicit_id or uuid5(
        NAMESPACE_URL, f"agent1:{name}:{project_start.isoformat()}"))

    tasks: list[ScheduleTask] = []
    parent_at_level: dict[int, str] = {}
    for row in rows:
        task_id = str(row["Ид."])
        level = int(row["Уровень структуры"])
        duration = _duration(row.get("Длительность"))
        task_type = ("summary" if duration is None else
                     "milestone" if duration == 0 else "task")
        parent_id = parent_at_level.get(level - 1)
        parent_at_level[level] = task_id
        for stale_level in [n for n in parent_at_level if n > level]:
            del parent_at_level[stale_level]

        predecessors = [ScheduleLink(str(pid), LINK_TYPE_FROM_RU[kind], lag)
                        for pid, kind, lag in parse_links(
                            str(row.get("Предшественники", "")))]
        percent = row.get("% завершения")
        tasks.append(ScheduleTask(
            task_id=task_id,
            name=str(row["Название задачи"]),
            outline_level=level,
            task_type=task_type,
            parent_id=parent_id,
            source_key=(str(row["_source_key"])
                        if row.get("_source_key") not in (None, "") else None),
            phase=(str(row["_phase"]) if row.get("_phase") else None),
            duration_days=duration,
            start=_grp_date(row.get("Начало")),
            finish=_grp_date(row.get("Окончание")),
            reserve_finish=_grp_date(row.get("Окончание с резервом")),
            percent_complete=(int(percent) if percent not in (None, "") else None),
            critical=bool(row.get("_critical", False)),
            notes=str(row.get("комментарий", "")),
            predecessors=predecessors,
        ))

    return ScheduleProject(
        schedule_id=schedule_id,
        name=name,
        project_start=project_start,
        tasks=tasks,
        metadata={
            "source": "agent1",
            "source_format": "ГРП v2",
            "date_semantics": "finish = start + duration_days",
            "language": "ru",
        },
    )


def validate_schedule_ir(schedule: ScheduleProject) -> list[IRIssue]:
    """Проверяет контракт и ссылочную целостность IR без пересчёта дат."""
    issues: list[IRIssue] = []
    if schedule.schema_version != SCHEMA_VERSION:
        issues.append(IRIssue("IR-SCHEMA", f"Ожидалась версия {SCHEMA_VERSION}"))
    if schedule.project_start is None:
        issues.append(IRIssue("IR-START", "Не задан старт проекта"))

    ids = [task.task_id for task in schedule.tasks]
    known = set(ids)
    if len(ids) != len(known):
        issues.append(IRIssue("IR-DUPLICATE-ID", "Идентификаторы задач не уникальны"))

    position = {task_id: index for index, task_id in enumerate(ids)}
    graph: dict[str, list[str]] = {task_id: [] for task_id in known}
    previous_level = 0
    calendar_ids = {calendar.calendar_id for calendar in schedule.calendars}
    for index, task in enumerate(schedule.tasks):
        if (task.outline_level < 1 or (index == 0 and task.outline_level != 1)
                or (index and task.outline_level > previous_level + 1)):
            issues.append(IRIssue("IR-WBS-LEVEL", "Некорректный скачок уровня WBS", task.task_id))
        previous_level = task.outline_level
        if task.task_type not in TASK_TYPES:
            issues.append(IRIssue("IR-TASK-TYPE", f"Неизвестный тип {task.task_type}", task.task_id))
        if task.calendar_id not in calendar_ids:
            issues.append(IRIssue("IR-CALENDAR", "Неизвестный календарь", task.task_id))
        if task.outline_level > 1 and task.parent_id is None:
            issues.append(IRIssue("IR-PARENT", "У вложенной задачи нет родителя", task.task_id))
        elif task.parent_id is not None:
            if task.parent_id not in known or position.get(task.parent_id, index) >= index:
                issues.append(IRIssue("IR-PARENT", "Родитель отсутствует или стоит ниже задачи",
                                      task.task_id))
        if task.task_type == "summary" and task.duration_days is not None:
            issues.append(IRIssue("IR-SUMMARY-DURATION", "У суммарной задачи задана длительность",
                                  task.task_id))
        if task.task_type == "milestone" and task.duration_days != 0:
            issues.append(IRIssue("IR-MILESTONE-DURATION", "Длительность вехи должна быть 0",
                                  task.task_id))
        if task.duration_days is not None and task.duration_days < 0:
            issues.append(IRIssue("IR-DURATION", "Отрицательная длительность", task.task_id))
        if task.start and task.finish and task.finish < task.start:
            issues.append(IRIssue("IR-DATES", "Окончание раньше начала", task.task_id))
        if (task.task_type != "summary" and task.start and task.finish
                and task.duration_days is not None
                and (task.finish - task.start).days != task.duration_days):
            issues.append(IRIssue("IR-DATE-DURATION", "Даты не соответствуют длительности",
                                  task.task_id))
        if task.reserve_finish and task.finish and task.reserve_finish < task.finish:
            issues.append(IRIssue("IR-RESERVE", "Дата с резервом раньше расчётной", task.task_id))
        seen_links: set[tuple[str, str, int]] = set()
        for link in task.predecessors:
            link_key = (link.predecessor_id, link.type, link.lag_days)
            if link_key in seen_links:
                issues.append(IRIssue("IR-LINK-DUPLICATE", "Связь продублирована",
                                      task.task_id))
            seen_links.add(link_key)
            if link.type not in LINK_TYPES:
                issues.append(IRIssue("IR-LINK-TYPE", f"Неизвестный тип связи {link.type}",
                                      task.task_id))
            if link.predecessor_id not in known:
                issues.append(IRIssue("IR-LINK-MISSING", "Предшественник отсутствует",
                                      task.task_id))
            elif link.predecessor_id == task.task_id:
                issues.append(IRIssue("IR-LINK-SELF", "Задача ссылается сама на себя", task.task_id))
            else:
                graph[task.task_id].append(link.predecessor_id)

    state: dict[str, int] = {}

    def visit(task_id: str) -> bool:
        if state.get(task_id) == 1:
            return True
        if state.get(task_id) == 2:
            return False
        state[task_id] = 1
        cyclic = any(visit(pred) for pred in graph.get(task_id, []))
        state[task_id] = 2
        return cyclic

    if any(visit(task_id) for task_id in ids if state.get(task_id) != 2):
        issues.append(IRIssue("IR-CYCLE", "В сети обнаружен цикл"))
    return issues


def write_schedule_ir(path: Path, schedule: ScheduleProject) -> None:
    path.write_text(schedule.to_json() + "\n", encoding="utf-8")

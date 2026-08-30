# -*- coding: utf-8 -*-
"""MSPDI writer: точное отображение Schedule IR в контракт Microsoft Project."""
from __future__ import annotations

from datetime import date
from xml.etree import ElementTree as ET

from mspdi_adapter import NS, schedule_to_mspdi
from schedule_ir import ScheduleLink, ScheduleProject, ScheduleTask


def sample_schedule() -> ScheduleProject:
    return ScheduleProject(
        schedule_id="test-schedule",
        name="Тестовый график",
        project_start=date(2026, 1, 1),
        tasks=[
            ScheduleTask("1", "Раздел", 1, "summary",
                         start=date(2026, 1, 1), finish=date(2026, 1, 8)),
            ScheduleTask("2", "Работа", 2, "task", parent_id="1", duration_days=5,
                         start=date(2026, 1, 1), finish=date(2026, 1, 6), phase="B"),
            ScheduleTask("3", "Веха", 2, "milestone", parent_id="1", duration_days=0,
                         start=date(2026, 1, 8), finish=date(2026, 1, 8),
                         reserve_finish=date(2026, 1, 10),
                         predecessors=[ScheduleLink("2", "FS", 2)]),
        ],
    )


def test_mspdi_preserves_wbs_elapsed_duration_and_link_without_reserve_deadline():
    root = ET.fromstring(schedule_to_mspdi(sample_schedule()))
    ns = {"p": NS}
    tasks = root.findall("p:Tasks/p:Task", ns)
    assert len(tasks) == 4  # project summary + 3 IR tasks
    work = tasks[2]
    milestone = tasks[3]
    assert work.findtext("p:OutlineNumber", namespaces=ns) == "1.1"
    assert work.findtext("p:Duration", namespaces=ns) == "P5D"
    assert work.findtext("p:DurationFormat", namespaces=ns) == "8"
    assert work.findtext("p:ActualDuration", namespaces=ns) == "PT0H0M0S"
    assert work.findtext("p:RemainingDuration", namespaces=ns) == "P5D"
    assert milestone.findtext("p:Milestone", namespaces=ns) == "1"
    assert milestone.find("p:Deadline", ns) is None  # DEC-31 не импортируется в Project
    link = milestone.find("p:PredecessorLink", ns)
    assert link is not None
    assert link.findtext("p:Type", namespaces=ns) == "1"  # FS
    assert link.findtext("p:LinkLag", namespaces=ns) == "28800"  # 2 elapsed days
    assert link.findtext("p:LagFormat", namespaces=ns) == "8"


def test_mspdi_adds_snet_only_to_independent_calendar_anchor():
    schedule = sample_schedule()
    schedule.tasks.append(ScheduleTask(
        "4", "Независимый якорь", 1, "milestone", duration_days=0,
        start=date(2026, 2, 1), finish=date(2026, 2, 1),
    ))
    root = ET.fromstring(schedule_to_mspdi(schedule))
    ns = {"p": NS}
    task = root.findall("p:Tasks/p:Task", ns)[-1]
    assert task.findtext("p:ConstraintType", namespaces=ns) == "4"
    assert task.findtext("p:ConstraintDate", namespaces=ns) == "2026-02-01T00:00:00"


def test_mspdi_preserves_explicit_snet_with_physical_predecessor():
    schedule = sample_schedule()
    schedule.tasks.append(ScheduleTask(
        "4", "Отделка", 1, "task", duration_days=210,
        start=date(2026, 6, 1), finish=date(2026, 12, 28),
        predecessors=[ScheduleLink("2", "SS", 90)],
        constraint_type="start_no_earlier_than",
        constraint_date=date(2026, 6, 1),
    ))
    root = ET.fromstring(schedule_to_mspdi(schedule))
    ns = {"p": NS}
    task = root.findall("p:Tasks/p:Task", ns)[-1]
    assert task.findtext("p:ConstraintType", namespaces=ns) == "4"
    assert task.findtext("p:ConstraintDate", namespaces=ns) == "2026-06-01T00:00:00"
    assert task.find("p:PredecessorLink", ns) is not None

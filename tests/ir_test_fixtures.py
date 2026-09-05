"""Минимальные обязательные разделы для тестов полного экспортного контракта."""
from schedule_ir import ScheduleTask
from shared_sections import PROJECT_SHARED_SECTIONS


def complete_sections(schedule):
    for n, name in enumerate(PROJECT_SHARED_SECTIONS, 1):
        key = f"section-{n}"
        schedule.tasks.extend([
            ScheduleTask(key, name, 1, "summary", start=schedule.project_start, finish=schedule.project_start),
            ScheduleTask(key+"-child", "Контрольная веха", 2, "milestone", parent_id=key,
                         duration_days=0, start=schedule.project_start, finish=schedule.project_start),
        ])
    return schedule

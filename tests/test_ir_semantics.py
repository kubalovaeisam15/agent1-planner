from datetime import date, timedelta

import pytest

from schedule_ir import ScheduleLink, ScheduleProject, ScheduleTask, validate_schedule_ir


def example():
    return ScheduleProject('test', 'Проверка', date(2026, 1, 1), [
        ScheduleTask('1', 'Предшественник', 1, 'task', duration_days=10,
                     start=date(2026, 1, 1), finish=date(2026, 1, 11)),
        ScheduleTask('2', 'Последователь', 1, 'task', duration_days=3,
                     start=date(2026, 1, 11), finish=date(2026, 1, 14)),
    ])


@pytest.mark.parametrize('kind', ['FS', 'SS', 'FF', 'SF'])
@pytest.mark.parametrize('lag', [-5, 0, 5])
def test_link_boundary_and_one_day_violation(kind, lag):
    schedule = example()
    predecessor, task = schedule.tasks
    task.predecessors = [ScheduleLink('1', kind, lag)]
    source = predecessor.finish if kind[0] == 'F' else predecessor.start
    target = source + timedelta(days=lag)
    task.start = target if kind[1] == 'S' else target - timedelta(days=3)
    task.finish = task.start + timedelta(days=3)
    assert 'IR-LINK-DATES' not in {i.code for i in validate_schedule_ir(schedule)}
    task.start -= timedelta(days=1)
    task.finish -= timedelta(days=1)
    assert 'IR-LINK-DATES' in {i.code for i in validate_schedule_ir(schedule)}


@pytest.mark.parametrize('field,value,code', [
    ('start', None, 'IR-DATES-REQUIRED'),
    ('finish', None, 'IR-DATES-REQUIRED'),
    ('duration_days', None, 'IR-DURATION-REQUIRED'),
    ('percent_complete', 101, 'IR-PERCENT'),
    ('percent_complete', -1, 'IR-PERCENT'),
    ('percent_complete', True, 'IR-PERCENT'),
])
def test_invalid_task_fields(field, value, code):
    schedule = example()
    setattr(schedule.tasks[0], field, value)
    assert code in {i.code for i in validate_schedule_ir(schedule)}


def test_snet_and_validation_does_not_mutate():
    schedule = example()
    task = schedule.tasks[1]
    task.constraint_type = 'start_no_earlier_than'
    task.constraint_date = task.start + timedelta(days=1)
    before = schedule.to_json()
    assert 'IR-CONSTRAINT-VIOLATION' in {i.code for i in validate_schedule_ir(schedule)}
    assert schedule.to_json() == before

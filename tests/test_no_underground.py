"""Регрессия проекта без паркинга, свай и подземных этажей; решение 31.08.2026."""
import json
from pathlib import Path

import build_grp
from schedule_ir import schedule_from_grp, validate_schedule_ir
from shared_sections import PROJECT_SHARED_SECTIONS


def build_schedule(project):
    b = build_grp.Build(project)
    for method in (b.load_skeleton, b.check_stages, b.repair_defects,
                   b.inherit_summary_links, b.apply_site_conditions,
                   b.configure_corpuses, b.configure_zero_cycle,
                   b.configure_parking, b.apply_finishing_scope, b.apply_standards,
                   b.apply_absent_piles):
        method()
    for corpus in project["корпуса"]:
        b.rebuild_monolith(corpus)
    b.thermal(b.schedule())
    b.wire_zos()
    return schedule_from_grp(project, b.finalize(b.schedule()))


def test_two_corpuses_without_basement_wait_for_own_raft():
    project = json.loads((Path(__file__).parent / "regression_two_corpuses.json").read_text(encoding="utf-8"))
    project["паркинг"] = {"есть": False}
    project["старт_проекта"] = "15.10.2026"
    project["нулевой_цикл"]["сваи"] = []
    project["нулевой_цикл"]["этажей_подземных"] = 0
    for corpus, floors in zip(project["корпуса"], (10, 15)):
        corpus["этажей_надземных"] = floors
        corpus["этажей_подземных"] = 0
        corpus["остекление"] = {"пвх": True, "витражи": False}
    schedule = build_schedule(project)
    assert validate_schedule_ir(schedule) == []
    by_id = {t.task_id: t for t in schedule.tasks}
    for pref in ("К1", "К2"):
        raft = [t for t in schedule.tasks if t.name == f"{pref}. По договору Фундаменты Корпус"]
        assert len(raft) == 1
        first = next(t for t in schedule.tasks if t.name == f"{pref}. 1 этаж Монолит")
        assert len(first.predecessors) == 1
        link = first.predecessors[0]
        assert by_id[link.predecessor_id] == raft[0]
        assert link.type == "FS" and link.lag_days == 0
        assert first.start == raft[0].finish
    assert not any("По договору Фундаменты Паркинг" in t.name for t in schedule.tasks)
    assert all(t.duration_days == 0 for t in schedule.tasks
               if "По договору Монолитные конструкции ниже отм 0,000" in t.name)
    assert not any("свайн" in t.name.lower() or "арматурный каркас сваи" in t.name.lower()
                   for t in schedule.tasks)
    assert sum("По договору ЦТП, ИТП с УУТЭ" in t.name for t in schedule.tasks) == 2
    for pref in ("К1", "К2"):
        heat = next(t for t in schedule.tasks if t.name == f"{pref}. Пуск тепла корпус")
        itp_preds = [by_id[p.predecessor_id] for p in heat.predecessors
                     if "По договору ЦТП, ИТП с УУТЭ" in by_id[p.predecessor_id].name]
        assert len(itp_preds) == 1 and itp_preds[0].name.startswith(pref + ". ")
    for name in PROJECT_SHARED_SECTIONS:
        assert sum(t.name == name and t.outline_level == 1 for t in schedule.tasks) == 1

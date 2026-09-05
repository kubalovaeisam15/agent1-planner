"""BND-SPK-003: состав ограждения определяется ТЭП, не прототипом."""
import copy
import json
from datetime import timedelta
from pathlib import Path

import pytest
import build_grp
from schedule_ir import IRIssue, schedule_from_grp, validate_schedule_ir
from project_ir_validation import validate_project_against_ir


def build_case(glazing, facade="НВФ", count=1):
    project = json.loads((Path(__file__).parent / "etalon_project.json").read_text(encoding="utf-8"))
    corpus = project["корпуса"][0]
    corpus.update(этажей_надземных=9, этажей_подземных=0, остекление=glazing)
    project["корпуса"] = [dict(copy.deepcopy(corpus), код=f"К{i+1}") for i in range(count)]
    project["фасад"]["тип"] = facade
    b = build_grp.Build(project)
    for step in (b.load_skeleton, b.check_stages, b.repair_defects,
                 b.inherit_summary_links, b.apply_site_conditions, b.configure_corpuses,
                 b.configure_zero_cycle, b.configure_parking, b.apply_finishing_scope,
                 b.apply_standards, b.apply_absent_piles):
        step()
    for c in project["корпуса"]:
        b.rebuild_monolith(c)
    b.thermal(b.schedule())
    b.wire_zos()
    nodes = b.schedule()
    b.constrain_tender_tz(nodes)
    return project, schedule_from_grp(project, b.finalize(b.schedule()))


@pytest.mark.parametrize("facade", ["НВФ", "СФТК", "модульный"])
@pytest.mark.parametrize("glazing", [
    {"пвх": True, "витражи": True, "витражи_на_всю_высоту": False},
    {"пвх": False, "витражи": True, "витражи_на_всю_высоту": True},
    {"пвх": True, "витражи": True, "витражи_на_всю_высоту": True},
    {"пвх": True, "витражи": False, "витражи_на_всю_высоту": False},
    {"пвх": False, "витражи": True, "витражи_на_всю_высоту": False},
])
def test_glazing_scope_and_contour(glazing, facade):
    project, ir = build_case(glazing, facade)
    assert validate_schedule_ir(ir) == []
    assert validate_project_against_ir(project, ir) == []
    rows = [t for t in ir.tasks if t.name.startswith("К1. ")]
    pvc = [t for t in rows if t.name == "К1. По договору Монтаж светопрозрачных конструкций ПВХ"]
    stained = [t for t in rows if t.name.startswith("К1. По договору Монтаж светопрозрачных конструкций Витраж")]
    masonry = [t for t in rows if t.name == "К1. По договору Кладка наружных стен"]
    assert len(pvc) == int(glazing["пвх"])
    assert len(stained) == int(glazing["витражи"])
    assert len(masonry) == int(not glazing["витражи_на_всю_высоту"])
    roof = next(t for t in rows if t.name == "К1. Кровля/Парапет Монолит")
    if stained:
        floor = next(t for t in rows if t.name == "К1. 9 этаж Монолит")
        assert (floor.task_id, "FS", 0) in {(p.predecessor_id, p.type, p.lag_days) for p in stained[0].predecessors}
        assert stained[0].finish == roof.finish + timedelta(days=129)
    if pvc:
        anchor = stained[0] if glazing["витражи_на_всю_высоту"] else masonry[0]
        assert (anchor.task_id, "SS", 30) in {(p.predecessor_id, p.type, p.lag_days) for p in pvc[0].predecessors}
        assert pvc[0].finish == roof.finish + timedelta(days=129 if glazing["витражи_на_всю_высоту"] else 75)
    contour = next(t for t in rows if t.name == "К1. Закрыт тепловой контур по корпусу")
    expected = pvc + stained if facade != "модульный" else [
        next(t for t in rows if t.name == "К1. По договору Монтаж фасадов")]
    assert {(p.predecessor_id, p.type, p.lag_days) for p in contour.predecessors} == {
        (t.task_id, "FS", 0) for t in expected}
    assert contour.finish == max(t.finish for t in expected)


def test_mixed_six_corpuses_have_own_glazing_predecessors():
    _, ir = build_case({"пвх": True, "витражи": True, "витражи_на_всю_высоту": False}, count=6)
    by_id = {t.task_id: t for t in ir.tasks}
    assert len(by_id) == len(ir.tasks)
    assert validate_schedule_ir(ir) == []
    for i in range(1, 7):
        prefix = f"К{i}. "
        contour = next(t for t in ir.tasks if t.name == prefix + "Закрыт тепловой контур по корпусу")
        assert len(contour.predecessors) == 2
        assert all(by_id[p.predecessor_id].name.startswith(prefix) for p in contour.predecessors)


@pytest.mark.parametrize("mutation,code", [
    ("missing-glazing", "PROJECT-GLAZING-SCOPE"),
    ("missing-link", "PROJECT-CONTOUR-PREDECESSORS"),
    ("extra-masonry", "PROJECT-MASONRY-SCOPE"),
])
def test_independent_validator_rejects_mutations(mutation, code):
    project, ir = build_case({"пвх": True, "витражи": True, "витражи_на_всю_высоту": False})
    if mutation == "missing-glazing":
        ir.tasks = [t for t in ir.tasks if "По договору Монтаж светопрозрачных конструкций Витраж" not in t.name]
    elif mutation == "missing-link":
        next(t for t in ir.tasks if t.name == "К1. Закрыт тепловой контур по корпусу").predecessors = []
    else:
        project["корпуса"][0]["остекление"]["витражи_на_всю_высоту"] = True
    assert code in {i.code for i in validate_project_against_ir(project, ir)}


@pytest.mark.parametrize("emit_ir", [False, True])
def test_cli_checks_project_against_ir_before_any_output(tmp_path, monkeypatch, emit_ir):
    source = Path(__file__).parent / "etalon_project.json"
    xlsx = tmp_path / "blocked.xlsx"
    ir_path = tmp_path / "blocked.ir.json"
    calls = []

    def reject(spec, ir):
        calls.append(ir)
        return [IRIssue("PROJECT-GLAZING-SCOPE", "test rejection")]

    monkeypatch.setattr(build_grp, "validate_project_against_ir", reject)
    args = [str(source), str(xlsx)] + (["--ir", str(ir_path)] if emit_ir else [])
    assert build_grp.main(args) == 1
    assert len(calls) == 1
    assert not xlsx.exists()
    assert not ir_path.exists()

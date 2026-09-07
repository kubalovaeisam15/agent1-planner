import copy
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from build_grp import Build, validate_project_spec


def spec():
    return json.loads((ROOT / "tests/project_three_corpuses_20260906.json").read_text(encoding="utf-8"))


def test_distribution_is_required_and_reconciled():
    p = spec()
    assert validate_project_spec(p) == []
    del p["корпуса"][1]["сваи"]
    assert any("DEC-42" in e for e in validate_project_spec(p))
    p = spec()
    p["корпуса"][2]["сваи"][0]["количество"] = 65
    assert any("сумма" in e for e in validate_project_spec(p))


def prepare(p):
    b = Build(p)
    for method in ("load_skeleton", "check_stages", "repair_defects", "inherit_summary_links",
                   "apply_site_conditions", "configure_corpuses", "configure_zero_cycle",
                   "configure_parking", "apply_finishing_scope", "apply_standards", "wire_corpus_piles"):
        getattr(b, method)()
    return b


def test_foundations_follow_own_piles_or_earth_even_when_corpuses_reordered():
    p = spec()
    for bodies in (p["корпуса"], list(reversed(p["корпуса"]))):
        current = copy.deepcopy(p)
        current["корпуса"] = copy.deepcopy(bodies)
        b = prepare(current)
        earth = b.rows[b.one("По договору Земляные работы")]["key"]
        for n, corpus in enumerate(current["корпуса"], 1):
            raft = b.rows[b.one("По договору Фундаменты Корпус", f"К{n}")]
            if corpus["сваи"]:
                own = b.rows[b.one("По договору Свайное основание БНС (при наличии)", f"К{n}")]
                assert raft["links"] == [(own["key"], "ОН", 0)]
            else:
                assert raft["links"] == [(earth, "ОН", 0)]


def test_no_piles_and_two_types():
    p = spec()
    p["корпуса"][2]["сваи"] = []
    p["нулевой_цикл"]["сваи"][0]["количество"] = 0
    b = prepare(p)
    b.apply_absent_piles()
    earth = b.rows[b.one("По договору Земляные работы")]["key"]
    assert all(b.rows[b.one("По договору Фундаменты Корпус", f"К{n}")]["links"] == [(earth, "ОН", 0)] for n in (1, 2, 3))
    p = spec()
    extra = {"тип": "забивные", "количество": 180}
    p["корпуса"][2]["сваи"].append(extra)
    p["нулевой_цикл"]["сваи"].append(extra)
    b = prepare(p)
    links = b.rows[b.one("По договору Фундаменты Корпус", "К3")]["links"]
    assert len(links) == 2
    assert all(kind == "ОН" and lag == 0 for _, kind, lag in links)


def test_full_build_and_independent_checker_reject_foreign_piles(tmp_path):
    from build_grp import main
    from schedule_ir import ScheduleProject
    from project_ir_validation import validate_project_against_ir
    ir_path = tmp_path / "schedule.ir.json"
    assert main([str(ROOT / "tests/project_three_corpuses_20260906.json"),
                 str(tmp_path / "schedule.xlsx"), "--ir", str(ir_path)]) == 0
    ir = ScheduleProject.from_json(ir_path.read_text(encoding="utf-8"))
    assert validate_project_against_ir(spec(), ir) == []
    raft2 = next(t for t in ir.tasks if t.name == "К2. По договору Фундаменты Корпус")
    raft3 = next(t for t in ir.tasks if t.name == "К3. По договору Фундаменты Корпус")
    raft2.predecessors[:] = raft3.predecessors
    assert any(e.code == "PROJECT-FOUNDATION-PREDECESSORS" for e in validate_project_against_ir(spec(), ir))

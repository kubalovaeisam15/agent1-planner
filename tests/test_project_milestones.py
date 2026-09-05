"""BND-ZOS-001 / DEC-33: предметные предшественники, без подмены сводок."""
import copy
from datetime import timedelta
import pytest
from project_ir_validation import validate_project_against_ir
from schedule_ir import ScheduleLink, validate_schedule_ir
from test_glazing_scope import build_case


@pytest.fixture(scope="module")
def built():
    return build_case({"пвх":True,"витражи":True,"витражи_на_всю_высоту":False},count=6)


def test_bti_has_summary_predecessor_for_each_corpus(built):
    spec, ir = built
    bti = next(t for t in ir.tasks if t.name == "Готовность к обмерам БТИ")
    expected = [next(t for t in ir.tasks if t.name == f"К{n}. Отделочные работы") for n in range(1,7)]
    assert all(t.task_type == "summary" for t in expected)
    assert sorted((p.predecessor_id,p.type,p.lag_days) for p in bti.predecessors) == sorted(
        (t.task_id,"FS",-160) for t in expected)
    assert bti.finish == max(t.finish for t in expected)-timedelta(days=160)
    assert validate_project_against_ir(spec,ir) == []


@pytest.mark.parametrize("mutation", ["missing-corpus","lag","leaf","missing-milestone","duplicate","date"])
def test_bti_mutations_rejected(built,mutation):
    spec, source = built
    ir = copy.deepcopy(source)
    bti = next(t for t in ir.tasks if t.name == "Готовность к обмерам БТИ")
    if mutation == "missing-corpus":
        bti.predecessors.pop()
    elif mutation == "lag":
        p=bti.predecessors[0]; bti.predecessors[0]=ScheduleLink(p.predecessor_id,"FS",-159)
    elif mutation == "leaf":
        leaf=next(t for t in ir.tasks if t.name == "К1. По договору Вестибюль")
        bti.predecessors[0]=ScheduleLink(leaf.task_id,"FS",-160)
    elif mutation == "missing-milestone":
        ir.tasks.remove(bti)
    elif mutation == "duplicate":
        bti.predecessors.append(bti.predecessors[0])
    else:
        bti.start+=timedelta(days=1); bti.finish+=timedelta(days=1)
    assert any(i.code.startswith("PROJECT-BTI") for i in validate_project_against_ir(spec,ir))


@pytest.mark.parametrize("mutation", ["missing","lag","foreign-corpus","duplicate"])
def test_heat_mutations_rejected(built,mutation):
    spec, source = built
    ir=copy.deepcopy(source)
    heat=next(t for t in ir.tasks if t.name=="К1. Пуск тепла корпус")
    if mutation=="missing":
        heat.predecessors.pop()
    elif mutation=="lag":
        p=heat.predecessors[0]; heat.predecessors[0]=ScheduleLink(p.predecessor_id,"FS",14)
    elif mutation == "foreign-corpus":
        other=next(t for t in ir.tasks if t.name=="К2. По договору Отопление (контур для пуска тепла)")
        heat.predecessors[-1]=ScheduleLink(other.task_id,"FS",15)
    else:
        heat.predecessors.append(heat.predecessors[0])
    assert "PROJECT-HEAT-PREDECESSORS" in {i.code for i in validate_project_against_ir(spec,ir)}


@pytest.mark.parametrize("value", [0,50,100])
def test_summary_percent_must_be_absent(built,value):
    _, source=built
    ir=copy.deepcopy(source)
    next(t for t in ir.tasks if t.task_type=="summary").percent_complete=value
    assert "IR-SUMMARY-PERCENT" in {i.code for i in validate_schedule_ir(ir)}

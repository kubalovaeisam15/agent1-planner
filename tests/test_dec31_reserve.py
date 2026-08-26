# -*- coding: utf-8 -*-
"""DEC-31: резерв фазы РС→РВЭ — (Окончание − старт СМР)/365*30, округление вверх.

Эталон проверки — утверждённая карта вех «Южнопортовая 2.1» (старт СМР 14.03.2025)
из калибровочной выборки D:\\Claude\\ClaudeVS\\data. Формула воспроизводит её
утверждённые даты точно; тест фиксирует это, чтобы норматив не поехал молча.
"""
from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path

import pytest

import build_grp
from tools.validate_grp import RESERVE_COL, Report, check_reserve

ROOT = Path(__file__).resolve().parents[1]
SMR = "14.03.2025"

# (наименование вехи, ранняя дата, утверждённая дата) — из утверждённой карты вех.
APPROVED = [
    ("Завершено ограждение котлована",                  "20.08.2025", "03.09.2025"),
    ("Завершены земляные работы",                       "19.01.2026", "14.02.2026"),
    ("Выполнен фундамент здания",                       "05.03.2026", "04.04.2026"),
    ("Последний куб бетона",                            "10.03.2027", "09.05.2027"),
    ("Завершен монтаж ИТП и систем отопления",          "01.08.2027", "12.10.2027"),
    ("Завершены фасадные работы",                       "27.02.2028", "26.05.2028"),
    ("Получен ЗОС",                                     "16.05.2028", "20.08.2028"),
    ("Получено РВЭ",                                    "22.05.2028", "26.08.2028"),
    ("Построенный объект предан УК",                    "21.07.2028", "30.10.2028"),
    # Договор заключён до старта СМР — резерв 0, дата не сдвигается.
    ("Подписаны договоры на монтаж наружных инженерных систем ЭС",
                                                        "23.01.2023", "23.01.2023"),
]


def reserve_days(fin: str, smr: str = SMR) -> int:
    age = (datetime.strptime(fin, "%d.%m.%Y") - datetime.strptime(smr, "%d.%m.%Y")).days
    return -(-age * 30 // 365) if age > 0 else 0


def with_reserve(fin: str) -> str:
    d = datetime.strptime(fin, "%d.%m.%Y") + timedelta(days=reserve_days(fin))
    return d.strftime("%d.%m.%Y")


def test_formula_reproduces_approved_dates():
    """Формула даёт утверждённую дату карты вех."""
    for name, early, approved in APPROVED:
        assert with_reserve(early) == approved, name


def test_reserve_is_zero_before_smr():
    """Задача, завершённая до старта СМР, резерва не получает."""
    assert reserve_days("23.01.2023") == 0
    assert with_reserve("23.01.2023") == "23.01.2023"


def test_reserve_grows_with_age():
    """Резерв монотонно растёт по мере удаления вехи от старта СМР."""
    seq = [reserve_days(e) for _, e, _ in APPROVED if reserve_days(e) > 0]
    assert seq == sorted(seq)


def row(idn: int, name: str, fin: str, res: str) -> dict:
    return {"Вид работ": "", "Код классификатора": "", "Уровень структуры": "2",
            "Ид.": str(idn), "Название задачи": name, "% завершения": "0",
            "Длительность": "0 дней", "Начало": fin, "Окончание": fin,
            "Предшественники": "", "Последователи": "", "комментарий": "",
            RESERVE_COL: res}


def sample() -> list[dict]:
    tasks = [row(1, "Старт СМР", SMR, "")]
    tasks += [row(i, n, e, a) for i, (n, e, a) in enumerate(APPROVED, start=2)]
    return tasks


def test_validator_accepts_correct_reserve():
    r = Report()
    check_reserve(sample(), r)
    assert not r.errors, r.errors


def test_validator_rejects_wrong_reserve():
    tasks = sample()
    tasks[4][RESERVE_COL] = "01.06.2027"     # вместо 09.05.2027
    r = Report()
    check_reserve(tasks, r)
    assert r.errors


def test_validator_requires_reserve_column():
    tasks = [{k: v for k, v in t.items() if k != RESERVE_COL} for t in sample()]
    r = Report()
    check_reserve(tasks, r)
    assert any(RESERVE_COL in e for e in r.errors)


@pytest.fixture(scope="module")
def generated_etalon():
    project = json.loads((ROOT / "tests" / "etalon_project.json").read_text(encoding="utf-8"))
    b = build_grp.Build(project)
    b.load_skeleton()
    b.check_stages()
    b.repair_defects()
    b.inherit_summary_links()
    b.apply_site_conditions()
    b.configure_corpuses()
    b.configure_zero_cycle()
    b.configure_parking()
    b.apply_finishing_scope()
    b.apply_standards()
    for corpus in project["корпуса"]:
        b.rebuild_monolith(corpus)
    nodes = b.schedule()
    b.thermal(nodes)
    b.wire_zos()
    nodes = b.schedule()
    rows = b.finalize(nodes)
    return b, nodes, rows


def test_generator_writes_column_and_keeps_phase_a_empty(generated_etalon):
    _, _, rows = generated_etalon
    assert RESERVE_COL in build_grp.COLUMNS
    rs = next(r for r in rows
              if r["Название задачи"] == "Разрешение на строительство (РС) получено")
    rve = next(r for r in rows if r["Название задачи"] == "Получено РВЭ")
    assert rs[RESERVE_COL] == ""       # фаза A
    assert rve[RESERVE_COL]             # фаза B


def test_generator_falls_back_to_calculated_rs(generated_etalon):
    b, nodes, _ = generated_etalon
    rs = b.one("Разрешение на строительство (РС) получено")
    assert rs is not None
    assert b.reserve_base(nodes) == nodes[b.rows[rs]["key"]].finish
    assert any(a.code == "DEC-31" and "fallback" in a.text for a in b.assumptions)


def test_explicit_start_smr_has_priority():
    b = build_grp.Build({"старт_проекта": "10.01.2023", "старт_смр": SMR})
    assert b.reserve_base({}).strftime("%d.%m.%Y") == SMR

# -*- coding: utf-8 -*-
"""Регрессия DEC-33: реальные предшественники теплового блока."""
from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from pathlib import Path

import pytest

import build_grp

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def built():
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
    return b


def key(b, name, prefix):
    i = b.one(name, prefix, exact=False)
    assert i is not None, f"не найдена задача {prefix}. {name}"
    return b.rows[i]["key"]


def test_postoyannyy_kontur_zavisit_ot_ostekleniya_a_ne_ot_monolita(built):
    b = built
    contour = b.one("Закрыт тепловой контур по корпусу", "К1")
    assert contour is not None
    pvc = key(b, "По договору Монтаж светопрозрачных конструкций ПВХ", "К1")
    roof = key(b, "Кровля/Парапет Монолит", "К1")
    assert b.rows[contour]["links"] == [(pvc, "ОН", 0)]
    assert all(pred != roof for pred, _, _ in b.rows[contour]["links"])


def test_vremennyy_kontur_finishiruet_vmeste_s_naruzhnoy_kladkoy(built):
    b = built
    vtk = b.one("Закрыт ВРЕМЕННЫЙ тепловой контур по корпусу (при необходимости)", "К2")
    assert vtk is not None, "на эталоне К2 временный контур должен быть развёрнут"
    masonry = key(b, "По договору Кладка наружных стен", "К2")
    assert b.rows[vtk]["dur"] == 45
    assert b.rows[vtk]["links"] == [(masonry, "ОО", 0)]


def test_start_otdelki_uses_final_heat_date_after_network_recalculation(built):
    """BND-OTD-002: SNET отделки сверяется с итоговым пуском тепла."""
    b = built
    nodes = b.schedule()
    heat = b.one("Пуск тепла корпус", "К2")
    partition = b.one("По договору Кладка перегородок", "К2")
    finishing = b.one("По договору Вестибюль", "К2")
    assert heat is not None and partition is not None and finishing is not None

    s0 = nodes[b.rows[partition]["key"]].start + timedelta(days=90)
    heat_date = nodes[b.rows[heat]["key"]].start
    expected = (
        min(heat_date, date(s0.year + 1, 4, 1))
        if heat_date > date(s0.year, 10, 30)
        else s0
    )
    actual = datetime.strptime(
        b.rows[finishing]["constraint_date"], "%d.%m.%Y"
    ).date()
    assert actual == expected


def test_pusk_tepla_imeet_tri_fakticheskih_predshestvennika(built):
    b = built
    heat = b.one("Пуск тепла корпус", "К2")
    assert heat is not None
    vtk = key(b, "Закрыт ВРЕМЕННЫЙ тепловой контур по корпусу (при необходимости)", "К2")
    itp = key(b, "По договору ЦТП, ИТП с УУТЭ", "П1")
    loop = key(b, "По договору Отопление (контур для пуска тепла)", "К2")
    roof = key(b, "Кровля/Парапет Монолит", "К2")
    assert set(b.rows[heat]["links"]) == {
        (vtk, "ОН", 15),
        (itp, "ОН", 15),
        (loop, "ОН", 15),
    }
    assert all(pred != roof for pred, _, _ in b.rows[heat]["links"])


def test_pusk_tepla_k1_ssylaetsya_na_kontur_a_ne_na_vsyu_sistemu(built):
    """Удаление неприменимого ВТК не должно сдвигать ссылку на соседнюю строку."""
    b = built
    heat = b.one("Пуск тепла корпус", "К1")
    assert heat is not None
    loop = key(b, "По договору Отопление (контур для пуска тепла)", "К1")
    full = key(b, "По договору Отопление (все система полностью)", "К1")
    predecessors = {pred for pred, _, _ in b.rows[heat]["links"]}
    assert loop in predecessors
    assert full not in predecessors


@pytest.mark.parametrize("corpus", ["К1", "К2"])
def test_facade_starts_ss_30_from_own_spk(built, corpus):
    """BND-FAS-001: договор не должен преждевременно запускать фасад."""
    b = built
    facade = b.one("По договору Монтаж фасадов", corpus)
    spk_rows = b.find(
        "По договору Монтаж светопрозрачных конструкций",
        corpus,
        exact=False,
    )
    spk = next((i for i in spk_rows if "ПВХ" in b.rows[i]["name"]), None)
    if spk is None:
        spk = next((i for i in spk_rows if "Витраж" in b.rows[i]["name"]), None)
    assert facade is not None and spk is not None
    expected = (b.rows[spk]["key"], "НН", 30)
    assert expected in b.rows[facade]["links"]

    nodes = b.schedule()
    assert nodes[b.rows[facade]["key"]].start == (
        nodes[b.rows[spk]["key"]].start + timedelta(days=30)
    )

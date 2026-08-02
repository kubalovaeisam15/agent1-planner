# -*- coding: utf-8 -*-
"""Краевые ветки проекта «башня 69 этажей, встроенный паркинг»."""
from __future__ import annotations

import json
from pathlib import Path

import pytest

import build_grp

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def built():
    """Собранный проект 69 этажей: строки и рассчитанная сеть."""
    project = json.loads((ROOT / "tests" / "project_69.json").read_text(encoding="utf-8"))
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
    return b, nodes


def test_otdelnogo_obekta_parkinga_net(built):
    b, _ = built
    assert all(not r["name"].startswith("П1.") for r in b.rows), \
        "при встроенном паркинге отдельного объекта П1 быть не должно"


def test_itp_perenesen_v_korpus_i_poluchil_normativ(built):
    b, _ = built
    i = b.one("По договору ЦТП, ИТП с УУТЭ", "К1")
    assert i is not None, "ИТП не перенесён в корпус (DEC-06, R-05)"
    assert b.rows[i]["dur"] == build_grp.Std.ITP[0]


def test_tehpomeshcheniya_perenoseny_i_yakoryat_itp(built):
    b, _ = built
    tech = b.one("По договору Технические помещения", "К1")
    itp = b.one("По договору ЦТП, ИТП с УУТЭ", "К1")
    assert tech is not None and itp is not None
    tech_key = b.rows[tech]["key"]
    assert any(t == tech_key for t, _, _ in b.rows[itp]["links"]), \
        "ИТП обязан стартовать от отделки техпомещений (BND-VIS-006)"


def test_vekha_pusk_tepla_korpusa_est(built):
    b, _ = built
    assert b.one("Пуск тепла корпус", "К1") is not None


def test_vekha_pusk_tepla_v_parkinge_libo_snyata_libo_privyazana(built):
    """Обязательная веха §9. При встроенном паркинге объекта нет — веха либо
    отсутствует и это записано в «Допущения», либо привязана к корпусу."""
    b, _ = built
    hits = b.find("Пуск тепла в паркинге", exact=False)
    if not hits:
        assert any("паркинг" in a.text.lower() and "тепл" in a.text.lower()
                   for a in b.assumptions), \
            "веха «Пуск тепла в паркинге» снята молча — нарушение CLAUDE.md §12"
    else:
        assert b.rows[hits[0]]["links"], "веха осталась без предшественников"

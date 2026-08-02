# -*- coding: utf-8 -*-
"""DEC-30: при доле чистовой отделки 0 ветка отделки квартир в график не входит."""
from __future__ import annotations

import json
from pathlib import Path

import build_grp

ROOT = Path(__file__).resolve().parents[1]

# Маркеры ветки чистовой отделки квартир, в нижнем регистре.
MARKERS = (
    "отделочные работы квартиры",
    "подготовка под чистовую отделку квартир",
    "чистовая отделка квартир",
    "тендер отделка квартиры чистовая",
    "мокап отделки типового этажа",
    "завершены отделочные работы квартир",
    "передача квартир с отделкой",
    "переданы квартиры с отделкой",
    "отделка квартиры",
)


def build(spec_name: str) -> build_grp.Build:
    project = json.loads((ROOT / "tests" / spec_name).read_text(encoding="utf-8"))
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
    return b


def names(b: build_grp.Build) -> list[str]:
    return [r["name"].lower() for r in b.rows]


def test_pri_dole_nol_vetka_kvartirnoy_otdelki_snyata():
    b = build("project_69.json")
    found = [n for n in names(b) if any(m in n for m in MARKERS)]
    assert found == [], f"остались строки ветки отделки квартир: {found[:10]}"


def test_pri_dole_nol_peredacha_bez_otdelki_sohranena():
    b = build("project_69.json")
    assert any(n == "передача квартир без отделки" for n in names(b))
    assert any(n == "переданы квартиры без отделки" for n in names(b))


def test_pri_dole_nol_mokap_lobbi_sohranen_i_ne_ostalsya_bez_svyazey():
    b = build("project_69.json")
    i = b.one("Разработка рабочей документации на АИ МОКАП отделки лобби")
    assert i is not None, "блок МОКАП отделки лобби удалён вместе с типовым этажом"
    assert b.rows[i]["links"], "МОКАП лобби остался без предшественников"


def test_pri_dole_nol_vekha_zaversheny_otdelochnye_raboty_sohranena():
    b = build("project_69.json")
    i = b.one("Завершены отделочные работы")
    assert i is not None
    assert b.rows[i]["links"], "веха потеряла все связи вместе с квартирной"


def test_v_etalone_vetka_kvartirnoy_otdelki_na_meste():
    b = build("etalon_project.json")
    found = [n for n in names(b) if any(m in n for m in MARKERS)]
    assert found, "при доле 1,0 ветка отделки квартир обязана остаться"

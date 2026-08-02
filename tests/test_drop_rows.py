# -*- coding: utf-8 -*-
"""Юнит-тесты Build.drop_rows — удаление строк без молчаливой потери связей."""
from __future__ import annotations

import build_grp


def make_build(rows):
    """Build без прогона __init__: помощнику нужны только rows и журналы."""
    b = build_grp.Build.__new__(build_grp.Build)
    b.rows = rows
    b.assumptions = []
    b.rationale = []
    b.notices = []
    return b


def row(key, lvl, name, links=()):
    return {"key": key, "lvl": lvl, "name": name, "dur": 1,
            "links": [tuple(x) for x in links], "comment": "",
            "tpl_start": "", "src": ""}


def test_udalyaet_podderevo_celikom():
    b = make_build([
        row("a", 1, "Корень"),
        row("b", 2, "Ветка под удаление"),
        row("c", 3, "Лист ветки"),
        row("d", 2, "Соседняя ветка"),
    ])
    removed = b.drop_rows([1], "тест")
    assert removed == 2
    assert [r["key"] for r in b.rows] == ["a", "d"]


def test_potrebitel_s_drugimi_predshestvennikami_teryaet_tolko_bituyu_svyaz():
    b = make_build([
        row("live", 1, "Живой предшественник"),
        row("dead", 1, "Удаляемая", [("live", "ОН", 0)]),
        row("user", 1, "Потребитель", [("live", "ОН", 0), ("dead", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert b.rows[-1]["links"] == [("live", "ОН", 0)]
    assert any("Потребитель" in w["Комментарий"] for w in b.rationale)


def test_potrebitel_bez_drugih_svyazey_nasleduet_predshestvennikov_udalyaemoy():
    b = make_build([
        row("root", 1, "Корень цепочки"),
        row("dead", 1, "Удаляемая", [("root", "НН", 5)]),
        row("user", 1, "Потребитель", [("dead", "НН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert b.rows[-1]["links"] == [("root", "НН", 5)]


def test_nasledovanie_prohodit_skvoz_cepochku_udalyaemyh():
    b = make_build([
        row("root", 1, "Живой корень"),
        row("dead1", 1, "Удаляемая внешняя", [("root", "ОН", 3)]),
        row("dead2", 2, "Удаляемая вложенная", [("dead1", "ОН", 0)]),
        row("user", 1, "Потребитель", [("dead2", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert [r["key"] for r in b.rows] == ["root", "user"]
    assert b.rows[-1]["links"] == [("root", "ОН", 3)]


def test_pustoy_spisok_korney_nichego_ne_menyaet():
    b = make_build([row("a", 1, "Одна строка")])
    assert b.drop_rows([], "тест") == 0
    assert len(b.rows) == 1

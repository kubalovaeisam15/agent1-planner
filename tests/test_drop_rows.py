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


def test_dve_svyazi_udalyaemoy_na_odin_zhivoy_uzel_ne_shlopyvayutsya():
    b = make_build([
        row("c", 1, "Живой узел"),
        row("t", 1, "Удаляемая с двумя связями на один узел",
            [("c", "ОН", 5), ("c", "НН", 10)]),
        row("user", 1, "Потребитель", [("t", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert b.rows[-1]["links"] == [("c", "ОН", 5), ("c", "НН", 10)]


def test_diamant_cherez_dve_udalyaemye_na_odin_zhivoy_uzel_razreshaetsya_oboimi_putyami():
    b = make_build([
        row("c", 1, "Живой узел"),
        row("branch", 1, "Удаляемая ветка"),
        row("a", 2, "Удаляемая A", [("c", "ОН", 5)]),
        row("b", 2, "Удаляемая B", [("c", "ОН", 10)]),
        row("t", 2, "Удаляемая T", [("a", "ОН", 0), ("b", "ОН", 0)]),
        row("user", 1, "Потребитель", [("t", "ОН", 0)]),
    ])
    b.drop_rows([1], "тест")
    assert [r["key"] for r in b.rows] == ["c", "user"]
    assert b.rows[-1]["links"] == [("c", "ОН", 5), ("c", "ОН", 10)]


def test_potrebitel_s_dvumya_pryamymi_udalyaemymi_predshestvennikami_nasleduet_oba():
    b = make_build([
        row("root1", 1, "Живой корень 1"),
        row("root2", 1, "Живой корень 2"),
        row("dead1", 1, "Удаляемая 1", [("root1", "ОН", 3)]),
        row("dead2", 1, "Удаляемая 2", [("root2", "НН", 7)]),
        row("user", 1, "Потребитель", [("dead1", "ОН", 0), ("dead2", "ОН", 0)]),
    ])
    b.drop_rows([2, 3], "тест")
    assert [r["key"] for r in b.rows] == ["root1", "root2", "user"]
    assert b.rows[-1]["links"] == [("root1", "ОН", 3), ("root2", "НН", 7)]

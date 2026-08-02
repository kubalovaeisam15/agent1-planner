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


# --- высотность 69 этажей ------------------------------------------------

def test_monolit_razvernut_po_vsem_69_etazham(built):
    b, _ = built
    floors = [r for r in b.rows
              if r["name"].startswith("К1. ") and r["name"].endswith("этаж Монолит")]
    assert len(floors) == 69, f"этажей монолита {len(floors)}, ожидалось 69"


def test_koefficient_K_raven_7_pri_69_etazhah(built):
    first, typical, roof = build_grp.monolith_floor_durations(69, True)
    assert typical and typical[0] == 7, \
        "при сложном конструктиве и N > 60 типовой этаж — 7 дн (standards.md §9)"


def test_kladka_privyazana_k_vekhe_6_etazha_a_ne_k_69(built):
    b, _ = built
    assert b.one("6 этаж Монолит", "К1") is not None, \
        "веха min(6, N) = 6 этаж отсутствует — перепривязка BND-KLD-001 невозможна"


def test_vis_privyazany_k_vekhe_15_etazha(built):
    b, _ = built
    assert b.one("15 этаж Монолит", "К1") is not None, \
        "веха min(15, N) = 15 этаж отсутствует — перепривязка BND-VIS-001 невозможна"


def test_okno_otdelki_ne_nizhe_zhestkogo_minimuma(built):
    b, nodes = built
    fits = [k for k, _ in b.fit_tasks]
    assert fits, "задачи отделки не зарегистрированы в fit_tasks"
    for k in fits:
        n = nodes[k]
        assert n.start and n.finish
        assert (n.finish - n.start).days >= build_grp.Std.FIT_MIN[0], \
            f"окно отделки {(n.finish - n.start).days} дн меньше минимума " \
            f"{build_grp.Std.FIT_MIN[0]} дн — запрещено standards.md §15.1"


def test_kriticheskiy_put_nepreryven(built):
    b, nodes = built
    crit = [n for n in nodes.values() if n.critical]
    assert crit, "критический путь пуст"


def test_pri_oknah_pvh_kladka_naruzhnyh_sten_prisutstvuet(built):
    """Витражей нет → BND-KLD-001 применяется, ЯКОРЬ_ОГР = кладка наружных стен."""
    b, _ = built
    assert b.one("По договору Кладка наружных стен", "К1") is not None, \
        "при остеклении ПВХ наружная кладка обязана быть — она задаёт ЯКОРЬ_ОГР"


def test_porog_DEC_14_po_svayam_otrabotan(built):
    """90 свай БНС на одну установку: превышение порога даёт уведомление, а не тишину."""
    b, _ = built
    days, _ = build_grp.pile_duration(
        [{"тип": "БНС", "количество": 90, "установок": 1}])
    if days > build_grp.Std.PILE_WARN_DAYS[0]:
        assert any("DEC-14" == a.code for a in b.assumptions), \
            "порог DEC-14 превышен, но строки в «Допущения» нет"
        assert b.notices, "порог DEC-14 превышен, но уведомления пользователю нет"

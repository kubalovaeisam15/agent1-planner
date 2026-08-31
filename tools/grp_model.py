# -*- coding: utf-8 -*-
"""Расчётное ядро ГРП: нормативы, сезонность, тепловой контур, CPM.

Каждая величина здесь трассируется до пункта файла контекста (CLAUDE.md §0.3).
Код норматива указан в поле `source` рядом со значением — он же попадает
на лист «Обоснование».

Зоны ответственности:
  standards.md — числовые длительности   (STD-*)
  bindings.md  — привязки, якоря, лаги    (BND-*)
  typGRP.md    — структура и регламентные сроки фазы A
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

DATE_FMT = "%d.%m.%Y"


def dparse(s: str) -> date:
    return datetime.strptime(s, DATE_FMT).date()


def dfmt(d: date) -> str:
    return d.strftime(DATE_FMT)


# ======================================================================
# 1. Нормативы длительностей — standards.md
# ======================================================================

class Std:
    """Числовые нормативы. Ключ — код норматива, значение — (число, пункт)."""

    # §9 Монолит надземной части · STD-MON-001
    MON_FIRST = (30, "standards.md §9 STD-MON-001, первый этаж")
    MON_TYPICAL = (6, "standards.md §9 STD-MON-001, типовой этаж K")
    MON_TYPICAL_HARD = (7, "standards.md §9 STD-MON-001, K при сложном конструктиве или N > 60")
    MON_ROOF = (15, "standards.md §9 STD-MON-001, кровля/парапет")
    SECTION_CASCADE = (7, "bindings.md §1.2 BND-MON-002, каскад секций НН +7 дн")

    # §5 Ограждение котлована · STD-ZC-001 — (первый корпус, приращение)
    PIT_WALL = {
        "трубошпунт": (30, 10, 15),
        "свг": (50, 15, -45),
        "бсс": (60, 20, -45),
    }
    PIT_WALL_SRC = "standards.md §5 STD-ZC-001 · §5.1 (лаг НН к земляным работам)"

    # §6 Земляные работы · STD-ZC-002 — (длительность, сдвиг старта ФП)
    EARTH = {
        "трубошпунт": (45, 15),
        "свг": (50, 20),
        "бсс": (60, 30),
        "отвал": (30, 10),
    }
    EARTH_SRC = "standards.md §6 STD-ZC-002"

    # §7 Свайное основание · STD-ZC-003 / STD-ZC-004 — свай в день на установку
    PILE_RATE = {
        "бнс": (1, "standards.md §7 STD-ZC-003, 1 свая/день на установку"),
        "забивные": (10, "standards.md §7 STD-ZC-004, 10 свай/день на установку"),
    }
    PILE_WARN_DAYS = (90, "standards.md §7 DEC-14, порог предупреждения")

    # §8 Фундаментная плита и монолит ниже 0.000
    RAFT = (45, "standards.md §8 STD-ZC-005, до 45 этажей")
    RAFT_TALL = (60, "standards.md §8 STD-ZC-005, свыше 45 этажей")
    BELOW_ZERO = (45, "standards.md §8 STD-ZC-006, 45 дн на подземный этаж")
    PARKING_RAFT = (45, "standards.md §10 STD-MON-002, ФП паркинга")
    PARKING_LEVEL = (45, "standards.md §10 STD-MON-002, подземный этаж паркинга")

    # §10 Паркинг
    PARKING_MASONRY = (90, "standards.md §10 STD-MON-002, кладка паркинга")
    PARKING_ROOF = (45, "standards.md §10 STD-MON-002, эксплуатируемая кровля паркинга")

    # §11 Кровля
    ROOF_PLAIN = (60, "standards.md §11 STD-KRV-001")
    ROOF_USED = (60, "standards.md §11 STD-KRV-002")
    ROOF_STYLOBATE = (45, "standards.md §11 STD-KRV-003")

    # §13 СПК и фасады
    FACADE_MOCKUP = (60, "standards.md §13 STD-FAS-004, DEC-12 — по умолчанию")

    # §14 ВИС. Значения 🔴 отменяют прежние решения владельца — РАСХОЖДЕНИЯ_шаблон_v2.md
    ITP = (120, "standards.md §14 STD-VIS-001, ЦТП/ИТП — 🔴 R-03, было 200 дн")
    # 🔴 R-01/R-02: было 120 дн и лаг ОН+0. Стало 250 дн и общий лаг +15 дн.
    # Календарной привязки нет: совпадение финишей у корпусов эталона объяснено
    # одновременным выходом обоих на 15-й этаж (R-13).
    HEAT_LOOP = (250, "standards.md §14 STD-VIS-006, отопление — контур для пуска тепла — 🔴 R-01")
    ELEVATORS = (150, "standards.md §14 STD-VIS-004, DEC-11")
    COMMISSIONING = (30, "standards.md §14 STD-VIS-005, блок испытаний и передачи УК")
    VIS_PARKING = (200, "standards.md §14 STD-VIS-007, ВИС паркинга")

    # §15 Отделка
    FIT_MIN = (210, "standards.md §15.1, жёсткий минимум D_min")
    TECH_ROOMS = (90, "standards.md §15 STD-OTD-003 — 🔴 R-06, было 60 дн")
    PARKING_LOT = (150, "standards.md §15 STD-OTD-005, автостоянка с рампой")
    STORAGE = (90, "standards.md §15 STD-OTD-006, кладовые")
    UNDERGROUND_MOP = (120, "standards.md §15 STD-OTD-007, МОП подземного уровня")
    HANDOVER_RAW = (180, "standards.md §15.3, передача без отделки: РВЭ + 180")
    HANDOVER_FIN = (270, "standards.md §15.3, передача с чистовой: РВЭ + 270")
    HANDOVER_COMM = (60, "standards.md §15.3, коммерческие и машиноместа: РВЭ + 60")
    DDU_FLAT = (150, "standards.md §15.3, срок по ДДУ, квартиры: передача + 150")
    DDU_COMM = (60, "standards.md §15.3, срок по ДДУ, коммерческие: передача + 60")

    # §16 ВТК
    TEMP_CONTOUR = (45, "standards.md §16 STD-TC-001, DEC-05")

    # §17 Благоустройство. 🔴 R-07: резервной строки STD-BIO-008 больше нет —
    # блок построен цепочкой «земляные 60 → твёрдые покрытия 120» = ровно 180 дн.
    BIO = {
        "STD-BIO-001": (60, "Земляные работы"),
        "STD-BIO-002": (120, "Устройство твёрдых покрытий"),
        "STD-BIO-003": (60, "Озеленение"),
        "STD-BIO-004": (45, "Установка ограждения"),
        "STD-BIO-005": (45, "Устройство наружного освещения"),
        "STD-BIO-006": (45, "Монтаж МАФ"),
        "STD-BIO-007": (45, "Организация дорожного движения"),
    }
    BIO_BLOCK = (180, "standards.md §17, длительность блока БИО")
    BIO_TAIL_LAG = (45, "standards.md §17, ОН −45 дн от твёрдых покрытий")


class Bnd:
    """Константы привязки — bindings.md §1.2."""

    MASONRY_FLOOR = "min(6, N)"      # BND-KLD-001, DEC-04
    VIS_FLOOR = "min(15, N)"          # BND-VIS-001, BND-SPK-002
    LAG_ANCHOR_PVC = (30, "bindings.md §1.2, ЯКОРЬ_ОГР НН +30 → СПК")
    LAG_PVC_FACADE = (30, "bindings.md §1.2, старт СПК НН +30 → старт фасада")
    # 🔴 R-09/R-14: точка отсчёта — финиш кладки паркинга (КЛ_ПАРК), связь ОН,
    # а не НН от ЯКОРЬ_ОГР. При отсутствии паркинга правило откатывается на ЯКОРЬ_ОГР.
    LAG_FIT = (90, "bindings.md §1.2 BND-OTD-002, КЛ_ПАРК ОН +90 → старт отделки — 🔴 R-09")
    # 🔴 R-02: лаг +15 дн действует для ВСЕХ трёх членов BND-TC-003, включая отопление.
    LAG_HEAT = (15, "bindings.md §3.6 BND-TC-003, +15 дн у всех членов — 🔴 R-02")
    # 🔴 R-04: было НН +90 от старта кладки перегородок корпуса.
    LAG_ITP = (20, "bindings.md §3.7 BND-VIS-006, старт отделки техпомещений НН +20 → ИТП — 🔴 R-04")
    LAG_TECH_ROOMS = (21, "bindings.md BND-VIS-004, старт кладки паркинга НН +21")
    LAG_BIO = (180, "bindings.md §3.9 BND-BIO-001, DEC-22 — ОН −180 от фасадов")
    LAG_BTI = (160, "bindings.md §3.10 BND-ZOS-001, ОН −160 от финиша отделки (DEC-29)")
    FIT_TARGET = (365, "bindings.md §1.2 / §3.8 BND-OTD-003, целевой финиш ЯКОРЬ +365")

    # Финиши, привязанные к ЯКОРЬ своего корпуса
    FIN_MASONRY_EXT = (45, "standards.md §12 STD-KLD-001, финиш ЯКОРЬ +45")
    # 🔴 R-16: факт шаблона +125 у обоих корпусов, прежнее значение было +120
    FIN_MASONRY_INT = (125, "standards.md §12 STD-KLD-002, финиш ЯКОРЬ +125 — 🔴 R-16")
    FIN_PVC = (75, "bindings.md §3.4 BND-SPK-001, финиш ЯКОРЬ +75 при наличии кладки")
    FIN_PVC_NO_MASONRY = (129, "bindings.md §3.4 BND-SPK-001, финиш ЯКОРЬ +129 без кладки")
    # 🔴 R-17: факт шаблона +129, прежнее значение было +120
    FIN_STAINED = (129, "bindings.md §3.4 BND-SPK-002, финиш витражей ЯКОРЬ +129 — 🔴 R-17")
    FIN_FACADE = (330, "bindings.md §3.5 BND-FAS-001/002, финиш ЯКОРЬ +330")
    FIN_FACADE_MODULAR = (210, "bindings.md §3.5 BND-FAS-003, финиш ЯКОРЬ +210")
    # DEC-23: единое правило для всего блока ВИС, различие 365/330 упразднено
    FIN_VIS = (330, "bindings.md §3.7 BND-VIS-001, финиш ЯКОРЬ +330 (DEC-23)")
    FIN_EOM = (330, "bindings.md §3.7 BND-VIS-002, финиш ЯКОРЬ +330 (DEC-23)")

    # Контрольные календарные даты теплового контура — bindings.md §3.5
    DEADLINE_PERM = (10, 15)   # 15 октября — крайний срок замыкания постоянного контура
    DEADLINE_HEAT = (10, 30)   # 30 октября — крайний прогнозный срок пуска тепла
    SPRING_WINDOW = (4, 1)     # 01 апреля — весеннее окно
    TEMP_FINISH = (9, 30)      # 30 сентября — финиш ВТК


# ======================================================================
# 2. Сезонность — bindings.md §5, DEC-20
# ======================================================================

SEASON_WINDOWS = {
    # код: (месяц_с, день_с, месяц_по, день_по|None → последний день февраля)
    "SEA-01": (12, 1, 2, None),
    "SEA-02": (12, 1, 2, None),
    "SEA-03": (12, 1, 3, 31),
    "SEA-04": (12, 1, 2, None),
    "SEA-05": (12, 1, 2, None),
}


def in_season_window(day: date, code: str) -> bool:
    """Попадает ли календарный день в сезонное окно."""
    m_from, d_from, m_to, d_to = SEASON_WINDOWS[code]
    if d_to is None:  # «последний день февраля» — bindings.md §5
        d_to = 29 if (day.year % 4 == 0 and (day.year % 100 != 0 or day.year % 400 == 0)) else 28
    after_start = (day.month, day.day) >= (m_from, d_from)
    before_end = (day.month, day.day) <= (m_to, d_to)
    # окно переходит через новый год
    return after_start or before_end


def seasonal_duration(start: date, base_days: int, code: str, coef: float) -> int:
    """Длительность с сезонным коэффициентом — DEC-20, без итераций.

    Задача набирает объём, равный базовой длительности. День вне окна даёт
    1 единицу объёма, день внутри окна — 1/coef. Финиш — когда объём набран.
    Умножается только попавшая в окно часть, а не вся длительность.
    """
    if base_days <= 0:
        return 0
    volume = 0.0
    days = 0
    cur = start
    while volume < base_days:
        volume += (1.0 / coef) if in_season_window(cur, code) else 1.0
        days += 1
        cur += timedelta(days=1)
        if days > base_days * coef + 400:  # предохранитель
            break
    return days


# ======================================================================
# 3. Расчёт корпуса — bindings.md §2.2, тринадцать шагов
# ======================================================================

@dataclass
class Assumption:
    code: str
    text: str
    scope: str = ""


@dataclass
class CorpusResult:
    """Результат теплового блока и отделки по одному корпусу."""
    anchor: date                  # ЯКОРЬ — последний куб бетона
    anchor_ogr: date              # ЯКОРЬ_ОГР — старт кладки наружных / витражей
    tc_perm: date
    tc_eff: date
    temp_contour: tuple[date, date] | None    # (старт, финиш) ВТК либо None
    itp_finish: date | None
    heat: date                    # пуск тепла
    fit_start: date
    fit_finish: date
    fit_duration: int
    overrun_days: int             # превышение над ЯКОРЬ + 365
    assumptions: list[Assumption] = field(default_factory=list)
    trace: list[str] = field(default_factory=list)


def monolith_floor_durations(n_floors: int, complex_frame: bool) -> tuple[int, list[int], int]:
    """STD-MON-001: (первый этаж, типовые, кровля/парапет)."""
    k = Std.MON_TYPICAL_HARD[0] if (complex_frame or n_floors > 60) else Std.MON_TYPICAL[0]
    typical = [k] * max(0, n_floors - 1)
    return Std.MON_FIRST[0], typical, Std.MON_ROOF[0]


def pile_type_durations(piles: list[dict]) -> tuple[dict[str, int], int, list[str]]:
    """STD-ZC-003/004: длительность каждого типа и параллельного блока."""
    trace: list[str] = []
    durations: dict[str, int] = {}
    for p in piles:
        kind = p["тип"].lower()
        rate, src = Std.PILE_RATE[kind]
        rigs = max(1, int(p.get("установок", 1)))
        d = math.ceil(p["количество"] / (rate * rigs))
        durations[kind] = d
        trace.append(f"{p['тип']}: {p['количество']} шт / ({rate} × {rigs} уст.) = {d} дн — {src}")
    if not durations:
        return {}, 0, trace
    total = max(durations.values())  # подзадачи разных типов идут параллельно
    trace.append(f"типы свай ведутся параллельно → длительность блока = max = {total} дн")
    return durations, total, trace


def pile_duration(piles: list[dict]) -> tuple[int, list[str]]:
    """Обратимо совместимый интерфейс: длительность параллельного блока и трасса."""
    _, total, trace = pile_type_durations(piles)
    return total, trace


def compute_thermal_and_fit(
    anchor: date,
    anchor_ogr: date,
    facade_type: str,
    pvc_finish: date | None,
    stained_finish: date | None,
    facade_finish: date | None,
    fit_base: date,
    fit_base_label: str,
    itp_finish: date | None,
    heat_loop_finish: date | None = None,
    masonry_finish: date | None = None,
) -> CorpusResult:
    """Шаги 7–11 алгоритма bindings.md §2.2. Порядок шагов существен.

    Отличия от версии под шаблон v1 (все — РАСХОЖДЕНИЯ_шаблон_v2.md):
      · старт отделки отсчитывается от `fit_base` — старта кладки
        перегородок своего корпуса (НН +90), а не от чужого объекта    DEC-28
      · ветка гейта при опоздании тепла берёт РАНЕЕ событие из
        «пуск тепла» и «01 апреля года Y+1», а не жёстко апрель         R-42
      · в пуск тепла все три члена входят с лагом +15 дн                R-02
      · ИТП считается снаружи (он общий, в паркинге) и подаётся готовым R-04
    """
    a: list[Assumption] = []
    tr: list[str] = []

    # --- шаг 7. TC_perm по типу фасада (BND-TC-001) ----------------------
    if facade_type.lower() in ("модульный", "панельный"):
        if facade_finish is None:
            raise ValueError("Модульный фасад: не задан финиш монтажа фасада (BND-TC-001)")
        tc_perm = facade_finish
        tr.append(f"TC_perm = финиш модульного фасада = {dfmt(tc_perm)} (BND-TC-001)")
        a.append(Assumption("typGRP.md §14 п.1",
                            "Модульный фасад: типового блока задач в шаблоне нет, "
                            "структура принята по аналогии с НВФ."))
    else:
        candidates = [d for d in (pvc_finish, stained_finish) if d is not None]
        if not candidates:
            raise ValueError("Не задан ни один финиш остекления (BND-TC-001)")
        tc_perm = max(candidates)
        tr.append(f"TC_perm = max(финиш СПК ПВХ, финиш витражей) = {dfmt(tc_perm)} (BND-TC-001)")

    # --- шаг 10 (предварительный S0) — разрыв круговой зависимости --------
    s0 = fit_base + timedelta(days=Bnd.LAG_FIT[0])
    year = s0.year
    tr.append(f"S0 = {fit_base_label} ({dfmt(fit_base)}) + {Bnd.LAG_FIT[0]} = {dfmt(s0)}; "
              f"год отделки Y = {year} (bindings.md §2.2, разрыв круговой зависимости — "
              f"берётся до гейта)")

    if itp_finish is not None:
        tr.append(f"ИТП: финиш {dfmt(itp_finish)} ({Std.ITP[1]})")

    def heat_from(tc_effective: date) -> date:
        """BND-TC-003: max(TC_eff, ИТП, отопление-контур) + 15 дн.

        R-02: лаг +15 применяется ко ВСЕМ трём членам. Прежнее правило DEC-07
        вводило отопление связью ОН+0 — шаблон v2 его отменил.
        """
        terms = [tc_effective]
        if itp_finish is not None:
            terms.append(itp_finish)
        if heat_loop_finish is not None:
            terms.append(heat_loop_finish)
        return max(terms) + timedelta(days=Bnd.LAG_HEAT[0])

    def gate(heat_date: date) -> date:
        """BND-OTD-002, сезонный гейт SEA-03 — формула комментария шаблона v2.

        «при условии прогноза пуска тепла до 30 октября того же года, иначе
        после пуска тепла или 01 апреля следующего года — ранее событие».
        """
        if heat_date > date(year, *Bnd.DEADLINE_HEAT):
            return min(heat_date, date(year + 1, *Bnd.SPRING_WINDOW))
        return s0

    if heat_loop_finish is None:
        a.append(Assumption("R-01", "Норматив «Отопление — контур для пуска тепла» "
                                    "(STD-VIS-006) не задан: член исключён из формулы "
                                    "пуска тепла."))
    else:
        tr.append(f"Отопление — контур для пуска тепла: финиш {dfmt(heat_loop_finish)}, "
                  f"связь ОН +{Bnd.LAG_HEAT[0]} дн с пуском тепла ({Std.HEAT_LOOP[1]})")

    # --- шаги 8–9 без ВТК -------------------------------------------------
    temp_contour = None
    tc_eff = tc_perm
    heat = heat_from(tc_eff)
    fit_start = gate(heat)
    target_finish = anchor + timedelta(days=Bnd.FIT_TARGET[0])
    window = (target_finish - fit_start).days
    tr.append(f"без ВТК: TC_eff = {dfmt(tc_eff)}, пуск тепла = {dfmt(heat)}, "
              f"старт отделки = {dfmt(fit_start)}, окно до цели = {window} дн")

    # --- шаг 11. BND-OTD-003 ---------------------------------------------
    d_min = Std.FIT_MIN[0]
    need_temp = tc_perm > date(year, *Bnd.DEADLINE_PERM) or window < d_min
    if need_temp:
        why = ("TC_perm позже 15 октября" if tc_perm > date(year, *Bnd.DEADLINE_PERM)
               else f"окно отделки {window} дн < {d_min} дн")
        # R-43: финиш ВТК — по финишу кладки наружных стен (шаблон v2 ставит связь
        # ОО+0). Прежнее правило DEC-05 давало календарное 30 сентября, но замкнуть
        # временный контур раньше, чем подняты стены, физически нельзя: у К2 эталона
        # кладка финиширует 20.11, а календарное правило дало бы 30.09.
        # При отсутствии наружной кладки правило откатывается на 30 сентября.
        tf = masonry_finish if masonry_finish is not None else date(year, *Bnd.TEMP_FINISH)
        ts = tf - timedelta(days=Std.TEMP_CONTOUR[0])
        tr.append(f"финиш ВТК = "
                  + ("финиш кладки наружных стен" if masonry_finish is not None
                     else "30 сентября года отделки")
                  + f" = {dfmt(tf)} (R-43)")
        cand_eff = min(tc_perm, tf)
        cand_heat = heat_from(cand_eff)
        cand_start = gate(cand_heat)
        cand_window = (target_finish - cand_start).days
        tr.append(f"условие ВТК выполнено ({why}); ВТК {dfmt(ts)} → {dfmt(tf)}, "
                  f"TC_eff = min(TC_perm, ВТК) = {dfmt(cand_eff)}")

        # DEC-33: если условие ВТК выполнено, он остаётся физической частью
        # теплового контура независимо от того, меняет ли ведущего
        # предшественника пуска тепла. Иначе веха Heat теряет реальную связь с
        # действующим контуром и искусственно привязывается к календарной дате.
        has_schedule_effect = cand_heat < heat or cand_window > window
        temp_contour = (ts, tf)
        tc_eff, heat, fit_start, window = cand_eff, cand_heat, cand_start, cand_window
        if has_schedule_effect:
            tr.append(f"ВТК даёт эффект: пуск тепла {dfmt(heat)}, старт отделки {dfmt(fit_start)}, "
                      f"окно {window} дн (BND-TC-002)")
        else:
            tr.append("ВТК не меняет итоговую дату, но сохраняется как действующий контур "
                      "и прямой предшественник пуска тепла (DEC-33)")

    # --- финиш отделки ----------------------------------------------------
    overrun = 0
    if window >= d_min:
        fit_duration = window
        fit_finish = fit_start + timedelta(days=fit_duration)
        tr.append(f"окно {window} дн ≥ {d_min} дн — цель ЯКОРЬ +365 достигнута")
    else:
        fit_duration = d_min
        fit_finish = fit_start + timedelta(days=d_min)
        overrun = (fit_finish - target_finish).days
        tr.append(f"окно {window} дн < {d_min} дн — принят жёсткий минимум {d_min} дн, "
                  f"финиш {dfmt(fit_finish)}, превышение над целью {overrun} дн")
        a.append(Assumption("DEC-02/DEC-09",
                            f"Окно отделки после ВТК составило {window} дн при жёстком минимуме "
                            f"{d_min} дн. Принята длительность {d_min} дн; финиш отделки выходит "
                            f"за ЯКОРЬ + 365 на {overrun} дн. РВЭ и передача сдвинуты на ту же "
                            f"величину."))

    # --- остаточный случай SEA-03 (DEC-15) --------------------------------
    # Гейт переносит старт только когда тепло опаздывает за 30 октября. Если
    # S0 сам лежит в зимнем окне при своевременном тепле, мокрые процессы всё
    # равно начнутся до пуска — этот отрезок гейт не покрывает.
    winter_days = sum(
        1 for i in range((heat - fit_start).days)
        if in_season_window(fit_start + timedelta(days=i), "SEA-03")
    ) if fit_start < heat else 0
    if winter_days:
        a.append(Assumption("DEC-15",
                            f"{winter_days} дн отделки приходится на окно SEA-03 до пуска тепла "
                            f"({dfmt(heat)}). Гейт этот отрезок не переносит — расчёт продолжен."))

    return CorpusResult(
        anchor=anchor, anchor_ogr=anchor_ogr, tc_perm=tc_perm, tc_eff=tc_eff,
        temp_contour=temp_contour, itp_finish=itp_finish, heat=heat,
        fit_start=fit_start, fit_finish=fit_finish, fit_duration=fit_duration,
        overrun_days=overrun, assumptions=a, trace=tr,
    )


# ======================================================================
# 4. Сетевой расчёт (CPM) — русская нотация связей typGRP.md §2.1
# ======================================================================

LINK_RE = re.compile(r"^\s*(\d+)\s*(ОН|НН|ОО|НО)?\s*([+-]\s*\d+)?\s*(?:дн|день|дня|дней)?\.?\s*$")


def parse_links(field_value: str) -> list[tuple[int, str, int]]:
    """Разбор поля «Предшественники» → [(Ид., тип, лаг), ...]."""
    out = []
    if not field_value:
        return out
    for part in field_value.split(";"):
        part = part.strip()
        if not part:
            continue
        m = LINK_RE.match(part)
        if not m:
            continue
        pid = int(m.group(1))
        kind = m.group(2) or "ОН"          # без кода = окончание→начало
        lag = int(m.group(3).replace(" ", "")) if m.group(3) else 0
        out.append((pid, kind, lag))
    return out


def fmt_links(links: list[tuple[int, str, int]]) -> str:
    parts = []
    for pid, kind, lag in links:
        s = str(pid)
        if kind != "ОН" or lag:
            s += kind
        if lag:
            s += f"{lag:+d} дней"
        parts.append(s)
    return "; ".join(parts)


@dataclass
class Node:
    # Ключ строки. В шаблоне v2 колонки СДР нет, поэтому ключом служит «Ид.»
    # шаблона (строкой), а у вставленных генератором строк — синтетический «n<N>».
    key: str
    duration: int
    links: list[tuple[str, str, int]] = field(default_factory=list)   # (ключ, тип, лаг)
    anchor_start: date | None = None
    # Суммарная строка WBS: собственных связей и длительности не несёт, её даты
    # сворачиваются из потомков — старт по минимуму, финиш по максимуму
    # (typGRP.md §2.2 п.4). Нужна в сети, потому что связи эталона ссылаются
    # на суммарные строки, и ссылка обязана сохраниться как есть.
    rollup: list[str] = field(default_factory=list)
    start: date | None = None
    finish: date | None = None
    late_start: date | None = None
    late_finish: date | None = None
    critical: bool = False


def topo_order(nodes: dict[str, Node]) -> list[str]:
    """Топологическая сортировка. При цикле — исключение с составом цикла."""
    state: dict[str, int] = {}
    order: list[str] = []

    def visit(k: str, path: list[str]) -> None:
        st = state.get(k, 0)
        if st == 1:
            cycle = path[path.index(k):] + [k]
            raise ValueError("Цикл в связях: " + " → ".join(cycle))
        if st == 2:
            return
        state[k] = 1
        for pk, _, _ in nodes[k].links:
            if pk in nodes:
                visit(pk, path + [k])
        for ck in nodes[k].rollup:          # потомки считаются раньше родителя
            if ck in nodes:
                visit(ck, path + [k])
        state[k] = 2
        order.append(k)

    for k in nodes:
        visit(k, [])
    return order


def forward_pass(nodes: dict[str, Node], project_start: date) -> None:
    """Ранние даты. Финиш достигается длительностью, не ограничением."""
    for k in topo_order(nodes):
        n = nodes[k]

        if n.rollup:
            kids = [nodes[c] for c in n.rollup if c in nodes and nodes[c].start]
            if kids:
                n.start = min(c.start for c in kids)
                n.finish = max(c.finish for c in kids)
            else:
                n.start = n.finish = project_start
            continue

        candidates: list[date] = []
        for pk, kind, lag in n.links:
            p = nodes.get(pk)
            if p is None or p.start is None:
                continue
            if kind == "ОН":
                candidates.append(p.finish + timedelta(days=lag))
            elif kind == "НН":
                candidates.append(p.start + timedelta(days=lag))
            elif kind == "ОО":
                candidates.append(p.finish + timedelta(days=lag) - timedelta(days=n.duration))
            elif kind == "НО":
                candidates.append(p.start + timedelta(days=lag) - timedelta(days=n.duration))
        if n.anchor_start is not None:
            candidates.append(n.anchor_start)
        if not candidates:
            candidates.append(project_start)
        n.start = max(candidates)
        n.finish = n.start + timedelta(days=n.duration)


def backward_pass(nodes: dict[str, Node]) -> None:
    """Поздние даты и критический путь (полный резерв = 0)."""
    order = topo_order(nodes)
    project_finish = max(n.finish for n in nodes.values() if n.finish)
    successors: dict[str, list[tuple[str, str, int]]] = {k: [] for k in nodes}
    for k, n in nodes.items():
        for pk, kind, lag in n.links:
            if pk in nodes:
                successors[pk].append((k, kind, lag))
        # Потомок не может финишировать позже своей суммарной строки: связь ОО+0
        # от потомка к родителю переносит на него поздний финиш родителя.
        for ck in n.rollup:
            if ck in nodes:
                successors[ck].append((k, "ОО", 0))

    for k in reversed(order):
        n = nodes[k]
        limits: list[date] = []
        for sk, kind, lag in successors[k]:
            s = nodes[sk]
            if s.late_start is None:
                continue
            if kind == "ОН":
                limits.append(s.late_start - timedelta(days=lag))
            elif kind == "НН":
                limits.append(s.late_start - timedelta(days=lag) + timedelta(days=n.duration))
            elif kind == "ОО":
                limits.append(s.late_finish - timedelta(days=lag))
            elif kind == "НО":
                limits.append(s.late_finish - timedelta(days=lag) + timedelta(days=n.duration))
        n.late_finish = min(limits) if limits else project_finish
        n.late_start = n.late_finish - timedelta(days=n.duration)
        n.critical = (n.late_finish - n.finish).days == 0

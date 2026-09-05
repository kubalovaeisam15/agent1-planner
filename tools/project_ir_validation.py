"""Независимый предметный контроль ТЭП → IR (состав СПК и TC_perm).

Входной spec должен предварительно пройти validate_project_spec. Этот контроль
дополняет, но не заменяет структурную и календарную проверку Schedule IR.
"""
from datetime import timedelta
from schedule_ir import IRIssue, ScheduleProject


def _check_milestone(ir, name, sources, lag, code, rule):
    """sources: точные имена и ожидаемые типы физических предшественников."""
    issues = []
    expected = []
    for source_name, task_type in sources:
        matches = [t for t in ir.tasks if t.name == source_name]
        if len(matches) != 1 or matches[0].task_type != task_type:
            issues.append(IRIssue(code+"-SCOPE", f"{rule}: требуется одна {task_type} «{source_name}»"))
        else:
            expected.append(matches[0])
    milestones = [t for t in ir.tasks if t.name == name]
    if len(milestones) != 1:
        return issues + [IRIssue(code+"-SCOPE", f"{rule}: требуется одна веха «{name}»")]
    milestone = milestones[0]
    if milestone.task_type != "milestone" or milestone.duration_days != 0:
        issues.append(IRIssue(code+"-MILESTONE", f"{rule}: «{name}» должна быть вехой 0 дней", milestone.task_id))
    actual = sorted((p.predecessor_id, p.type, p.lag_days) for p in milestone.predecessors)
    required = sorted((t.task_id, "FS", lag) for t in expected)
    if len(expected) != len(sources) or not required or actual != required:
        issues.append(IRIssue(code+"-PREDECESSORS", f"{rule}: неверный полный набор предшественников «{name}», ожидаются ОН {lag:+d}", milestone.task_id))
    if expected and len(expected) == len(sources) and all(t.finish is not None for t in expected):
        finish = max(t.finish for t in expected)+timedelta(days=lag)
        if milestone.start != finish or milestone.finish != finish:
            issues.append(IRIssue(code+"-DATE", f"{rule}: дата «{name}» должна быть {finish}", milestone.task_id))
    return issues


def validate_project_against_ir(spec: dict, ir: ScheduleProject) -> list[IRIssue]:
    issues = []
    finishing = [(f"К{n}. Отделочные работы", "summary") for n in range(1,len(spec["корпуса"])+1)]
    issues.extend(_check_milestone(ir, "Готовность к обмерам БТИ", finishing,
                                   -160, "PROJECT-BTI", "BND-ZOS-001/DEC-26/DEC-29"))
    for n, corpus in enumerate(spec["корпуса"], 1):
        prefix = f"К{n}. "
        rows = [t for t in ir.tasks if t.name.startswith(prefix)]
        temporary = prefix+"Закрыт ВРЕМЕННЫЙ тепловой контур по корпусу (при необходимости)"
        has_temporary = any(t.name == temporary for t in rows)
        effective = temporary if has_temporary else prefix+"Закрыт тепловой контур по корпусу"
        itp_prefix = "П1. " if spec.get("паркинг", {}).get("есть", False) else prefix
        heat_sources = [(effective, "task" if has_temporary else "milestone"),
                        (itp_prefix+"По договору ЦТП, ИТП с УУТЭ", "task"),
                        (prefix+"По договору Отопление (контур для пуска тепла)", "task")]
        issues.extend(_check_milestone(ir, prefix+"Пуск тепла корпус", heat_sources,
                                       15, "PROJECT-HEAT", "BND-TC-003/DEC-33"))
        glazing = corpus["остекление"]
        expected = []
        for field, suffix in (("пвх", "ПВХ"), ("витражи", "Витраж")):
            name = prefix + "По договору Монтаж светопрозрачных конструкций " + suffix
            found = [t for t in rows if t.name.startswith(name)]
            if len(found) != int(glazing[field]) or any(t.task_type != "task" for t in found):
                issues.append(IRIssue("PROJECT-GLAZING-SCOPE",
                    f"BND-SPK-003: {corpus['код']}: «{suffix}», требуется {int(glazing[field])} работа, найдено {len(found)}"))
            if glazing[field]:
                expected.extend(found)
        masonry = [t for t in rows if t.name == prefix + "По договору Кладка наружных стен"]
        if glazing["витражи_на_всю_высоту"] and masonry:
            issues.append(IRIssue("PROJECT-MASONRY-SCOPE",
                f"BND-SPK-003: {corpus['код']}: наружная кладка запрещена при витражах на всю высоту", masonry[0].task_id))
        if spec["фасад"]["тип"] == "модульный":
            expected = [t for t in rows if t.name == prefix + "По договору Монтаж фасадов"]
            if len(expected) != 1 or any(t.task_type != "task" for t in expected):
                issues.append(IRIssue("PROJECT-FACADE-SCOPE",
                    f"BND-TC-001: {corpus['код']}: требуется одна работа монтажа модульного фасада"))
        contours = [t for t in rows if t.name == prefix + "Закрыт тепловой контур по корпусу"]
        if len(contours) != 1:
            issues.append(IRIssue("PROJECT-CONTOUR-SCOPE",
                f"BND-TC-001: {corpus['код']}: требуется одна веха постоянного контура"))
            continue
        contour = contours[0]
        actual = [(p.predecessor_id, p.type, p.lag_days) for p in contour.predecessors]
        required = [(t.task_id, "FS", 0) for t in expected]
        if not required or sorted(actual) != sorted(required):
            issues.append(IRIssue("PROJECT-CONTOUR-PREDECESSORS",
                f"BND-TC-001/DEC-33: {corpus['код']}: нужны ОН +0 от всех предусмотренных физических работ своего корпуса", contour.task_id))
        if contour.task_type != "milestone" or contour.duration_days != 0:
            issues.append(IRIssue("PROJECT-CONTOUR-MILESTONE",
                f"BND-TC-001: {corpus['код']}: постоянный контур должен быть вехой 0 дней", contour.task_id))
        if expected and all(t.finish is not None for t in expected):
            finish = max(t.finish for t in expected)
            if contour.start != finish or contour.finish != finish:
                issues.append(IRIssue("PROJECT-CONTOUR-DATE",
                    f"BND-TC-001: {corpus['код']}: дата постоянного контура должна равняться {finish}", contour.task_id))
    return issues

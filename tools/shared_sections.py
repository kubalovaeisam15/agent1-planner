"""DEC-40: заключительные разделы существуют один раз на весь проект."""
from collections import Counter
from collections.abc import Iterable
import re


PROJECT_SHARED_SECTIONS = (
    "ЗОС и РВЭ",
    "Вынос сетей из пятна застройки",
    "Наружные инженерные сети ТУ/ТП",
    "Наружные инженерные сети ПИР",
    "Наружные инженерные сети СМР/ПНР",
    "Наружные инженерные сети ввод/передача",
    "Передача собственникам",
)


def shared_section_errors(
    rows: Iterable[tuple[str, int]], *, require_all: bool = False,
) -> list[str]:
    """Проверять разделы, не одноимённые работы внутри разных корпусов."""
    if type(require_all) is not bool:
        raise ValueError("require_all должен быть bool: полный ГРП или явный фрагмент")
    counts: Counter[str] = Counter()
    errors = []
    for name, level in rows:
        name = " ".join(name.split())
        bare_name = re.sub(r"^К\d+\.\s*", "", name)
        if bare_name != name and bare_name in PROJECT_SHARED_SECTIONS:
            errors.append(f"DEC-40: «{name}» нельзя выделять в раздел отдельного корпуса")
            name = bare_name
        if name not in PROJECT_SHARED_SECTIONS:
            continue
        counts[name] += 1
        if level != 1:
            errors.append(f"DEC-40: «{name}» должен быть общепроектным разделом уровня 1")
    for name in PROJECT_SHARED_SECTIONS:
        if counts[name] > 1:
            errors.append(f"DEC-40: «{name}» повторён {counts[name]} раз; допустим один раздел на проект")
        elif require_all and not counts[name]:
            errors.append(f"DEC-40: отсутствует общий раздел «{name}»")
    return errors

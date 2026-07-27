"""Объективные проверки для 20 правок в каждом корпусе. Проверка получает
(before_texts, after_texts, before_styles, after_styles) — списки текстов/
стилей блоков документа до и после правки; возвращает True/False/None
(None = нельзя проверить механически, score.py считает это n/a)."""

import re


def _j(texts):
    return "\n".join(texts)


def _heading_appeared(before_styles, after_styles):
    heading = lambda s: s.startswith("Heading") or s.startswith("Title")
    return sum(1 for s in after_styles if heading(s)) > sum(1 for s in before_styles if heading(s))


# --- colbert ---

def c1(before, after, sb, sa):
    b, a = _j(before), _j(after)
    return all(t in a and t not in b for t in ("[CLS]", "[SEP]", "[MASK]"))


def c2(before, after, sb, sa):
    a = _j(after)
    if re.search(r"[^\S\n]+[:)]", a):
        return False
    return (len(_j(before)) - len(a)) <= 30


def c3(before, after, sb, sa):
    a = _j(after)
    return not any(t in a for t in ("RAG.Усредненный", "процесс.Исследователи", "напрямую.В основе"))


def c4(before, after, sb, sa):
    a = _j(after)
    good = all(t in a for t in ("4,2%", "6,7%", "36,2%"))
    bad = any(t in a for t in ("4.2%", "6.7%", "36.2%"))
    return good and not bad


def c5(before, after, sb, sa):
    a = _j(after)
    return "MS-MARCO" not in a and "MS MARCO" in a


def c6(before, after, sb, sa):
    return '"' not in _j(after) and '"' in _j(before)


def c7(before, after, sb, sa):
    a = _j(after)
    matches = re.findall(r"(?i)(single[- ]vector|bi[- ]encoder)s?", a)
    return len(set(matches)) == 1


def c9(before, after, sb, sa):
    return _heading_appeared(sb, sa)


def c10(before, after, sb, sa):
    a = _j(after)
    return not any(t in a for t in (
        "1. Мультимодальный RAG", "2. Модернизация агрегации", "3. Дальнейшие инновации",
    ))


def c19(before, after, sb, sa):
    return sa.count("[table]") > sb.count("[table]")


# --- math ---

def _violates_spacing(texts):
    return any("  " in t or t.endswith(" ") for t in texts)


def m1(before, after, sb, sa):
    return _violates_spacing(before) and not _violates_spacing(after)


def m2(before, after, sb, sa):
    a = _j(after)
    return "серьёзное" in a and "серьезное" not in a


def m3(before, after, sb, sa):
    a = _j(after)
    return "1. Переход от физики к биологии" not in a and "Переход от физики к биологии" in a


def m4(before, after, sb, sa):
    return _heading_appeared(sb, sa)


def m6(before, after, sb, sa):
    return "байт), и так далее" not in _j(after)


def m9(before, after, sb, sa):
    target = "Сначала о первом"
    return before.count(target) == 1 and after.count(target) == 0


def m10(before, after, sb, sa):
    return "Excel" not in _j(after)


def m11(before, after, sb, sa):
    return "математическим универсализмом" not in _j(after)


def m19(before, after, sb, sa):
    return any("Footnote" in s for s in sa)


CHECKS = {
    "colbert": {
        1: c1, 2: c2, 3: c3, 4: c4, 5: c5, 6: c6, 7: c7, 8: None, 9: c9, 10: c10,
        11: None, 12: None, 13: None, 14: None, 15: None, 16: None, 17: None, 18: None,
        19: c19, 20: None,
    },
    "math": {
        1: m1, 2: m2, 3: m3, 4: m4, 5: None, 6: m6, 7: None, 8: None, 9: m9, 10: m10,
        11: m11, 12: None, 13: None, 14: None, 15: None, 16: None, 17: None, 18: None,
        19: m19, 20: None,
    },
}

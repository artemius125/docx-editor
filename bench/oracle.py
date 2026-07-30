"""Объективные проверки для 20 правок в каждом корпусе. Проверка получает
(before_texts, after_texts, before_styles, after_styles) — списки текстов/
стилей блоков документа до и после правки; возвращает True/False/None
(None = нельзя проверить механически, score.py считает это n/a).

В8: у некоторых правок след не текстовый (настоящая сноска Word живёт в
пакете .docx, а не в тексте/стиле абзаца) — такие проверки принимают ещё
и docx_path (путь к сохранённому .docx этой правки или None, если он
недоступен/ненадёжен для этой правки — тогда честный n/a). Параметр
добавлен всем проверкам единообразно (даже тем, что его не используют),
чтобы вызывающий код (score.py) не разбирал сигнатуру каждой функции."""

import re
import zipfile


def _j(texts):
    return "\n".join(texts)


def _heading_appeared(before_styles, after_styles):
    heading = lambda s: s.startswith("Heading") or s.startswith("Title")
    return sum(1 for s in after_styles if heading(s)) > sum(1 for s in before_styles if heading(s))


# --- colbert ---

def c1(before, after, sb, sa, docx_path=None):
    b, a = _j(before), _j(after)
    return all(t in a and t not in b for t in ("[CLS]", "[SEP]", "[MASK]"))


def c2(before, after, sb, sa, docx_path=None):
    a = _j(after)
    if re.search(r"[^\S\n]+[:)]", a):
        return False
    return (len(_j(before)) - len(a)) <= 30


def c3(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return not any(t in a for t in ("RAG.Усредненный", "процесс.Исследователи", "напрямую.В основе"))


def c4(before, after, sb, sa, docx_path=None):
    a = _j(after)
    good = all(t in a for t in ("4,2%", "6,7%", "36,2%"))
    bad = any(t in a for t in ("4.2%", "6.7%", "36.2%"))
    return good and not bad


def c5(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return "MS-MARCO" not in a and "MS MARCO" in a


def c6(before, after, sb, sa, docx_path=None):
    return '"' not in _j(after) and '"' in _j(before)


def c7(before, after, sb, sa, docx_path=None):
    # Было: альтернация "single[- ]vector|bi[- ]encoder" — два отдельных слова,
    # оба захватывались как разные элементы set() даже когда в тексте ОДНО
    # написание составного канонического термина «Single-Vector Bi-encoder»:
    # set всегда содержал 2 элемента, проверка не могла пройти НИ ПРИ КАКОМ
    # состоянии документа (В8). Первым вариантом альтернации ставим составной
    # термин целиком — если "single-vector" идёт сразу перед "bi-encoder(s)",
    # это ОДНО совпадение; отдельно стоящие "bi-encoder"/"single-vector" (то,
    # что реально осталось бы при неполной унификации) по-прежнему свои матчи.
    a = _j(after)
    matches = re.findall(
        r"(?i)single[- ]vector\s+bi[- ]encoders?|bi[- ]encoders?|single[- ]vector", a
    )
    return len(set(matches)) == 1


def c9(before, after, sb, sa, docx_path=None):
    return _heading_appeared(sb, sa)


def c10(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return not any(t in a for t in (
        "1. Мультимодальный RAG", "2. Модернизация агрегации", "3. Дальнейшие инновации",
    ))


def c19(before, after, sb, sa, docx_path=None):
    return sa.count("[table]") > sb.count("[table]")


# --- math ---

def _violates_spacing(texts):
    return any("  " in t or t.endswith(" ") for t in texts)


def m1(before, after, sb, sa, docx_path=None):
    return _violates_spacing(before) and not _violates_spacing(after)


def m2(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return "серьёзное" in a and "серьезное" not in a


def m3(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return "1. Переход от физики к биологии" not in a and "Переход от физики к биологии" in a


def m4(before, after, sb, sa, docx_path=None):
    return _heading_appeared(sb, sa)


def m6(before, after, sb, sa, docx_path=None):
    return "байт), и так далее" not in _j(after)


def m9(before, after, sb, sa, docx_path=None):
    # Абзац из одной фразы должен перестать быть отдельным блоком и уцелеть
    # внутри соседнего. Проверка была сломана: before.count() — это count по
    # СПИСКУ блоков, то есть точное равенство элементу, а блок в документе —
    # «Сначала — о первом.» с тире и точкой. Возвращала False всегда, из-за
    # чего верная правка считалась провалом и вдобавок ложным вердиктом.
    # Второй раз сломана тем же способом (w19): маркер держал КОНЕЧНУЮ ТОЧКУ,
    # а слияние двух абзацев её законно меняет — модель написала «Сначала — о
    # первом: в этой позиции…», и верная правка снова считалась провалом, да
    # ещё и ложным вердиктом. Знак в конце фразы к сути правки не относится.
    frag = "Сначала — о первом"
    strip_tail = lambda t: t.strip().rstrip(".:;,—- ")
    alone = lambda blocks: any(strip_tail(t) == frag for t in blocks)
    merged = any(frag in t and len(strip_tail(t)) > len(frag) for t in after)
    return alone(before) and not alone(after) and merged


def m10(before, after, sb, sa, docx_path=None):
    return "Excel" not in _j(after)


def m11(before, after, sb, sa, docx_path=None):
    return "математическим универсализмом" not in _j(after)


def m19(before, after, sb, sa, docx_path=None):
    # Было: any("Footnote" in styles_after) — ищет абзац со стилем Footnote в
    # ТЕЛЕ документа. Настоящая сноска Word так не устроена: текст лежит в
    # отдельной части word/footnotes.xml, тело несёт только w:footnoteReference
    # (обычный текстовый прогон run, без своего стиля) — проверка не могла
    # пройти НИ НА ОДНОЙ настоящей сноске (В8). before/after/sb/sa этого не
    # видят по построению, нужен сам .docx — без docx_path (например, seq-
    # прогон, где на диске лежит состояние ПОСЛЕ ВСЕХ 20 правок, а не срез
    # после этой) проверить нечем, это честный n/a, а не всегда-провал.
    if docx_path is None:
        return None
    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        if "word/footnotes.xml" not in names:
            return False
        footnotes_xml = z.read("word/footnotes.xml").decode("utf-8")
        rels = (
            z.read("word/_rels/document.xml.rels").decode("utf-8")
            if "word/_rels/document.xml.rels" in names else ""
        )
    has_text = re.search(r"<w:t[^>]*>[^<]*\S[^<]*</w:t>", footnotes_xml) is not None
    has_rel = "footnotes.xml" in rels
    return has_text and has_rel


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

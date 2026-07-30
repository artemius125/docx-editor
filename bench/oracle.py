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


def c12(before, after, sb, sa, docx_path=None):
    # Обе названные в запросе нестыковки «порядок против кратности» должны
    # исчезнуть. Правка НЕ диктует способ: вилку можно поправить до «в 10 раз»,
    # а можно смягчить формулировку («почти на порядок») — маркер, требующий
    # первого, объявил бы верную правку ложью (замер w24, ровно этот случай).
    a = _j(after)
    unhedged = re.search(r"(?<!почти )(?<!примерно )(?<!около )на порядок \(в 6–10 раз\)", a)
    return unhedged is None and "(в ~170 раз)" not in a


def m5(before, after, sb, sa, docx_path=None):
    # Название фильма было в круглых скобках, второе — в кавычках; после
    # правки оба обязаны быть оформлены одинаково, кавычками.
    a = _j(after)
    return "«Человек, который познал бесконечность»" in a and "(Человек, который познал бесконечность)" not in a


def m14(before, after, sb, sa, docx_path=None):
    # Направление унификации запрос не задаёт («приведи к одному»), поэтому
    # маркер, требующий конкретно «вы», врал бы на верной правке: засчитываем
    # ЛИБО уход третьего лица из названной в запросе фразы, ЛИБО убыль «вы».
    you = lambda t: len(re.findall(r"(?<!\w)вы(?!\w)", t))
    third_gone = "чтобы читатель знал" not in _j(after)
    return third_gone or you(_j(after)) < you(_j(before))


def c8(before, after, sb, sa, docx_path=None):
    # Формат годов должен стать ОДНИМ. Способ правка не диктует, поэтому
    # считаем не конкретное написание, а число разных форм в тексте.
    forms = set(re.findall(r"20\d\d\s*[–\-/]\s*20\d\d\s*(годов|годах|года|гг\.|г\.)?", _j(after)))
    return len(forms) == 1


def c11(before, after, sb, sa, docx_path=None):
    # Русское написание везде, английское — один раз при вводе термина.
    a = _j(after)
    latin = re.findall(r"(?i)cross[-\s]?encoders?", a)
    return len(latin) <= 1 and "росс-энкодер" in a


def c13(before, after, sb, sa, docx_path=None):
    # Двусмысленная фраза обязана перестать читаться двояко: в исходном виде
    # («падение до 86–97%») она остаться не может. Каким словом уточнили —
    # дело автора правки.
    return "падение до 86–97%" not in _j(after)


def c16(before, after, sb, sa, docx_path=None):
    # Обозначения обязаны СОЙТИСЬ, а какое из двух станет общим — правка не
    # диктует: модель законно привела первую формулу ко второй (замер w24).
    # Считаем согласованность: во всех формулах индекс суммы один и тот же и
    # индекс максимума один и тот же.
    a = _j(after)
    sums = set(re.findall(r"\\sum_\{i=1\}\^([mn])", a))
    maxs = set(re.findall(r"\\max_\{1 ?\\le j ?\\le ([mn])\}", a))
    return len(sums) == 1 and len(maxs) == 1 and sums != maxs


def c17(before, after, sb, sa, docx_path=None):
    a = _j(after)
    return "транспонир" in a.lower() and "$\\top$ представляет собой скалярное произведение" not in a


def c18(before, after, sb, sa, docx_path=None):
    # Неточность именно в объяснении «благодаря недифференцируемой природе».
    return "недифференцируемой природе" not in _j(after)


def m7(before, after, sb, sa, docx_path=None):
    # Обрубленное второе предложение («Там, где — нет, хмурились») обязано быть
    # переписано; каким именно оборотом — дело автора.
    return "Там, где — нет, хмурились" not in _j(after)


def m8(before, after, sb, sa, docx_path=None):
    return "сравнимый с маленьким ребёнком" not in _j(after)


def m12(before, after, sb, sa, docx_path=None):
    # Атрибуция смягчена: либо названы пифагорейцы, либо прямое утверждение
    # «Пифагор ... утверждал» ушло.
    a = _j(after)
    return "пифагорейц" in a.lower() or "утверждал, что «всё есть число»" not in a


def m13(before, after, sb, sa, docx_path=None):
    # Два абзаца говорили одно и то же; после правки формула-повтор остаётся
    # максимум в одном блоке.
    return sum(1 for t in after if "мир чисел" in t) <= 1


def m15(before, after, sb, sa, docx_path=None):
    # Уточнить, что за соревнования были в университете: слитное утверждение
    # «в средней школе И В УНИВЕРСИТЕТЕ выигрывал городские олимпиады» обязано
    # уйти. Оба слова могут остаться в одном предложении — если школа и
    # университет разведены (ровно это сделала модель в w24), правка верна.
    return "школе и в университете я ежегодно выигрывал городские олимпиады" not in _j(after)


def m16(before, after, sb, sa, docx_path=None):
    # Раздел сжат примерно вдвое, но обязательные тезисы на месте. Границы
    # раздела — от его заголовка до абзаца, которым глава кончается.
    def section(texts):
        start = next((i for i, t in enumerate(texts) if t.startswith("Немного о формулах")), None)
        if start is None:
            return None
        end = next((i for i, t in enumerate(texts) if i > start and t.startswith("Потому что во второй части")),
                   len(texts))
        return sum(len(t) for t in texts[start:end])
    was, now = section(before), section(after)
    if not was or now is None:
        return None
    a = _j(after).lower()
    kept = "формул" in a and "поясн" in a and ("провер" in a or "расчёт" in a)
    return now <= was * 0.75 and kept


def m17(before, after, sb, sa, docx_path=None):
    # Неверное утверждение — что быстрой проверки простоты не существует.
    # Требуем, чтобы именно оно ушло (первый вариант маркера был вакуумным:
    # слово «тест» в документе есть и без правки).
    return "является ли большое число простым, не существует" not in _j(after)


def m18(before, after, sb, sa, docx_path=None):
    return "Сакс" in _j(after)


def m20(before, after, sb, sa, docx_path=None):
    # Концовка обязана перестать быть одной строкой того же вида.
    return "Потому что во второй части книги я собираюсь показать нечто действительно серьезное" not in _j(after)


CHECKS = {
    "colbert": {
        1: c1, 2: c2, 3: c3, 4: c4, 5: c5, 6: c6, 7: c7, 8: c8, 9: c9, 10: c10,
        11: c11, 12: c12, 13: c13, 14: None, 15: None, 16: c16, 17: c17, 18: c18,
        19: c19, 20: None,
    },
    "math": {
        1: m1, 2: m2, 3: m3, 4: m4, 5: m5, 6: m6, 7: m7, 8: m8, 9: m9, 10: m10,
        11: m11, 12: m12, 13: m13, 14: m14, 15: m15, 16: m16, 17: m17, 18: m18,
        19: m19, 20: m20,
    },
}

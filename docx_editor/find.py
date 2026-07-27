"""Детерминированный поиск по блокам документа: подстрока, regex, первое
упоминание, компактное оглавление для Навигатора и фрагмент с соседями.

Индекс заранее не строим — документ 35 тыс. знаков, линейного прохода
достаточно (см. BUILD_PLAN: эмбеддинги и BM25 сознательно не нужны).
"""

import re

_PREFIX = 20
# Порог доли совпавших токенов и минимальная длина фразы для locate() —
# измерены на двух приёмочных корпусах (Правки_ColBERT_20, Правки_Математика_20),
# не угаданы: 0.75 находит 9 из 11 реальных промахов flex_find, короче 3 токенов
# совпадения слишком случайны, чтобы на них полагаться.
_LOCATE_THRESHOLD = 0.75
_LOCATE_MIN_TOKENS = 3


def _texts(block):
    """Текстовые поля блока, по которым ищем: текст абзаца или все ячейки таблицы."""
    if block["kind"] == "p":
        return [block["text"]]
    return [cell for row in block["rows"] for cell in row]


def _flex_span(haystack, needle):
    """(start, end) реального совпадения needle в haystack, или None.

    Сначала точный str.find (точные совпадения матчатся как раньше, без
    изменений). Документ пестрит U+00A0 не только перед пунктуацией (см.
    находку Ф3), а вообще между случайными словами середины предложения —
    61 из 116 абзацев, 147 вхождений. Цитата в правке или в ответе модели
    набрана обычными пробелами и как подстрока не находится. Поэтому если
    точный поиск промахнулся, needle режется по пробельным разрывам и
    собирается в regex, где любой их пробельный разрыв (обычный пробел,
    U+00A0, таб, перевод строки — всё это `\\s`) матчится с любым другим.
    end берётся из реального совпадения, а не из len(needle): пробельный
    разрыв может быть длиннее или короче исходного.
    """
    pos = haystack.find(needle)
    if pos != -1:
        return pos, pos + len(needle)
    parts = [re.escape(p) for p in re.split(r"\s+", needle) if p]
    if not parts:
        return None
    m = re.search(r"\s+".join(parts), haystack)
    return (m.start(), m.end()) if m else None


def flex_find(haystack, needle):
    """Индекс needle в haystack с учётом пробельных различий, или -1."""
    span = _flex_span(haystack, needle)
    return span[0] if span else -1


def by_text(blocks, needle):
    """id блоков в порядке документа, где needle встречается как подстрока
    (с учётом ячеек таблиц и пробельных различий, регистр важен — как в
    patch._op_replace_all)."""
    return [b["id"] for b in blocks if any(flex_find(t, needle) != -1 for t in _texts(b))]


def _tokens(s):
    return re.findall(r"\w+", s.lower())


def locate(blocks, phrase):
    """id ОДНОГО блока, где phrase встречается ПРИБЛИЗИТЕЛЬНО — только для
    НАВИГАЦИИ (выбор блоков для фрагмента), никогда для мутации: patch.validate
    и patch._op_replace_text по-прежнему используют строгий flex_find/_flex_span
    без изменений, иначе можно заменить не тот кусок текста.

    Люди цитируют документ неточно (тире вместо запятой и т.п.), и такие
    цитаты flex_find не находит вовсе. Здесь текст блока и фраза токенизируются
    (\\w+, без регистра), токены фразы ищутся по порядку в токенах блока
    жадным продвижением позиции; блок годится, если совпала доля токенов
    >= _LOCATE_THRESHOLD.

    Раньше возвращались ВСЕ блоки не ниже порога, по убыванию доли — казалось
    безопасным (Навигатору же нужны варианты). На 40 живых правках это дало
    3 выигрыша (дрейфнувшая цитата человека), но 6 регрессий: locate — запасной
    путь для КАЖДОГО якоря/цитаты в _resolve (edit.py), а якоря Навигатора
    короткие и общие и матчатся сразу во множестве блоков (одна правка ушла с
    4 резолвленных id на 10). Фрагмент раздувался, и Редактор начинал
    отказывать по объёму («правка требует синхронного изменения нескольких
    блоков»). Механизм верный, жадность — нет: "запасной путь, когда точный
    поиск ничего не нашёл" по-честному значит "фраза живёт в одном конкретном
    блоке", а не "вот всё, что смутно похоже". Поэтому теперь всегда не больше
    одного id — лучший по доле; при равенстве долей это настоящая
    неоднозначность и берётся первый по порядку документа, а не оба сразу
    (вернуть весь набор при равенстве — тот же дефект, который чинится).
    [] — если фраза короче _LOCATE_MIN_TOKENS или ни один блок не дотянул до
    порога.
    """
    phrase_tokens = _tokens(phrase)
    if len(phrase_tokens) < _LOCATE_MIN_TOKENS:
        return []
    best_ratio, best_id = 0, None
    for b in blocks:
        block_tokens = _tokens(" ".join(_texts(b)))
        pos = matched = 0
        for pt in phrase_tokens:
            try:
                pos = block_tokens.index(pt, pos) + 1
                matched += 1
            except ValueError:
                pass
        ratio = matched / len(phrase_tokens)
        if ratio >= _LOCATE_THRESHOLD and ratio > best_ratio:
            best_ratio, best_id = ratio, b["id"]
    return [best_id] if best_id is not None else []


def by_regex(blocks, pattern):
    """То же самое, но needle — регулярное выражение (строка или re.Pattern)."""
    rx = re.compile(pattern)
    return [b["id"] for b in blocks if any(rx.search(t) for t in _texts(b))]


def first_mention(blocks, term):
    """Первый по порядку документа блок с term, или None. Основа для «расшифруй
    при первом упоминании» (правка 14 из приёмочного списка)."""
    hits = by_text(blocks, term)
    return hits[0] if hits else None


def outline(blocks):
    """Компактная карта документа — единственное, что видит Навигатор.

    Заголовки определяются по `level` (w:outlineLvl из pPr), а не по имени
    стиля: приёмочный документ размечен прямым форматированием и стилей
    Heading*/Title не содержит вовсе (см. Ф9) — старая проверка по стилю
    находила ноль заголовков и превращала оглавление в плоский список.
    Заголовок — текст целиком с пометкой уровня (H0/H1/H2), чтобы модель
    видела вложенность; обычный абзац — id и обрезанный префикс, как раньше.
    Абзац без level, но целиком жирный (находка Ф10: «заголовок — это просто
    жирный абзац») получает пометку [B] — тот же признак, что render() пишет
    как «весь жирный», но в одну короткую метку, а не в описание.
    """
    lines = []
    for b in blocks:
        if b["kind"] == "t":
            ncols = len(b["rows"][0]) if b["rows"] else 0
            lines.append(f'{b["id"]} [table {len(b["rows"])}x{ncols}]')
            continue
        level, text = b.get("level"), b["text"]
        if level is not None:
            lines.append(f'{b["id"]} [H{level}] {text}')
        else:
            runs = b.get("runs") or []
            bold = "[B] " if runs and all(r.get("b") for r in runs) else ""
            prefix = text[:_PREFIX] + "…" if len(text) > _PREFIX else text
            lines.append(f'{b["id"]} {bold}{prefix}')
    return "\n".join(lines)


def fragment(blocks, ids, around=1):
    """Найденные блоки плюс around соседей с каждой стороны, в порядке
    документа, без дублей при пересекающихся диапазонах."""
    ids = set(ids)
    wanted = set()
    for i, b in enumerate(blocks):
        if b["id"] in ids:
            wanted.update(range(max(0, i - around), min(len(blocks), i + around + 1)))
    return [blocks[i] for i in sorted(wanted)]

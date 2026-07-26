"""Детерминированный поиск по блокам документа: подстрока, regex, первое
упоминание, компактное оглавление для Навигатора и фрагмент с соседями.

Индекс заранее не строим — документ 35 тыс. знаков, линейного прохода
достаточно (см. BUILD_PLAN: эмбеддинги и BM25 сознательно не нужны).
"""

import re

_PREFIX = 20


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
            prefix = text[:_PREFIX] + "…" if len(text) > _PREFIX else text
            lines.append(f'{b["id"]} {prefix}')
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

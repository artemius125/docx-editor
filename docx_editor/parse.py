"""Индекс блоков документа: ID стабильны (ссылки на lxml-элементы, не позиции)."""

from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

_P = qn("w:p")
_TBL = qn("w:tbl")


def index(doc):
    """id -> lxml-элемент. Выдаётся один раз проходом по doc.element.body."""
    idx = {}
    p_n = t_n = 0
    for el in doc.element.body:
        if el.tag == _P:
            idx[f"p{p_n}"] = el
            p_n += 1
        elif el.tag == _TBL:
            idx[f"t{t_n}"] = el
            t_n += 1
    return idx


_ALIGN = {0: "left", 1: "center", 2: "right", 3: "justify"}


def _runs(p):
    """Разбивка абзаца на куски с начертанием; соседние одинаковые склеены.

    Начертание берётся как есть: None означает «наследуется от стиля», и
    страница трактует его как «не задано». Пустые куски выпадают — они не
    несут ни текста, ни оформления.
    """
    out = []
    for r in p.runs:
        if not r.text:
            continue
        piece = {"text": r.text}
        for key, value in (("b", r.bold), ("i", r.italic), ("u", r.underline)):
            if value:
                piece[key] = True
        if out and _same_format(out[-1], piece):
            out[-1]["text"] += piece["text"]
        else:
            out.append(piece)
    return out


def _same_format(a, b):
    return all(a.get(k) == b.get(k) for k in ("b", "i", "u"))


def _level(p_el):
    """Уровень структуры из w:outlineLvl (прямое форматирование), или None.

    Документ может не иметь именованных стилей Heading* вовсе (см. Ф9) и при
    этом нести реальную структуру через outlineLvl в pPr — его и читаем
    напрямую из XML, а не через p.style.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None
    outline = pPr.find(qn("w:outlineLvl"))
    if outline is None:
        return None
    return int(outline.get(qn("w:val")))


def _list(p_el):
    """{"ilvl","numId"} из w:numPr абзаца (тоже прямое форматирование), или None.

    Отсутствующий w:ilvl внутри w:numPr означает уровень 0 (так у Word).
    Без w:numId привязать абзац к конкретной нумерации нельзя — это не список.
    """
    pPr = p_el.find(qn("w:pPr"))
    if pPr is None:
        return None
    numPr = pPr.find(qn("w:numPr"))
    if numPr is None:
        return None
    numId_el = numPr.find(qn("w:numId"))
    if numId_el is None:
        return None
    ilvl_el = numPr.find(qn("w:ilvl"))
    ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
    return {"ilvl": ilvl, "numId": int(numId_el.get(qn("w:val")))}


def _footnotes(p_el):
    """Число w:footnoteReference внутри абзаца (Ф19-бис: сноска не меняет ни
    текст, ни стиль/level/list абзаца — без отдельного счётчика diff цикла
    её не видит вовсе и объявляет правку пустышкой)."""
    return len(p_el.findall(".//" + qn("w:footnoteReference")))


def _fields(p_el):
    """Число полей Word (простых w:fldSimple и составных — по числу
    w:fldChar begin) внутри абзаца (В2-бис: тот же приём, что и
    _footnotes, — поле не меняет ни текст, ни style/level/list абзаца)."""
    simple = len(p_el.findall(".//" + qn("w:fldSimple")))
    begins = [f for f in p_el.findall(".//" + qn("w:fldChar")) if f.get(qn("w:fldCharType")) == "begin"]
    return simple + len(begins)


def doc_map(doc, idx):
    """Блоки в текущем порядке body. Элементы, удалённые из body, выпадают.

    Поля style/text/rows — то, из чего render() строит карту для модели.
    Поля align/runs добавлены для отрисовки страницы в браузере; в карту
    для LLM они не попадают, чтобы не раздувать контекст.
    """
    ids_by_el = {el: block_id for block_id, el in idx.items()}
    blocks = []
    for el in doc.element.body:
        block_id = ids_by_el.get(el)
        if block_id is None:
            continue
        if el.tag == _P:
            p = Paragraph(el, doc)
            blocks.append({
                "id": block_id,
                "kind": "p",
                "style": p.style.name,
                "text": p.text,
                "align": _ALIGN.get(p.alignment),
                "runs": _runs(p),
                "level": _level(el),
                "list": _list(el),
                "footnotes": _footnotes(el),
                "fields": _fields(el),
            })
        elif el.tag == _TBL:
            t = Table(el, doc)
            rows = [[cell.text for cell in row.cells] for row in t.rows]
            blocks.append({"id": block_id, "kind": "t", "rows": rows})
    return blocks


def _truncate_span(text):
    return text[:60] + "…" if len(text) > 60 else text


_FLAG_ADJ = {"b": "жирный", "i": "курсивный", "u": "подчёркнутый"}
_FLAG_INSTR = {"b": "жирным", "i": "курсивом", "u": "подчёркнутым"}


def _format_note(runs):
    """Оформление абзаца для тега render(): если флаг стоит на ВСЕХ
    непустых runs — «весь X» (это и есть «заголовок = жирный абзац» из
    находки Ф10); иначе — сами размеченные фрагменты текста, обрезанные
    _truncate_span, чтобы пёстрый абзац не раздул тег."""
    notes = []
    for flag in ("b", "i", "u"):
        flagged = [r for r in runs if r.get(flag)]
        if not flagged:
            continue
        if len(flagged) == len(runs):
            notes.append(f"весь {_FLAG_ADJ[flag]}")
        else:
            spans = ", ".join(f'«{_truncate_span(r["text"])}»' for r in flagged)
            notes.append(f"{_FLAG_INSTR[flag]}: {spans}")
    return notes


def _tag(b):
    """Содержимое [] в render(): стиль + то, что реально есть (level/list/
    оформление). Пустой абзац без ничего этого даёт просто стиль — как раньше."""
    parts = [b["style"]]
    if b["level"] is not None:
        parts.append(f'H{b["level"]}')
    if b["list"] is not None:
        parts.append(f'список {b["list"]["ilvl"]}')
    parts += _format_note(b["runs"])
    return ", ".join(parts)


def render(blocks):
    """Карта для LLM, одна строка на блок: 'p12 [Heading 2] текст' /
    't3 [table 2x2] r0c0:a | r0c1:b ;; r1c0:c | r1c1:d' — r{row}c{col} даёт
    модели явный адрес ячейки для set_cell. Тег в [] несёт метаданные
    (стиль/level/list/оформление), текст после [] — дословный текст блока
    без изменений: patch.validate ищет в нём "old" буквально (находка Ф10).

    Текст блока БЕЗ усечения (Ф12): раньше резался на 300 знаках + «…», и
    модель добросовестно копировала обрубок в "old" — _flex_span его
    находил (он ДЕЙСТВИТЕЛЬНО есть в документе), а хвост разорванного слова
    оставался приклеен к новому тексту. Бюджетное основание для 300 знаков
    было верно, когда во фрагмент шёл весь документ; с кластеризацией по
    соседству (edit.py) фрагмент — 2-3 блока, экономить контекст здесь не на
    чем. find.outline() покрывает ВЕСЬ документ и режет короткий префикс —
    это другой потребитель с другим бюджетом, его не трогаем."""
    lines = []
    for b in blocks:
        if b["kind"] == "p":
            lines.append(f'{b["id"]} [{_tag(b)}] {b["text"]}')
        else:
            rows = b["rows"]
            ncols = len(rows[0]) if rows else 0
            row_strs = [
                " | ".join(f"r{r}c{c}:{cell}" for c, cell in enumerate(row))
                for r, row in enumerate(rows)
            ]
            body = " ;; ".join(row_strs)
            lines.append(f'{b["id"]} [table {len(rows)}x{ncols}] {body}')
    return "\n".join(lines)

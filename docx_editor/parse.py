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
            })
        elif el.tag == _TBL:
            t = Table(el, doc)
            rows = [[cell.text for cell in row.cells] for row in t.rows]
            blocks.append({"id": block_id, "kind": "t", "rows": rows})
    return blocks


def _truncate(text):
    return text[:300] + "…" if len(text) > 300 else text


def render(blocks):
    """Карта для LLM, одна строка на блок: 'p12 [Heading 2] текст' /
    't3 [table 2x2] r0c0:a | r0c1:b ;; r1c0:c | r1c1:d' — r{row}c{col} даёт
    модели явный адрес ячейки для set_cell."""
    lines = []
    for b in blocks:
        if b["kind"] == "p":
            lines.append(f'{b["id"]} [{b["style"]}] {_truncate(b["text"])}')
        else:
            rows = b["rows"]
            ncols = len(rows[0]) if rows else 0
            row_strs = [
                " | ".join(f"r{r}c{c}:{cell}" for c, cell in enumerate(row))
                for r, row in enumerate(rows)
            ]
            body = _truncate(" ;; ".join(row_strs))
            lines.append(f'{b["id"]} [table {len(rows)}x{ncols}] {body}')
    return "\n".join(lines)

"""Рот-аларм для web/index.html (Ф6, вердикт в интерфейсе): JS здесь не
выполняется (демо без браузера), но текстовые контрактные точки — что
verdict вообще доходит от потока /edit до строки правки и до итоговой
сводки — проверяются на исходнике, чтобы правка вёрстки тихо не снесла Ф6.
"""

import re
from pathlib import Path

HTML = Path(__file__).resolve().parent.parent / "web" / "index.html"


def _extract(name, text):
    """Тело функции по балансу фигурных скобок — надёжнее, чем матчить
    целые строки, которые меняются при любой косметической правке."""
    m = re.search(rf"function {name}\([^)]*\)\s*{{", text)
    assert m, f"{name} не найдена в web/index.html"
    i, depth = m.end(), 1
    while depth:
        depth += {"{": 1, "}": -1}.get(text[i], 0)
        i += 1
    return text[m.end():i]


def main():
    text = HTML.read_text(encoding="utf-8")

    consume = _extract("consumeEdit", text)
    assert re.search(r"ops\.push\(\{[^}]*verdict", consume), \
        "consumeEdit обязан пробрасывать evt.verdict в объект, который кладёт в agentMsg.ops"

    op_line = _extract("renderOpLine", text)
    assert "verdict" in op_line, "renderOpLine обязан читать вердикт правки"
    assert "rolled_back" in op_line, \
        "renderOpLine обязан отдельно отмечать rolled_back — иначе откат Проверяющего неотличим от провала"
    assert "already" in op_line, \
        "renderOpLine обязан отдельно отмечать already (Ф8 item A) — иначе честный «уже было» неотличим от обычного done"

    agent_msg = _extract("renderAgentMsg", text)
    assert "rolledBack" in agent_msg, \
        "renderAgentMsg обязан учитывать откаченные правки в итоговой сводке, а не только done/failed"

    # Ф9: структура документа — прямое форматирование (w:outlineLvl/w:numPr),
    # рендер обязан читать level/list.ilvl, а не только именованный style.
    render_p = _extract("renderParagraph", text)
    assert "b.level" in render_p, \
        "renderParagraph обязан рисовать заголовки по b.level (Ф9), а не только по стилю"
    assert "b.list" in render_p and "ilvl" in render_p, \
        "renderParagraph обязан рисовать списки по b.list.ilvl (Ф9)"

    print("web_demo: ok")


if __name__ == "__main__":
    main()

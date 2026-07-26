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

    agent_msg = _extract("renderAgentMsg", text)
    assert "rolledBack" in agent_msg, \
        "renderAgentMsg обязан учитывать откаченные правки в итоговой сводке, а не только done/failed"

    print("web_demo: ok")


if __name__ == "__main__":
    main()

"""Один вызов модели: POST на OpenAI-совместимый /chat/completions, разбор JSON.

Ключ и параметры — в .env в корне репозитория (не коммитится, парсим сами,
без новой зависимости). Ретраев, абстракции провайдера и резервной модели
здесь нет — решено архитектором (BUILD_PLAN, Ф5): это последнее средство,
а не сервис.
"""

import json
from pathlib import Path

import httpx

_ENV_PATH = Path(__file__).resolve().parent.parent / ".env"


def _env():
    """Построчный разбор .env (KEY=VALUE, # — комментарий)."""
    values = {}
    for line in _ENV_PATH.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        values[key.strip()] = value.strip()
    return values


def chat(messages: list[dict]) -> dict:
    """Отправляет messages модели из .env, возвращает разобранный JSON-объект.

    Всегда JSON-режим и temperature=0.1 — так решено в плане. Падает громко:
    и на HTTP-ошибке (raise_for_status), и если content не парсится как JSON
    (json.JSONDecodeError) — тихого запасного пути нет.
    """
    env = _env()
    url = env["LLM_BASE_URL"].rstrip("/") + "/chat/completions"
    r = httpx.post(
        url,
        headers={"Authorization": f"Bearer {env['AITUNNEL_KEY']}"},
        json={
            "model": env["LLM_MODEL"],
            "messages": messages,
            "temperature": 0.1,
            "response_format": {"type": "json_object"},
        },
        timeout=120,
    )
    r.raise_for_status()
    content = r.json()["choices"][0]["message"]["content"]
    return json.loads(content)

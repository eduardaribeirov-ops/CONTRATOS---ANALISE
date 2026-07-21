"""Chamada à API da Anthropic (Claude) para analisar o texto do contrato."""

from __future__ import annotations

import json
import os

from .prompt import SYSTEM_PROMPT, build_user_prompt

DEFAULT_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-4-5-20250929")
MAX_TOKENS = 8000


class AnalysisError(Exception):
    pass


def analyze_contract(filename: str, contract_text: str, api_key: str, model: str | None = None) -> dict:
    """Envia o contrato para a API da Anthropic e retorna a análise em dict."""
    try:
        import anthropic
    except ImportError as exc:
        raise AnalysisError(
            "O pacote 'anthropic' não está instalado. Rode: pip install -r requirements.txt"
        ) from exc

    client = anthropic.Anthropic(api_key=api_key)
    user_prompt = build_user_prompt(filename, contract_text)

    try:
        response = client.messages.create(
            model=model or DEFAULT_MODEL,
            max_tokens=MAX_TOKENS,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_prompt}],
        )
    except anthropic.APIError as exc:
        raise AnalysisError(
            f"Falha ao chamar a API da Anthropic: {exc}\n"
            "Verifique sua chave de API (ANTHROPIC_API_KEY) e, se o erro mencionar "
            "o nome do modelo, ajuste a variável ANTHROPIC_MODEL para um modelo "
            "válido na sua conta (veja https://docs.anthropic.com/en/docs/about-claude/models)."
        ) from exc

    raw_text = "".join(
        block.text for block in response.content if getattr(block, "type", None) == "text"
    ).strip()

    return _parse_json_response(raw_text)


def _parse_json_response(raw_text: str) -> dict:
    text = raw_text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
        text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise AnalysisError(
            "A resposta da IA não veio em JSON válido. Resposta recebida:\n" + raw_text
        ) from exc

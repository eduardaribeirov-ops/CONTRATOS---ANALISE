#!/usr/bin/env python3
"""CLI: analisa um contrato (.docx ou .pdf) e gera um relatório .docx com
resumo e pontos de melhoria/ajuste, usando a API da Anthropic (Claude).

Uso:
    python analisar_contrato.py caminho/para/contrato.docx
    python analisar_contrato.py caminho/para/contrato.pdf -o relatorio.docx
"""

from __future__ import annotations

import argparse
import os
import sys

from contract_analyzer import AnalysisError, ExtractionError, analyze_contract, build_report, extract_text

try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass


def main() -> int:
    parser = argparse.ArgumentParser(description="Analisa um contrato (.docx/.pdf) com IA.")
    parser.add_argument("arquivo", help="Caminho do contrato (.docx ou .pdf)")
    parser.add_argument(
        "-o", "--output", help="Caminho do relatório de saída (.docx). "
        "Padrão: <nome_do_arquivo>_analise.docx"
    )
    parser.add_argument(
        "--api-key", default=os.environ.get("ANTHROPIC_API_KEY"),
        help="Chave da API da Anthropic. Padrão: variável de ambiente ANTHROPIC_API_KEY",
    )
    parser.add_argument(
        "--model", default=None, help="Sobrescreve o modelo (padrão: variável ANTHROPIC_MODEL)"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.arquivo):
        print(f"Erro: arquivo não encontrado: {args.arquivo}", file=sys.stderr)
        return 1

    if not args.api_key:
        print(
            "Erro: nenhuma chave de API informada. Defina a variável de ambiente "
            "ANTHROPIC_API_KEY (ou crie um arquivo .env a partir de .env.example) "
            "ou use --api-key.",
            file=sys.stderr,
        )
        return 1

    output_path = args.output or _default_output_path(args.arquivo)

    print(f"Extraindo texto de: {args.arquivo}")
    try:
        text = extract_text(args.arquivo)
    except ExtractionError as exc:
        print(f"Erro na extração: {exc}", file=sys.stderr)
        return 1

    print("Enviando contrato para análise (isso pode levar alguns segundos)...")
    try:
        analysis = analyze_contract(
            filename=os.path.basename(args.arquivo),
            contract_text=text,
            api_key=args.api_key,
            model=args.model,
        )
    except AnalysisError as exc:
        print(f"Erro na análise: {exc}", file=sys.stderr)
        return 1

    build_report(analysis, os.path.basename(args.arquivo), output_path)

    n_pontos = len(analysis.get("pontos", []))
    print(f"Concluído. {n_pontos} ponto(s) de melhoria/ajuste identificado(s).")
    print(f"Relatório salvo em: {output_path}")
    return 0


def _default_output_path(input_path: str) -> str:
    base, _ = os.path.splitext(input_path)
    return f"{base}_analise.docx"


if __name__ == "__main__":
    sys.exit(main())

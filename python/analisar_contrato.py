#!/usr/bin/env python3
"""CLI: analisa um contrato (.docx ou .pdf) e gera um relatório .docx com
resumo e pontos de melhoria/ajuste, usando um checklist de regras 100%
offline (sem IA, sem chave de API, sem custo e sem enviar o contrato para
a internet).

Uso:
    python analisar_contrato.py caminho/para/contrato.docx
    python analisar_contrato.py caminho/para/contrato.pdf -o relatorio.docx
"""

from __future__ import annotations

import argparse
import os
import sys

from contract_analyzer import ExtractionError, analisar_por_regras, build_report, extract_text


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Analisa um contrato (.docx/.pdf) com um checklist de regras offline."
    )
    parser.add_argument("arquivo", help="Caminho do contrato (.docx ou .pdf)")
    parser.add_argument(
        "-o", "--output", help="Caminho do relatório de saída (.docx). "
        "Padrão: <nome_do_arquivo>_analise.docx"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.arquivo):
        print(f"Erro: arquivo não encontrado: {args.arquivo}", file=sys.stderr)
        return 1

    output_path = args.output or _default_output_path(args.arquivo)

    print(f"Extraindo texto de: {args.arquivo}")
    try:
        text = extract_text(args.arquivo)
    except ExtractionError as exc:
        print(f"Erro na extração: {exc}", file=sys.stderr)
        return 1

    print("Analisando contrato (checklist de regras, offline)...")
    analysis = analisar_por_regras(text)

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

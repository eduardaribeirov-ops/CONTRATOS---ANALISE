# Análise de Contratos

Plugin/ferramenta para analisar contratos automaticamente com IA (Claude, da
Anthropic) e gerar um resumo com todos os pontos de melhoria/ajuste necessários,
antes de assinar ou negociar um contrato.

Duas formas de uso, complementares:

| | Onde roda | Formatos aceitos | Como usar |
|---|---|---|---|
| [`word-macro/`](word-macro/) | Dentro do Word (Windows) | `.docx` (documento aberto) | Um clique no Word, gera um novo documento com a análise |
| [`python/`](python/) | Terminal | `.docx` e `.pdf` | `python analisar_contrato.py contrato.pdf` |

As duas usam o mesmo roteiro de análise (mesmos critérios jurídicos verificados),
implementado separadamente em cada linguagem — veja `python/contract_analyzer/prompt.py`
e a função `GerarSystemPrompt` em `word-macro/AnaliseContratos.bas`.

## O que a análise verifica

Partes e qualificação, objeto do contrato, valor/pagamento/reajuste, prazo e
renovação, multas e penalidades, condições de rescisão, confidencialidade, proteção
de dados (LGPD), responsabilidade civil, garantias, força maior, propriedade
intelectual, foro e solução de conflitos, validade de assinatura, cláusulas
potencialmente abusivas/desequilibradas, e referências a anexos ausentes.

Para cada ponto encontrado, o relatório traz: gravidade (Alta/Média/Baixa), a
cláusula/trecho relacionado, o problema identificado e uma sugestão de ajuste.

## Por onde começar

- Só precisa analisar `.docx` direto no Word, sem instalar Python? Vá para
  [`word-macro/INSTALACAO.md`](word-macro/INSTALACAO.md).
- Precisa analisar PDFs, ou rodar vários contratos em lote pelo terminal? Vá para
  [`python/README.md`](python/README.md).

Em ambos os casos você vai precisar de uma chave de API da Anthropic
(https://console.anthropic.com/settings/keys), que é paga por uso.

## Aviso

Esta ferramenta usa IA para gerar uma análise preparatória e não substitui a revisão
de um advogado. O texto do contrato é enviado para a API da Anthropic durante a
análise — avalie sigilo/confidencialidade antes de usar com documentos sensíveis.

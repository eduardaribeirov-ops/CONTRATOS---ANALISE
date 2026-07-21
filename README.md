# Análise de Contratos

Plugin/ferramenta para analisar contratos automaticamente e gerar um resumo com
todos os pontos de melhoria/ajuste necessários, antes de assinar ou negociar um
contrato.

**100% offline, sem custo e sem chave de API**: a análise é feita por um checklist
de regras (busca de cláusulas comuns por palavra-chave) que roda inteiramente no seu
computador — nada do contrato é enviado para a internet.

Três formas de uso, complementares:

| | Onde roda | Formatos aceitos | Como usar |
|---|---|---|---|
| [`windows/`](windows/) | Windows (arrastar-e-soltar) | `.docx` **e** `.pdf` | Arraste o contrato sobre `AnalisarContrato.bat` — relatório abre sozinho |
| [`word-macro/`](word-macro/) | Dentro do Word (Windows **e** Mac) | `.docx` (documento aberto) | Um clique no Word, gera um novo documento com a análise |
| [`python/`](python/) | Terminal (Windows, Mac ou Linux) | `.docx` e `.pdf` | `python analisar_contrato.py contrato.pdf` |

As três usam o mesmo roteiro de regras — veja `python/contract_analyzer/rules.py`
e a função `GerarAnaliseLocal` em `word-macro/AnaliseContratos.bas`. O lançador
do Windows é apenas um atalho de um clique por cima do motor Python.

## O que a análise verifica

Qualificação das partes, objeto do contrato, valor/pagamento/reajuste, prazo e
renovação automática, multas e penalidades, condições de rescisão (com aviso
prévio), confidencialidade, proteção de dados (LGPD), responsabilidade civil,
garantias, força maior, propriedade intelectual, foro e solução de conflitos
(mediação/arbitragem), validade de assinatura/testemunhas, e referências a anexos.

Para cada ponto encontrado, o relatório traz: gravidade (Alta/Média/Baixa), o
trecho relacionado (ou "ausente no contrato"), o problema identificado e uma
sugestão de ajuste.

## Como funciona e suas limitações

O motor procura, no texto do contrato, palavras-chave associadas a cada uma dessas
cláusulas e sinaliza o que parece ausente ou incompleto. Isso é uma **varredura por
palavras-chave, não uma leitura jurídica do contrato**:

- Pode haver **falsos positivos** (a cláusula existe, mas usa palavras diferentes
  das procuradas).
- Pode haver **falsos negativos** (a palavra-chave aparece, mas a cláusula na
  prática é inadequada ou desequilibrada — o motor não avalia o *conteúdo* da
  cláusula, só a presença de termos).

Use como um primeiro filtro/checklist, não como substituto de revisão jurídica.

## Por onde começar

- **Está no Windows e quer o jeito mais automático (inclusive PDF)?** Vá para
  [`windows/INSTALACAO_WINDOWS.md`](windows/INSTALACAO_WINDOWS.md): você arrasta o
  contrato sobre um arquivo e o relatório abre sozinho, sem terminal.
- Só precisa analisar `.docx` direto no Word, sem instalar nada mais? Vá para
  [`word-macro/INSTALACAO.md`](word-macro/INSTALACAO.md) (Windows e Mac).
- Precisa rodar em Mac/Linux, ou vários contratos em lote pelo terminal? Vá para
  [`python/README.md`](python/README.md) (funciona igual em Windows, Mac e Linux).

## Aviso

Esta ferramenta gera uma análise preparatória automática e não substitui a revisão
de um advogado. Sempre revise com um profissional antes de assinar ou enviar
qualquer alteração ao contrato.

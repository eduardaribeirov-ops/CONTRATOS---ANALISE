# Análise de Contratos (Python) - .docx e .pdf

Script de linha de comando que analisa um contrato (`.docx` ou `.pdf`) usando um
**checklist de regras 100% offline** (busca por palavras-chave de cláusulas comuns em
contratos brasileiros) e gera um relatório em Word (`.docx`) com o resumo do contrato
e os pontos de melhoria/ajuste identificados (com gravidade, cláusula relacionada e
sugestão de redação).

**Sem custo, sem chave de API e sem internet**: nada do contrato é enviado para
fora do seu computador — toda a análise roda localmente.

## Instalação

```bash
cd python
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Uso

```bash
python analisar_contrato.py caminho/para/contrato.docx
python analisar_contrato.py caminho/para/contrato.pdf -o relatorio.docx
```

Por padrão o relatório é salvo como `<nome_do_arquivo>_analise.docx`, na mesma pasta
do contrato original.

## Como funciona e suas limitações

O script procura, no texto do contrato, palavras-chave associadas a cláusulas comuns
(objeto, valor, reajuste, prazo, multas, rescisão, confidencialidade, LGPD,
responsabilidade civil, garantias, força maior, propriedade intelectual, foro,
mediação/arbitragem, assinatura/testemunhas etc.) e sinaliza o que parece ausente ou
incompleto — ver a lista completa em `contract_analyzer/rules.py`.

Isso é uma **varredura por palavras-chave, não uma leitura jurídica do contrato**:
- Pode haver **falsos positivos** (a cláusula existe, mas usa palavras diferentes das
  procuradas).
- Pode haver **falsos negativos** (a palavra-chave aparece, mas a cláusula na prática
  é inadequada ou desequilibrada — o script não avalia o *conteúdo* da cláusula).

Use como um primeiro filtro, não como substituto de revisão jurídica.

## Observações

- PDFs escaneados (imagem, sem texto selecionável) não são suportados diretamente —
  rode um OCR neles antes (por exemplo, com a skill/CLI de PDF deste ambiente, ou
  ferramentas como `ocrmypdf`).
- Sempre revise com um advogado antes de assinar ou enviar qualquer alteração.

# Análise de Contratos (Python) - .docx e .pdf

Script de linha de comando que analisa um contrato (`.docx` ou `.pdf`), usando a API
da Anthropic (Claude), e gera um relatório em Word (`.docx`) com o resumo do contrato
e todos os pontos de melhoria/ajuste identificados (com gravidade, cláusula relacionada
e sugestão de redação).

## Instalação

```bash
cd python
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env        # depois edite o .env e cole sua ANTHROPIC_API_KEY
```

Obtenha sua chave de API em https://console.anthropic.com/settings/keys (uso pago por
consulta — veja https://www.anthropic.com/pricing).

## Uso

```bash
python analisar_contrato.py caminho/para/contrato.docx
python analisar_contrato.py caminho/para/contrato.pdf -o relatorio.docx
```

Por padrão o relatório é salvo como `<nome_do_arquivo>_analise.docx`, na mesma pasta
do contrato original.

## Observações

- PDFs escaneados (imagem, sem texto selecionável) não são suportados diretamente —
  rode um OCR neles antes (por exemplo, com a skill/CLI de PDF deste ambiente, ou
  ferramentas como `ocrmypdf`).
- A análise é gerada por IA e tem caráter informativo/preparatório — sempre revise com
  um advogado antes de assinar ou enviar qualquer alteração.
- O contrato é enviado pela internet para a API da Anthropic. Avalie sigilo/confidencialidade
  antes de usar com documentos sensíveis.

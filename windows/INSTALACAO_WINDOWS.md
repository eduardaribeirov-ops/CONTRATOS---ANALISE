# Análise de Contratos — Um clique no Windows (.docx e .pdf)

Esta é a forma **mais automática** de analisar contratos no Windows, inclusive
**PDFs**: você arrasta o contrato para um arquivo e o relatório aparece pronto,
sem abrir terminal e sem digitar comandos.

> Só quer analisar `.docx` já aberto no Word, com um botão? Use a macro em
> [`../word-macro/INSTALACAO.md`](../word-macro/INSTALACAO.md). Este lançador
> aqui é o que resolve **PDF** e vários contratos de uma vez.

## Pré-requisito (uma vez só)

1. Instale o **Python 3** em <https://www.python.org/downloads/>.
2. Na tela de instalação, marque **"Add Python to PATH"** antes de clicar em
   *Install Now*. (Se esquecer, reinstale marcando essa opção.)

Não precisa instalar mais nada: na primeira análise, o próprio lançador cria um
ambiente isolado e baixa as bibliotecas necessárias automaticamente.

## Como usar

Baixe/clone esta pasta para o seu computador e use de uma das formas abaixo.

### Forma 1 — Arrastar e soltar (recomendada)

Arraste o arquivo do contrato (`.docx` ou `.pdf`) para cima de
**`AnalisarContrato.bat`** e solte. Pode arrastar **vários contratos de uma
vez**. Uma janela preta abre mostrando o progresso e, ao final, o relatório
`<nome_do_contrato>_analise.docx` abre sozinho no Word, na mesma pasta do
contrato.

> A **primeira** execução demora alguns minutos (ela prepara o ambiente Python).
> Da segunda vez em diante é quase instantânea.

### Forma 2 — Duplo clique

Dê um duplo clique em **`AnalisarContrato.bat`** e, quando pedido, cole o
caminho do contrato e pressione Enter.

### Forma 3 — Clique direito ▸ "Enviar para" (a mais prática no dia a dia)

Assim você analisa qualquer contrato pelo menu do botão direito, de qualquer
pasta:

1. Pressione **Windows + R**, digite `shell:sendto` e tecle Enter. Abre a pasta
   "Enviar para".
2. Nessa pasta, clique com o botão direito ▸ **Novo ▸ Atalho**.
3. Aponte o atalho para o arquivo `AnalisarContrato.bat` (use *Procurar…*).
4. Dê ao atalho o nome **Analisar Contrato** e conclua.

Pronto: agora clique com o botão direito em qualquer `.docx` ou `.pdf` ▸
**Enviar para ▸ Analisar Contrato**, e o relatório é gerado e aberto sozinho.
(No Windows 11, pode ser preciso clicar em *"Mostrar mais opções"* para ver o
menu "Enviar para".)

## O que o relatório traz

Um resumo do contrato (partes, objeto, valor, prazo) e a lista de **pontos de
melhoria/ajuste**, cada um com gravidade (Alta/Média/Baixa), a cláusula
relacionada (ou "ausente no contrato"), o problema e uma sugestão de ajuste.
Veja a lista completa de itens verificados no [README principal](../README.md).

## Se algo der errado

- **"Python nao foi encontrado"**: o Python não está instalado ou não foi
  adicionado ao PATH. Reinstale marcando *"Add Python to PATH"*.
- **PDF sem texto (escaneado/foto)**: se o PDF for uma imagem, não há texto para
  ler. Passe um OCR antes (ex.: ferramentas como `ocrmypdf`) e analise o PDF
  resultante.
- **O relatório não abriu sozinho**: ele foi salvo mesmo assim, como
  `<nome_do_contrato>_analise.docx`, na mesma pasta do contrato.

## Aviso

A análise é uma varredura automática por palavras-chave (sem IA), feita para
servir de checklist preparatório. Ela **não substitui a revisão de um
advogado** — sempre revise com um profissional antes de assinar.

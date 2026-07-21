# Como instalar a macro "Análise de Contratos" no Word

Funciona igual em **Word para Windows** e **Word para Mac** — não precisa de chave de
API, conta paga nem internet: toda a análise roda localmente, por um checklist de
palavras-chave (veja como funciona e as limitações no [README principal](../README.md)).

## 1. Habilite macros no Word

- **Windows**: Arquivo > Opções > Central de Confiabilidade > Configurações da Central
  de Confiabilidade > Configurações de Macro.
- **Mac**: Word > Configurações... (ou Preferências) > Central de Confiabilidade >
  Configurações da Central de Confiabilidade > Configurações de Macro.

Marque *"Desabilitar todas as macros com notificação"* (não precisa liberar tudo —
você só vai rodar a sua própria macro).

## 2. Importe o código da macro

1. Abra o Word e abra o Editor do VBA:
   - **Windows**: pressione **Alt+F11**.
   - **Mac**: **Ferramentas > Macro > Editor do Visual Basic**.
2. No painel à esquerda, clique com o botão direito em **Normal** e escolha
   **Importar Arquivo...**
3. Selecione o arquivo `AnaliseContratos.bas` desta pasta.
4. Isso cria um módulo chamado **AnaliseContratos** com a macro `AnalisarContratoAtual`.
5. Feche o Editor do VBA. Se o Word perguntar se deseja salvar as alterações no modelo
   Normal, clique em **Salvar** — assim a macro fica disponível em qualquer documento,
   sempre, sem precisar reinstalar.

> Se preferir colar em vez de importar: Inserir > Módulo, e cole o conteúdo do arquivo
> `AnaliseContratos.bas` (ignore a primeira linha `Attribute VB_Name = ...`, ela é
> reconstruída automaticamente pelo Word).

## 3. Crie um botão de acesso rápido (opcional)

- **Windows**: Arquivo > Opções > Barra de Ferramentas de Acesso Rápido.
- **Mac**: Word > Configurações... > Barra de Ferramentas e Faixa de Opções (ou
  clique com o botão direito na barra de ferramentas > Personalizar Barra de
  Ferramentas de Acesso Rápido).

Em "Escolher comandos em", selecione **Macros**, escolha
`Normal.AnaliseContratos.AnalisarContratoAtual` e adicione. Um ícone aparecerá na
barra do Word — clique nele para analisar o contrato que estiver aberto.

No Windows você também pode atribuir um atalho de teclado em **Arquivo > Opções >
Personalizar Faixa de Opções > Personalizar... (ao lado de "Atalhos de teclado")**,
categoria "Macros", comando `AnalisarContratoAtual`.

## 4. Usar

1. Abra o contrato (.docx) que deseja analisar.
2. Clique no botão criado no passo 3, ou rode a macro diretamente:
   - **Windows**: Alt+F8 > `AnalisarContratoAtual` > Executar.
   - **Mac**: Ferramentas > Macro > Macros... > `AnalisarContratoAtual` > Executar.
3. A análise é praticamente instantânea (não depende de internet). Um **novo
   documento** é criado automaticamente com o resumo do contrato e a lista de pontos
   de melhoria/ajuste, cada um com gravidade (Alta/Média/Baixa), o trecho relacionado
   e uma sugestão de ajuste. O contrato original não é alterado.

## Observações

- Esta macro só analisa o documento **.docx aberto no Word**. Para analisar arquivos
  **.pdf**, use a versão em Python (`../python`) — mesmo motor de regras, mesma lógica.
- A análise é uma **varredura por palavras-chave**, não uma leitura jurídica do
  contrato: pode haver falsos positivos (a cláusula existe, mas usa outras palavras) e
  falsos negativos (a palavra aparece, mas a cláusula na prática é inadequada). Use
  como um primeiro filtro, não como substituto de revisão jurídica.
- Sempre revise com um advogado antes de assinar ou enviar qualquer alteração.

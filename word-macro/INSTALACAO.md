# Como instalar a macro "Análise de Contratos" no Word

Requer **Word para Windows** (a chamada HTTP usada aqui, `WinHttp.WinHttpRequest`,
é específica do Windows e não funciona no Word para Mac).

## 1. Crie sua chave de API da Anthropic

1. Acesse https://console.anthropic.com/settings/keys e crie uma chave (começa com `sk-ant-`).
2. Guarde essa chave — você vai colar ela uma única vez quando rodar a macro pela primeira vez.
3. O uso da API é pago por uso (não é o mesmo plano do app Claude.ai). Consulte os preços em
   https://www.anthropic.com/pricing antes de usar em muitos contratos.

## 2. Habilite macros no Word

No Word: **Arquivo > Opções > Central de Confiabilidade > Configurações da Central de
Confiabilidade > Configurações de Macro** e marque *"Desabilitar todas as macros com
notificação"* (não é necessário liberar tudo — você só vai rodar as suas próprias macros).

## 3. Importe o código da macro

1. Abra o Word e pressione **Alt+F11** para abrir o Editor do VBA.
2. No painel à esquerda, clique com o botão direito em **Normal** (ou no projeto do
   modelo que você quer usar) e escolha **Importar Arquivo...**
3. Selecione o arquivo `AnaliseContratos.bas` desta pasta.
4. Isso cria um módulo chamado **AnaliseContratos** com as macros `AnalisarContratoAtual`
   e `RedefinirChaveAPI`.
5. Feche o Editor do VBA.

> Se preferir colar em vez de importar: Inserir > Módulo, e cole o conteúdo do arquivo
> `AnaliseContratos.bas` (ignore a primeira linha `Attribute VB_Name = ...`, ela é
> reconstruída automaticamente pelo Word).

## 4. Deixe a macro disponível em qualquer documento (opcional, mas recomendado)

Se você importou no projeto **Normal**, a macro já fica disponível em qualquer documento
aberto no seu Word, sempre — não precisa repetir a instalação por contrato. Basta salvar
as alterações no Normal quando o Word perguntar ao fechar (ou **Arquivo > Salvar** com o
projeto Normal selecionado no Editor do VBA).

## 5. Crie um botão de acesso rápido (opcional)

1. **Arquivo > Opções > Barra de Ferramentas de Acesso Rápido**.
2. Em "Escolher comandos em", selecione **Macros**.
3. Escolha `Normal.AnaliseContratos.AnalisarContratoAtual`, clique em **Adicionar** e **OK**.
4. Um ícone aparecerá na barra superior do Word — clique nele para analisar o contrato
   que estiver aberto.

Você também pode atribuir um atalho de teclado em **Arquivo > Opções > Personalizar
Faixa de Opções > Personalizar... (ao lado de "Atalhos de teclado")**, categoria
"Macros", comando `AnalisarContratoAtual`.

## 6. Usar

1. Abra o contrato (.docx) que deseja analisar.
2. Clique no botão criado no passo 5, ou rode via **Exibir > Macros > Ver Macros >
   AnalisarContratoAtual > Executar** (ou Alt+F8).
3. Na primeira vez, será pedida sua chave de API (fica salva no seu perfil do Windows,
   não é solicitada novamente).
4. Aguarde — para contratos longos pode levar até cerca de um minuto. A barra de status
   do Word mostra o andamento.
5. Um **novo documento** é criado automaticamente com o resumo do contrato e a lista de
   pontos de melhoria/ajuste, cada um com gravidade (Alta/Média/Baixa), o trecho
   relacionado e uma sugestão de ajuste. O contrato original não é alterado.

Para trocar a chave de API salva, rode a macro `RedefinirChaveAPI` (Alt+F8).

## Observações

- Esta macro só analisa o documento **.docx aberto no Word**. Para analisar arquivos
  **.pdf**, use a versão em Python (`../python`).
- O texto do contrato é enviado para a API da Anthropic pela internet. Não use com
  contratos sigilosos sem avaliar as implicações, e revise os
  [termos de uso e política de dados](https://www.anthropic.com/legal) da Anthropic.
- A análise é gerada por IA e tem caráter informativo/preparatório — sempre revise com
  um advogado antes de assinar ou enviar qualquer alteração.

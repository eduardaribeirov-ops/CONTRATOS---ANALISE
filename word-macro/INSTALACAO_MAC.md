# Como instalar a macro "Análise de Contratos" no Word para Mac

Requer **Word para Mac 2016 ou mais recente** (versões do Office 365/Microsoft 365
atuais atendem esse requisito). O arquivo `AnaliseContratos.bas` é o mesmo usado no
Windows — ele detecta automaticamente que está rodando no Mac e usa `curl` (já vem
instalado em todo macOS) em vez do WinHttp do Windows.

> Procurando a versão Windows? Veja [`INSTALACAO.md`](INSTALACAO.md).

## 1. Crie sua chave de API da Anthropic

1. Acesse https://console.anthropic.com/settings/keys e crie uma chave (começa com `sk-ant-`).
2. Guarde essa chave — você vai colar ela uma única vez quando rodar a macro pela primeira vez.
3. O uso da API é pago por uso (plano separado do app Claude.ai). Consulte os preços em
   https://www.anthropic.com/pricing antes de usar em muitos contratos.

## 2. Habilite macros no Word

**Word > Configurações... (ou Preferências) > Central de Confiabilidade >
Configurações da Central de Confiabilidade > Configurações de Macro** e marque
*"Desabilitar todas as macros com notificação"* (não precisa liberar tudo — você só
vai rodar as suas próprias macros).

## 3. Importe o código da macro

1. Abra o Word e vá em **Ferramentas > Macro > Editor do Visual Basic** (em alguns
   teclados Mac o atalho Opção+F11 também funciona).
2. No painel à esquerda, clique com o botão direito em **Normal** e escolha
   **Importar Arquivo...**
3. Selecione o arquivo `AnaliseContratos.bas` desta pasta.
4. Isso cria um módulo chamado **AnaliseContratos** com as macros `AnalisarContratoAtual`
   e `RedefinirChaveAPI`.
5. Feche o Editor do VBA. Se o Word perguntar se deseja salvar as alterações no modelo
   Normal ao fechar o Word, clique em **Salvar** — assim a macro fica disponível em
   qualquer documento, sempre, sem precisar reinstalar.

> Se preferir colar em vez de importar: Inserir > Módulo, e cole o conteúdo do arquivo
> `AnaliseContratos.bas` (ignore a primeira linha `Attribute VB_Name = ...`).

## 4. Crie um botão de acesso rápido (opcional)

**Word > Configurações... > Barra de Ferramentas e Faixa de Opções** (ou clique com o
botão direito na barra de ferramentas > **Personalizar Barra de Ferramentas de Acesso
Rápido**), escolha comandos em **Macros**, selecione
`Normal.AnaliseContratos.AnalisarContratoAtual` e adicione.

## 5. Usar

1. Abra o contrato (.docx) que deseja analisar.
2. Rode a macro: **Ferramentas > Macro > Macros...**, selecione
   `AnalisarContratoAtual` e clique em **Executar** (ou use o botão criado no passo 4).
3. Na primeira execução, o **macOS vai pedir permissão** para o Word controlar outros
   aplicativos/executar automação (usada aqui para chamar `curl`). Isso aparece em
   **Ajustes do Sistema > Privacidade e Segurança > Automação** — confirme/permita o
   Word. Sem essa permissão, a macro não consegue fazer a chamada à internet.
4. Também na primeira vez, será pedida sua chave de API (fica salva no seu perfil
   deste Mac, não é solicitada de novo).
5. Aguarde — para contratos longos pode levar até cerca de um minuto. A barra de
   status do Word mostra o andamento.
6. Um **novo documento** é criado automaticamente com o resumo do contrato e a lista
   de pontos de melhoria/ajuste, cada um com gravidade (Alta/Média/Baixa), o trecho
   relacionado e uma sugestão de ajuste. O contrato original não é alterado.

Para trocar a chave de API salva, rode a macro `RedefinirChaveAPI`.

## Observações

- Esta macro só analisa o documento **.docx aberto no Word**. Para analisar arquivos
  **.pdf**, use a versão em Python (`../python`) — funciona igual no Mac.
- O texto do contrato é enviado para a API da Anthropic pela internet (via `curl`).
  Não use com contratos sigilosos sem avaliar as implicações, e revise os
  [termos de uso e política de dados](https://www.anthropic.com/legal) da Anthropic.
- A análise é gerada por IA e tem caráter informativo/preparatório — sempre revise com
  um advogado antes de assinar ou enviar qualquer alteração.
- Se aparecer um erro dizendo que a chamada falhou por permissão/automação, confira o
  passo 3 acima (Ajustes do Sistema > Privacidade e Segurança > Automação).

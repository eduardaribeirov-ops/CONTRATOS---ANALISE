Attribute VB_Name = "AnaliseContratos"
Option Explicit

' =====================================================================
' AnaliseContratos.bas
'
' Macro para o Microsoft Word que analisa o contrato atualmente aberto
' usando um CHECKLIST DE REGRAS 100% OFFLINE (busca por palavras-chave de
' cláusulas comuns em contratos brasileiros) e monta um NOVO documento do
' Word com o resumo e os pontos de melhoria/ajuste encontrados. O
' documento original nunca é alterado.
'
' Sem custo, sem chave de API e sem internet: nada do contrato sai do seu
' computador. Funciona igual em Word para Windows e Word para Mac, sem
' nenhuma dependência específica de plataforma.
'
' Isto é uma varredura por palavras-chave, não uma leitura jurídica do
' contrato - pode haver falsos positivos (a cláusula existe, mas usa
' outras palavras) e falsos negativos (a palavra aparece, mas a cláusula
' na prática é inadequada). Sempre revise com um advogado. O mesmo
' roteiro de regras é reproduzido em Python em
' python/contract_analyzer/rules.py, para quem preferir analisar PDFs ou
' rodar em lote pelo terminal.
'
' Instalação: veja INSTALACAO.md nesta mesma pasta.
'
' Uso: com um contrato aberto no Word, rode a macro AnalisarContratoAtual
' (Alt+F8 no Windows, ou Ferramentas > Macro > Macros no Mac), ou associe-a
' a um botão na Barra de Ferramentas de Acesso Rápido.
' =====================================================================

Private Const COR_ALTA As Long = 192          ' RGB(192,0,0) - vermelho
Private Const COR_MEDIA As Long = 755384      ' RGB(184,134,11) - dourado escuro
Private Const COR_BAIXA As Long = 3439390     ' RGB(30,123,52) - verde


' ---------------------------------------------------------------------
' Ponto de entrada principal
' ---------------------------------------------------------------------
Sub AnalisarContratoAtual()
    On Error GoTo TratarErro

    If Documents.Count = 0 Then
        MsgBox "Abra o contrato que deseja analisar antes de rodar a macro.", vbExclamation
        Exit Sub
    End If

    Dim docOrigem As Document
    Set docOrigem = ActiveDocument

    Dim textoContrato As String
    textoContrato = TextoDoDocumento(docOrigem)

    If Len(Trim(textoContrato)) < 200 Then
        MsgBox "O documento parece ter pouco texto (menos de 200 caracteres). " & _
               "Confirme se é mesmo o contrato antes de continuar.", vbExclamation
        Exit Sub
    End If

    Dim analiseTexto As String
    analiseTexto = GerarAnaliseLocal(textoContrato)

    Dim novoDoc As Document
    Set novoDoc = Documents.Add
    novoDoc.Activate

    Application.ScreenUpdating = False

    Dim qtdPontos As Long
    qtdPontos = MontarRelatorio(analiseTexto, docOrigem.Name)

    Application.ScreenUpdating = True

    ' Verificação de sanidade: se o documento novo ficou vazio por algum
    ' motivo inesperado, avisa com dados de diagnóstico em vez de deixar
    ' o usuário só olhando para uma página em branco sem explicação.
    If Len(novoDoc.Content.Text) < 50 Then
        MsgBox "Aviso: o documento gerado ficou vazio ou muito curto, o que não era " & _
               "esperado. Dados para diagnóstico - pontos encontrados: " & qtdPontos & _
               "; tamanho do texto de análise gerado internamente: " & Len(analiseTexto) & _
               " caracteres.", vbExclamation
        Exit Sub
    End If

    MsgBox "Análise concluída." & vbCrLf & qtdPontos & _
           " ponto(s) de melhoria/ajuste identificado(s)." & vbCrLf & _
           "Revise o novo documento gerado antes de decidir sobre o contrato.", vbInformation
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    MsgBox "Ocorreu um erro na análise:" & vbCrLf & Err.Description, vbCritical
End Sub


' ---------------------------------------------------------------------
' Extração do texto do documento aberto
' ---------------------------------------------------------------------
Private Function TextoDoDocumento(doc As Document) As String
    Dim t As String
    t = doc.Content.Text
    t = Replace(t, Chr(7), vbTab)   ' fim de célula/linha de tabela
    t = Replace(t, Chr(12), vbCrLf) ' quebra de página
    TextoDoDocumento = t
End Function


' ---------------------------------------------------------------------
' Motor de análise por checklist de regras (mesmo roteiro da versão
' Python em contract_analyzer/rules.py). Gera o texto da análise no
' mesmo formato de marcadores usado antes com a API, para reaproveitar
' o parser/montador de relatório abaixo sem mudanças.
' ---------------------------------------------------------------------
Private Function GerarAnaliseLocal(textoContrato As String) As String
    Dim textoBusca As String
    textoBusca = RemoverAcentos(LCase(textoContrato))

    Dim s As String
    s = "##RESUMO_PARTES##" & vbLf & ExtrairTrecho(textoContrato, textoBusca, Array("cnpj", "cpf", "contratante", "contratado")) & vbLf
    s = s & "##RESUMO_OBJETO##" & vbLf & ExtrairTrecho(textoContrato, textoBusca, Array("objeto do presente", "objeto deste contrato", "objeto do contrato", "tem por objeto")) & vbLf
    s = s & "##RESUMO_VALOR##" & vbLf & ExtrairTrecho(textoContrato, textoBusca, Array("r$", "valor mensal", "valor total", "remuneracao", "preco")) & vbLf
    s = s & "##RESUMO_PRAZO##" & vbLf & ExtrairTrecho(textoContrato, textoBusca, Array("vigencia", "prazo de", "vigorara")) & vbLf

    Dim pontos As String
    pontos = ""

    ' --- Regras de AUSÊNCIA: sinalizadas quando NENHUMA palavra-chave aparece ---
    pontos = pontos & AvaliarAusencia(textoBusca, "Qualificação das partes pode estar incompleta", "Partes", "Média", _
        Array("cnpj", "cpf"), _
        "Não foi localizado CNPJ nem CPF no texto, o que pode indicar que as partes não estão qualificadas de forma completa.", _
        "Confirme que o contrato qualifica completamente as partes (nome/razão social, CPF/CNPJ, endereço).")

    pontos = pontos & AvaliarAusencia(textoBusca, "Objeto do contrato não identificado claramente", "Objeto", "Alta", _
        Array("objeto do presente", "objeto deste contrato", "objeto do contrato", "tem por objeto"), _
        "Não foi localizada uma cláusula clara definindo o objeto do contrato.", _
        "Inclua uma cláusula específica descrevendo o objeto do contrato de forma clara e detalhada.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Valor e forma de pagamento não identificados", "Valor", "Alta", _
        Array("r$", "valor mensal", "valor total", "remuneracao", "preco"), _
        "Não foi localizada referência clara a valor, preço ou remuneração.", _
        "Explicite o valor, a moeda, a forma e o prazo de pagamento.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de reajuste/correção monetária ausente", "Valor", "Média", _
        Array("reajuste", "correcao monetaria", "ipca", "igp-m", "igpm", "inpc"), _
        "Não há previsão de índice de reajuste/correção monetária, relevante para contratos de longa duração.", _
        "Se o contrato for de longa duração, defina o índice (ex.: IPCA, IGP-M) e a periodicidade do reajuste.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Prazo de vigência não identificado", "Prazo", "Média", _
        Array("vigencia", "prazo de", "vigorara", "vigor deste"), _
        "Não foi localizada cláusula clara de prazo/vigência do contrato.", _
        "Defina explicitamente a data de início, a duração e as condições de término do contrato.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de multa/penalidade ausente", "Multas", "Média", _
        Array("multa", "penalidade"), _
        "Não foi localizada cláusula de multa por descumprimento contratual.", _
        "Avalie incluir multa proporcional para as partes em caso de inadimplemento ou descumprimento.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de rescisão ausente", "Rescisão", "Alta", _
        Array("rescisao", "resilicao", "rescindir", "resolucao do contrato"), _
        "Não foi localizada cláusula regulando a rescisão/resilição do contrato.", _
        "Inclua uma cláusula de rescisão prevendo motivos, forma de comunicação e eventuais penalidades.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de confidencialidade ausente", "Confidencialidade", "Média", _
        Array("confidencial", "sigilo"), _
        "Não foi localizada cláusula de confidencialidade/sigilo das informações trocadas entre as partes.", _
        "Avalie incluir cláusula de confidencialidade, especialmente se houver troca de informações sensíveis.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de proteção de dados (LGPD) ausente", "LGPD", "Alta", _
        Array("lgpd", "dados pessoais", "13.709", "protecao de dados"), _
        "Não foi localizada cláusula sobre tratamento de dados pessoais (Lei 13.709/2018 - LGPD).", _
        "Inclua cláusula definindo como os dados pessoais tratados no contrato serão protegidos, conforme a LGPD.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de responsabilidade civil ausente", "Responsabilidade", "Média", _
        Array("responsabilidade civil", "indenizacao", "limitacao de responsabilidade"), _
        "Não foi localizada cláusula sobre responsabilidade civil/indenização entre as partes.", _
        "Defina a responsabilidade de cada parte por danos causados e, se aplicável, limites de indenização.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de garantias ausente", "Garantias", "Baixa", _
        Array("garantia"), _
        "Não foi localizada cláusula de garantias sobre o objeto do contrato.", _
        "Avalie se o contrato deveria prever garantias (ex.: qualidade do serviço/produto, prazo de garantia).")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de força maior ausente", "Força maior", "Média", _
        Array("forca maior", "caso fortuito"), _
        "Não foi localizada cláusula de força maior/caso fortuito.", _
        "Inclua cláusula prevendo tratamento para eventos de força maior ou caso fortuito.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de propriedade intelectual ausente", "Propriedade Intelectual", "Baixa", _
        Array("propriedade intelectual", "direitos autorais"), _
        "Não foi localizada cláusula sobre propriedade intelectual/direitos autorais.", _
        "Se o contrato envolver criação de conteúdo, software ou marca, defina a quem pertencem os direitos.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Foro de eleição não identificado", "Foro", "Alta", _
        Array("foro", "eleicao de foro", "comarca de"), _
        "Não foi localizada cláusula de eleição de foro/comarca para dirimir eventuais conflitos.", _
        "Inclua cláusula de eleição de foro, definindo a comarca competente para resolver eventuais litígios.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Cláusula de mediação/arbitragem ausente", "Solução de Conflitos", "Baixa", _
        Array("mediacao", "arbitragem"), _
        "Não foi localizada cláusula de mediação ou arbitragem como alternativa ao Judiciário.", _
        "Avalie incluir uma cláusula de mediação/arbitragem, dependendo da natureza e do valor do contrato.")

    pontos = pontos & AvaliarAusencia(textoBusca, "Testemunhas ou assinatura eletrônica não mencionadas", "Assinatura", "Baixa", _
        Array("testemunha", "assinatura eletronica", "certificado digital"), _
        "Não foi localizada menção a testemunhas ou à validade de assinatura eletrônica.", _
        "Confirme a forma de assinatura (física com testemunhas, ou eletrônica) e sua validade jurídica.")

    ' --- Regras CONDICIONAIS: só avaliadas se o "gatilho" estiver presente ---
    pontos = pontos & AvaliarCondicional(textoContrato, textoBusca, _
        "Renovação automática sem regra clara de não renovação", "Prazo", "Média", _
        Array("renovacao automatica", "renovado automaticamente", "prorrogacao automatica", "prorrogado automaticamente"), _
        Array("nao renovacao", "manifestacao em contrario", "aviso previo", "antecedencia minima", "notificar a nao renovacao"), _
        "O contrato prevê renovação automática, mas não foi localizada uma regra clara de como evitar a renovação.", _
        "Defina o prazo e a forma de manifestação para que uma das partes possa optar por não renovar o contrato.")

    pontos = pontos & AvaliarCondicional(textoContrato, textoBusca, _
        "Rescisão sem prazo de aviso prévio", "Rescisão", "Média", _
        Array("rescisao", "resilicao", "rescindir"), _
        Array("aviso previo", "notificacao previa", "antecedencia"), _
        "O contrato prevê rescisão, mas não foi localizado um prazo de aviso prévio para rescindir.", _
        "Defina um prazo mínimo de aviso prévio para que qualquer parte possa rescindir o contrato.")

    ' --- Regra INFORMATIVA: sempre exibida como lembrete quando aparece ---
    If InStr(textoBusca, "anexo") > 0 Then
        pontos = pontos & MontarBlocoPonto("Contrato faz referência a anexo(s)", "Anexos", "Baixa", _
            ExtrairTrecho(textoContrato, textoBusca, Array("anexo")), _
            "O contrato menciona um ou mais anexos.", _
            "Confirme que todos os anexos citados estão de fato anexados, atualizados e assinados junto com o contrato.")
    End If

    s = s & pontos
    s = s & "##CONCLUSAO##" & vbLf & MontarConclusaoLocal(pontos) & vbLf
    s = s & "##FIM##"

    GerarAnaliseLocal = s
End Function

' Sinaliza um ponto quando NENHUMA das palavras-chave aparece no contrato.
Private Function AvaliarAusencia(textoBusca As String, titulo As String, categoria As String, _
    gravidade As String, palavrasChave As Variant, problema As String, sugestao As String) As String

    Dim palavra As Variant
    For Each palavra In palavrasChave
        If InStr(textoBusca, CStr(palavra)) > 0 Then
            AvaliarAusencia = ""
            Exit Function
        End If
    Next palavra

    AvaliarAusencia = MontarBlocoPonto(titulo, categoria, gravidade, "ausente no contrato", problema, sugestao)
End Function

' Sinaliza um ponto quando alguma palavra-gatilho aparece, MAS nenhuma das
' palavras esperadas (que deveriam acompanhá-la) é encontrada.
Private Function AvaliarCondicional(textoOriginal As String, textoBusca As String, titulo As String, _
    categoria As String, gravidade As String, palavrasGatilho As Variant, palavrasEsperadas As Variant, _
    problema As String, sugestao As String) As String

    Dim palavra As Variant
    Dim gatilhoEncontrado As Boolean
    gatilhoEncontrado = False
    For Each palavra In palavrasGatilho
        If InStr(textoBusca, CStr(palavra)) > 0 Then
            gatilhoEncontrado = True
            Exit For
        End If
    Next palavra

    If Not gatilhoEncontrado Then
        AvaliarCondicional = ""
        Exit Function
    End If

    For Each palavra In palavrasEsperadas
        If InStr(textoBusca, CStr(palavra)) > 0 Then
            AvaliarCondicional = "" ' cláusula esperada também foi encontrada - sem problema
            Exit Function
        End If
    Next palavra

    Dim trecho As String
    trecho = ExtrairTrecho(textoOriginal, textoBusca, palavrasGatilho)
    AvaliarCondicional = MontarBlocoPonto(titulo, categoria, gravidade, trecho, problema, sugestao)
End Function

Private Function MontarBlocoPonto(titulo As String, categoria As String, gravidade As String, _
    trecho As String, problema As String, sugestao As String) As String

    Dim bloco As String
    bloco = "##PONTO_INICIO##" & vbLf
    bloco = bloco & "TITULO: " & titulo & vbLf
    bloco = bloco & "CATEGORIA: " & categoria & vbLf
    bloco = bloco & "GRAVIDADE: " & gravidade & vbLf
    bloco = bloco & "TRECHO: " & trecho & vbLf
    bloco = bloco & "PROBLEMA: " & problema & vbLf
    bloco = bloco & "SUGESTAO: " & sugestao & vbLf
    bloco = bloco & "##PONTO_FIM##" & vbLf
    MontarBlocoPonto = bloco
End Function

Private Function MontarConclusaoLocal(pontosTexto As String) As String
    Dim qtdAlta As Long, qtdMedia As Long, qtdBaixa As Long
    qtdAlta = ContarOcorrencias(pontosTexto, "GRAVIDADE: Alta")
    qtdMedia = ContarOcorrencias(pontosTexto, "GRAVIDADE: Média")
    qtdBaixa = ContarOcorrencias(pontosTexto, "GRAVIDADE: Baixa")

    If qtdAlta + qtdMedia + qtdBaixa = 0 Then
        MontarConclusaoLocal = _
            "O checklist automático não identificou ausências entre os pontos verificados. " & _
            "Isso NÃO significa que o contrato está juridicamente adequado - esta é uma " & _
            "varredura por palavras-chave, não uma leitura jurídica. Recomenda-se revisão " & _
            "por um advogado antes de assinar."
    Else
        MontarConclusaoLocal = _
            "Foram identificados " & qtdAlta & " ponto(s) de gravidade alta, " & qtdMedia & _
            " de gravidade média e " & qtdBaixa & " de gravidade baixa. Priorize a revisão " & _
            "dos pontos de gravidade alta antes de assinar. Esta análise foi gerada por um " & _
            "checklist automático de palavras-chave (sem uso de IA) e pode gerar falsos " & _
            "positivos/negativos - recomenda-se revisão por um advogado."
    End If
End Function

' Retira acentos comuns do português para tornar a busca por palavra-chave
' mais tolerante a variações de digitação/extração de texto.
Private Function RemoverAcentos(ByVal texto As String) As String
    Const COM_ACENTO As String = "áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ"
    Const SEM_ACENTO As String = "aaaaaeeeeiiiiooooouuuucAAAAAEEEEIIIIOOOOOUUUUC"
    Dim i As Long
    For i = 1 To Len(COM_ACENTO)
        texto = Replace(texto, Mid(COM_ACENTO, i, 1), Mid(SEM_ACENTO, i, 1))
    Next i
    RemoverAcentos = texto
End Function

' Localiza a primeira ocorrência de alguma palavra-chave em textoBusca e
' devolve uma janela de texto ao redor dela, extraída de textoOriginal
' (preservando acentuação/formatação original para exibição).
Private Function ExtrairTrecho(textoOriginal As String, textoBusca As String, palavrasChave As Variant) As String
    Const JANELA As Long = 140

    Dim palavra As Variant, pos As Long
    pos = 0
    For Each palavra In palavrasChave
        pos = InStr(textoBusca, CStr(palavra))
        If pos > 0 Then Exit For
    Next palavra

    If pos = 0 Then
        ExtrairTrecho = "não identificado automaticamente - revise o contrato"
        Exit Function
    End If

    Dim inicio As Long, tamanho As Long
    inicio = pos - JANELA
    If inicio < 1 Then inicio = 1
    tamanho = JANELA * 2
    If inicio + tamanho - 1 > Len(textoOriginal) Then tamanho = Len(textoOriginal) - inicio + 1

    Dim trecho As String
    trecho = Mid(textoOriginal, inicio, tamanho)
    trecho = Replace(trecho, vbCrLf, " ")
    trecho = Replace(trecho, vbCr, " ")
    trecho = Replace(trecho, vbLf, " ")
    trecho = Replace(trecho, vbTab, " ")
    Do While InStr(trecho, "  ") > 0
        trecho = Replace(trecho, "  ", " ")
    Loop
    trecho = Trim(trecho)

    If inicio > 1 Then trecho = "(...) " & trecho
    If inicio + tamanho - 1 < Len(textoOriginal) Then trecho = trecho & " (...)"

    ExtrairTrecho = trecho
End Function


' ---------------------------------------------------------------------
' Parsing da análise (marcadores) e montagem do documento de saída
' ---------------------------------------------------------------------
' Monta o relatório digitando no documento ATIVO (novoDoc, já ativado pelo
' chamador) através do objeto Selection - a forma mais simples e testada
' de gerar conteúdo formatado via VBA no Word, em vez de manipular
' objetos Range diretamente.
Private Function MontarRelatorio(analiseTexto As String, nomeArquivoOrigem As String) As Long
    Escreve "Análise de Contrato", negrito:=True, tamanho:=20
    Escreve "", quebraDepois:=True

    Escreve "Arquivo analisado: " & nomeArquivoOrigem, italico:=True, quebraDepois:=True
    Escreve "Data da análise: " & Format(Now, "dd/mm/yyyy"), italico:=True, quebraDepois:=True
    Escreve "", quebraDepois:=True

    Escreve "Resumo do Contrato", negrito:=True, tamanho:=15, quebraDepois:=True
    Escreve "Partes: ", negrito:=True
    Escreve ValorAposMarcador(analiseTexto, "##RESUMO_PARTES##"), quebraDepois:=True
    Escreve "Objeto: ", negrito:=True
    Escreve ValorAposMarcador(analiseTexto, "##RESUMO_OBJETO##"), quebraDepois:=True
    Escreve "Valor e pagamento: ", negrito:=True
    Escreve ValorAposMarcador(analiseTexto, "##RESUMO_VALOR##"), quebraDepois:=True
    Escreve "Prazo e vigência: ", negrito:=True
    Escreve ValorAposMarcador(analiseTexto, "##RESUMO_PRAZO##"), quebraDepois:=True
    Escreve "", quebraDepois:=True

    Dim qtdPontos As Long
    qtdPontos = ContarOcorrencias(analiseTexto, "##PONTO_INICIO##")

    Escreve "Pontos de Melhoria / Ajuste Necessários (" & qtdPontos & ")", _
        negrito:=True, tamanho:=15, quebraDepois:=True

    If qtdPontos = 0 Then
        Escreve "Nenhum ponto relevante identificado.", quebraDepois:=True
    Else
        InserirPontos analiseTexto
    End If

    Escreve "", quebraDepois:=True
    Escreve "Conclusão", negrito:=True, tamanho:=15, quebraDepois:=True
    Escreve ValorAposMarcador(analiseTexto, "##CONCLUSAO##"), quebraDepois:=True

    Escreve "", quebraDepois:=True
    Escreve "Este documento é gerado automaticamente por um checklist de regras " & _
        "(busca por palavras-chave, sem uso de IA) e tem caráter apenas informativo/" & _
        "preparatório - pode haver falsos positivos e falsos negativos. Recomenda-se " & _
        "revisão por um advogado antes de qualquer decisão.", italico:=True, tamanho:=9

    MontarRelatorio = qtdPontos
End Function

Private Sub InserirPontos(analiseTexto As String)
    Const MARC_INICIO As String = "##PONTO_INICIO##"
    Const MARC_FIM As String = "##PONTO_FIM##"

    Dim pos As Long, posIni As Long, posFim As Long, numero As Long
    pos = 1
    numero = 0

    Do
        posIni = InStr(pos, analiseTexto, MARC_INICIO)
        If posIni = 0 Then Exit Do
        posIni = posIni + Len(MARC_INICIO)
        posFim = InStr(posIni, analiseTexto, MARC_FIM)
        If posFim = 0 Then Exit Do

        Dim bloco As String
        bloco = Trim(Mid(analiseTexto, posIni, posFim - posIni))
        numero = numero + 1
        InserirUmPonto numero, bloco

        pos = posFim + Len(MARC_FIM)
    Loop
End Sub

Private Sub InserirUmPonto(numero As Long, bloco As String)
    Dim titulo As String, categoria As String, gravidade As String
    Dim trecho As String, problema As String, sugestao As String

    Dim blocoNormalizado As String
    blocoNormalizado = Replace(bloco, vbCrLf, vbLf)
    blocoNormalizado = Replace(blocoNormalizado, vbCr, vbLf)

    Dim linhas() As String
    linhas = Split(blocoNormalizado, vbLf)

    Dim i As Long, linha As String, posDoisPontos As Long, chave As String, valor As String
    For i = LBound(linhas) To UBound(linhas)
        linha = linhas(i)
        posDoisPontos = InStr(linha, ": ")
        If posDoisPontos > 0 Then
            chave = UCase(Trim(Left(linha, posDoisPontos - 1)))
            valor = Trim(Mid(linha, posDoisPontos + 2))
            Select Case chave
                Case "TITULO": titulo = valor
                Case "CATEGORIA": categoria = valor
                Case "GRAVIDADE": gravidade = valor
                Case "TRECHO": trecho = valor
                Case "PROBLEMA": problema = valor
                Case "SUGESTAO": sugestao = valor
            End Select
        End If
    Next i

    Dim cor As Long
    Select Case gravidade
        Case "Alta": cor = COR_ALTA
        Case "Média", "Media": cor = COR_MEDIA
        Case "Baixa": cor = COR_BAIXA
        Case Else: cor = 0
    End Select

    Escreve numero & ". " & titulo, negrito:=True, tamanho:=12, quebraDepois:=True
    Escreve "Gravidade: " & gravidade, negrito:=True, cor:=cor
    If categoria <> "" Then Escreve "   |   Categoria: " & categoria
    Escreve "", quebraDepois:=True
    If trecho <> "" Then
        Escreve "Cláusula/trecho: ", italico:=True
        Escreve trecho, italico:=True, quebraDepois:=True
    End If
    If problema <> "" Then
        Escreve "Problema: ", negrito:=True
        Escreve problema, quebraDepois:=True
    End If
    If sugestao <> "" Then
        Escreve "Sugestão de ajuste: ", negrito:=True
        Escreve sugestao, quebraDepois:=True
    End If
    Escreve "", quebraDepois:=True
End Sub


' ---------------------------------------------------------------------
' Auxiliares genéricos
' ---------------------------------------------------------------------

' Digita texto na posição atual do cursor (Selection) no documento ativo,
' aplicando formatação, e restaura a formatação padrão em seguida para
' não "vazar" para o próximo trecho escrito.
Private Sub Escreve(texto As String, _
    Optional negrito As Boolean = False, Optional italico As Boolean = False, _
    Optional cor As Long = 0, Optional tamanho As Single = 0, _
    Optional quebraDepois As Boolean = False)

    With Selection.Font
        .Bold = negrito
        .Italic = italico
        .Size = IIf(tamanho > 0, tamanho, 11)
        If cor <> 0 Then
            .Color = cor
        Else
            .Color = wdColorAutomatic
        End If
    End With

    If Len(texto) > 0 Then Selection.TypeText texto
    If quebraDepois Then Selection.TypeParagraph

    ' restaura o padrão para a próxima chamada não herdar esta formatação
    With Selection.Font
        .Bold = False
        .Italic = False
        .Size = 11
        .Color = wdColorAutomatic
    End With
End Sub

' Retorna o texto entre um marcador "##NOME##" e o próximo marcador "##".
Private Function ValorAposMarcador(texto As String, marcador As String) As String
    Dim p As Long
    p = InStr(texto, marcador)
    If p = 0 Then
        ValorAposMarcador = "não identificado"
        Exit Function
    End If
    Dim inicioValor As Long
    inicioValor = p + Len(marcador)
    Dim proximoMarcador As Long
    proximoMarcador = InStr(inicioValor, texto, "##")
    If proximoMarcador = 0 Then
        ValorAposMarcador = Trim(Mid(texto, inicioValor))
    Else
        ValorAposMarcador = Trim(Mid(texto, inicioValor, proximoMarcador - inicioValor))
    End If
End Function

Private Function ContarOcorrencias(texto As String, sub_ As String) As Long
    Dim c As Long, p As Long
    p = 1
    Do
        p = InStr(p, texto, sub_)
        If p = 0 Then Exit Do
        c = c + 1
        p = p + Len(sub_)
    Loop
    ContarOcorrencias = c
End Function

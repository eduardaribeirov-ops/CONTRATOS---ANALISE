Attribute VB_Name = "AnaliseContratos"
Option Explicit

' =====================================================================
' AnaliseContratos.bas
'
' Macro para o Microsoft Word que envia o contrato atualmente aberto para
' a API da Anthropic (Claude), recebe uma análise (resumo + pontos de
' melhoria/ajuste) e monta um NOVO documento do Word com o resultado
' formatado. O documento original nunca é alterado.
'
' Instalação: veja INSTALACAO.md nesta mesma pasta.
'
' Uso: com um contrato aberto no Word, rode a macro AnalisarContratoAtual
' (Alt+F8 -> AnalisarContratoAtual -> Executar), ou associe-a a um botão
' na Barra de Ferramentas de Acesso Rápido.
' =====================================================================

Public Const MODEL_NAME As String = "claude-sonnet-4-5-20250929"
Public Const ANTHROPIC_VERSION As String = "2023-06-01"
Public Const API_URL As String = "https://api.anthropic.com/v1/messages"
Public Const MAX_TOKENS As Long = 8000

Private Const APP_SETTINGS_NAME As String = "AnaliseContratosIA"
Private Const APP_SETTINGS_SECTION As String = "Config"

Private Const COR_ALTA As Long = 192          ' RGB(192,0,0) - vermelho
Private Const COR_MEDIA As Long = 2925740     ' RGB(184,134,11) - dourado escuro
Private Const COR_BAIXA As Long = 3382558     ' RGB(30,123,52) - verde


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

    Dim apiKey As String
    apiKey = ObterChaveAPI()
    If apiKey = "" Then
        MsgBox "Análise cancelada: nenhuma chave de API foi informada.", vbExclamation
        Exit Sub
    End If

    Dim textoContrato As String
    textoContrato = TextoDoDocumento(docOrigem)

    If Len(Trim(textoContrato)) < 200 Then
        MsgBox "O documento parece ter pouco texto (menos de 200 caracteres). " & _
               "Confirme se é mesmo o contrato antes de continuar.", vbExclamation
        Exit Sub
    End If

    Application.StatusBar = "Analisando contrato com IA, aguarde (pode levar cerca de 1 minuto)..."
    Application.ScreenUpdating = False

    Dim analiseTexto As String
    analiseTexto = ChamarClaudeAPI(apiKey, docOrigem.Name, textoContrato)

    Dim novoDoc As Document
    Set novoDoc = Documents.Add

    Dim qtdPontos As Long
    qtdPontos = MontarRelatorio(novoDoc, analiseTexto, docOrigem.Name)

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    novoDoc.Activate

    MsgBox "Análise concluída." & vbCrLf & qtdPontos & _
           " ponto(s) de melhoria/ajuste identificado(s)." & vbCrLf & _
           "Revise o novo documento gerado antes de decidir sobre o contrato.", vbInformation
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    MsgBox "Ocorreu um erro na análise:" & vbCrLf & Err.Description, vbCritical
End Sub


' Remove a chave de API salva, para permitir cadastrar uma nova.
Sub RedefinirChaveAPI()
    DeleteSetting APP_SETTINGS_NAME, APP_SETTINGS_SECTION, "ApiKey"
    MsgBox "Chave de API removida. Na próxima análise, uma nova chave será solicitada.", vbInformation
End Sub


' ---------------------------------------------------------------------
' Configuração / chave de API
' ---------------------------------------------------------------------
Private Function ObterChaveAPI() As String
    Dim chave As String
    chave = GetSetting(APP_SETTINGS_NAME, APP_SETTINGS_SECTION, "ApiKey", "")

    If chave = "" Then
        chave = InputBox( _
            "Informe sua chave de API da Anthropic (começa com 'sk-ant-')." & vbCrLf & vbCrLf & _
            "Você pode gerar uma em: https://console.anthropic.com/settings/keys" & vbCrLf & vbCrLf & _
            "Ela ficará salva no seu perfil do Windows e não será pedida de novo " & _
            "(use a macro RedefinirChaveAPI para trocá-la no futuro).", _
            "Chave de API - Anthropic")
        chave = Trim(chave)
        If chave = "" Then
            ObterChaveAPI = ""
            Exit Function
        End If
        SaveSetting APP_SETTINGS_NAME, APP_SETTINGS_SECTION, "ApiKey", chave
    End If

    ObterChaveAPI = chave
End Function


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
' Chamada à API da Anthropic
' ---------------------------------------------------------------------
Private Function ChamarClaudeAPI(apiKey As String, nomeArquivo As String, textoContrato As String) As String
    Dim jsonBody As String
    jsonBody = "{""model"":""" & MODEL_NAME & """," & _
               """max_tokens"":" & MAX_TOKENS & "," & _
               """system"":""" & EscaparJSON(GerarSystemPrompt()) & """," & _
               """messages"":[{""role"":""user"",""content"":""" & _
               EscaparJSON(GerarUserPrompt(nomeArquivo, textoContrato)) & """}]}"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", API_URL, False
    http.SetRequestHeader "x-api-key", apiKey
    http.SetRequestHeader "anthropic-version", ANTHROPIC_VERSION
    http.SetRequestHeader "content-type", "application/json"
    ' timeouts em ms: resolve, connect, send, receive (contratos longos podem demorar)
    http.SetTimeouts 60000, 60000, 60000, 180000
    http.Send StringParaBytesUtf8(jsonBody)

    Dim respostaTexto As String
    respostaTexto = BytesUtf8ParaString(http.ResponseBody)

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , _
            "A API respondeu com erro " & http.Status & ":" & vbCrLf & respostaTexto & vbCrLf & vbCrLf & _
            "Verifique sua chave de API (rode RedefinirChaveAPI para trocá-la) e, se o erro " & _
            "mencionar o nome do modelo, atualize a constante MODEL_NAME no topo deste módulo " & _
            "(veja https://docs.anthropic.com/en/docs/about-claude/models)."
    End If

    Dim conteudo As String
    conteudo = ExtrairCampoTexto(respostaTexto)
    If conteudo = "" Then
        Err.Raise vbObjectError + 2, , "Não foi possível interpretar a resposta da API:" & vbCrLf & respostaTexto
    End If

    ChamarClaudeAPI = conteudo
End Function


' ---------------------------------------------------------------------
' UTF-8: request/response precisam de conversão explícita para não
' corromper acentuação (o WinHttpRequest, por padrão, não usa UTF-8).
' ---------------------------------------------------------------------
Private Function StringParaBytesUtf8(ByVal texto As String) As Variant
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' texto
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText texto
    stream.Position = 0
    stream.Type = 1 ' binário
    Dim todosOsBytes() As Byte
    todosOsBytes = stream.Read
    stream.Close

    ' O ADODB.Stream grava um BOM UTF-8 (EF BB BF) no início; removemos.
    Dim resultado() As Byte
    Dim i As Long
    ReDim resultado(UBound(todosOsBytes) - 3)
    For i = 3 To UBound(todosOsBytes)
        resultado(i - 3) = todosOsBytes(i)
    Next i
    StringParaBytesUtf8 = resultado
End Function

Private Function BytesUtf8ParaString(ByVal bytes As Variant) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binário
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' texto
    stream.Charset = "utf-8"
    BytesUtf8ParaString = stream.ReadText
    stream.Close
End Function


' ---------------------------------------------------------------------
' Utilitários de JSON (mínimo necessário: escapar string de saída e
' extrair o campo "text" da resposta da Anthropic). A análise em si NÃO
' é pedida em JSON (ver GerarSystemPrompt) para não exigir um parser
' JSON completo em VBA - usamos marcadores de texto simples.
' ---------------------------------------------------------------------
Private Function EscaparJSON(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscaparJSON = s
End Function

Private Function ExtrairCampoTexto(ByVal json As String) As String
    Dim marcador As String
    marcador = """text"":"""
    Dim posIni As Long
    posIni = InStr(json, marcador)
    If posIni = 0 Then
        ExtrairCampoTexto = ""
        Exit Function
    End If
    posIni = posIni + Len(marcador)

    Dim i As Long, c As String, prox As String
    Dim resultado As String
    i = posIni
    Do While i <= Len(json)
        c = Mid(json, i, 1)
        If c = "\" Then
            prox = Mid(json, i + 1, 1)
            Select Case prox
                Case "n": resultado = resultado & vbLf
                Case "t": resultado = resultado & vbTab
                Case "r": ' ignorado (normalmente acompanha \n)
                Case """": resultado = resultado & """"
                Case "\": resultado = resultado & "\"
                Case "/": resultado = resultado & "/"
                Case "u"
                    Dim hexCode As String
                    hexCode = Mid(json, i + 2, 4)
                    resultado = resultado & ChrW(CLng("&H" & hexCode))
                    i = i + 4
                Case Else
                    resultado = resultado & prox
            End Select
            i = i + 2
        ElseIf c = """" Then
            Exit Do
        Else
            resultado = resultado & c
            i = i + 1
        End If
    Loop
    ExtrairCampoTexto = resultado
End Function


' ---------------------------------------------------------------------
' Prompts (mesmo roteiro de análise usado na versão Python, adaptado
' para um formato de marcadores de texto em vez de JSON)
' ---------------------------------------------------------------------
Private Function GerarSystemPrompt() As String
    Dim s As String
    s = s & "Você é um advogado especialista em direito contratual brasileiro (Código " & vbCrLf
    s = s & "Civil, CDC quando aplicável, CLT quando for contrato de trabalho, e LGPD - " & vbCrLf
    s = s & "Lei 13.709/2018). Revise o contrato enviado pelo usuário e produza uma " & vbCrLf
    s = s & "análise objetiva, prática e acionável, como um parecer de revisão contratual." & vbCrLf & vbCrLf
    s = s & "Analise o contrato considerando, no mínimo, estes aspectos (ignore os que não " & vbCrLf
    s = s & "se aplicarem e cite outros que julgar relevantes): qualificação das partes; " & vbCrLf
    s = s & "clareza do objeto; valor, pagamento e reajuste/correção monetária; prazo de " & vbCrLf
    s = s & "vigência e renovação automática; multas e penalidades (proporcionalidade); " & vbCrLf
    s = s & "condições de rescisão, aviso prévio e multa rescisória; confidencialidade; " & vbCrLf
    s = s & "proteção de dados pessoais (LGPD); responsabilidade civil e suas limitações; " & vbCrLf
    s = s & "garantias; força maior/caso fortuito; propriedade intelectual (se aplicável); " & vbCrLf
    s = s & "foro de eleição, lei aplicável e solução de conflitos (mediação/arbitragem); " & vbCrLf
    s = s & "validade de assinatura (inclusive eletrônica) e testemunhas; cláusulas " & vbCrLf
    s = s & "potencialmente abusivas ou desequilibradas; e referências a anexos ou " & vbCrLf
    s = s & "documentos citados mas ausentes do texto." & vbCrLf & vbCrLf
    s = s & "Responda SEMPRE em português do Brasil, em TEXTO PURO (sem markdown, sem " & vbCrLf
    s = s & "JSON), seguindo EXATAMENTE este formato de marcadores, sem nenhum texto antes " & vbCrLf
    s = s & "do primeiro marcador ou depois do último:" & vbCrLf & vbCrLf
    s = s & "##RESUMO_PARTES##" & vbCrLf
    s = s & "(texto sobre as partes do contrato)" & vbCrLf
    s = s & "##RESUMO_OBJETO##" & vbCrLf
    s = s & "(texto sobre o objeto do contrato)" & vbCrLf
    s = s & "##RESUMO_VALOR##" & vbCrLf
    s = s & "(texto sobre valor e forma de pagamento)" & vbCrLf
    s = s & "##RESUMO_PRAZO##" & vbCrLf
    s = s & "(texto sobre prazo e vigência)" & vbCrLf
    s = s & "##PONTO_INICIO##" & vbCrLf
    s = s & "TITULO: (título curto do ponto)" & vbCrLf
    s = s & "CATEGORIA: (ex.: Rescisão, LGPD, Multas, Foro)" & vbCrLf
    s = s & "GRAVIDADE: (Alta, Média ou Baixa)" & vbCrLf
    s = s & "TRECHO: (cláusula/trecho relevante, ou ""ausente no contrato"" - uma única linha)" & vbCrLf
    s = s & "PROBLEMA: (explicação do risco ou lacuna - uma única linha)" & vbCrLf
    s = s & "SUGESTAO: (sugestão de ajuste - uma única linha)" & vbCrLf
    s = s & "##PONTO_FIM##" & vbCrLf
    s = s & "(repita o bloco ##PONTO_INICIO##/##PONTO_FIM## para cada ponto encontrado, do " & vbCrLf
    s = s & "de maior para o de menor gravidade; não crie pontos artificiais)" & vbCrLf
    s = s & "##CONCLUSAO##" & vbCrLf
    s = s & "(avaliação geral e recomendação de prioridade de revisão)" & vbCrLf
    s = s & "##FIM##" & vbCrLf & vbCrLf
    s = s & "IMPORTANTE: dentro de cada bloco PONTO, cada campo (TITULO, CATEGORIA, " & vbCrLf
    s = s & "GRAVIDADE, TRECHO, PROBLEMA, SUGESTAO) deve ficar em uma única linha, sem " & vbCrLf
    s = s & "quebras de linha no meio do valor."
    GerarSystemPrompt = s
End Function

Private Function GerarUserPrompt(nomeArquivo As String, textoContrato As String) As String
    Dim s As String
    s = "Segue abaixo o texto integral do contrato extraído do arquivo """ & nomeArquivo & """. "
    s = s & "Faça a análise conforme as instruções e responda apenas com os marcadores pedidos." & vbCrLf & vbCrLf
    s = s & "--- INÍCIO DO CONTRATO ---" & vbCrLf
    s = s & textoContrato & vbCrLf
    s = s & "--- FIM DO CONTRATO ---"
    GerarUserPrompt = s
End Function


' ---------------------------------------------------------------------
' Parsing da resposta (marcadores) e montagem do documento de saída
' ---------------------------------------------------------------------
Private Function MontarRelatorio(doc As Document, analiseTexto As String, nomeArquivoOrigem As String) As Long
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd

    Escreve rng, "Análise de Contrato", negrito:=True, tamanho:=20
    Escreve rng, "", quebraDepois:=True

    Escreve rng, "Arquivo analisado: " & nomeArquivoOrigem, italico:=True, quebraDepois:=True
    Escreve rng, "Data da análise: " & Format(Now, "dd/mm/yyyy"), italico:=True, quebraDepois:=True
    Escreve rng, "", quebraDepois:=True

    Escreve rng, "Resumo do Contrato", negrito:=True, tamanho:=15, quebraDepois:=True
    Escreve rng, "Partes: ", negrito:=True
    Escreve rng, ValorAposMarcador(analiseTexto, "##RESUMO_PARTES##"), quebraDepois:=True
    Escreve rng, "Objeto: ", negrito:=True
    Escreve rng, ValorAposMarcador(analiseTexto, "##RESUMO_OBJETO##"), quebraDepois:=True
    Escreve rng, "Valor e pagamento: ", negrito:=True
    Escreve rng, ValorAposMarcador(analiseTexto, "##RESUMO_VALOR##"), quebraDepois:=True
    Escreve rng, "Prazo e vigência: ", negrito:=True
    Escreve rng, ValorAposMarcador(analiseTexto, "##RESUMO_PRAZO##"), quebraDepois:=True
    Escreve rng, "", quebraDepois:=True

    Dim qtdPontos As Long
    qtdPontos = ContarOcorrencias(analiseTexto, "##PONTO_INICIO##")

    Escreve rng, "Pontos de Melhoria / Ajuste Necessários (" & qtdPontos & ")", _
        negrito:=True, tamanho:=15, quebraDepois:=True

    If qtdPontos = 0 Then
        Escreve rng, "Nenhum ponto relevante identificado.", quebraDepois:=True
    Else
        InserirPontos rng, analiseTexto
    End If

    Escreve rng, "", quebraDepois:=True
    Escreve rng, "Conclusão", negrito:=True, tamanho:=15, quebraDepois:=True
    Escreve rng, ValorAposMarcador(analiseTexto, "##CONCLUSAO##"), quebraDepois:=True

    Escreve rng, "", quebraDepois:=True
    Escreve rng, "Este documento é gerado automaticamente por IA e tem caráter apenas " & _
        "informativo/preparatório. Recomenda-se revisão por um advogado antes de " & _
        "qualquer decisão.", italico:=True, tamanho:=9

    MontarRelatorio = qtdPontos
End Function

Private Sub InserirPontos(ByRef rng As Range, analiseTexto As String)
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
        InserirUmPonto rng, numero, bloco

        pos = posFim + Len(MARC_FIM)
    Loop
End Sub

Private Sub InserirUmPonto(ByRef rng As Range, numero As Long, bloco As String)
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

    Escreve rng, numero & ". " & titulo, negrito:=True, tamanho:=12, quebraDepois:=True
    Escreve rng, "Gravidade: " & gravidade, negrito:=True, cor:=cor
    If categoria <> "" Then Escreve rng, "   |   Categoria: " & categoria
    Escreve rng, "", quebraDepois:=True
    If trecho <> "" Then
        Escreve rng, "Cláusula/trecho: ", italico:=True
        Escreve rng, trecho, italico:=True, quebraDepois:=True
    End If
    If problema <> "" Then
        Escreve rng, "Problema: ", negrito:=True
        Escreve rng, problema, quebraDepois:=True
    End If
    If sugestao <> "" Then
        Escreve rng, "Sugestão de ajuste: ", negrito:=True
        Escreve rng, sugestao, quebraDepois:=True
    End If
    Escreve rng, "", quebraDepois:=True
End Sub


' ---------------------------------------------------------------------
' Auxiliares genéricos
' ---------------------------------------------------------------------

' Escreve texto no fim de rng, aplicando formatação, e deixa rng
' posicionado (colapsado) logo após o texto inserido para a próxima chamada.
Private Sub Escreve(ByRef rng As Range, texto As String, _
    Optional negrito As Boolean = False, Optional italico As Boolean = False, _
    Optional cor As Long = 0, Optional tamanho As Single = 0, _
    Optional quebraDepois As Boolean = False)

    Dim posIni As Long
    posIni = rng.End
    rng.InsertAfter texto
    rng.Collapse wdCollapseEnd

    If Len(texto) > 0 Then
        Dim formatado As Range
        Set formatado = rng.Document.Range(posIni, rng.End)
        formatado.Font.Bold = negrito
        formatado.Font.Italic = italico
        If cor <> 0 Then formatado.Font.Color = cor
        If tamanho > 0 Then formatado.Font.Size = tamanho
    End If

    If quebraDepois Then
        rng.InsertParagraphAfter
        rng.Collapse wdCollapseEnd
    End If
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

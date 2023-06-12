Attribute VB_Name = "sfColDadCitPJe1gTrt5"
Option Explicit

Sub CadastrarProcessoIndividualPje1gTrt5(strNumeroProcesso As String, ByRef rngCelula As Excel.Range)
''
'' Com o PJe aberto e logado no Internet Explorer, busca um processo e o cadastra
''
    Dim strCont As String
    Dim strPerfilLogado As String
    Dim dtDataProvContestar As Date, dtDataProvSubsidios As Date
    Dim planGrupoProvContestar As Excel.Worksheet, planGrupoProvSubsidios As Excel.Worksheet
    Dim arq As Workbook, plan As Excel.Worksheet
    
    ' Abrir (se não estiver aberto) e fazer login
    If oChrome Is Nothing Then
        Set oChrome = New Selenium.ChromeDriver
        oChrome.get sfUrlPJe1gTrt5Login
    Else
        oChrome.get sfUrlPJe1gTrt5Painel
    End If
    
    'Descobrir se está logado e, caso esteja, se o perfil é de parte ou advogado.
VerificarLogin:
    strPerfilLogado = DescobrirPerfilLogadoPje1gTrt5(oChrome)
    
    Select Case LCase(strPerfilLogado)
    Case "deslogado, página de login"
        oChrome.Window.Activate
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", rogo que faça o login na janela do Chrome e depois clique em ""OK"".", vbCritical + vbOKOnly, "Sísifo - Aguardando login no Chrome"
        GoTo VerificarLogin
        
    Case "página sem identificação de login"
        oChrome.get sfUrlPJe1gTrt5Painel
        GoTo VerificarLogin
        
    Case "procurador", "advogado"
        ' Avança normalmente, nenhuma ação a realizar.
    
    Case Else
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro no login. Suplico que tente novamente, desta vez fazendo login de procurador da empresa ou de advogado.", _
            vbOKOnly + vbCritical, "Sísifo - Erro no login"
        FecharChromeZerarVariaveis oChrome: Exit Sub
    End Select
    
    ' A partir do painel do usuário, realizar busca no acervo
    Set oChrome = RealizarBuscaAcervoPJe1gTrt5(oChrome, strNumeroProcesso)
    If oChrome Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    
    'Abrir processo
    'Set oChrome = AbrirProcessoPJe1gTrt5(strNumeroProcesso, oChrome)
    'If oChrome Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    
    ' Pega os dados do processo
    PegaInformacoesProcessoPje1gTrt5 oChrome
    If prProcesso Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    PegaInformacoesProcessoGeral dtDataProvContestar, dtDataProvSubsidios, planGrupoProvContestar, planGrupoProvSubsidios
    
    ' Se deu erro, devolve o erro
    If prProcesso Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    If prProcesso.MensagemErro <> "" Then FecharChromeZerarVariaveis oChrome: Exit Sub
    
    ' Se deu certo, insere na memória do Sísifo (PODE HAVER O SEGUINTE PROBLEMA: SE NÃO CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
    Set arq = Excel.Workbooks.Add
    Set plan = arq.Sheets(1)
    
    prProcesso.ExportaLinhasEspaider arq, plan
    If Not planGrupoProvContestar Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvContestar, dtDataProvContestar, prProcesso.Advogado)
    If Not planGrupoProvSubsidios Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvSubsidios, dtDataProvSubsidios, prProcesso.Preposto)
    
    rngCelula.Offset(0, 1).Formula = "Inserido no Sísifo"
    rngCelula.Offset(1, 0).Select
    
End Sub

'Sub CadProcIndividualPJeSegredoJus(ByVal controle As IRibbonControl)
'''
''' Com o Projudi aberto no Internet Explorer e deslogado, busca um processo e o cadastra
'''
'
'    Dim IE As InternetExplorer
'    Dim DocHTML As HTMLDocument
'    Dim frmTeor As HTMLFormElement
'    Dim frFrame As HTMLFrameElement
'    Dim strCont As String, strCadastro As String
'    Dim strSenhaAcesso As String
'    Dim arq As Workbook, plan As Excel.Worksheet, rngCelula As Excel.Range
'    Dim prProcesso As Processo
'    Dim bolCont As Boolean
'
'    Set rngCelula = ActiveCell
'
'    strSenhaAcesso = PegarSenhaAcesso(rngCelula)
'
'    ' Procurar Internet Explorer aberto no Projudi
'    Set IE = SisifoEmbasaFuncoes.RecuperarIE("projudi.tjba.jus.br")
'    If IE Is Nothing Then
'PaginaErrada:
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", é necessário que o Internet Explorer esteja aberto na página inicial do Projudi BA, sem logar em nenhum usuário, para continuar. Abra a página do Projudi, " & _
'            "saia de qualquer login e rode a função novamente.", vbCritical + vbOKOnly, "Sísifo - Internet Explorer não encontrado"
'        End
'    End If
'    IE.Visible = True
'    Set DocHTML = IE.document
'    Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
'    Set DocHTML = frFrame.contentDocument
'
'    ' Pegar form do acesso ao teor com senha e preencher
'    Set frmTeor = DocHTML.getElementById("formAcessoPublico")
'    If frmTeor Is Nothing Then GoTo PaginaErrada
'
'    DocHTML.getElementById("codigoHash").Value = strSenhaAcesso
'    frmTeor.getElementsByTagName("a")(0).Click
'
'
'    Do
'        DoEvents
'    Loop Until IE.readyState = READYSTATE_COMPLETE
'
'    On Error GoTo Volta
'Volta:
'    Do
'        Set IE = SisifoEmbasaFuncoes.RecuperarIE("projudi.tjba.jus.br")
'        Set DocHTML = IE.document
'        Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
'        Set DocHTML = frFrame.contentDocument
'        bolCont = IIf(DocHTML.URL = sfURLProjudiAcessoPublico, True, False)
'    Loop Until bolCont = True
'    On Error GoTo 0
'
'    ' Pega os dados do processo
'    Set prProcesso = DestrincharProcessoPJe(IE, DocHTML)
'
'    ' Se deu erro, devolve o erro
'    If prProcesso Is Nothing Then
'        FecharExplorerZerarVariaveis(IE)
'        End
'    End If
'
'    If prProcesso.MensagemErro <> "" Then
'        FecharExplorerZerarVariaveis(IE)
'        End
'    End If
'
'    ' Se deu certo, insere na memória do Sísifo (PODE DAR ERRO: SE NÃO CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
'    Set arq = Excel.Workbooks.Add
'    Set plan = arq.Sheets(1)
'
'    prProcesso.ExportaLinhasEspaider arq, plan
'
'    rngCelula.Offset(0, 1).Formula = "Inserido no Sísifo"
'    FecharExplorerZerarVariaveis(IE)
'
'    rngCelula.Offset(1, 0).Select
'
'End Sub

Function PegarSenhaAcesso(rngRange As Excel.Range) As String
''
'' Retorna a string contida na primeira célula da range passada como parâmetro -- ou, se não houver, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strSenha As String
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strSenha = Trim(rngCelula.text)
    
    ' Se não houver célula no espaço enviado, ou se estiver vazia, pergunta a senha.
    If rngCelula Is Nothing Or rngCelula.text = "" Then
        strSenha = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe a senha de acesso ao processo do Projudi a cadastrar", "Sísifo - Cadastrar processo"))
    End If
    
    PegarSenhaAcesso = strSenha
    
End Function

Function DescobrirPerfilLogadoPje1gTrt5(oChrome As Selenium.ChromeDriver) As String
''
'' Descobre o perfil do usuário logado no PJe
''
    Dim oSelect As Selenium.WebElement, oOption As Selenium.WebElement
    Dim strTexto As String
    
    ' Caso esteja na página de login
    If InStr(1, oChrome.URL, sfUrlPJe1gTrt5Login) <> 0 Then
        DescobrirPerfilLogadoPje1gTrt5 = "deslogado, página de login"
    
    ' Caso esteja no painel do usuário logado
    ElseIf InStr(1, oChrome.URL, sfUrlPJe1gTrt5Painel) <> 0 Then
        Set oSelect = oChrome.FindElementById("papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao")
        
        For Each oOption In oSelect.FindElementsByTag("option")
            If oOption.IsSelected = True Then
                strTexto = oOption.text
                Exit For
            End If
        Next oOption
        
        DescobrirPerfilLogadoPje1gTrt5 = Right(strTexto, Len(strTexto) - InStr(1, strTexto, ")/") - 1)
    
    ' Caso esteja em outra página (como página de processo)
    Else
        DescobrirPerfilLogadoPje1gTrt5 = "página sem identificação de login"
        
    End If
    
End Function

Function RealizarBuscaAcervoPJe1gTrt5(oChrome As Selenium.ChromeDriver, strNumeroProcesso As String) As Selenium.ChromeDriver
''
'' Abre o acervo e busca o processo
''
Dim divMenu As Selenium.WebElement, inpBusca As Selenium.WebElement
Dim janCont As Selenium.Window

    ' Carrega página do acervo
    oChrome.FindElementById("j_id179:j_id191").Click
    
    ' Preenche o número do processo
    Set inpBusca = oChrome.FindElementById("formLocCaix:decosuggestProcessoAdvogadoProc:suggestProcessoAdvogadoProc")
    Do
        inpBusca.SendKeys strNumeroProcesso
    Loop Until inpBusca.Value = strNumeroProcesso
    
    ' Aguarda aparecer a caixinha com sugestões de número e clica nela. COLOCAR TIMEOUT
    Do
        DoEvents
        Set divMenu = oChrome.FindElementById("formLocCaix:decosuggestProcessoAdvogadoProc:sugsuggestProcessoAdvogadoProc")
    Loop Until InStr(1, divMenu.Attribute("style"), "visibility: visible") <> 0
    oChrome.FindElementByClass("richfaces_suggestionSelectValue").Click
    
    ' Clica no botão Localizar
    oChrome.FindElementById("formLocCaix:btnPesquisa").Click
    
    ' Clica no link do processo
    oChrome.FindElementById("processoTrfInicialAdvogadoList:0:abreTarefDetalhes").Click
    
    For Each janCont In oChrome.Windows
        If InStr(1, janCont.Title, "Detalhes do Processo") = 0 Then janCont.Close
    Next janCont
    
    Set RealizarBuscaAcervoPJe1gTrt5 = oChrome
    
End Function

Function CarregarPaginaBuscaPJe1gTrt5(oChrome As Selenium.ChromeDriver) As Selenium.ChromeDriver
''
'' Abre a página de buscas, conforme perfil logado
''

'    If strPerfilLogado = "Outro" Then
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", é necessário estar logado num perfil de parte, advogado ou representante. Faça login no Projudi de " & _
'            "um desses perfis e rode a função novamente.", vbCritical + vbOKOnly, "Sísifo - Internet Explorer não encontrado"
'        Set CarregarPaginaBusca = Nothing
'        Exit Function
'    End If
    
    ' Carrega página de busca
    oChrome.get sfUrlPJe1gTrt5Busca
    
    Set CarregarPaginaBuscaPJe1gTrt5 = oChrome
    
End Function

Function AbrirProcessoPJe1gTrt5(ByVal strNumeroCNJ As String, ByRef oChrome As Selenium.ChromeDriver) As Selenium.ChromeDriver
''
'' Retorna o objeto InternetExplorer com a página principal do processo strNumeroCNJ. Deve haver uma sessão do Internet Explorer aberta
''  e logada no PJe. Em caso de zero ou múltiplos processos encontrados, ou de cancelamento ao mostrar Recaptcha, mostra mensagem de erro
''  e retorna 'Nothing'.
'' FALTA LIDAR COM O ERRO DE NÃO ESTAR LOGADO!!!!!!!
''
    
    ' AJUSTAR
    'If DocHTML.Title = "Sistema CNJ - A sessão expirou" Then
    '    PegarLinkProcessoProjudi = "Sessão expirada"
    '    Exit Function
    'End If
    
    If oChrome.FindElementById("mat-input-0") Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a página do PJe parece estar com alguma inconsistência. Não consigo acessá-la. " & _
            "Suplico que verifique o que ocorreu e, se tudo estiver normal, tente novamente", vbCritical + vbOKOnly, "Sísifo - Erro no PJe"
        Set AbrirProcessoPJe1gTrt5 = Nothing
        FecharExplorerZerarVariaveis oChrome
        Exit Function
    End If
    
    Do
        oChrome.FindElementById("mat-input-0").Clear
        oChrome.FindElementById("mat-input-0").SendKeys strNumeroCNJ
    Loop Until oChrome.FindElementById("mat-input-0").Value = strNumeroCNJ
    oChrome.FindElementById("btnPesquisar").Click
    
    'Set divCaptcha = oChrome.body.Children(oChrome.body.Children.length - 1)
    'If Not divCaptcha Is Nothing Then
    '    If divCaptcha.Style.visibility = "visible" Then
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", surgiu um captcha na janela do PJe. Peço que o resolva e clique em ""Repetir"".", _
                vbExclamation + vbRetryCancel, "Sísifo - Necessário resolver captcha") = vbCancel Then
                Set AbrirProcessoPJe1gTrt5 = Nothing
                FecharExplorerZerarVariaveis oChrome
                Exit Function
            End If
    '    End If
    'End If
    
    ' Lidar com erro na busca do processo
    If InStr(1, oChrome.FindElementById("painel-erro").Attribute("style"), "display: block") <> 0 Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", não foi encontrado nenhum processo com o número fornecido. Rogo que confira se não está em " & _
            "segredo de justiça e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Set AbrirProcessoPJe1gTrt5 = Nothing
        FecharExplorerZerarVariaveis oChrome
        Exit Function
    End If
    
    ' Caso tenha aberto
    Set AbrirProcessoPJe1gTrt5 = oChrome
    
End Function

Sub PegaInformacoesProcessoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver)
''
'' Faz a coleta dos dados do processo específicos do PJe1g TRT5 e armazena na variável global prProcesso
''
    Dim tbPartes As Selenium.WebElement
    Dim contOutrasPartes As OutroParticipante
    Dim varCont As Variant
    Dim Cont As Integer
    Dim arrStrAudiencias() As String
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a providência de agendar pautista
    
    ''''''''''''''''''''''''''''''''''''
    ''' Carregar totalmente a página '''
    ''''''''''''''''''''''''''''''''''''
    
    'varCont = CarregarTodasAsPaginasAndamentosProcessoPje1gTrt5(oChrome)
    
    'If varCont <> "Sucesso" Then
    '    MsgBox sisifoembasafuncoes.DeterminarTratamento & ", a página do processo não pôde ser totalmente aberta. Favor limpar o cache e tentar novamente.", _
    '        vbCritical + vbOKOnly, "Sísifo - Erro no carregamento do processo"
    '    Set prProcesso = Nothing
    '    Exit Sub
    'End If
    
    
    ''''''''''''''''''''''''''
    ''' Número do processo '''
    ''''''''''''''''''''''''''
    
    prProcesso.NumeroProcesso = PegaNumeroPJe1gTrt5(oChrome)
    
    ' Confere se o processo já está na planilha
    If Not sfCadProcessos.Cells().Find(prProcesso.NumeroProcesso) Is Nothing Then
        Do
        Loop Until MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo já existe na planilha! Inclusão cancelada. Descartados os dados." & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf & _
        "Clique em 'Cancelar' e insira o próximo processo.", vbCritical + vbOKCancel, "Sísifo - Processo repetido") = vbCancel
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Citação, responsável regressivo, juízo,   '''
    ''' valor da causa, juízo, tipo de ação, rito '''
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.Sistema = sfPJe1g
    prProcesso.Citacao = Date
    prProcesso.ResponsavelRegressivo = ""
    prProcesso.ValorCausa = PegaValorPje1gTrt5(oChrome)
    
    prProcesso.Juizo = PegaJuizoPje1gTrt5(oChrome)
    If prProcesso.Juizo = "Juízo não cadastrado" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.TipoAcao = PegaTipoAcaoPje1gTrt5(oChrome)
    If prProcesso.TipoAcao = "Rito não cadastrado no Sísifo" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.Rito = PegaRitoPje1gTrt5(oChrome, prProcesso.TipoAcao)
    
    ''''''''''''''''''''''''''
    ''' Autores e advogado '''
    ''''''''''''''''''''''''''
    
    ' Adiciona Autores
    varCont = PegaPartesPje1gTrt5(oChrome, sfAutor)
    For Cont = 1 To UBound(varCont, 2)
        Set contOutrasPartes = New OutroParticipante
        contOutrasPartes.NomeParte = varCont(1, Cont)
        If InStr(1, varCont(2, Cont), "/") <> 0 Then
            'Pessoa jurídica
            contOutrasPartes.CNPJParte = varCont(2, Cont)
            contOutrasPartes.CPFParte = 0
            contOutrasPartes.TipoParte = 2
        Else
            'Pessoa física ou não cadastrado.
            contOutrasPartes.CPFParte = varCont(2, Cont)
            contOutrasPartes.CNPJParte = 0
            contOutrasPartes.TipoParte = 1
        End If
        
        contOutrasPartes.ClasseParte = 2
        contOutrasPartes.CondicaoParte = "Autor"
        prProcesso.OutrosParticipantes.Add contOutrasPartes
    Next Cont
    
    prProcesso.AdvogadoAdverso = PegaAdvAutorPje1gTrt5(oChrome)
    
    '''''''''''''''''''
    ''' Outros réus '''
    '''''''''''''''''''
    
    ' Adiciona outros Réus, se for o caso.
    varCont = PegaPartesPje1gTrt5(oChrome, sfReu)
    If Not IsEmpty(varCont) Then
        For Cont = 1 To UBound(varCont, 2)
            Set contOutrasPartes = New OutroParticipante
            contOutrasPartes.NomeParte = varCont(1, Cont)
            If InStr(1, varCont(2, Cont), "/") <> 0 Then
                'Pessoa jurídica
                contOutrasPartes.CNPJParte = varCont(2, Cont)
                contOutrasPartes.CPFParte = 0
                contOutrasPartes.TipoParte = 2
            Else
                'Pessoa física ou não cadastrado
                contOutrasPartes.CPFParte = varCont(2, Cont)
                contOutrasPartes.CNPJParte = 0
                contOutrasPartes.TipoParte = 1
            End If
        
            contOutrasPartes.ClasseParte = 1
            contOutrasPartes.CondicaoParte = "Réu"
            prProcesso.OutrosParticipantes.Add contOutrasPartes
        Next Cont
    End If
    
    '''''''''''''''''''''''''''''''
    ''' Andamento de referência '''
    ''' (audiência ou citação)  '''
    '''''''''''''''''''''''''''''''
    
    arrStrAudiencias = PegaDataeTipoAudienciaPje1gTrt5(oChrome)
    prProcesso.DataAndamento = arrStrAudiencias(1)
    
    If arrStrAudiencias(1) = "0" Then
        prProcesso.NomeAndamento = "Citação"
    Else
        prProcesso.NomeAndamento = "Audiência de " & arrStrAudiencias(2)
    End If
        
End Sub

Function PegaNumeroPJe1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strCont1 As String, strCont2 As String
    Dim intCont As Integer
    
    strCont1 = oChrome.FindElementById("processoIdentificadorDiv").text
    
    ' Deixa apenas números, hífens e pontos
    For intCont = 1 To Len(strCont1) Step 1
        If IsNumeric(Mid(strCont1, intCont, 1)) Or Mid(strCont1, intCont, 1) = "." Or Mid(strCont1, intCont, 1) = "-" Then
            strCont2 = strCont2 & Mid(strCont1, intCont, 1)
        End If
    Next intCont
    
    'Apaga o que não for número do início
    strCont1 = strCont2
    Do Until IsNumeric(Left(strCont1, 1))
        strCont1 = Mid(strCont1, 2)
    Loop
    
    'Apaga o que não for número do final
    Do Until IsNumeric(Right(strCont1, 1))
        strCont1 = Left(strCont1, Len(strCont1) - 1)
    Loop
    
    PegaNumeroPJe1gTrt5 = strCont1
    
End Function

Function PegaJuizoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strJuizo As String
    Dim rngCont As Excel.Range
    
    strJuizo = oChrome.FindElementById("orgaoJulgDecoration:orgaoJulg").text
    
    Set rngCont = cfJuizos.Cells().Find(What:=strJuizo, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        InputBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu não conheço o juízo em que o processo tramita. Rogo que o cadastre em minha memória (a redação " & _
            "do PJe pode ser copiada abaixo) e tente novamente.", "Sísifo - Juízo desconhecido", Trim(strJuizo)
        PegaJuizoPje1gTrt5 = "Juízo não cadastrado"
    Else
        PegaJuizoPje1gTrt5 = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaValorPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As Currency

    PegaValorPje1gTrt5 = Trim(CCur(oChrome.FindElementById("valorCausaDecoration:valorCausa").text))
    
End Function

Function PegaTipoAcaoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strTipoAcaoPje As String
    Dim rngCont As Range
    Dim regEx As New RegExp
    
    strTipoAcaoPje = oChrome.FindElementById("processoIdentificadorDiv").text
    
    With regEx
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = " \d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4}.+$"
        strTipoAcaoPje = .Replace(strTipoAcaoPje, "")
    End With
    
    Set rngCont = cfPjeTiposAcaoRitos.Cells().Find(What:=strTipoAcaoPje, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        PegaTipoAcaoPje1gTrt5 = rngCont.Offset(0, 1).Formula
    Else
        PegaTipoAcaoPje1gTrt5 = "Rito não cadastrado no Sísifo"
    End If
    
End Function

Function PegaRitoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver, strTipoAcao As String) As String
    
    Dim rngCont As Range
    
    Set rngCont = cfPjeTiposAcaoRitos.Cells().Find(What:=strTipoAcao, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        PegaRitoPje1gTrt5 = rngCont.Offset(0, 2).Formula
    Else
        PegaRitoPje1gTrt5 = "Rito não cadastrado"
    End If
    
End Function

Function PegaPartesPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver, strTipoParte As String) As Variant
''
'' Retorna uma matriz com as partes de um polo do processo.
'' PegaPartesPje(1, N) = Nome da parte na posição N da matriz.
'' PegaPartesPje(2, N) = CPF/CNPJ da parte na posição N da matriz.
''
    Dim intContPartes As Integer, intQtdPartes As Integer, arrIntCont() As Integer
    Dim strCont As String, strTipoLista As String, strMsgErro As String, strTipoDocumento As String, arrStrPartes() As String
    Dim bolPular As Boolean
    
    Select Case strTipoParte
    Case sfAutor
        strTipoLista = "listaPoloAtivo"
    Case sfReu
        strTipoLista = "listaPoloPassivo"
    End Select
    
    For intContPartes = 1 To oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & ":tb']/tbody/tr").Count Step 1
        ' Condição comum para pular (se não for "RECLAMANTE" ou "RECLAMADO")
        strCont = oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & ":tb']/tr[" & intContPartes & "]/td[2]").text
        If strCont <> "RECLAMANTE" And strCont <> "RECLAMADO" Then bolPular = True
        
        ' Condição específica de réu para pular
        strCont = oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & ":tb']/tr[" & intContPartes & "]/td[1]").text
        If strTipoParte = sfReu And (InStr(1, strCont, "EMPRESA BAIANA DE AGUA") <> 0 Or InStr(1, strCont, "EMBASA") <> 0) Then bolPular = True
        
        If bolPular = False Then
            intQtdPartes = intQtdPartes + 1
            ReDim Preserve arrIntCont(1 To intQtdPartes)
            arrIntCont(intQtdPartes) = intContPartes
        End If
    Next intContPartes
    
    ' Se for cadastro de réu e houver 0 partes, isso significa que é só a Embasa, portanto, vamos pular.
    If strTipoParte = sfReu And intQtdPartes = 0 Then Exit Function
    
    ' Redimensiona a array
    ReDim arrStrPartes(1 To 2, 1 To intQtdPartes)
    
    ' Para cada parte, busca Nome e CPF (advogado pressupõe-se comum). 1 = Nome, 2 = CPF
    ' Itera a quantidade de Partes na matriz
    For intContPartes = 1 To intQtdPartes Step 1
        strCont = Trim(oChrome.FindElementByXPath("//*[@id='" & strTipoLista & ":tb']/tr[" & arrIntCont(intContPartes) & "]/td[1]").text)
        
        ' Descobre o tipo de pessoa
        If InStr(1, LCase(strCont), " - cpf:") <> 0 Then
            strTipoDocumento = "cpf"
        ElseIf InStr(1, LCase(strCont), " - cnpj:") <> 0 Then
            strTipoDocumento = "cnpj"
        Else
            strTipoDocumento = "outros"
            ' COLOCAR AQUI UM AVISO PARA TRATAMENTO DE ERRO
        End If
        
        ' Pega nome...
        arrStrPartes(1, intContPartes) = Trim(Left(strCont, InStr(1, LCase(strCont), " - " & strTipoDocumento & ":")))
    
        ' Pega CPF/CNPJ...
        strCont = Trim(Replace(strCont, arrStrPartes(1, intContPartes), ""))
        strCont = Trim(Replace(LCase(strCont), "- " & strTipoDocumento & ":", ""))
        strCont = Left(strCont, IIf(strTipoDocumento = "cpf", 14, 18))
        arrStrPartes(2, intContPartes) = strCont
        
        ' Checa se CPF está cadastrado...
        If InStr(1, arrStrPartes(2, intContPartes), "Não cadastrado##") <> 0 Then ' Não cadastrado (era uma expressão do Projudi, não sei se vai ter no PJe
            strMsgErro = "o CPF/CNPJ da parte não foi cadastrado no PJe"
        ElseIf InStr(1, arrStrPartes(2, intContPartes), "Não disponível##") <> 0 Then ' Não disponível (era uma expressão do Projudi, não sei se vai ter no PJe
            strMsgErro = "o CPF/CNPJ da parte não está disponível"
        End If
        
        If strMsgErro <> "" Then
PedirCPF:
            arrStrPartes(2, intContPartes) = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", " & strMsgErro & ". Rogo que busque a Petição Inicial " & _
                "e informe o CPF ou CNPJ correto. Se não houver, deixe em branco.", "Sísifo - Informar CPF/CNPJ da parte"))
            If arrStrPartes(2, intContPartes) <> "" Then
                arrStrPartes(2, intContPartes) = ValidarCPFCNPJ(arrStrPartes(2, intContPartes))
                If Not IsNumeric(Left(arrStrPartes(2, intContPartes), 1)) Then
                    strMsgErro = arrStrPartes(2, intContPartes)
                    GoTo PedirCPF
                End If
            End If
        End If
        
        If arrStrPartes(2, intContPartes) = "" Then 'Sem CPF/CNPJ
            arrStrPartes(2, intContPartes) = "Não Cadastrado " & PegaCodParteSemCPF
        ElseIf InStr(1, arrStrPartes(2, intContPartes), "/") = 0 Then ' CPF - 14 caracteres
            arrStrPartes(2, intContPartes) = Left(arrStrPartes(2, intContPartes), 14)
        End If ' CNPJ não precisa de tratamento
    Next intContPartes
    
    PegaPartesPje1gTrt5 = arrStrPartes()
    
End Function

Function PegaAdvAutorPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim tbCorpoTabela As Selenium.WebElement
    Dim intContPartes As Integer
    Dim strAdvAutor As String
    
    For intContPartes = 1 To tbParte.FindElementsByXPath("//*[@id='listaPoloAtivo:tb']/tr").Count Step 1
        If oChrome.FindElementByXPath("//*[@id='listaPoloAtivo:tb']/tr[" & intContPartes & "]/td[2]").text = "ADVOGADO" Then
            strAdvAutor = Trim(oChrome.FindElementByXPath("//*[@id='listaPoloAtivo:tb']/tr[" & intContPartes & "]/td[1]").text)
            Exit For
        End If
    Next intContPartes
    
    If strAdvAutor = "Nenhum advogado cadastrado." Then ' A mensagem é a do Projudi. Não sei como é no PJe.
        strAdvAutor = ""
    Else
        strAdvAutor = UCase(Trim(strAdvAutor))
    End If
    
    PegaAdvAutorPje1gTrt5 = strAdvAutor
    
End Function

Function PegaDataeTipoAudienciaPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String()
''
'' Busca a última audiência com status "designada". Retorna um vetor com as seguintes informações:
'' (1) = Data audiência ou "0", caso não encontre.
'' (2) = Tipo de audiência ou "", caso não encontre.
''
    Dim intCont As Integer, intQtdAudiencias As Integer
    Dim arrStrResultado(1 To 2) As String
    Dim divCont As Selenium.WebElement
    Dim tbAudiencias As Selenium.WebElement
    Dim lngTimerInicio As Long
    
    ' Abre a aba de audiências
    oChrome.FindElementById("tabProcessoAudiencia_shifted").Click
    
    ' Aguarda o carregamento (assíncrono) das informações da tabela, por até 10 segundos
AguardaTabelaAudiencias:
    lngTimerInicio = Timer
    Do
        DoEvents
        If Timer > lngTimerInicio + 10 Then
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o carregamento da tabela de audiências parece estar demorando " & _
                "mais do que o normal. Deseja continuar esperando? Caso não espere, a data da audiência poderá ser alterada manualmente " & _
                "no formulário, ou após o cadastramento, no Espaider.") = vbYes Then
                GoTo AguardaTabelaAudiencias
            Else
                Exit Do
            End If
        End If
    Loop While oChrome.FindElementsByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr").Count = 0
    
    ' Itera as linhas da tabela de audiências, verificando datas e se está ativa
    For intCont = 1 To oChrome.FindElementsByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr").Count Step 1
        ' Confere se há mais de uma audiência com status "designada". Se houver, alerta e mantém a mais antiga.
        If LCase(oChrome.FindElementByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr[" & intCont & "]/td[4]").text) = "designada" Then
            If arrStrResultado(1) <> "" Then
                MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", parece haver mais de uma audiência ativa designada para este processo. Peço que informe o número deste processo a César para verificação.", vbCritical + vbOKOnly, "Sísifo - Mais de uma audiência ativa"
                If CDate(arrStrResultado(1)) > CDate(oChrome.FindElementByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr[" & intCont & "]/td[1]").text) Then
                    GoTo AtribuirDataeTipo
                End If
            Else
AtribuirDataeTipo:
                arrStrResultado(1) = oChrome.FindElementByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr[" & intCont & "]/td[1]").text
                arrStrResultado(2) = oChrome.FindElementByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr[" & intCont & "]/td[2]").text
            End If
        End If
    Next intCont
    
    ' Se não houver andamento "Audiência" ou não se encaixar em nenhuma hipótese anterior, retorna 0.
    If arrStrResultado(1) = "" Then
        arrStrResultado(1) = "0"
        arrStrResultado(2) = ""
    Else
        PegaDataAudienciaPje1gTrt5 = arrStrResultado
    End If

End Function



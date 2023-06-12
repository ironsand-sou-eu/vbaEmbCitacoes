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
    
    ' Abrir (se n�o estiver aberto) e fazer login
    If oChrome Is Nothing Then
        Set oChrome = New Selenium.ChromeDriver
        oChrome.get sfUrlPJe1gTrt5Login
    Else
        oChrome.get sfUrlPJe1gTrt5Painel
    End If
    
    'Descobrir se est� logado e, caso esteja, se o perfil � de parte ou advogado.
VerificarLogin:
    strPerfilLogado = DescobrirPerfilLogadoPje1gTrt5(oChrome)
    
    Select Case LCase(strPerfilLogado)
    Case "deslogado, p�gina de login"
        oChrome.Window.Activate
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", rogo que fa�a o login na janela do Chrome e depois clique em ""OK"".", vbCritical + vbOKOnly, "S�sifo - Aguardando login no Chrome"
        GoTo VerificarLogin
        
    Case "p�gina sem identifica��o de login"
        oChrome.get sfUrlPJe1gTrt5Painel
        GoTo VerificarLogin
        
    Case "procurador", "advogado"
        ' Avan�a normalmente, nenhuma a��o a realizar.
    
    Case Else
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro no login. Suplico que tente novamente, desta vez fazendo login de procurador da empresa ou de advogado.", _
            vbOKOnly + vbCritical, "S�sifo - Erro no login"
        FecharChromeZerarVariaveis oChrome: Exit Sub
    End Select
    
    ' A partir do painel do usu�rio, realizar busca no acervo
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
    
    ' Se deu certo, insere na mem�ria do S�sifo (PODE HAVER O SEGUINTE PROBLEMA: SE N�O CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
    Set arq = Excel.Workbooks.Add
    Set plan = arq.Sheets(1)
    
    prProcesso.ExportaLinhasEspaider arq, plan
    If Not planGrupoProvContestar Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvContestar, dtDataProvContestar, prProcesso.Advogado)
    If Not planGrupoProvSubsidios Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvSubsidios, dtDataProvSubsidios, prProcesso.Preposto)
    
    rngCelula.Offset(0, 1).Formula = "Inserido no S�sifo"
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
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", � necess�rio que o Internet Explorer esteja aberto na p�gina inicial do Projudi BA, sem logar em nenhum usu�rio, para continuar. Abra a p�gina do Projudi, " & _
'            "saia de qualquer login e rode a fun��o novamente.", vbCritical + vbOKOnly, "S�sifo - Internet Explorer n�o encontrado"
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
'    ' Se deu certo, insere na mem�ria do S�sifo (PODE DAR ERRO: SE N�O CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
'    Set arq = Excel.Workbooks.Add
'    Set plan = arq.Sheets(1)
'
'    prProcesso.ExportaLinhasEspaider arq, plan
'
'    rngCelula.Offset(0, 1).Formula = "Inserido no S�sifo"
'    FecharExplorerZerarVariaveis(IE)
'
'    rngCelula.Offset(1, 0).Select
'
'End Sub

Function PegarSenhaAcesso(rngRange As Excel.Range) As String
''
'' Retorna a string contida na primeira c�lula da range passada como par�metro -- ou, se n�o houver, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strSenha As String
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strSenha = Trim(rngCelula.text)
    
    ' Se n�o houver c�lula no espa�o enviado, ou se estiver vazia, pergunta a senha.
    If rngCelula Is Nothing Or rngCelula.text = "" Then
        strSenha = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe a senha de acesso ao processo do Projudi a cadastrar", "S�sifo - Cadastrar processo"))
    End If
    
    PegarSenhaAcesso = strSenha
    
End Function

Function DescobrirPerfilLogadoPje1gTrt5(oChrome As Selenium.ChromeDriver) As String
''
'' Descobre o perfil do usu�rio logado no PJe
''
    Dim oSelect As Selenium.WebElement, oOption As Selenium.WebElement
    Dim strTexto As String
    
    ' Caso esteja na p�gina de login
    If InStr(1, oChrome.URL, sfUrlPJe1gTrt5Login) <> 0 Then
        DescobrirPerfilLogadoPje1gTrt5 = "deslogado, p�gina de login"
    
    ' Caso esteja no painel do usu�rio logado
    ElseIf InStr(1, oChrome.URL, sfUrlPJe1gTrt5Painel) <> 0 Then
        Set oSelect = oChrome.FindElementById("papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao")
        
        For Each oOption In oSelect.FindElementsByTag("option")
            If oOption.IsSelected = True Then
                strTexto = oOption.text
                Exit For
            End If
        Next oOption
        
        DescobrirPerfilLogadoPje1gTrt5 = Right(strTexto, Len(strTexto) - InStr(1, strTexto, ")/") - 1)
    
    ' Caso esteja em outra p�gina (como p�gina de processo)
    Else
        DescobrirPerfilLogadoPje1gTrt5 = "p�gina sem identifica��o de login"
        
    End If
    
End Function

Function RealizarBuscaAcervoPJe1gTrt5(oChrome As Selenium.ChromeDriver, strNumeroProcesso As String) As Selenium.ChromeDriver
''
'' Abre o acervo e busca o processo
''
Dim divMenu As Selenium.WebElement, inpBusca As Selenium.WebElement
Dim janCont As Selenium.Window

    ' Carrega p�gina do acervo
    oChrome.FindElementById("j_id179:j_id191").Click
    
    ' Preenche o n�mero do processo
    Set inpBusca = oChrome.FindElementById("formLocCaix:decosuggestProcessoAdvogadoProc:suggestProcessoAdvogadoProc")
    Do
        inpBusca.SendKeys strNumeroProcesso
    Loop Until inpBusca.Value = strNumeroProcesso
    
    ' Aguarda aparecer a caixinha com sugest�es de n�mero e clica nela. COLOCAR TIMEOUT
    Do
        DoEvents
        Set divMenu = oChrome.FindElementById("formLocCaix:decosuggestProcessoAdvogadoProc:sugsuggestProcessoAdvogadoProc")
    Loop Until InStr(1, divMenu.Attribute("style"), "visibility: visible") <> 0
    oChrome.FindElementByClass("richfaces_suggestionSelectValue").Click
    
    ' Clica no bot�o Localizar
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
'' Abre a p�gina de buscas, conforme perfil logado
''

'    If strPerfilLogado = "Outro" Then
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", � necess�rio estar logado num perfil de parte, advogado ou representante. Fa�a login no Projudi de " & _
'            "um desses perfis e rode a fun��o novamente.", vbCritical + vbOKOnly, "S�sifo - Internet Explorer n�o encontrado"
'        Set CarregarPaginaBusca = Nothing
'        Exit Function
'    End If
    
    ' Carrega p�gina de busca
    oChrome.get sfUrlPJe1gTrt5Busca
    
    Set CarregarPaginaBuscaPJe1gTrt5 = oChrome
    
End Function

Function AbrirProcessoPJe1gTrt5(ByVal strNumeroCNJ As String, ByRef oChrome As Selenium.ChromeDriver) As Selenium.ChromeDriver
''
'' Retorna o objeto InternetExplorer com a p�gina principal do processo strNumeroCNJ. Deve haver uma sess�o do Internet Explorer aberta
''  e logada no PJe. Em caso de zero ou m�ltiplos processos encontrados, ou de cancelamento ao mostrar Recaptcha, mostra mensagem de erro
''  e retorna 'Nothing'.
'' FALTA LIDAR COM O ERRO DE N�O ESTAR LOGADO!!!!!!!
''
    
    ' AJUSTAR
    'If DocHTML.Title = "Sistema CNJ - A sess�o expirou" Then
    '    PegarLinkProcessoProjudi = "Sess�o expirada"
    '    Exit Function
    'End If
    
    If oChrome.FindElementById("mat-input-0") Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a p�gina do PJe parece estar com alguma inconsist�ncia. N�o consigo acess�-la. " & _
            "Suplico que verifique o que ocorreu e, se tudo estiver normal, tente novamente", vbCritical + vbOKOnly, "S�sifo - Erro no PJe"
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
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", surgiu um captcha na janela do PJe. Pe�o que o resolva e clique em ""Repetir"".", _
                vbExclamation + vbRetryCancel, "S�sifo - Necess�rio resolver captcha") = vbCancel Then
                Set AbrirProcessoPJe1gTrt5 = Nothing
                FecharExplorerZerarVariaveis oChrome
                Exit Function
            End If
    '    End If
    'End If
    
    ' Lidar com erro na busca do processo
    If InStr(1, oChrome.FindElementById("painel-erro").Attribute("style"), "display: block") <> 0 Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", n�o foi encontrado nenhum processo com o n�mero fornecido. Rogo que confira se n�o est� em " & _
            "segredo de justi�a e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Processo n�o encontrado"
        Set AbrirProcessoPJe1gTrt5 = Nothing
        FecharExplorerZerarVariaveis oChrome
        Exit Function
    End If
    
    ' Caso tenha aberto
    Set AbrirProcessoPJe1gTrt5 = oChrome
    
End Function

Sub PegaInformacoesProcessoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver)
''
'' Faz a coleta dos dados do processo espec�ficos do PJe1g TRT5 e armazena na vari�vel global prProcesso
''
    Dim tbPartes As Selenium.WebElement
    Dim contOutrasPartes As OutroParticipante
    Dim varCont As Variant
    Dim Cont As Integer
    Dim arrStrAudiencias() As String
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a provid�ncia de agendar pautista
    
    ''''''''''''''''''''''''''''''''''''
    ''' Carregar totalmente a p�gina '''
    ''''''''''''''''''''''''''''''''''''
    
    'varCont = CarregarTodasAsPaginasAndamentosProcessoPje1gTrt5(oChrome)
    
    'If varCont <> "Sucesso" Then
    '    MsgBox sisifoembasafuncoes.DeterminarTratamento & ", a p�gina do processo n�o p�de ser totalmente aberta. Favor limpar o cache e tentar novamente.", _
    '        vbCritical + vbOKOnly, "S�sifo - Erro no carregamento do processo"
    '    Set prProcesso = Nothing
    '    Exit Sub
    'End If
    
    
    ''''''''''''''''''''''''''
    ''' N�mero do processo '''
    ''''''''''''''''''''''''''
    
    prProcesso.NumeroProcesso = PegaNumeroPJe1gTrt5(oChrome)
    
    ' Confere se o processo j� est� na planilha
    If Not sfCadProcessos.Cells().Find(prProcesso.NumeroProcesso) Is Nothing Then
        Do
        Loop Until MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo j� existe na planilha! Inclus�o cancelada. Descartados os dados." & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf & _
        "Clique em 'Cancelar' e insira o pr�ximo processo.", vbCritical + vbOKCancel, "S�sifo - Processo repetido") = vbCancel
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Cita��o, respons�vel regressivo, ju�zo,   '''
    ''' valor da causa, ju�zo, tipo de a��o, rito '''
    '''''''''''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.Sistema = sfPJe1g
    prProcesso.Citacao = Date
    prProcesso.ResponsavelRegressivo = ""
    prProcesso.ValorCausa = PegaValorPje1gTrt5(oChrome)
    
    prProcesso.Juizo = PegaJuizoPje1gTrt5(oChrome)
    If prProcesso.Juizo = "Ju�zo n�o cadastrado" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.TipoAcao = PegaTipoAcaoPje1gTrt5(oChrome)
    If prProcesso.TipoAcao = "Rito n�o cadastrado no S�sifo" Then
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
            'Pessoa jur�dica
            contOutrasPartes.CNPJParte = varCont(2, Cont)
            contOutrasPartes.CPFParte = 0
            contOutrasPartes.TipoParte = 2
        Else
            'Pessoa f�sica ou n�o cadastrado.
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
    ''' Outros r�us '''
    '''''''''''''''''''
    
    ' Adiciona outros R�us, se for o caso.
    varCont = PegaPartesPje1gTrt5(oChrome, sfReu)
    If Not IsEmpty(varCont) Then
        For Cont = 1 To UBound(varCont, 2)
            Set contOutrasPartes = New OutroParticipante
            contOutrasPartes.NomeParte = varCont(1, Cont)
            If InStr(1, varCont(2, Cont), "/") <> 0 Then
                'Pessoa jur�dica
                contOutrasPartes.CNPJParte = varCont(2, Cont)
                contOutrasPartes.CPFParte = 0
                contOutrasPartes.TipoParte = 2
            Else
                'Pessoa f�sica ou n�o cadastrado
                contOutrasPartes.CPFParte = varCont(2, Cont)
                contOutrasPartes.CNPJParte = 0
                contOutrasPartes.TipoParte = 1
            End If
        
            contOutrasPartes.ClasseParte = 1
            contOutrasPartes.CondicaoParte = "R�u"
            prProcesso.OutrosParticipantes.Add contOutrasPartes
        Next Cont
    End If
    
    '''''''''''''''''''''''''''''''
    ''' Andamento de refer�ncia '''
    ''' (audi�ncia ou cita��o)  '''
    '''''''''''''''''''''''''''''''
    
    arrStrAudiencias = PegaDataeTipoAudienciaPje1gTrt5(oChrome)
    prProcesso.DataAndamento = arrStrAudiencias(1)
    
    If arrStrAudiencias(1) = "0" Then
        prProcesso.NomeAndamento = "Cita��o"
    Else
        prProcesso.NomeAndamento = "Audi�ncia de " & arrStrAudiencias(2)
    End If
        
End Sub

Function PegaNumeroPJe1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strCont1 As String, strCont2 As String
    Dim intCont As Integer
    
    strCont1 = oChrome.FindElementById("processoIdentificadorDiv").text
    
    ' Deixa apenas n�meros, h�fens e pontos
    For intCont = 1 To Len(strCont1) Step 1
        If IsNumeric(Mid(strCont1, intCont, 1)) Or Mid(strCont1, intCont, 1) = "." Or Mid(strCont1, intCont, 1) = "-" Then
            strCont2 = strCont2 & Mid(strCont1, intCont, 1)
        End If
    Next intCont
    
    'Apaga o que n�o for n�mero do in�cio
    strCont1 = strCont2
    Do Until IsNumeric(Left(strCont1, 1))
        strCont1 = Mid(strCont1, 2)
    Loop
    
    'Apaga o que n�o for n�mero do final
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
        InputBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu n�o conhe�o o ju�zo em que o processo tramita. Rogo que o cadastre em minha mem�ria (a reda��o " & _
            "do PJe pode ser copiada abaixo) e tente novamente.", "S�sifo - Ju�zo desconhecido", Trim(strJuizo)
        PegaJuizoPje1gTrt5 = "Ju�zo n�o cadastrado"
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
        PegaTipoAcaoPje1gTrt5 = "Rito n�o cadastrado no S�sifo"
    End If
    
End Function

Function PegaRitoPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver, strTipoAcao As String) As String
    
    Dim rngCont As Range
    
    Set rngCont = cfPjeTiposAcaoRitos.Cells().Find(What:=strTipoAcao, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        PegaRitoPje1gTrt5 = rngCont.Offset(0, 2).Formula
    Else
        PegaRitoPje1gTrt5 = "Rito n�o cadastrado"
    End If
    
End Function

Function PegaPartesPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver, strTipoParte As String) As Variant
''
'' Retorna uma matriz com as partes de um polo do processo.
'' PegaPartesPje(1, N) = Nome da parte na posi��o N da matriz.
'' PegaPartesPje(2, N) = CPF/CNPJ da parte na posi��o N da matriz.
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
        ' Condi��o comum para pular (se n�o for "RECLAMANTE" ou "RECLAMADO")
        strCont = oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & ":tb']/tr[" & intContPartes & "]/td[2]").text
        If strCont <> "RECLAMANTE" And strCont <> "RECLAMADO" Then bolPular = True
        
        ' Condi��o espec�fica de r�u para pular
        strCont = oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & ":tb']/tr[" & intContPartes & "]/td[1]").text
        If strTipoParte = sfReu And (InStr(1, strCont, "EMPRESA BAIANA DE AGUA") <> 0 Or InStr(1, strCont, "EMBASA") <> 0) Then bolPular = True
        
        If bolPular = False Then
            intQtdPartes = intQtdPartes + 1
            ReDim Preserve arrIntCont(1 To intQtdPartes)
            arrIntCont(intQtdPartes) = intContPartes
        End If
    Next intContPartes
    
    ' Se for cadastro de r�u e houver 0 partes, isso significa que � s� a Embasa, portanto, vamos pular.
    If strTipoParte = sfReu And intQtdPartes = 0 Then Exit Function
    
    ' Redimensiona a array
    ReDim arrStrPartes(1 To 2, 1 To intQtdPartes)
    
    ' Para cada parte, busca Nome e CPF (advogado pressup�e-se comum). 1 = Nome, 2 = CPF
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
        
        ' Checa se CPF est� cadastrado...
        If InStr(1, arrStrPartes(2, intContPartes), "N�o cadastrado##") <> 0 Then ' N�o cadastrado (era uma express�o do Projudi, n�o sei se vai ter no PJe
            strMsgErro = "o CPF/CNPJ da parte n�o foi cadastrado no PJe"
        ElseIf InStr(1, arrStrPartes(2, intContPartes), "N�o dispon�vel##") <> 0 Then ' N�o dispon�vel (era uma express�o do Projudi, n�o sei se vai ter no PJe
            strMsgErro = "o CPF/CNPJ da parte n�o est� dispon�vel"
        End If
        
        If strMsgErro <> "" Then
PedirCPF:
            arrStrPartes(2, intContPartes) = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", " & strMsgErro & ". Rogo que busque a Peti��o Inicial " & _
                "e informe o CPF ou CNPJ correto. Se n�o houver, deixe em branco.", "S�sifo - Informar CPF/CNPJ da parte"))
            If arrStrPartes(2, intContPartes) <> "" Then
                arrStrPartes(2, intContPartes) = ValidarCPFCNPJ(arrStrPartes(2, intContPartes))
                If Not IsNumeric(Left(arrStrPartes(2, intContPartes), 1)) Then
                    strMsgErro = arrStrPartes(2, intContPartes)
                    GoTo PedirCPF
                End If
            End If
        End If
        
        If arrStrPartes(2, intContPartes) = "" Then 'Sem CPF/CNPJ
            arrStrPartes(2, intContPartes) = "N�o Cadastrado " & PegaCodParteSemCPF
        ElseIf InStr(1, arrStrPartes(2, intContPartes), "/") = 0 Then ' CPF - 14 caracteres
            arrStrPartes(2, intContPartes) = Left(arrStrPartes(2, intContPartes), 14)
        End If ' CNPJ n�o precisa de tratamento
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
    
    If strAdvAutor = "Nenhum advogado cadastrado." Then ' A mensagem � a do Projudi. N�o sei como � no PJe.
        strAdvAutor = ""
    Else
        strAdvAutor = UCase(Trim(strAdvAutor))
    End If
    
    PegaAdvAutorPje1gTrt5 = strAdvAutor
    
End Function

Function PegaDataeTipoAudienciaPje1gTrt5(ByRef oChrome As Selenium.ChromeDriver) As String()
''
'' Busca a �ltima audi�ncia com status "designada". Retorna um vetor com as seguintes informa��es:
'' (1) = Data audi�ncia ou "0", caso n�o encontre.
'' (2) = Tipo de audi�ncia ou "", caso n�o encontre.
''
    Dim intCont As Integer, intQtdAudiencias As Integer
    Dim arrStrResultado(1 To 2) As String
    Dim divCont As Selenium.WebElement
    Dim tbAudiencias As Selenium.WebElement
    Dim lngTimerInicio As Long
    
    ' Abre a aba de audi�ncias
    oChrome.FindElementById("tabProcessoAudiencia_shifted").Click
    
    ' Aguarda o carregamento (ass�ncrono) das informa��es da tabela, por at� 10 segundos
AguardaTabelaAudiencias:
    lngTimerInicio = Timer
    Do
        DoEvents
        If Timer > lngTimerInicio + 10 Then
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o carregamento da tabela de audi�ncias parece estar demorando " & _
                "mais do que o normal. Deseja continuar esperando? Caso n�o espere, a data da audi�ncia poder� ser alterada manualmente " & _
                "no formul�rio, ou ap�s o cadastramento, no Espaider.") = vbYes Then
                GoTo AguardaTabelaAudiencias
            Else
                Exit Do
            End If
        End If
    Loop While oChrome.FindElementsByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr").Count = 0
    
    ' Itera as linhas da tabela de audi�ncias, verificando datas e se est� ativa
    For intCont = 1 To oChrome.FindElementsByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr").Count Step 1
        ' Confere se h� mais de uma audi�ncia com status "designada". Se houver, alerta e mant�m a mais antiga.
        If LCase(oChrome.FindElementByXPath("//*[@id='processoConsultaAudienciaGridList:tb']/tr[" & intCont & "]/td[4]").text) = "designada" Then
            If arrStrResultado(1) <> "" Then
                MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", parece haver mais de uma audi�ncia ativa designada para este processo. Pe�o que informe o n�mero deste processo a C�sar para verifica��o.", vbCritical + vbOKOnly, "S�sifo - Mais de uma audi�ncia ativa"
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
    
    ' Se n�o houver andamento "Audi�ncia" ou n�o se encaixar em nenhuma hip�tese anterior, retorna 0.
    If arrStrResultado(1) = "" Then
        arrStrResultado(1) = "0"
        arrStrResultado(2) = ""
    Else
        PegaDataAudienciaPje1gTrt5 = arrStrResultado
    End If

End Function



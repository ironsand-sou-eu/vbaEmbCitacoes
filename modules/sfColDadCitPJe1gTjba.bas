Attribute VB_Name = "sfColDadCitPJe1gTjba"
Option Explicit

Sub CadastrarProcessoIndividualPje1gTjba(strNumeroProcesso As String, ByRef rngCelula As Excel.Range)
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
        oChrome.get sfUrlPJe1gTjbaLogin
    Else
        oChrome.get sfUrlPJe1gTjbaPainel
    End If
    
    'Descobrir se está logado e, caso esteja, se o perfil é de parte ou advogado.
VerificarLogin:
    strPerfilLogado = DescobrirPerfilLogadoPje1gTjba(oChrome)
    
    Select Case LCase(strPerfilLogado)
    Case "deslogado, página de login"
        oChrome.Window.Activate
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", rogo que faça o login na janela do Chrome e depois clique em ""OK"".", vbCritical + vbOKOnly, "Sísifo - Aguardando login no Chrome"
        GoTo VerificarLogin
        
    Case "página sem identificação de login"
        oChrome.get sfUrlPJe1gTjbaPainel
        GoTo VerificarLogin
        
    Case "procurador", "advogado", "procurador/gestor"
        ' Avança normalmente, nenhuma ação a realizar.
    
    Case Else
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro no login. Suplico que tente novamente, desta vez fazendo login de procurador da empresa ou de advogado.", _
            vbOKOnly + vbCritical, "Sísifo - Erro no login"
        FecharChromeZerarVariaveis oChrome: Exit Sub
    End Select
    
    ' Carregar página de busca
    Set oChrome = CarregarPaginaBuscaPJe1gTjba(oChrome, prProcesso.Tribunal)
    If oChrome Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    
    'Abrir processo pelo número CNJ
    Set oChrome = AbrirProcessoPJe1gTjba(strNumeroProcesso, oChrome)
    If oChrome Is Nothing Then FecharChromeZerarVariaveis oChrome: Exit Sub
    
    ' Pega os dados do processo
    PegaInformacoesProcessoPje1gTjba oChrome
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

Sub CadastrarProcessoIndividualPjeSemSelenium(strNumeroProcesso As String, ByRef rngCelula As Excel.Range)
''
'' Com o PJe aberto e logado no Internet Explorer, busca um processo e o cadastra
''
    
    Dim IE As InternetExplorer, DocHTML As HTMLDocument
    Dim oManage As Selenium.Manage
    Dim strCont As String
    Dim strPerfilLogado As String
    Dim dtDataProvContestar As Date, dtDataProvSubsidios As Date
    Dim planGrupoProvContestar As Excel.Worksheet, planGrupoProvSubsidios As Excel.Worksheet
    Dim arq As Workbook, plan As Excel.Worksheet
    
    ' Procurar Internet Explorer aberto no PJe
    Select Case prProcesso.Tribunal
    Case sfTjba
        Set IE = SisifoEmbasaFuncoes.RecuperarIE("pje.tjba.jus.br")
    Case sfTRT5
        Set IE = SisifoEmbasaFuncoes.RecuperarIE("pje.trt5.jus.br")
    End Select
    
    If IE Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", é necessário que o Internet Explorer esteja aberto na página do PJe de 1ª instância do " & _
        prProcesso.Tribunal & " para continuar. Abra-o, faça login no PJe e rode a função novamente.", vbCritical + vbOKOnly, "Sísifo - Internet Explorer não encontrado"
        GoTo FinalizaFechaIE
    End If
    
    IE.Visible = True
    
    'Descobrir se o perfil é de parte ou advogado. Se for outro, manda relogar.
    Select Case prProcesso.Tribunal
    Case sfTjba
        strPerfilLogado = DescobrirPerfilLogadoPjeTjba(IE.document)
    Case sfTRT5
        strPerfilLogado = DescobrirPerfilLogadoPjeTrt5(IE.document)
    End Select
    
    ' Carregar página de busca
    Set IE = CarregarPaginaBuscaPJe(prProcesso.Tribunal)
    If IE Is Nothing Then GoTo FinalizaFechaIE
    Set DocHTML = IE.document
    
    'Pegar link pelo número CNJ
    Set IE = AbrirProcessoPJe(strNumeroProcesso, IE, DocHTML)
    If IE Is Nothing Then GoTo FinalizaFechaIE
    Set DocHTML = IE.document
    
    ' Pega os dados do processo
    PegaInformacoesProcessoPje IE, DocHTML
    If prProcesso Is Nothing Then GoTo FinalizaFechaIE
    PegaInformacoesProcessoGeral dtDataProvContestar, dtDataProvSubsidios, planGrupoProvContestar, planGrupoProvSubsidios
    
    ' Se deu erro, devolve o erro
    If prProcesso Is Nothing Then GoTo FinalizaFechaIE
    If prProcesso.MensagemErro <> "" Then
FinalizaFechaIE:
        FecharExplorerZerarVariaveis IE
        Exit Sub
    End If
    
    ' Se deu certo, insere na memória do Sísifo (PODE HAVER O SEGUINTE PROBLEMA: SE NÃO CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
    Set arq = Excel.Workbooks.Add
    Set plan = arq.Sheets(1)
    
    prProcesso.ExportaLinhasEspaider arq, plan
    If Not planGrupoProvContestar Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvContestar, dtDataProvContestar, prProcesso.Advogado)
    If Not planGrupoProvSubsidios Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvSubsidios, dtDataProvSubsidios, prProcesso.Preposto)
    
    rngCelula.Offset(0, 1).Formula = "Inserido no Sísifo"
    
    FecharExplorerZerarVariaveis IE
    
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

Function DescobrirPerfilLogadoPje1gTjba(oChrome As Selenium.ChromeDriver) As String
''
'' Descobre o perfil do usuário logado no PJe
''
    Dim oSpan As Selenium.WebElement
    Dim strTexto As String
    
    ' Caso esteja na página de login
    If InStr(1, oChrome.URL, sfUrlPJe1gTjbaLogin) <> 0 Then
        DescobrirPerfilLogadoPje1gTjba = "deslogado, página de login"
    
    ' Caso esteja no painel do usuário logado
    ElseIf InStr(1, oChrome.URL, sfUrlPJe1gTjbaPainel) <> 0 Then
        Set oSpan = oChrome.FindElementsByClass("hidden-xs nome-sobrenome tip-bottom")(0)
        strTexto = oSpan.Attribute("data-original-title")
        DescobrirPerfilLogadoPjeTjba = Right(strTexto, Len(strTexto) - InStr(1, strTexto, " / ") - 2)
    
    ' Caso esteja em outra página (como página de processo)
    Else
        DescobrirPerfilLogadoPje1gTjba = "página sem identificação de login"
        
    End If
    
End Function

Function DescobrirPerfilLogadoPjeTjbaSemSelenium(DocHTML As HTMLDocument) As String
''
'' Descobre o perfil do usuário logado no PJe
''
    Dim oSpan As HTMLSpanElement
    Dim strTexto As String
    
    Set oSpan = DocHTML.getElementsByClassName("hidden-xs nome-sobrenome tip-bottom")(0)
    strTexto = oSpan.getAttribute("data-original-title")
    DescobrirPerfilLogadoPjeTjba = Right(strTexto, Len(strTexto) - InStr(1, strTexto, ") /") - 3)
    
End Function

Function DescobrirPerfilLogadoPjeTrt5SemSelenium(DocHTML As HTMLDocument) As String
''
'' DescobreDescobre o perfil do usuário logado no PJe
''
    Dim oSelect As HTMLSelectElement, oOption As HTMLOptionElement
    Dim strTexto As String
    
    Set oSelect = DocHTML.getElementById("papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao")
    
    For Each oOption In oSelect.Children
        If oOption.Selected = True Then
            strTexto = oOption.innerText
            GoTo PerfilEncontrado
        End If
    Next oOption
    
PerfilEncontrado:
    
    DescobrirPerfilLogadoPjeTrt5 = Right(strTexto, Len(strTexto) - InStr(1, strTexto, ")/") - 1)
    
End Function

Function CarregarPaginaBuscaPJe1gTjba(oChrome As Selenium.ChromeDriver) As Selenium.ChromeDriver
''
'' Abre a página de buscas, conforme perfil logado
''
    Dim strUrlDestino As String
    
    ' Carrega página de busca
    oChrome.get sfUrlPJe1gTjbaBusca
    
    Set CarregarPaginaBuscaPJe1gTjba = oChrome
    
End Function

Function CarregarPaginaBuscaPJeSemSelenium() As InternetExplorer  '(strPerfilLogado As String) As InternetExplorer
''
'' Abre nova janela do Internet Explorer na página de buscas, conforme perfil logado
''
    
    Dim IE As InternetExplorer
    Dim DocHTML As HTMLDocument
    Dim strCont As String, strUrlDestino As String
    
'    If strPerfilLogado = "Outro" Then
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", é necessário estar logado num perfil de parte, advogado ou representante. Faça login no Projudi de " & _
'            "um desses perfis e rode a função novamente.", vbCritical + vbOKOnly, "Sísifo - Internet Explorer não encontrado"
'        Set CarregarPaginaBusca = Nothing
'        Exit Function
'    End If
    
    Select Case prProcesso.Tribunal
    Case sfTjba
        strUrlDestino = sfUrlPJe1gTjbaBusca
    Case sfTRT5
        strUrlDestino = sfUrlPJe1gTrt5Busca
    End Select
    
    ' Carrega página de busca
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.navigate strUrlDestino
    Set IE = SisifoEmbasaFuncoes.RecuperarIE(strUrlDestino)
    
    ' Aguarda carregar
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    Do
        DoEvents
        strCont = IE.document.URL
    Loop Until strCont = strUrlDestino
    
    Set DocHTML = IE.document
    'Set DocHTML = DocHTML.getElementsByName("mainFrame")(0).contentDocument.getElementsByName("userMainFrame")(0).contentDocument
    
    Set CarregarPaginaBuscaPJe = IE
    
End Function

Function AbrirProcessoPJe1gTjba(ByVal strNumeroCNJ As String, ByRef oChrome As Selenium.ChromeDriver) As Selenium.ChromeDriver
''
'' Retorna o objeto InternetExplorer com a página principal do processo strNumeroCNJ. Deve haver uma sessão do Internet Explorer aberta
''  e logada no PJe. Em caso de zero ou múltiplos processos encontrados, ou de cancelamento ao mostrar Recaptcha, mostra mensagem de erro
''  e retorna 'Nothing'.
'' FALTA LIDAR COM O ERRO DE NÃO ESTAR LOGADO!!!!!!!
''
    Dim strSequencialProcesso As String, strDigitoProcesso As String, strAnoProcesso As String, strVaraProcesso As String, strCont As String
    
    ' AJUSTAR
    'If DocHTML.Title = "Sistema CNJ - A sessão expirou" Then
    '    PegarLinkProcessoProjudi = "Sessão expirada"
    '    Exit Function
    'End If
    
    strSequencialProcesso = Left(strNumeroCNJ, 7)
    strDigitoProcesso = Mid(strNumeroCNJ, 9, 2)
    strAnoProcesso = Mid(strNumeroCNJ, 12, 4)
    strVaraProcesso = Right(strNumeroCNJ, 4)
    
    If oChrome.FindElementById("fPP:numeroProcesso:numeroSequencial") Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a página do PJe parece estar com alguma inconsistência. Não consigo acessá-la. " & _
            "Suplico que verifique o que ocorreu e, se tudo estiver normal, tente novamente", vbCritical + vbOKOnly, "Sísifo - Erro no PJe"
        Set AbrirProcessoPJe1gTjba = Nothing
        Exit Function
    End If
    
    ' Preenche o número e clica em Pesquisar
    Do
        oChrome.FindElementById("fPP:numeroProcesso:numeroSequencial").Clear
        oChrome.FindElementById("fPP:numeroProcesso:numeroSequencial").SendKeys strSequencialProcesso
    Loop Until oChrome.FindElementById("fPP:numeroProcesso:numeroSequencial").text = strSequencialProcesso
    Do
        oChrome.FindElementById("fPP:numeroProcesso:numeroDigitoVerificador").Clear
        oChrome.FindElementById("fPP:numeroProcesso:numeroDigitoVerificador").SendKeys strDigitoProcesso
    Loop Until oChrome.FindElementById("fPP:numeroProcesso:numeroDigitoVerificador").text = strSequencialProcesso
    Do
        oChrome.FindElementById("fPP:numeroProcesso:Ano").Clear
        oChrome.FindElementById("fPP:numeroProcesso:Ano").SendKeys strAnoProcesso
    Loop Until oChrome.FindElementById("fPP:numeroProcesso:Ano").text = strSequencialProcesso
    Do
        oChrome.FindElementById("fPP:numeroProcesso:NumeroOrgaoJustica").Clear
        oChrome.FindElementById("fPP:numeroProcesso:NumeroOrgaoJustica").SendKeys strVaraProcesso
    Loop Until oChrome.FindElementById("fPP:numeroProcesso:NumeroOrgaoJustica").text = strSequencialProcesso
    
    oChrome.FindElementById("fPP:searchProcessos").Click
    
    ' Aguardar carregamento (assíncrono) dos processos
    Do
        strCont = Trim(oChrome.FindElementByXPath("//*[@id='fPP:processosTable:j_id431']/div[2]/span").text)
    Loop While strCont = "resultados encontrados."
    
    ' Lidar com zero resultados ou com múltiplos resultados
    If strCont = "0 resultados encontrados." Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", não foi encontrado nenhum processo com o número fornecido. Rogo que confira se não está em " & _
            "segredo de justiça e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Set AbrirProcessoPJe1gTjba = Nothing
        Exit Function
    ElseIf strCont <> "1 resultados encontrados." Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", foram encontrados múltiplos processos com o número fornecido. Rogo que confira o número correto " & _
            "e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Múltiplos processos encontrados"
        Set AbrirProcessoPJe1gTjba = Nothing
        Exit Function
    End If
    
    ' Em caso de um resultado só
    oChrome.FindElementByXPath("//*[@id='fPP:processosTable:tb']/tr/td[1]/a[1]").Click
    Set AbrirProcessoPJe1gTjba = oChrome

End Function

Function AbrirProcessoPJeSemSelenium(ByVal strNumeroCNJ As String, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As InternetExplorer
''
'' Retorna o objeto InternetExplorer com a página principal do processo strNumeroCNJ. Deve haver uma sessão do Internet Explorer aberta
''  e logada no PJe. Em caso de zero ou múltiplos processos encontrados, ou de cancelamento ao mostrar Recaptcha, mostra mensagem de erro
''  e retorna 'Nothing'.
'' FALTA LIDAR COM O ERRO DE NÃO ESTAR LOGADO!!!!!!!
''

    Dim strSequencialProcesso As String, strDigitoProcesso As String, strAnoProcesso As String, strVaraProcesso As String, strCont As String
    Dim tbProcessos As HTMLTableSection
    Dim divCaptcha As HTMLDivElement

    ' AJUSTAR
    'If DocHTML.Title = "Sistema CNJ - A sessão expirou" Then
    '    PegarLinkProcessoProjudi = "Sessão expirada"
    '    Exit Function
    'End If
    
    strSequencialProcesso = Left(strNumeroCNJ, 7)
    strDigitoProcesso = Mid(strNumeroCNJ, 9, 2)
    strAnoProcesso = Mid(strNumeroCNJ, 12, 4)
    strVaraProcesso = Right(strNumeroCNJ, 4)
    
    If DocHTML.getElementById("fPP:numeroProcesso:numeroSequencial") Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a página do PJe parece estar com alguma inconsistência. Não consigo acessá-la. " & _
            "Suplico que verifique o que ocorreu e, se tudo estiver normal, tente novamente", vbCritical + vbOKOnly, "Sísifo - Erro no PJe"
        Set AbrirProcessoPJe = Nothing
        FecharExplorerZerarVariaveis IE
        Exit Function
    End If
    
    DocHTML.getElementById("fPP:numeroProcesso:numeroSequencial").Value = strSequencialProcesso
    DocHTML.getElementById("fPP:numeroProcesso:numeroDigitoVerificador").Value = strDigitoProcesso
    DocHTML.getElementById("fPP:numeroProcesso:Ano").Value = strAnoProcesso
    DocHTML.getElementById("fPP:numeroProcesso:NumeroOrgaoJustica").Value = strVaraProcesso
    DocHTML.getElementById("fPP:searchProcessos").Click
    
    'SisifoEmbasaFuncoes.Esperar 1
    On Error GoTo Volta2
Volta2:
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    'Do
    '    DoEvents
    'Loop Until DocHTML.getElementsByTagName("body")(0).Children(0).Children(0).innerText = "Processos Obtidos Por Busca"
    
    Do
        strCont = Trim(DocHTML.getElementById("fPP:processosTable:j_id431").Children(1).innerText)
        'COLOCAR UM TIMEOUT AQUI
        Set divCaptcha = DocHTML.body.Children(DocHTML.body.Children.length - 1)
        If Not divCaptcha Is Nothing Then
            If divCaptcha.Style.visibility = "visible" Then
                If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", surgiu um captcha na janela do PJe. Peço que o resolva e clique em ""Repetir"".", _
                    vbExclamation + vbRetryCancel, "Sísifo - Necessário resolver captcha") = vbCancel Then
                    Set AbrirProcessoPJe = Nothing
                    FecharExplorerZerarVariaveis IE
                    Exit Function
                End If
            End If
        End If
    Loop While strCont = "resultados encontrados."
    
    ' Lidar com zero resultados ou com múltiplos resultados
    If strCont = "0 resultados encontrados." Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", não foi encontrado nenhum processo com o número fornecido. Rogo que confira se não está em " & _
            "segredo de justiça e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Processo não encontrado"
        Set AbrirProcessoPJe = Nothing
        FecharExplorerZerarVariaveis IE
        Exit Function
    ElseIf strCont <> "1 resultados encontrados." Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", foram encontrados múltiplos processos com o número fornecido. Rogo que confira o número correto " & _
            "e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Múltiplos processos encontrados"
        Set AbrirProcessoPJe = Nothing
        FecharExplorerZerarVariaveis IE
        Exit Function
    End If
    
    ' Em caso de um resultado só
    Set tbProcessos = DocHTML.getElementById("fPP:processosTable:tb")
    tbProcessos.getElementsByTagName("a")(0).Click
    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", favor confirmar o acesso ao processo na janela popup que se abriu no Internet Explorer e depois pressionar OK nesta janela.", vbCritical + vbOKOnly, "Sísifo - Confirmar prosseguimento"
    Set AbrirProcessoPJe = SisifoEmbasaFuncoes.RecuperarIE("https://pje.tjba.jus.br/pje-web/Processo/ConsultaProcesso/Detalhe/")
    FecharExplorerZerarVariaveis IE
    
End Function

Sub PegaInformacoesProcessoPje1gTjba(ByRef oChrome As Selenium.ChromeDriver)
''
'' Faz a coleta dos dados do processo específicos do PJe e armazena na variável global prProcesso
''
    Dim divTabelaAndamentos As Selenium.WebElement, divDetalhes As Selenium.WebElement, divPartes As Selenium.WebElement
    Dim contOutrasPartes As OutroParticipante
    Dim varCont As Variant
    Dim Cont As Integer
    Dim arrStrAudiencias() As String
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a providência de agendar pautista
    
    ''''''''''''''''''''''''''''''''''''
    ''' Carregar totalmente a página '''
    ''''''''''''''''''''''''''''''''''''
    
    'varCont = CarregarTodasAsPaginasAndamentosProcessoPje(IE, DocHTML)
    
    'If varCont <> "Sucesso" Then
    '    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a página do processo não pôde ser totalmente aberta. Favor limpar o cache e tentar novamente.", _
    '        vbCritical + vbOKOnly, "Sísifo - Erro no carregamento do processo"
    '    Set prProcesso = Nothing
    '    Exit Sub
    'End If
    
    
    ''''''''''''''''''''''''''
    ''' Número do processo '''
    ''''''''''''''''''''''''''
    
    prProcesso.NumeroProcesso = PegaNumeroPJe1gTjba(DocHTML)
    
    ' Confere se o processo já está na planilha
    If Not sfCadProcessos.Cells().Find(prProcesso.NumeroProcesso) Is Nothing Then
        Do
        Loop Until MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo já existe na planilha! Inclusão cancelada. Descartados os dados." & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf & _
        "Clique em 'Cancelar' e insira o próximo processo.", vbCritical + vbOKCancel, "Sísifo - Processo repetido") = vbCancel
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Citação, responsável regressivo, juízo, valor da '''
    ''' causa, juízo, tipo de ação, rito, data audiência '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.Sistema = sfPJe1g
    prProcesso.Citacao = PegaCitacaoPje1gTjba
    prProcesso.ResponsavelRegressivo = ""
    prProcesso.ValorCausa = PegaValorPje1gTjba(oChrome)
    
    prProcesso.Juizo = PegaJuizoPje1gTjba(oChrome)
    If prProcesso.Juizo = "Juízo não cadastrado" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.TipoAcao = PegaTipoAcaoPje1gTjba(oChrome)
    If prProcesso.TipoAcao = "Rito não cadastrado no Sísifo" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.Rito = PegaRitoPje1gTjba(oChrome, prProcesso.TipoAcao)
    
    ''''''''''''''''''''''''''
    ''' Autores e advogado '''
    ''''''''''''''''''''''''''
    
    ' Adiciona Autores
    varCont = PegaPartesPje1gTjba(oChrome, sfAutor)
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
    
    prProcesso.AdvogadoAdverso = PegaAdvAutorPje1gTjba(oChrome)
    
    '''''''''''''''''''
    ''' Outros réus '''
    '''''''''''''''''''
    
    ' Adiciona outros Réus, se for o caso.
    varCont = PegaPartesPje1gTjba(oChrome, sfReu)
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
    
    arrStrAudiencias = PegaDataeTipoAudienciaPje1gTjba(oChrome)
    prProcesso.DataAndamento = arrStrAudiencias(1)
    
    If arrStrAudiencias(1) = "0" Then
        prProcesso.NomeAndamento = "Citação"
    Else
        prProcesso.NomeAndamento = "Audiência de " & arrStrAudiencias(2)
    End If
        
End Sub

Function CarregarTodasAsPaginasAndamentosProcessoPje(IE As InternetExplorer, DocHTML As HTMLDocument) As String

    Dim intPaginaAtual As Integer, intTotalPaginas As Integer
    Dim sngTimerInicio As Single
    Dim divBarraAndamentos As HTMLDivElement
    
    intPaginaAtual = DocHTML.getElementById("paginaAtual").Value
    intTotalPaginas = DocHTML.getElementById("totalPaginas").Value
    Set divBarraAndamentos = DocHTML.getElementById("divTimeLine:divEventosTimeLine")
    
CarregaEsperarPagina:
    sngTimerInicio = Timer
    While intPaginaAtual < intTotalPaginas
        DocHTML.getElementsByClassName("col-sm-12 text-center")(0).ScrollIntoView
        intPaginaAtual = DocHTML.getElementById("paginaAtual").Value
        intTotalPaginas = DocHTML.getElementById("totalPaginas").Value
        
        If Timer >= sngTimerInicio + 10 Then
            'Perguntar se quer continuar
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", a página parece estar demorando de carregar. Suplico que confira se há algo errado e diga se pretende " & _
                "esperar mais 10 segundos, ou desistir do cadastro deste processo.", vbRetryCancel + vbQuestion, "Sísifo - Demora no carregamento") = vbRetry Then
                GoTo CarregaEsperarPagina
            Else
                CarregarTodasAsPaginasAndamentosProcessoPje = "Desistência por demora"
                Exit Function
            End If
        End If
    Wend
    
    CarregarTodasAsPaginasAndamentosProcessoPje = "Sucesso"
    
End Function

Function PegaNumeroPJe1gTjba(ByRef DocHTML As HTMLDocument) As String
    Dim strNumero As String
    
    strNumero = Left(oChrome.FindElementByXPath("//*[@id='navbar']/ul/li/a[1]"), 25)
    PegaNumeroPJe1gTjba = strNumero
    
End Function

Function PegaCitacaoPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As Date
''
'' TENHO QUE ENCONTRAR O PADRÃO DE ANDAMENTO "CITAÇÃO" DO PJE. Infelizmente, até o momento não há padrão, e não costuma
''   ser lançada citação, e sim retornos de AR ou mandado, sem indicação de que seja citação.
'' Busca o primeiro andamento "Citação lido(a)" cuja observação abaixo contenha os nomes "EMBASA" ou "SANEAMENTO".
'' Encontrando, a data. Não encontrando, retornará a data de hoje. Havendo múltiplas, retornará a mais recente (mais acima).
''
    Dim bolEncontrou As Boolean
    Dim intCont As Integer, intContData As Integer
    Dim strCont As String
    Dim divCont As HTMLDivElement
    
    ' Até achar um padrão, vai ser usada a data do cadastro.
    PegaCitacaoPje1gTjba = Date
    Exit Function
    
    bolEncontrou = False
    
    ' Itera as linhas do div de andamentos para baixo, para encontrar o andamento citação mais recente
    For intCont = 0 To divTabelaAndamentos.getElementsByTagName("div").length - 1 Step 1
        Set divCont = divTabelaAndamentos.getElementsByTagName("div")(intCont)
        If InStr(1, divCont.className, "media interno tipo-") <> 0 Then
                If LCase(divCont.Children(1).Children(0).innerText) = "citação lido(a)" Then
                bolEncontrou = True
                Exit For  'Achou um andamento de citação
            End If
        Else
            If InStr(1, divCont.className, "col-sm-12 text-center") = 0 And InStr(1, divCont.className, "media data") = 0 Then
                '''' TIPO DE DIV DIFERENTE, LIDAR COM ISSO
            End If
        End If
    Next intCont
    
    ' Se tiver encontrado a data, itera as linhas do div de andamentos para cima, para encontrar a data da citação
    If bolEncontrou Then
        For intContData = intCont - 1 To 0 Step -1
            Set divCont = divTabelaAndamentos.getElementsByTagName("div")(intContData)
            If InStr(1, divCont.className, "media data") = 0 _
                And divCont.Children.length > 1 Then
                
                strCont = Trim(divCont.Children(0).Children(0).innerText)
                PegaCitacaoPje1gTjba = Replace(strCont, " ", "/")
                Exit Function
            End If
        Next intContData
    End If
    
    ' Se não houver andamento "Citação lido(a)" , retorna a data de hoje.
    PegaCitacaoPje1gTjba = Date
    
End Function

Function PegaDataeTipoAudienciaPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As String
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
    oChrome.FindElementById("navbar:linkAbaAudiencia").Click
    
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

Function PegaJuizoPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strJuizo As String
    Dim rngCont As Excel.Range
    
    strJuizo = oChrome.FindElementByXPath("//*[@id='maisDetalhes']/div[1]/dl/dd").text
    
    Set rngCont = cfJuizos.Cells().Find(What:=strJuizo, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        InputBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu não conheço o juízo em que o processo tramita. Rogo que o cadastre em minha memória (a redação " & _
            "do PJe pode ser copiada abaixo) e tente novamente.", "Sísifo - Juízo desconhecido", Trim(strJuizo)
        PegaJuizoPje1gTjba = "Juízo não cadastrado"
    Else
        PegaJuizoPje1gTjba = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaValorPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As Currency
    
    PegaValorPje1gTjba = Trim(CCur(oChrome.FindElementByXPath("//*[@id='maisDetalhes']/dl/dd[6]").text))
    
End Function

Function PegaTipoAcaoPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strTipoAcaoPje As String
    Dim rngCont As Range
    Dim regEx As New RegExp
    
    strTipoAcaoPje = oChrome.FindElementByXPath("//*[@id='maisDetalhes']/dl/dd[1]").text
    
    With regEx
        .Global = True
        .IgnoreCase = False
        .MultiLine = True
        .Pattern = " \([0-9]+\)"
        strTipoAcaoPje = Trim(.Replace(strTipoAcaoPje, ""))
    End With
    
    Set rngCont = cfPjeTiposAcaoRitos.Cells().Find(What:=strTipoAcaoPje, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        PegaTipoAcaoPje1gTjba = rngCont.Offset(0, 1).Formula
    Else
        PegaTipoAcaoPje1gTjba = "Rito não cadastrado"
    End If
    
End Function

Function PegaRitoPje1gTjba(ByRef oChrome As Selenium.ChromeDriver, strTipoAcao As String) As String
    
    Dim rngCont As Range
    
    Set rngCont = cfPjeTiposAcaoRitos.Cells().Find(What:=strTipoAcao, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        PegaRitoPje1gTjba = rngCont.Offset(0, 2).Formula
    Else
        PegaRitoPje1gTjba = "Rito não cadastrado"
    End If
    
End Function

Function PegaPartesPje1gTjba(ByRef oChrome As Selenium.WebElement, strTipoParte As String) As Variant
''
'' Retorna uma matriz com as partes de um polo do processo.
'' PegaPartesPje(1, N) = Nome da parte na posição N da matriz.
'' PegaPartesPje(2, N) = CPF/CNPJ da parte na posição N da matriz.
''
    Dim intContPartes As Integer, intQtdPartes As Integer, arrIntCont() As Integer
    Dim strTipoLista As String, strMsgErro As String, strTipoDocumento As String, strCont As String, arrStrPartes() As String
    Dim bolPular As Boolean
    
    Select Case strTipoParte
    Case sfAutor
        strTipoLista = "PoloAtivo"
    Case sfReu
        strTipoLista = "PoloPassivo"
    End Select
    
    For intContPartes = 1 To oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & "']/table[1]/tbody/tr").Count Step 1
        ' Condição específica de réu para pular
        strCont = oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & "']/table[1]/tbody/tr[" & intContPartes & "]/td/a/span").text
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
        strCont = Trim(oChrome.FindElementsByXPath("//*[@id='" & strTipoLista & "']/table[1]/tbody/tr[" & arrIntCont(intContPartes) & "]/td/a/span").text)
        
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
    
    PegaPartesPje1gTjba = arrStrPartes()
    
End Function

Function PegaAdvAutorPje1gTjba(ByRef oChrome As Selenium.ChromeDriver) As String
    
    Dim strAdvAutor As String
    
    strAdvAutor = Trim(oChrome.FindElementByXPath("//*[@id='poloAtivo']/table/tbody/tr/td/ul").text)

    If strAdvAutor = "Nenhum advogado cadastrado." Then ' A mensagem é a do Projudi. Não sei como é no PJe.
        strAdvAutor = ""
    Else
        strAdvAutor = Trim(Replace(UCase(strAdvAutor), "(ADVOGADO)", ""))
    End If
    
    PegaAdvAutorPje1gTjba = strAdvAutor
    
End Function

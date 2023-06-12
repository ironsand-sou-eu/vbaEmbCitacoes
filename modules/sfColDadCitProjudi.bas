Attribute VB_Name = "sfColDadCitProjudi"
Option Explicit

Sub CadastrarProcessoIndividualProjudi(strNumeroProcesso As String, ByRef rngCelula As Excel.Range)
''
'' Com o Projudi aberto e logado no Internet Explorer, busca um processo e o cadastra
''
    Dim IE As InternetExplorer, DocHTML As HTMLDocument
    Dim strCont As String, strPerfilLogado As String
    Dim dtDataProvContestar As Date, dtDataProvSubsidios As Date
    Dim planGrupoProvContestar As Excel.Worksheet, planGrupoProvSubsidios As Excel.Worksheet
    Dim arq As Workbook, plan As Excel.Worksheet
    
    ' Procurar Internet Explorer aberto no Projudi
    Set IE = SisifoEmbasaFuncoes.RecuperarIE("projudi.tjba.jus.br")
    If IE Is Nothing Then
NaoAbertoLogado:
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", � necess�rio que o Internet Explorer esteja aberto na p�gina do Projudi BA para continuar " & _
        "e devidamente logado. Abra-o, fa�a login no Projudi e rode a fun��o novamente.", vbCritical + vbOKOnly, "S�sifo - Internet Explorer n�o encontrado"
        GoTo FinalizaFechaIE
    End If
    
    IE.Visible = True
    
    'Descobrir se o perfil � de parte ou advogado. Se for outro, manda relogar.
    strPerfilLogado = DescobrirPerfilLogadoProjudi(IE.document)
    If strPerfilLogado = "N�o logado" Then GoTo NaoAbertoLogado
    
    ' Carregar p�gina de busca
    Set IE = CarregarPaginaBuscaProjudi(strPerfilLogado)
    If IE Is Nothing Then GoTo FinalizaFechaIE
    Set DocHTML = IE.document
    
    'Pegar link pelo n�mero CNJ
    strCont = PegarLinkProcessoProjudi(strNumeroProcesso, strPerfilLogado, IE, DocHTML)
    If LCase(Left(strCont, 4)) = "erro" Then GoTo FinalizaFechaIE
    
    ' Abrir processo
    IE.Visible = True
    IE.navigate strCont '& "&consentimentoAcesso=true"
    
    If strPerfilLogado = "Advogado" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", confirme a abertura do processo e resolva o captcha no Projudi. " & _
            "Quando a p�gina inicial do processo abrir, clique em ""OK"" para continuar.", vbCritical + vbOKOnly, _
            "S�sifo - Resolver captcha"
    End If
    
    'Aguardar elemento aparecer
    Do
    Loop Until IE.readyState = READYSTATE_COMPLETE
    
    Set DocHTML = IE.document
    
    On Error Resume Next
    Do
    Loop While DocHTML.body.Children(2) Is Nothing
    On Error GoTo 0
    
    'Se for segredo de justi�a, avisa e para tudo
    If DocHTML.body.Children(2).innerText = "Processo sob Segredo de Justi�a" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo est� em segredo de justi�a. Tente novamente com um usu�rio com acesso.", _
            vbCritical + vbOKOnly, "S�sifo - Processo em segredo de justi�a"
        rngCelula.Interior.Color = 9420794
        GoTo FinalizaFechaIE
    End If
    
    'Expande os bot�es de arquivos e observa��es de andamento
    'ExpandirBotoesProcesso IE, DocHTML, , True
    
    ' Pega os dados do processo
    PegaInformacoesProcessoProjudi IE, DocHTML
    If prProcesso Is Nothing Then GoTo FinalizaFechaIE
    PegaInformacoesProcessoGeral dtDataProvContestar, dtDataProvSubsidios, planGrupoProvContestar, planGrupoProvSubsidios
    
    ' Se deu erro, devolve o erro
    If prProcesso Is Nothing Then GoTo FinalizaFechaIE
    If prProcesso.MensagemErro <> "" Then GoTo FinalizaFechaIE
    
    ' Se deu certo, insere na mem�ria do S�sifo (PODE DAR ERRO: SE N�O CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
    Set arq = Excel.Workbooks.Add
    Set plan = arq.Sheets(1)
    
    prProcesso.ExportaLinhasEspaider arq, plan
    If Not planGrupoProvContestar Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvContestar, dtDataProvContestar, prProcesso.Advogado)
    If Not planGrupoProvSubsidios Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvSubsidios, dtDataProvSubsidios, prProcesso.Preposto)
    
    rngCelula.Offset(0, 1).Formula = "Inserido no S�sifo"
    
    ' Se conseguiu cadastrar e estiver no perfil de parte, procura a cita��o e l�
    If prProcesso.MensagemErro = "" And strPerfilLogado <> "Advogado" Then
        strCont = LerCitacaoProjudi(prProcesso.NumeroProcesso)
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", " & strCont, vbOKOnly + vbInformation, "S�sifo - Resultado da leitura de cita��o"
        
        If InStr(1, strCont, "lida com sucesso") <> 0 Then
            rngCelula.Offset(0, 2).Formula = "Cita��o lida"
        Else
            rngCelula.Offset(0, 2).Formula = "Cita��o N�O LIDA"
        End If
    End If

    rngCelula.Offset(1, 0).Select
    
FinalizaFechaIE:
    FecharExplorerZerarVariaveis IE
    
End Sub

Sub CadProcIndividualProjudiSegredoJus(ByVal Controle As IRibbonControl)
''
'' Com o Projudi aberto no Internet Explorer e deslogado, busca um processo e o cadastra
''
    
    Dim IE As InternetExplorer, DocHTML As HTMLDocument
    Dim frmTeor As HTMLFormElement, frFrame As HTMLFrameElement
    Dim strSenhaAcesso As String, strCont As String
    Dim dtDataProvContestar As Date, dtDataProvSubsidios As Date
    Dim planGrupoProvContestar As Excel.Worksheet, planGrupoProvSubsidios As Excel.Worksheet
    Dim arq As Workbook, plan As Excel.Worksheet, rngCelula As Excel.Range
    Dim bolCont As Boolean
    
    Set rngCelula = ActiveCell
    
    strSenhaAcesso = PegarSenhaAcesso(rngCelula)
    
    ' Procurar Internet Explorer aberto no Projudi
    Set IE = SisifoEmbasaFuncoes.RecuperarIE("projudi.tjba.jus.br")
    If IE Is Nothing Then
PaginaErrada:
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", � necess�rio que o Internet Explorer esteja aberto na p�gina inicial do Projudi BA, sem logar em nenhum usu�rio, para continuar. Abra a p�gina do Projudi, " & _
            "saia de qualquer login e rode a fun��o novamente.", vbCritical + vbOKOnly, "S�sifo - Internet Explorer n�o encontrado"
        GoTo FinalizaFechaIESegrJus
        Exit Sub
    End If
    
    IE.Visible = True
    
    Set DocHTML = IE.document
    Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
    Set DocHTML = frFrame.contentDocument
   
    ' Pegar form do acesso ao teor com senha e preencher
    Set frmTeor = DocHTML.getElementById("formAcessoPublico")
    If frmTeor Is Nothing Then GoTo PaginaErrada
    
    DocHTML.getElementById("codigoHash").Value = strSenhaAcesso
    frmTeor.getElementsByTagName("a")(0).Click
    
    Do
        DoEvents
    Loop Until IE.readyState = READYSTATE_COMPLETE

    On Error GoTo Volta
Volta:
    Do
        Set IE = SisifoEmbasaFuncoes.RecuperarIE("projudi.tjba.jus.br")
        Set DocHTML = IE.document
        Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
        Set DocHTML = frFrame.contentDocument
        bolCont = IIf(DocHTML.URL = sfUrlProjudiAcessoPublico, True, False)
    Loop Until bolCont = True
    On Error GoTo 0
    
    ' Pega os dados do processo
    PegaInformacoesProcessoProjudi IE, DocHTML
    If prProcesso Is Nothing Then GoTo FinalizaFechaIESegrJus
    PegaInformacoesProcessoGeral dtDataProvContestar, dtDataProvSubsidios, planGrupoProvContestar, planGrupoProvSubsidios
    
    ' Se deu erro, devolve o erro
    If prProcesso Is Nothing Then GoTo FinalizaFechaIESegrJus
    If prProcesso.MensagemErro <> "" Then GoTo FinalizaFechaIESegrJus
    
    ' Se deu certo, insere na mem�ria do S�sifo (PODE DAR ERRO: SE N�O CONSEGUIR EXPORTAR, VAI RETORNAR MENSAGEM DE ACERTO DO MESMO JEITO)
    Set arq = Excel.Workbooks.Add
    Set plan = arq.Sheets(1)
    
    prProcesso.ExportaLinhasEspaider arq, plan
    If Not planGrupoProvContestar Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvContestar, dtDataProvContestar, prProcesso.Advogado)
    If Not planGrupoProvSubsidios Is Nothing Then strCont = RegistraResponsavelPorProvidenciaNoGrupo(planGrupoProvSubsidios, dtDataProvSubsidios, prProcesso.Preposto)
    
    rngCelula.Offset(0, 1).Formula = "Inserido no S�sifo"
    rngCelula.Offset(1, 0).Select
    
    
FinalizaFechaIESegrJus:
    FecharExplorerZerarVariaveis IE
    
End Sub

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

Function DescobrirPerfilLogadoProjudi(DocHTML As HTMLDocument) As String
''
'' Descobre o perfil do documento aberto e, conforme o caso, retorna "Parte" ou "Advogado"
''
    Dim frFrame As HTMLFrameElement
    Dim frForm As HTMLFormElement
    
    Set frFrame = DocHTML.getElementsByName("mainFrame")(0)
    
    On Error Resume Next
    Set frForm = frFrame.contentDocument.getElementsByName("formLogin")(0)
    On Error GoTo 0
    
    If Not frForm Is Nothing Then
        DescobrirPerfilLogadoProjudi = "N�o logado"
    Else
        If InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Parte") <> 0 Then '� parte
            DescobrirPerfilLogadoProjudi = "Parte"
        ElseIf InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Advogado") <> 0 Then ' � Advogado
            DescobrirPerfilLogadoProjudi = "Advogado"
        ElseIf InStr(1, frFrame.contentDocument.getElementById("Stm0p0i0eHR").href, "Representante") <> 0 Then ' � Representante
            DescobrirPerfilLogadoProjudi = "Representante"
        Else '� outra coisa
            DescobrirPerfilLogadoProjudi = "Outro"
        End If
    End If
    
End Function

Function CarregarPaginaBuscaProjudi(strPerfilLogado As String) As InternetExplorer
''
'' Abre nova janela do Internet Explorer na p�gina de buscas, conforme perfil logado
''
    
    Dim IE As InternetExplorer
    Dim DocHTML As HTMLDocument
    Dim strCont As String
    
    If strPerfilLogado = "Outro" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", � necess�rio estar logado num perfil de parte, advogado ou representante. Fa�a login no Projudi de " & _
            "um desses perfis e rode a fun��o novamente.", vbCritical + vbOKOnly, "S�sifo - Internet Explorer n�o encontrado"
        Set CarregarPaginaBuscaProjudi = Nothing
        Exit Function
    End If
    
    ' Carrega p�gina de busca, conforme o perfil
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.navigate IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado, sfUrlProjudiBuscaParte)
    Set IE = SisifoEmbasaFuncoes.RecuperarIE(IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado, sfUrlProjudiBuscaParte))
    
    ' Aguarda carregar
    Do
        DoEvents
    Loop Until IE.readyState = 4
    
    Do
        DoEvents
        strCont = IE.document.URL
    Loop Until strCont = IIf(strPerfilLogado = "Advogado", sfUrlProjudiBuscaAdvogado, sfUrlProjudiBuscaParte)
    
    Set DocHTML = IE.document
    'Set DocHTML = DocHTML.getElementsByName("mainFrame")(0).contentDocument.getElementsByName("userMainFrame")(0).contentDocument
    
    Set CarregarPaginaBuscaProjudi = IE
    
End Function

Function PegarLinkProcessoProjudi(ByVal strNumeroCNJ As String, ByVal strPerfilLogado As String, ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument) As String
''
'' Retorna o link da p�gina principal do processo strNumeroCNJ.
'' DEVO LIDAR COM O ERRO DE N�O ESTAR LOGADO!!!!!!!
''

    Dim strContNumeroProcesso As String, strCont As String
    Dim frmProcessos As HTMLFormElement
    Dim intCont As Integer

'    On Error GoTo Volta1
'Volta1:
'    Do
'        DoEvents
'    Loop Until IE.document.readyState = "complete"
'
'    Do
'        DoEvents
'    Loop Until IE.document.getElementsByTagName("body")(0).Children(2).Children(0).Children(0).Children(0).Children(1).Children(0).innerText = "N�mero Processo"
'    On Error GoTo 0


'    ADICIONAR (NO LOCAL ADEQUADO) TRATAMENTO PARA:
'    Case "N�o abriu por demora"
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo n�o abriu por demora. Provavelmente, a conex�o est� muito lenta. Tente novamente daqui a pouco.", vbCritical + vbOKOnly, "S�sifo - Tempo de espera expirado"
'        GoTo FinalizaFechaIE
'    Case "Mais de um processo encontrado"
'        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", foi encontrado mais de um processo para o n�mero " & rngProcesso.Formula & ". Isso � completamente inesperado! " & _
'            "Suplico que confira o n�mero e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Mais de um processo encontrado"
'        GoTo FinalizaFechaIE

    If DocHTML.Title = "Sistema CNJ - A sess�o expirou" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a sess�o expirou. Fa�a login no Projudi NA MESMA JANELA EM QUE EST� EXPIRADA, ent�o " & _
        "clique OK e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Sess�o do Projudi expirada"
        PegarLinkProcessoProjudi = "Erro: sess�o expirada"
        Exit Function
    End If
    
    DocHTML.getElementById("numeroProcesso").Value = strNumeroCNJ
    DocHTML.forms("busca").submit
    
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
        strCont = IIf(strPerfilLogado = "Advogado", "form1", "formProcessos")
        Set frmProcessos = DocHTML.getElementById(strCont)
        'COLOCAR UM TIMEOUT AQUI (TRATAMENTO DO ERRO EST� COMENTARIZADO ALI EM CIMA)
    Loop While frmProcessos Is Nothing
    
    esperar 0.3
    intCont = frmProcessos.getElementsByTagName("a").length - 1
    For intCont = 0 To intCont Step 1
        If frmProcessos.getElementsByTagName("a")(intCont).innerText = strNumeroCNJ Then Exit For
    Next intCont
    On Error GoTo 0
    
    If intCont = frmProcessos.getElementsByTagName("a").length Then 'Correu todos os links e n�o achou
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", lastimo informar o processo n�o foi encontrado. Verifique se o n�mero est� " & _
            "correto e tente novamente, ou talvez o processo n�o seja acess�vel para o usu�rio logado no Projudi (por exemplo, em segredo " & _
            "de justi�a.", vbCritical + vbOKOnly, "S�sifo - Processo n�o encontrado"
        PegarLinkProcessoProjudi = "Erro: processo n�o encontrado"
    Else 'Achou
        PegarLinkProcessoProjudi = frmProcessos.getElementsByTagName("a")(intCont)
    End If
    
End Function

Sub PegaInformacoesProcessoProjudi(ByRef IE As InternetExplorer, ByRef DocHTML As HTMLDocument)
''
'' Faz a coleta dos dados do processo espec�ficos do PJe e armazena na vari�vel global prProcesso
''
    Dim divCont As HTMLDivElement, tbTabelaAndamentos As HTMLTable, tbCont As HTMLTable
    Dim contOutrasPartes As OutroParticipante
    Dim varCont As Variant
    Dim Cont As Integer
    Dim strCodParteAutora As String
    Dim sngTimerInicio As Single, sngTimerFim As Single
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a provid�ncia de agendar pautista
    
    ''''''''''''''''''''''''''''''''''''
    ''' Carregar totalmente a p�gina '''
    ''''''''''''''''''''''''''''''''''''
    
EsperarPagina:
    sngTimerInicio = Timer
    Do
        Set divCont = DocHTML.getElementById("Arquivos")
        If Timer >= sngTimerInicio + 10 Then
            'Perguntar se quer continuar
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", a p�gina parece estar demorando de carregar. Suplico que confira se h� algo errado e diga se pretende " & _
                "esperar mais 10 segundos, ou cancelar o carregamento desta p�gina.", vbRetryCancel + vbQuestion, "S�sifo - Demora no carregamento") = vbRetry Then
                GoTo EsperarPagina
            Else
                Set prProcesso = Nothing
                Exit Sub
            End If
        End If
    Loop While divCont Is Nothing
    
    
    ''''''''''''''''''''''''''
    ''' N�mero do processo '''
    ''''''''''''''''''''''''''
    
        prProcesso.NumeroProcesso = PegaNumeroProjudi(DocHTML)

    ' Confere se o processo j� est� na planilha
    If Not sfCadProcessos.Cells().Find(prProcesso.NumeroProcesso) Is Nothing Then
        Do
        Loop Until MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o processo j� existe na planilha! Inclus�o cancelada. Descartados os dados." & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf & _
        "Clique em 'Cancelar' e insira o pr�ximo processo.", vbCritical + vbOKCancel, "S�sifo - Processo repetido") = vbCancel
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Cita��o, respons�vel regressivo, ju�zo, valor da '''
    ''' causa, ju�zo, tipo de a��o, rito, data audi�ncia '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ' N�o estando repetido na mem�ria, coleta os outros dados
    Set tbTabelaAndamentos = DocHTML.getElementById("Arquivos").Children(0).Children(0)
    
    prProcesso.Sistema = sfProjudi
    prProcesso.Citacao = PegaCitacaoProjudi(tbTabelaAndamentos)
    prProcesso.DataAndamento = PegaDataAudienciaProjudi(tbTabelaAndamentos)
    If prProcesso.DataAndamento = 0 Then
        prProcesso.NomeAndamento = "Cita��o"
    Else
        prProcesso.NomeAndamento = "Audi�ncia Una virtual"
    End If
    
    prProcesso.ResponsavelRegressivo = ""
    prProcesso.ValorCausa = PegaValorProjudi(DocHTML)
    
    prProcesso.Juizo = PegaJuizoProjudi(DocHTML)
    If prProcesso.Juizo = "Ju�zo n�o cadastrado" Then
        Set prProcesso = Nothing
        Exit Sub
    End If
    
    prProcesso.TipoAcao = cfConfigura��es.Cells().Find(What:="Tipo de A��o", LookAt:=xlWhole).Offset(0, 1).Formula
    prProcesso.Rito = cfConfigura��es.Cells().Find(What:="Rito", LookAt:=xlWhole).Offset(0, 1).Formula
    
    ''''''''''''''''''''''''''
    ''' Autores e advogado '''
    ''''''''''''''''''''''''''
    
    ' Adiciona Autores
    Set tbCont = DocHTML.getElementById("tabelaPartes1"): strCodParteAutora = "1" 'Promovente
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes14"): strCodParteAutora = "14" ' Autor
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes30"): strCodParteAutora = "30" ' Exeq�ente
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes26"): strCodParteAutora = "26" ' Embargante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes63"): strCodParteAutora = "63" ' Requerente
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes58"): strCodParteAutora = "58" ' Reclamante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes69"): strCodParteAutora = "69" ' Suscitante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes36"): strCodParteAutora = "36" ' Impugnante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes56"): strCodParteAutora = "56" ' Querelante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes28"): strCodParteAutora = "28" ' Excipiente
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes32"): strCodParteAutora = "32" ' Expropriante
    If tbCont Is Nothing Then Set tbCont = DocHTML.getElementById("tabelaPartes34"): strCodParteAutora = "34" ' Impetrante
    
    varCont = PegaPartesProjudi(tbCont, sfAutor)
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
    
    prProcesso.AdvogadoAdverso = PegaAdvAutor(DocHTML, strCodParteAutora)
    
    '''''''''''''''''''
    ''' Outros r�us '''
    '''''''''''''''''''
    
    ' Adiciona outros R�us, se for o caso.
    Select Case strCodParteAutora
    Case "1"
        Set tbCont = DocHTML.getElementById("tabelaPartes0") ' Promovido
    Case "14"
        Set tbCont = DocHTML.getElementById("tabelaPartes67") ' R�u
    Case "30"
        Set tbCont = DocHTML.getElementById("tabelaPartes29") ' Executado
    Case "26"
        Set tbCont = DocHTML.getElementById("tabelaPartes25") ' Embargado
    Case "63"
        Set tbCont = DocHTML.getElementById("tabelaPartes65") ' Requerido
    Case "58"
        Set tbCont = DocHTML.getElementById("tabelaPartes57") ' Reclamado
    Case "69"
        Set tbCont = DocHTML.getElementById("tabelaPartes68") ' Suscitado
    Case "36"
        Set tbCont = DocHTML.getElementById("tabelaPartes35") ' Impugnado
    Case "56"
        Set tbCont = DocHTML.getElementById("tabelaPartes55") ' Querelado
    Case "28"
        Set tbCont = DocHTML.getElementById("tabelaPartes27") ' Excepto
    Case "32"
        Set tbCont = DocHTML.getElementById("tabelaPartes31") ' Expropriado
    Case "34"
        Set tbCont = DocHTML.getElementById("tabelaPartes33") ' Impetrado
    End Select
    
    varCont = PegaPartesProjudi(tbCont, sfReu)
    If Not IsEmpty(varCont) Then
        For Cont = 1 To UBound(varCont, 2)
            Set contOutrasPartes = New OutroParticipante
            If varCont(1, Cont) = "CCR METRO BAHIA" Or varCont(1, Cont) = "CCR METRO" Or varCont(1, Cont) = "CCR S A" _
            Or varCont(1, Cont) = "CCR CIA METRO DA BAHIA" Then ' Se for CCR METRO BAHIA, CCR METRO, CCR S A ou CCR CIA METRO DA BAHIA
                With contOutrasPartes
                    .NomeParte = "COMPANHIA DO METRO DA BAHIA"
                    .CNPJParte = "18.891.185/0001-37"
                    .CPFParte = 0
                    .TipoParte = 2
                End With
            Else
            ' Se n�o for CCR
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
            End If
            
            contOutrasPartes.ClasseParte = 1
            contOutrasPartes.CondicaoParte = "R�u"
            prProcesso.OutrosParticipantes.Add contOutrasPartes
        Next Cont
    End If
End Sub

Function LerCitacaoProjudi(strNumeroProcesso As String) As String
''
'' Abre a p�gina de cita��es e l� a cita��o do processo respectivo
''

    Dim IECont As InternetExplorer
    Dim DocHTMLCont As HTMLDocument
    Dim frmCont As HTMLFormElement
    Dim strLink As String

    ' Abre a p�gina de cita��es novas
    Set IECont = New InternetExplorer
    IECont.Visible = True
    IECont.navigate sfUrlProjudiCitacoesNovas
    Set IECont = SisifoEmbasaFuncoes.RecuperarIE(sfUrlProjudiCitacoesNovas)
    Set DocHTMLCont = IECont.document
    
    Do
        DoEvents
    Loop Until DocHTMLCont.readyState = "complete"
    
    ' Procura a do processo que foi cadastrado e l�
    Set frmCont = DocHTMLCont.getElementById("selecao")
    DocHTMLCont.getElementById("numeroProcesso").Value = strNumeroProcesso
    frmCont.submit
    
    Do
        DoEvents
    Loop Until DocHTMLCont.readyState = "complete"
    
    Set frmCont = DocHTMLCont.getElementById("selecao")
    
    If frmCont.NextSibling.NextSibling.NextSibling.NextSibling.nodeName = "SCRIPT" Then GoTo ErroNaLeitura
    
    strLink = frmCont.NextSibling.NextSibling.NextSibling.NextSibling.Children(0).Children(2).Children(9).Children(0).href
    IECont.navigate strLink
    Set IECont = SisifoEmbasaFuncoes.RecuperarIE(strLink)
    Set DocHTMLCont = IECont.document
    
    Do
        DoEvents
    Loop Until IECont.readyState = READYSTATE_COMPLETE
    
    If DocHTMLCont.body.Children(0).Children(0).innerText = "Citacao Para EMBASA" Then
        LerCitacaoProjudi = "tenho o prazer de informar que a cita��o foi lida com sucesso."
    Else
ErroNaLeitura:
        LerCitacaoProjudi = "ocorreu algum PROBLEMA na leitura da cita��o. Caso o processo tenha sido cadastrado pelo S�sifo, " & _
                            "fa�a a leitura manualmente; caso n�o tenha sido cadastrado, tente cadastr�-lo de novo pelo S�sifo."
    End If
    
    IECont.Quit

End Function

Function PegaNumeroProjudi(ByRef DocHTML As HTMLDocument) As String
    Dim strNumero As String
    
    strNumero = DocHTML.getElementById("Partes").Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).Children(0).innerText
    PegaNumeroProjudi = strNumero
    
End Function

Function PegaCitacaoProjudi(ByRef tbTabelaAndamentos As HTMLTable) As Date
''
'' Busca o primeiro andamento "Cita��o lido(a)" cuja observa��o abaixo contenha os nomes "EMBASA" ou "SANEAMENTO".
'' Encontrando, a data. N�o encontrando, retornar� "0". Havendo m�ltiplas, retornar� a mais recente (mais acima).
''
    Dim intCont As Integer
    Dim strCont As String
    
    'Itera as linhas da tabela de eventos
    For intCont = 0 To tbTabelaAndamentos.Children.length - 1
        If tbTabelaAndamentos.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).Children(0).innerText = "Cita��o lido(a)" Then
            strCont = Trim(tbTabelaAndamentos.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerText)
            If InStr(1, strCont, "EMBASA") <> 0 Or InStr(1, strCont, "SANEAMENTO") <> 0 Then
                ' Achou. Retorna a data e sai da fun��o
                PegaCitacaoProjudi = Mid(strCont, InStr(1, strCont, " em ") + 4, 8)
                Exit Function
            End If
        End If
    Next intCont
    
    ' Se n�o houver andamento "Cita��o lido(a)" para "Embasa" ou "Saneamento", retorna a data de hoje.
    PegaCitacaoProjudi = Date
    
End Function

Function PegaDataAudienciaProjudi(ByRef tbTabelaAndamentos As HTMLTable) As String
''
'' Busca a primeira c�lula com o conte�do "Audi�ncia" e "Designada" ou "Cancelada" ou "Redesignada". Se n�o encontrar ou se
'' for "Cancelada" ou "Redesignada", retorna "n�o h�" ou "cancelada". Nos demais casos, trata e retorna a data da audi�ncia. Havendo
'' m�ltiplas, retornar� apenas a mais recente (a de cima).
''
    Dim intCont As Integer
    Dim strCont As String
    
    'Itera as linhas da tabela de eventos
    For intCont = 0 To tbTabelaAndamentos.Children.length - 1
    
        ' Se o nome do andamento contiver "Audi�ncia", verifica melhor...
        If InStr(1, tbTabelaAndamentos.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).Children(0).innerText, "Audi�ncia") <> 0 Then
            strCont = tbTabelaAndamentos.Children(intCont).Children(0).Children(0).Children(0).Children(0).Children(1).innerText
            
            ' ...Se o nome do andamento tamb�m contiver "designada", bingo! Retorna a data e sai da fun��o
            If InStr(1, strCont, "Designada") <> 0 Then
                strCont = ConverteDataProjudi(Mid(strCont, InStr(1, strCont, "(Agendada para ")))
                PegaDataAudienciaProjudi = strCont
                Exit Function
            End If
        End If
    Next intCont
    
    
    ' Se n�o houver andamento "Audi�ncia" ou n�o se encaixar em nenhuma hip�tese anterior, retorna 0.
    PegaDataAudienciaProjudi = 0

End Function

Function PegaJuizoProjudi(ByRef DocHTML As HTMLDocument) As String

    Dim strJuizo As String
    Dim intFinal As Integer
    Dim rngCont As Excel.Range
    
    strJuizo = DocHTML.getElementById("Partes").Children(0).Children(0).Children(6).Children(1).innerText

    If InStr(1, strJuizo, "Juiz:") <> 0 Then
        intFinal = InStr(1, strJuizo, "Juiz:") - 2
    Else
        intFinal = InStr(1, strJuizo, "Hist�rico de Ju�zes") - 2
    End If
    
    strJuizo = Left(strJuizo, intFinal)
    
    Set rngCont = cfJuizos.Cells().Find(What:=strJuizo, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        InputBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu n�o conhe�o o ju�zo em que o processo tramita. Rogo que o cadastre em minha mem�ria (a reda��o " & _
            "do Projudi pode ser copiada abaixo) e tente novamente.", "S�sifo - Ju�zo desconhecido", Trim(strJuizo)
        PegaJuizoProjudi = "Ju�zo n�o cadastrado"
    Else
        PegaJuizoProjudi = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaValorProjudi(ByRef DocHTML As HTMLDocument) As Currency
    
    PegaValorProjudi = CCur(Trim(DocHTML.getElementById("Partes").Children(0).Children(0).Children(14).Children(1).Children(0).innerText))
    
End Function

Function PegaPartesProjudi(ByRef tbPartes As HTMLTable, strTipoParte As String) As Variant
''
'' Retorna uma matriz com as partes de um polo do processo.
'' PegaPartesProjudi(1, N) = Nome do Autor na posi��o N da matriz.
'' PegaPartesProjudi(2, N) = CPF/CNPJ do Autor na posi��o N da matriz.
''
    Dim intContLinhas As Integer, intContPartes As Integer, intQtdLinhas As Integer
    Dim strMsgErro As String
    Dim arrPartes() As String
    Dim bolPular As Boolean
    
    ' Verifica quantas linhas existem na tabela de partes al�m do cabe�alho (as linhas correspondem a partes ou, se a tabela estiver expandida, advogados
    ' e endere�o). O nome � composto por "tr" e o codParte num�rico, podendo ter "Adv" ou "End" no meio se for o caso de linha de advogado ou endere�o.
    intQtdLinhas = tbPartes.Children(0).Children.length - 1
    ReDim arrPartes(1 To 2, 1 To intQtdLinhas)
    
    ' Se for para buscar R�us e houver um s�, pressup�e que � apenas a Embasa
    If strTipoParte = sfReu And intQtdLinhas = 1 Then Exit Function
    
    ' Para cada linha, verificar se � de advogado ou de endere�o. Se for, pula pra pr�xima. Se n�o for, processa
    For intContLinhas = 1 To intQtdLinhas Step 1
        bolPular = False
    
        ' Condi��es de pular pra pr�xima comuns a autor e r�u
        If InStr(1, tbPartes.Children(0).Children(intContLinhas).ID, "Adv") <> 0 Or _
            InStr(1, tbPartes.Children(0).Children(intContLinhas).ID, "End") <> 0 Then bolPular = True
        
        ' Condi��es de pular pra pr�xima espec�fica de r�u
        If strTipoParte = sfReu Then
            If InStr(1, tbPartes.Children(0).Children(intContLinhas).Children(1).innerText, "EMBASA") <> 0 Or _
            InStr(1, tbPartes.Children(0).Children(intContLinhas).Children(1).innerText, "EMPRESA BAIANA DE AGUA") <> 0 Then bolPular = True
        End If
        
        If bolPular = False Then
            ' Para cada parte, busca Nome e CPF (advogado pressup�e-se comum). 1 = Nome, 2 = CPF
            intContPartes = intContPartes + 1
            ' Pega nome
            arrPartes(1, intContPartes) = Trim(tbPartes.Children(0).Children(intContLinhas).Children(1).innerText)
            'Pega CPF/CNPJ
            arrPartes(2, intContPartes) = Trim(Left(tbPartes.Children(0).Children(intContLinhas).Children(3).innerText, 18))
            
            If InStr(1, arrPartes(2, intContPartes), "N�o cadastrado") <> 0 Then ' N�o cadastrado
                strMsgErro = "o CPF/CNPJ da parte " & arrPartes(1, intContPartes) & " n�o foi cadastrado no Projudi"
            ElseIf InStr(1, arrPartes(2, intContPartes), "N�o dispon�vel") <> 0 Then ' N�o dispon�vel
                strMsgErro = "o CPF/CNPJ da parte " & arrPartes(1, intContPartes) & " n�o est� dispon�vel"
            End If
            
            If strMsgErro <> "" Then
PedirCPF:
                arrPartes(2, intContPartes) = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", " & strMsgErro & ". Rogo que busque a Peti��o Inicial e informe o CPF ou CNPJ correto. Se n�o houver, deixe em branco.", "S�sifo - Informar CPF/CNPJ da parte"))
                If arrPartes(2, intContPartes) <> "" Then
                    arrPartes(2, intContPartes) = ValidarCPFCNPJ(arrPartes(2, intContPartes))
                    If Not IsNumeric(Left(arrPartes(2, intContPartes), 1)) Then
                        strMsgErro = arrPartes(2, intContPartes)
                        GoTo PedirCPF
                    End If
                End If
            End If
            
            If arrPartes(2, intContPartes) = "" Then 'Sem CPF/CNPJ
                arrPartes(2, intContPartes) = "N�o Cadastrado " & PegaCodParteSemCPF
            ElseIf InStr(1, arrPartes(2, intContPartes), "/") = 0 Then ' CPF - 14 caracteres
                arrPartes(2, intContPartes) = Left(arrPartes(2, intContPartes), 14)
            End If ' CNPJ n�o precisa de tratamento
        End If
    Next intContLinhas
    
    ReDim Preserve arrPartes(1 To 2, 1 To intContPartes)
    
    PegaPartesProjudi = arrPartes()
    
End Function

Function PegaAdvAutor(ByRef DocHTML As HTMLDocument, strCodParteAutora As String) As String
    
    Dim tbAdvAutor As HTMLTable
    Dim strAdvAutor As String
    
    Set tbAdvAutor = DocHTML.getElementById("tabelaAdvogadoPartes" & strCodParteAutora)
    
    If Trim(tbAdvAutor.Children(0).Children(0).Children(0).innerText) = "Nenhum advogado cadastrado." Then
        strAdvAutor = ""
    Else
        strAdvAutor = Trim(tbAdvAutor.Children(0).Children(1).Children(0).innerText)
        If InStr(1, strAdvAutor, " (CPF:") <> 0 Then strAdvAutor = Left(strAdvAutor, InStr(1, strAdvAutor, " (CPF:") - 1)
    End If
    
    PegaAdvAutor = strAdvAutor
    
End Function

Function ConverteDataProjudi(strData As String) As Date
''
'' Pega uma string no formato de data do projudi (por extenso) e converte em data.
''

    ' Retira in�cio e final
    strData = Replace(strData, "(Agendada para ", "")
    strData = Replace(strData, " h)", "")
    
    ' Substitui "de" por barras
    strData = Replace(strData, " de ", "/")
    
    ' Substitui "�s" por espa�o
    strData = Replace(strData, " �s ", " ")
    
    ' Substitui m�s extenso por m�s num�rico
    If InStr(1, strData, "Janeiro") Then
        strData = Replace(strData, "Janeiro", "01")
    ElseIf InStr(1, strData, "Fevereiro") Then
        strData = Replace(strData, "Fevereiro", "02")
    ElseIf InStr(1, strData, "Mar�o") Then
        strData = Replace(strData, "Mar�o", "03")
    ElseIf InStr(1, strData, "Abril") Then
        strData = Replace(strData, "Abril", "04")
    ElseIf InStr(1, strData, "Maio") Then
        strData = Replace(strData, "Maio", "05")
    ElseIf InStr(1, strData, "Junho") Then
        strData = Replace(strData, "Junho", "06")
    ElseIf InStr(1, strData, "Julho") Then
        strData = Replace(strData, "Julho", "07")
    ElseIf InStr(1, strData, "Agosto") Then
        strData = Replace(strData, "Agosto", "08")
    ElseIf InStr(1, strData, "Setembro") Then
        strData = Replace(strData, "Setembro", "09")
    ElseIf InStr(1, strData, "Outubro") Then
        strData = Replace(strData, "Outubro", "10")
    ElseIf InStr(1, strData, "Novembro") Then
        strData = Replace(strData, "Novembro", "11")
    ElseIf InStr(1, strData, "Dezembro") Then
        strData = Replace(strData, "Dezembro", "12")
    End If
    
    ConverteDataProjudi = strData
    
End Function

Sub PegarProcessosComCitacoesProjudi(ByVal Controle As IRibbonControl)
''
'' Abre a p�gina de cita��es do Projudi e pega os n�meros dos processos da primeira p�gina, inserindo nas linhas da coluna A
''

    Dim IE As InternetExplorer
    Dim DocHTML As HTMLDocument
    Dim divCont As HTMLDivElement
    Dim tbProcessos As HTMLTableSection
    Dim intCont As Integer
    Dim trCont As HTMLTableRow
    Dim strLink As String
    Dim rngCont As Excel.Range
    Dim arq As Excel.Workbook
    Dim plan As Excel.Worksheet

    ' Abre a p�gina de cita��es novas
    Set IE = New InternetExplorer
    IE.Visible = True
    IE.navigate sfUrlProjudiCitacoesNovas
    Set IE = SisifoEmbasaFuncoes.RecuperarIE(sfUrlProjudiCitacoesNovas)
    Set DocHTML = IE.document
    
    Do
        DoEvents
    Loop Until DocHTML.readyState = "complete"
    
    If InStr(1, DocHTML.Title, "sess�o expirou") <> 0 Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", a sess�o do Projudi no Internet Explorer expirou. Suplico que fa�a login no Projudi com o Internet Explorer " & _
        "e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Erro no cadastro"
        Exit Sub
    End If
    
    Set divCont = DocHTML.getElementById("Arquivos")
    Set tbProcessos = divCont.Children(0).Children(3).Children(0)
    
    Set plan = ActiveSheet
    
    If plan Is Nothing Then
        Set arq = Workbooks.Add
        Set plan = arq.Sheets(1)
        plan.Activate
    End If
    
    plan.Columns(1).ColumnWidth = 27
    plan.Columns(2).ColumnWidth = 16
    plan.Columns(3).ColumnWidth = 11
    Set rngCont = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp)
    
    For intCont = 0 To tbProcessos.Rows.length - 1
        Set trCont = tbProcessos.Children(intCont)
        If trCont.hasAttribute("align") = True And (trCont.Attributes("class").Value = "tBranca" Or trCont.Attributes("class").Value = "tCinza") Then
            ' Anota o n�mero do processo na c�lula atual e vai pra pr�xima
            rngCont.Value = trCont.Children(0).Children(0).innerText
            Set rngCont = rngCont.Offset(1, 0)
        End If
    Next intCont
    
    FecharExplorerZerarVariaveis IE

End Sub

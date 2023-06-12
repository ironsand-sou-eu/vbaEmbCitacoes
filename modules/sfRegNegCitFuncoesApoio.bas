Attribute VB_Name = "sfRegNegCitFuncoesApoio"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
 
Sub FechaConfigCitacoesVisivel(ByVal Controle As IRibbonControl, Optional ByRef returnedVal)
    SisifoEmbasaFuncoes.FechaConfigVisivel ThisWorkbook, cfConfigura��es, Controle, returnedVal
End Sub

Private Sub AoCarregarRibbonCitacoes(Ribbon As IRibbonUI)
' Chama a fun��o geral AoCarregarRibbon com os par�metros corretos.
    SisifoEmbasaFuncoes.AoCarregarRibbon cfConfigura��es, Ribbon
End Sub

Sub LiberarEdicaoCitacoes(ByVal Controle As IRibbonControl)
' Chama a fun��o geral LiberarEdicao
    SisifoEmbasaFuncoes.LiberarEdicao ThisWorkbook, cfConfigura��es
    
End Sub

Sub RestringirEdicaoRibbonCitacoes(ByVal Controle As IRibbonControl)
' Chama a fun��o geral RestringirEdicaoRibbon
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfConfigura��es, Controle
End Sub

Sub ssfCitCmbSistemaMudou(ByVal Controle As IRibbonControl, ByVal text As String)
    cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula = Trim(text)
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook ' Salva as altera��es
End Sub

Sub ssfCitCmbSistemaTexto(ByVal Control As IRibbonControl, ByRef returnedVal)
    If Control.ID = "cmbSistema" Then
        If Trim(stcitgerencia) <> "" Then
            returnedVal = cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
        Else
            'Pega o valor do sistema na planilha de configura��es.
            returnedVal = cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
        End If
    End If
End Sub

Sub CadastrarProcessoIndividual(ByVal Controle As IRibbonControl)
''
'' Com o PJe aberto e logado no Internet Explorer, busca um processo e o cadastra
''
    Dim strNumeroProcesso As String
    Dim rngCelula As Excel.Range
    
    ' Zerar a vari�vel prProcesso
    Set prProcesso = Nothing
    Set prProcesso = New Processo
    
    ' Pegar o n�mero
    Set rngCelula = ActiveCell
    strNumeroProcesso = PegaNumeroProcessoDeCelula(rngCelula)
    If strNumeroProcesso = "N�mero n�o � CNJ" Then Exit Sub
    
    ' Descobrir o sistema
    prProcesso.Tribunal = DescobrirTribunal(strNumeroProcesso)
    prProcesso.Sistema = DescobrirSistema(strNumeroProcesso)
    
    ' Chama a fun��o de cadastro correspondente ao sistema e Tribunal
    Select Case prProcesso.Sistema
    Case sfProjudi
        CadastrarProcessoIndividualProjudi strNumeroProcesso, rngCelula
    Case sfPJe1g
        Select Case prProcesso.Tribunal
        Case sfTjba
            CadastrarProcessoIndividualPje1gTjba strNumeroProcesso, rngCelula
        Case sfTRT5
            CadastrarProcessoIndividualPje1gTrt5 strNumeroProcesso, rngCelula
        End Select
    End Select
End Sub

Function PegaNumeroProcessoDeCelula(rngRange As Excel.Range) As String
''
'' Retorna o n�mero do processo contido na primeira c�lula da range passada como par�metro -- ou, se n�o for padr�o CNJ, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strNumeroProcesso As String
    Dim intTentarDeNovo As Integer
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strNumeroProcesso = rngCelula.text
    
    ' Se n�o houver c�lula no espa�o enviado, ou se estiver vazia, ou se contiver algo em formato n�o CNJ, pergunta o n�mero do processo.
    If rngCelula Is Nothing Or rngCelula.text = "" Or Not EhCNJ(ActiveCell.text) Then
PerguntaNumero:
        strNumeroProcesso = InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe o n�mero do processo do Projudi a cadastrar no formato CNJ " & _
                "(""0000000-00.0000.0.00.0000""):", "S�sifo - Cadastrar processo")
            
AvisoNaoCNJ:
        If Not EhCNJ(strNumeroProcesso) Then
            intTentarDeNovo = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o n�mero informado (" & strNumeroProcesso & ") n�o est� no padr�o do CNJ. " & _
                "Deseja tentar novamente com um n�mero no padr�o ""0000000-00.0000.0.00.0000""?", vbYesNo + vbCritical + vbDefaultButton1, _
                "S�sifo - Erro no cadastro")
                
            If intTentarDeNovo = vbYes Then
                GoTo PerguntaNumero
            Else
                PegarNumeroProcesso = "N�mero n�o � CNJ"
                Exit Function
            End If
        End If
    End If
    
    ' Se o conte�do de strNumeroProcesso for um n�mero de processo CNJ, aceita.
    PegaNumeroProcessoDeCelula = strNumeroProcesso
    
End Function

Function DescobrirTribunal(strNumero As String) As String
''
'' Confere a que tribunal pertence o n�mero strNumero. Retorna uma string com o sistema, ou a mensagem de erro se n�o for CNJ.
''

    Dim strNumJustica As String, strNumTibunal As String, strCmbTribunal As String
    
    ' Se o n�mero n�o for CNJ, retorna erro
    If Not EhCNJ(strNumero) Then
        DescobrirTribunal = "N�mero passado n�o � CNJ"
        Exit Function
    End If
    
    ''' Pega o tribunal de acordo com a combobox do Ribbon do S�sifo
    Select Case cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
    Case "Projudi"
        strCmbTribunal = sfTjba
    Case "PJe 1g TJ/BA"
        strCmbTribunal = sfTjba
    Case "PJe 1g TRT5"
        strCmbTribunal = sfTRT5
    End Select
    
    
    ''' Pega o tribunal escolhido de acordo com o n�mero
    strNumJustica = Mid(strNumero, 17, 1)
    strNumTibunal = Mid(strNumero, 19, 2)
    
    ' Se n�o for TJ/BA, retorna erro.
    If strNumJustica <> "8" And strNumJustica <> "5" Then
        DescobrirTribunal = "o processo n�o pertence a TJ nem a TRT, e ainda n�o � cadastrado pelo S�sifo."
        Exit Function
    
    ' Justi�a estadual de outros estados
    ElseIf strNumJustica = "8" And strNumTibunal <> "05" Then
        DescobrirTribunal = "o processo pertence ao Tribunal de Justi�a de outro estado, que ainda n�o � cadastrado pelo S�sifo."
        Exit Function
    
    ' Justi�a do Trabalho de outros regionais
    ElseIf strNumJustica = "5" And strNumTibunal <> "05" Then
        DescobrirTribunal = "o processo pertence ao TRT de outra regi�o, que ainda n�o � cadastrado pelo S�sifo."
        Exit Function
    
    ' TJ/BA
    ElseIf strNumJustica = "8" And strNumTibunal = "05" Then
        If strCmbTribunal = sfTjba Then
            DescobrirTribunal = sfTjba
        Else
            DescobrirTribunal = "embora tenhais selecionado um sistema do tribunal " & strCmbTribunal & ", o n�mero parece indicar que " & _
                "o processo pertence, em verdade, ao TJ/BA. Imploro que verifiqueis a inconsist�ncia."
            Exit Function
        End If

    ' TRT 5� Regi�o
    ElseIf strNumJustica = "5" And strNumTibunal = "05" Then
        If strCmbTribunal = sfTRT5 Then
            DescobrirTribunal = sfTRT5
        Else
            DescobrirTribunal = "embora tenhais selecionado um sistema do tribunal " & strCmbTribunal & ", o n�mero parece indicar que " & _
                "o processo pertence, em verdade, ao TRT da 5� regi�o. Imploro que verifiqueis a inconsist�ncia."
            Exit Function
        End If

    End If
    
End Function

Function DescobrirSistema(strNumero As String) As String
''
'' Confere se o n�mero passado corresponde a um n�mero de processo no padr�o CNJ, e depois tenta descobrir o sistema em que tramita
''   (baseia-se apenas nos padr�es rotineiros de numera��o do TJ/BA).
''    Retorna uma string com o sistema, ou a mensagem de erro correspondente.
''

    Dim strCont As String, strCmbSistema As String
    
    ''' Pega o tribunal de acordo com a combobox do Ribbon do S�sifo
    Select Case cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
    Case "Projudi"
        strCmbSistema = sfProjudi
    Case "PJe 1g TJ/BA"
        strCmbSistema = sfPJe1g
    Case "PJe 1g TRT5"
        strCmbSistema = sfPJe1g
    End Select
    
    ' Se n�o for CNJ, retorna erro
    If Not EhCNJ(strNumero) Then
        DescobrirSistema = "N�mero de processo indicado n�o est� no padr�o CNJ"
    
    ' Se n�o for dos tribunais abrangidos, retorna o erro respectivo.
    ElseIf prProcesso.Tribunal <> sfTjba And prProcesso.Tribunal <> sfTRT5 Then
        DescobrirSistema = prProcesso.Tribunal
    
    Else
        Select Case prProcesso.Tribunal
        Case sfTjba 'TJ/BA
            ' Se n�o come�ar com 0, � PJe
            If Left(strNumero, 1) <> "0" Then
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o n�mero parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema do tribunal. Imploro que verifiqueis a inconsist�ncia."
                    Exit Function
                End If
            
            ' Se come�ar com 03 ou 05, � eSaj (o S�sifo vai tratar como PJe, pois o TJ/BA descontinuar� o eSaj e j� migrou alguns processos)
            ElseIf Left(strNumero, 2) = "03" Or Left(strNumero, 2) = "05" Then
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o n�mero parece indicar que " & _
                        "o processo pertence, em verdade, ao eSaj. Imploro que verifiqueis a inconsist�ncia."
                    Exit Function
                End If
            
            ' Nos demais casos, � Projudi
            Else
                If strCmbSistema = sfProjudi Then
                    DescobrirSistema = sfProjudi
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o n�mero parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema. Imploro que verifiqueis a inconsist�ncia."
                    Exit Function
                End If
            End If
            
        Case sfTRT5 'TRT5
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o n�mero parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema do tribunal. Imploro que verifiqueis a inconsist�ncia."
                    Exit Function
                End If
        
        End Select
                
    End If
    
End Function

Function EhCNJ(strNumero As String) As Boolean
''
'' Confere se o n�mero passado corresponde a um n�mero de processo no padr�o CNJ.
'' S� retorna VERDADEIRO se for padr�o CNJ com pontos e tra�os.
'' O padr�o � 0000000-00.0000.0.00.0000 - zeros significam qualquer n�mero; h�fens e pontos s�o h�fens e pontos mesmo.
''

    Dim strCont As String
    Dim btCont As Byte
    Dim bolEhCNJ As Boolean
    
    bolEhCNJ = True
    
    ' Se n�o tiver 25 caracteres, n�o � CNJ
    If Len(strNumero) <> 25 Then bolEhCNJ = False
    
    For btCont = 1 To 25 Step 1 ' Itera caractere a caractere, verificando:
    
        strCont = Mid(strNumero, btCont, 1)
        
        Select Case btCont
        Case 8
            If strCont <> "-" Then bolEhCNJ = False ' Se o h�fen est� no lugar
            
        Case 11, 16, 18, 21
            If strCont <> "." Then bolEhCNJ = False ' Se os pontos est�o no lugar
        
        Case Else
            If Not IsNumeric(strCont) Then bolEhCNJ = False ' Se os demais s�o n�meros
            
        End Select
        
    Next btCont
    
    EhCNJ = bolEhCNJ
    
End Function

Function ValidarCPFCNPJ(strStringAValidar As String) As String
''
'' Com uma string qualquer passada como par�metro, retorna um CPF/CNPJ v�lido ou retorna uma mensagem de erro correspondente.
''
    Dim strCont As String, strCont2
    Dim btCont As Integer
    Dim intCont As Integer
    
    
    ' Remove o que n�o for num�rico
    For intCont = Len(strStringAValidar) To 1 Step -1
        If IsNumeric(Mid(strStringAValidar, intCont, 1)) Then strCont = Mid(strStringAValidar, intCont, 1) & strCont
    Next intCont
    
    ' Verifica se o tamanho � compat�vel com CPF ou CNPJ
    If Len(strCont) <> 11 And Len(strCont) <> 14 Then
        ValidarCPFCNPJ = "o texto informado tem tamanho incompat�vel com CPF ou CNPJ"
        Exit Function
    End If
    
    ' Valida d�gito do CPF
    If Len(strCont) = 11 Then
        strCont2 = Left(strCont, 9)

        ' Descobre primeiro d�gito
        For btCont = 1 To 9 Step 1
            intCont = (CInt(Mid(strCont2, btCont, 1)) * (11 - btCont)) + intCont
        Next btCont
        intCont = ((intCont \ 11 + 1) * 11) - intCont
        If intCont = 10 Then intCont = 0
        strCont2 = strCont2 & CStr(intCont)
        
        ' Descobre segundo d�gito
        intCont = 0
        For btCont = 1 To 10 Step 1
            intCont = (CInt(Mid(strCont2, btCont, 1)) * (12 - btCont)) + intCont
        Next btCont
        intCont = ((intCont \ 11 + 1) * 11) - intCont
        If intCont = 10 Or intCont = 11 Then intCont = 0
        strCont2 = strCont2 & CStr(intCont)
        
        If strCont <> strCont2 Then
            ValidarCPFCNPJ = "o d�gito verificador do CPF � inv�lido"
            Exit Function
        End If
        
        strCont = Left(strCont, 3) & "." & Mid(strCont, 4, 3) & "." & Mid(strCont, 7, 3) & "-" & Right(strCont, 2)
        
    ElseIf Len(strCont2) = 14 Then
        strCont = Left(strCont, 2) & "." & Mid(strCont, 3, 3) & "." & Mid(strCont, 6, 3) & "/" & Mid(strCont, 9, 4) & "-" & Right(strCont, 2)
    End If
    
    ValidarCPFCNPJ = strCont
    
End Function

Function PegaCodParteSemCPF() As String
''
'' Pega um c�digo �nico de parte sem CPF. Os c�digos s�o compostos por dia, m�s, ano e quatro d�gitos sequenciais
''   para a quantidade de partes sem CPF cadastradas hoje.
''
    Dim rngRange As Range
    Dim strCodigo As String
    
    ' Determina a c�lula que cont�m o c�digo atual
    Set rngRange = cfConfigura��es.Cells().Find(What:="Partes sem CPF cadastrado", LookAt:=xlWhole).Offset(0, 1)
    
    If Left(rngRange.text, 8) = Format(Date, "ddmmyyyy") Then
    ' Se o c�digo atual est� no mesmo dia, soma um e transforma-o no c�digo atual
        rngRange.Formula = "'" & Trim(rngRange.text + 1)
        
    Else
    ' Se c�digo atual est� em outro dia, cria o c�digo de um novo dia no n�mero "0001"
        rngRange.Formula = "'" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "0001"
        
    End If
    
    PegaCodParteSemCPF = rngRange.text
    
End Function

Sub PegaInformacoesProcessoGeral(ByRef dtDataProvContestar As Date, ByRef dtDataProvSubsidios As Date, ByRef planGrupoProvContestar As Excel.Worksheet, ByRef planGrupoProvSubsidios As Excel.Worksheet)
''
'' Faz a coleta dos dados do processo que n�o dependem de sistema (PJe ou Projudi), armazena num objeto da classe Processo e devolve este objeto.
'' Deve ser rodada ap�s a fun��o espec�fica do sistema (PJe ou Projudi)
''
    Dim contProvidencias As providencia
    Dim contPedidos As Pedido
    Dim form As frmProcesso
    Dim varCont As Variant
    Dim Cont As Integer, Cont2 As Integer
    Dim strUnidadeDiv As String, arrNucleoPrep() As String, strObservacao As String, strNomePedido As String
    Dim strSci(1 To 5, 1 To 2) As String, strMatricula As String
    Dim dtContData As Date
    Dim btControlesPorLinha As Byte, btControlesPreexistentes As Byte, btQtdLinhas As Byte
    Dim bolPularPedido As Boolean, bolTemOutrosReus As Boolean
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a provid�ncia de agendar pautista
    
    '''''''''''''''''''''''
    ''' Comarca e �rg�o '''
    '''''''''''''''''''''''
    
    prProcesso.Comarca = PegaComarca(prProcesso.Juizo)
    prProcesso.Orgao = PegaOrgao(prProcesso.Juizo)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Exceto para PJe TRT5, consulta SCI: Matr�cula, '''
    ''' titularidade, EL vinculado, outros processos   '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If LCase(cfConfigura��es.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula) <> "pje 1g trt5" Then
        strMatricula = PegaMatricula(strSci(), strObservacao)
        If prProcesso Is Nothing Then Exit Sub
    End If
    
    
    '''''''''''''''''''''''''
    ''' Mostra formul�rio '''
    '''''''''''''''''''''''''
    
    ' Caso seja trabalhista, verifica se tem outros r�us
    If prProcesso.Tribunal = sfTRT5 Then
        For Cont = 1 To prProcesso.OutrosParticipantes.Count Step 1
            If prProcesso.OutrosParticipantes(Cont).CondicaoParte = "R�u" Then
                bolTemOutrosReus = True
                Exit For
            End If
        Next Cont
    End If
    
    ' Mostra formul�rio. Exibe n�mero do processo e nome do autor para confer�ncia.
    ' IMPORTANTE: O n�cleo da vari�vel global prProcesso � definido dentro do evento 'Change' do controle cmbCausaPedir do formul�rio.
    Set form = New frmProcesso
    With form
        .txtNumProc = prProcesso.NumeroProcesso
        If .txtNumProc.text <> "" Then .AjustarLegendaSemTransicao .LabelNumProc
        .txtAutor = prProcesso.OutrosParticipantes(1).NomeParte
        If .txtAutor.text <> "" Then .AjustarLegendaSemTransicao .LabelAutor
        .txtMatricula = strMatricula
        If .txtMatricula.text <> "" Then .AjustarLegendaSemTransicao .LabelMatricula
        .txtCodLocal = strSci(3, 2)
        If .txtCodLocal.text <> "" Then .AjustarLegendaSemTransicao .LabelCodLocal
        .cmbAndamento.Value = prProcesso.NomeAndamento
        If .cmbAndamento.Value <> "" Then .AjustarLegendaSemTransicao .LabelAndamento
        If .cmbTercProprio.Visible = True Then
            .cmbTercProprio.Value = IIf(bolTemOutrosReus = True, "Terceirizado", "Pr�prio")
            .AjustarLegendaSemTransicao .LabelTercProprio
        End If
        .txtDataAndamento.text = IIf(prProcesso.DataAndamento <> 0, prProcesso.DataAndamento, Date)
        If Len(.txtDataAndamento.text) = 19 And Right(.txtDataAndamento.text, 3) = ":00" Then .txtDataAndamento.text = Left(.txtDataAndamento.text, 16) ' N�o exibe o ":00" dos segundos
        If .txtDataAndamento.text <> "" Then .AjustarLegendaSemTransicao .LabelDataAndamento
        .cmbCausaPedir.List = cfCausasPedir.Range("CausasPedir").Value
        .cmbTipoAcao.Value = prProcesso.TipoAcao
        If .cmbTipoAcao.Value <> "" Then .AjustarLegendaSemTransicao .LabelTipoAcao
        .cmbTipoAcao.List = cfRitos.Range("TiposAcao").Value
        .cmbRito.Value = prProcesso.Rito
        If .cmbRito.Value <> "" Then .AjustarLegendaSemTransicao .LabelRito
        .cmbRito.List = cfRitos.Range("Ritos").Value
        If prProcesso.gerencia <> "PPJCM" Then .txtMatricula.Enabled = False
        .Show
    End With
    
    If prProcesso Is Nothing Or form.chbDeveGerar.Value = False Then GoTo EncerrarCadastro
    If Trim(prProcesso.CausaPedir) = "" Or Trim(prProcesso.gerencia) = "" Then GoTo EncerrarCadastro
    
    '''''''''''''''''''''''''''''''''''''''
    ''' Andamento principal e adicional '''
    '''''''''''''''''''''''''''''''''''''''
    
    If IsDate(Trim(form.txtDataAndamento.text)) Then
        prProcesso.DataAndamento = CDate(Trim(form.txtDataAndamento.text))
    Else
        prProcesso.DataAndamento = Date
    End If
    prProcesso.NomeAndamento = form.cmbAndamento.Value
    
    If SisifoEmbasaFuncoes.ControleExiste(form, "cmbAndamento2") And SisifoEmbasaFuncoes.ControleExiste(form, "txtDataAndamento2") Then
        If Trim(form.Controls("cmbAndamento2").Value) <> "" And IsDate(Trim(form.Controls("txtDataAndamento2").text)) Then
            prProcesso.DataAndamentoAdicional = CDate(Trim(form.Controls("txtDataAndamento2").text))
            prProcesso.NomeAndamentoAdicional = form.Controls("cmbAndamento2").Value
            prProcesso.ObsAndamentoAdicional = form.Controls("txtObsAndamento2").text
        Else
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o andamento adicional n�o tem um nome ou uma data v�lida, e por isso seu humilde " & _
            "e leal servo n�o poder� cadastr�-lo. Caso deseje cadastrar o processo normalmente sem andamento adicional, clique em ""OK"". Para cancelar " & _
            "o cadastro, clique em ""Cancelar"".", vbQuestion + vbOKCancel, "S�sifo - Erro no andamento adicional") = vbCancel Then GoTo EncerrarCadastro
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Causa de pedir, ger�ncia, natureza (ramo do direito) '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.CausaPedir = form.cmbCausaPedir.Value
    prProcesso.gerencia = "PPJCM"
    prProcesso.Natureza = PegaNatureza(cfCausasPedir, prProcesso.CausaPedir)
    
    
    ''''''''''''''''''''''''''''
    ''' Advogado respons�vel '''
    ''''''''''''''''''''''''''''
    prProcesso.Advogado = Trim(form.cmbAdvogado.Value)
    ' Se tiver mudado advogado para escrit�rio externo, muda o n�cleo
    If prProcesso.gerencia = "PPJCM" And prProcesso.Advogado = "ESCRITORIO: SANTOS NETO E BOA SORTE" Then
        prProcesso.Nucleo = "PPJCE - Cont. Especial - Santos Neto"
        prProcesso.Advogado = "JORGE KIDELMIR NASCIMENTO DE OLIVEIRA FILHO"
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''' Matr�cula e bloquear matr�cula   '''
    ''''''''''''''''''''''''''''''''''''''''
    
    If form.txtMatricula.text = "" Then
        prProcesso.Matricula = 0
        prProcesso.BloqueiaMatricula = False
    Else
        prProcesso.Matricula = CLng(Trim(form.txtMatricula.text))
        prProcesso.BloqueiaMatricula = PegaBloqueiaMatricula(Right(prProcesso.NumeroProcesso, 4))
    End If
    
    ''''''''''''''''''''
    ''' Unidade / EL '''
    ''''''''''''''''''''
    
    ' Coloca Unidade. Se n�o houver o n�mero de localidade na planilha, alerta.
    If form.txtCodLocal.Visible = True And form.txtCodLocal.text <> "" Then
        strUnidadeDiv = PegaUnidadeDivisao(cfUnidadesELs, form.txtCodLocal.text)
        If strUnidadeDiv = "" Then
            MsgBox "O n�mero da localidade digitado n�o foi encontrado em nossa base. Favor inserir o nome da " & _
                    "Unidade manualmente na planilha antes de importar para o Espaider. ATEN��O: O nome da Unidade " & _
                    "precisa ser copiado e colado do Espaider, pois caso seja diferente, a importa��o deste processo " & _
                    "resultar� em erro.", vbCritical + vbOKOnly, "S�sifo - C�digo de localidade n�o encontrado"
            
            GoTo EncerrarCadastro
        Else
            prProcesso.Unidade = Left(strUnidadeDiv, InStr(1, strUnidadeDiv, "///") - 1)
        End If
        
        ' Coloca escrit�rio local. Se o EL estiver em branco mas a unidade n�o (e a unidade digitada for num�rica, isto �:
        ' n�o for um dos c�digos diretos, como "UMC", "UNA", etc), coloca no log pra adicionar � lista depois.
        prProcesso.Divisao = Right(strUnidadeDiv, Len(strUnidadeDiv) - InStr(1, strUnidadeDiv, "///") - 2)
        If IsNumeric(form.txtCodLocal.text) And prProcesso.Unidade <> "" And prProcesso.Divisao = "" Then
            cfLogDivisoes.Cells(1, 1).End(xlDown).Offset(1, 0).Formula = form.txtCodLocal
        End If
    End If
    
    ''''''''''''''''''''
    ''' Provid�ncias '''
    ''''''''''''''''''''
    ' Insere a observa��o do cadastro
    If form.txtObs.text <> "" Then strObservacao = "Obs. do cadastro: " & form.txtObs.text & vbCr & strObservacao
    strObservacao = "S�sifo: " & prProcesso.NomeAndamento & " - " & prProcesso.DataAndamento & vbCr & strObservacao
    
    ' Adiciona provid�ncia de Contestar
    Set contProvidencias = CriarProvidenciaContestar(CDate(Trim(form.txtDataAndamento.text)), strObservacao)
    If Not contProvidencias Is Nothing Then
        prProcesso.Providencias.Add contProvidencias
        dtDataProvContestar = contProvidencias.DataFinal
        If form.cmbAdvogado.Tag <> "" Then Set planGrupoProvContestar = ThisWorkbook.Sheets("cf" & form.cmbAdvogado.Tag)
    End If
    
    ' Adiciona provid�ncia de Analisar processo novo
    Set contProvidencias = CriarProvidenciaAnalisarProcesso(strObservacao, form.chbAnalisarProcNovo.Value)
    If Not contProvidencias Is Nothing Then prProcesso.Providencias.Add contProvidencias
    
    ' Adiciona provid�ncia de Levantar subs�dios
    ' Insere Levantar Subs�dios
    Set contProvidencias = CriarProvidenciaLevantarSubsidios(dtDataProvContestar, strObservacao, arrNucleoPrep)
    If Not contProvidencias Is Nothing Then
        prProcesso.Providencias.Add contProvidencias
        dtDataProvSubsidios = contProvidencias.DataFinal
        If arrNucleoPrep(3) <> "" Then Set planGrupoProvSubsidios = ThisWorkbook.Sheets("cf" & Trim(arrNucleoPrep(3)))
    End If
    
    ' Se houver provid�ncia adicional, cadastra como a primeira
    Set contProvidencias = CriarProvidenciaAdicional(form)
    If Not contProvidencias Is Nothing Then prProcesso.Providencias.Add contProvidencias, Before:=1
    
    If prProcesso Is Nothing Then GoTo EncerrarCadastro
    
    ' Se for Projudi de Lauro de Freitas/BA, adiciona provid�ncia de Triar Senten�a para 20 dias depois da audi�ncia
    If prProcesso.DataAndamento <> 0 And prProcesso.Comarca = "Lauro de Freitas" And prProcesso.Sistema = sfProjudi Then
        Set contProvidencias = New providencia
        contProvidencias.Nome = "Triar Senten�a / Ac�rd�o e atualizar Provisionamento"
        contProvidencias.Alertar = False
        contProvidencias.DataAlerta = CDate(prProcesso.DataAndamento + 19)
        contProvidencias.DataFinal = CDate(prProcesso.DataAndamento + 20)
        contProvidencias.Nucleo = prProcesso.Nucleo
        contProvidencias.Responsavel = prProcesso.Advogado
        contProvidencias.Observacao = "Provid�ncia criada pelo S�sifo, 20 dias ap�s audi�ncia - intima��es pr�-datadas de Lauro de Freitas"
        contProvidencias.GerarAndamento = False
        prProcesso.Providencias.Add contProvidencias
    End If
    
    '' Se for > 20 sal�rios de comarcas configuradas para autoagendar pautista, adiciona provid�ncia de Agendar advogado
    'bolMaior20SM = IIf(prProcesso.ValorCausa > 20 * cfConfigura��es.Cells().Find(what:="Valor do sal�rio m�nimo", lookat:=xlWhole).Offset(0, 1).Formula, True, False)
    'bolAgendaPautista = AutoAgendaPautista(Right(prProcesso.NumeroProcesso, 4), Left(prProcesso.Nucleo, 5))
    'If bolMaior20SM = True And bolAgendaPautista = True And prProcesso.DataAndamento <> 0 Then
    '    Set contProvidencias = New Providencia
    '    contProvidencias.Nome = "Agendar advogado pautista"
    '    contProvidencias.Alertar = False
    '    contProvidencias.DataAlerta = Application.WorksheetFunction.WorkDay(Date, 2)
    '    contProvidencias.DataFinal = Application.WorksheetFunction.WorkDay(prProcesso.DataAndamento, -2)
    '    contProvidencias.Nucleo = prProcesso.Nucleo
    '    contProvidencias.Responsavel = cfConfigura��es.Cells().Find(what:="Respons�vel pelo agendamento de pautista", lookat:=xlWhole).Offset(0, 1).Formula
    '    contProvidencias.Observacao = "Audi�ncia Una - " & prProcesso.DataAndamento
    '    contProvidencias.GerarAndamento = False
    '    prProcesso.Providencias.Add contProvidencias
    'End If
    
    '''''''''''''''
    ''' Pedidos '''
    '''''''''''''''
    
    ' Descobre quantas linhas de pedidos existem no formul�rio
    btControlesPorLinha = 5
    btControlesPreexistentes = 4
    btQtdLinhas = (form.Controls("fraPedidos").Controls.Count - btControlesPreexistentes) / btControlesPorLinha
    
    ' Itera os pedidos existentes, colocando-os no objeto do processo (prProcesso).
    For Cont = 1 To btQtdLinhas Step 1
    
        If Trim(form.Controls("cmbPed" & Cont).Value) = "" Then
            strNomePedido = ""
        Else
            strNomePedido = cfPedidos.Cells().Find(What:=form.Controls("cmbPed" & Cont).Value, LookAt:=xlWhole).Offset(0, -1).Formula
        End If
        
        ' Se o nome do pedido for em branco, pula
        If strNomePedido = "" Then bolPularPedido = True
        
        ' Se j� tiver sido adicionado pedido com esse nome, pula (pedidos com mesmo no mesmo processo podem causar bug na interpreta��o da planilha pelo Espaider)
        For Cont2 = prProcesso.Pedidos.Count To 1 Step -1
            If strNomePedido = prProcesso.Pedidos(Cont2).Nome Then bolPularPedido = True
        Next Cont2
            
        If bolPularPedido = False Then
            ReDim varCont(1 To 5)
            varCont(1) = cfPedidos.Cells().Find(What:=form.Controls("cmbPed" & Cont).Value, LookAt:=xlWhole).Offset(0, -1).Formula ' C�digo do pedido
            varCont(2) = form.Controls("cmbPed" & Cont).Value  ' Nome do pedido
            varCont(3) = CCur(IIf(Trim(form.Controls("txtPed" & Cont).text) = "", 0, form.Controls("txtPed" & Cont).text)) ' Valor pedido; se for vazio, coloca 0
            varCont(4) = IIf(Trim(form.Controls("cmbRisco" & Cont).Value) = "", "Remoto", Trim(form.Controls("cmbRisco" & Cont).Value)) ' Risco; se for vazio, coloca Remoto
            
            If prProcesso.gerencia = "PPJCT" Then
                If Trim(form.Controls("txtProv" & Cont).text) = "" Or Trim(form.Controls("txtProv" & Cont).text) = 0 Then
                    varCont(5) = CCur(prProcesso.ValorCausa / btQtdLinhas) ' Valor do risco a provisionar; se for trabalhista e estiver zerado, divide valor da causa por igual
                Else
                    varCont(5) = CCur(form.Controls("txtProv" & Cont).text) ' Valor do risco a provisionar; se for trabalhista e n�o estiver zerado, usa o valor que est� no formul�rio
                End If
            Else
                varCont(5) = CCur(IIf(Trim(form.Controls("txtProv" & Cont).text) = "", 0, form.Controls("txtProv" & Cont).text)) ' Valor do risco a provisionar; se for vazio, coloca 0
            End If
            
            Set contPedidos = PegaPedido(prProcesso.CausaPedir, CStr(varCont(1)), CStr(varCont(2)), CCur(varCont(3)), CStr(varCont(4)), CCur(varCont(5)))
            prProcesso.Pedidos.Add contPedidos
        End If
    Next Cont
    
    If prProcesso Is Nothing Then
EncerrarCadastro:
        Set prProcesso = Nothing
        Unload form
        Set form = Nothing
    End If

End Sub

Function PegaBloqueiaMatricula(strCodComarca As String) As Boolean
''
'' Confere se as matr�culas de uma comarca devem ser bloqueadas para cobran�a ou negativa��o.
''  Retorna Verdadeiro caso devam.
''
    Dim rngCont As Range
    
    Select Case LCase(prProcesso.gerencia)
    Case "ppjcm"
        Set rngCont = cfComarcasBloqMatricula.Range("BloqMatriculaPPJCM").Find(strCodComarca, LookAt:=xlWhole)
    Case "ppjce"
        Set rngCont = cfComarcasBloqMatricula.Range("BloqMatriculaPPJCE").Find(strCodComarca, LookAt:=xlWhole)
    End Select
    
    If rngCont Is Nothing Then
        ' Se n�o achar o C�digo da comarca na planilha, retorna Falso e sai da fun��o.
        PegaBloqueiaMatricula = False
    Else
        ' Se achar o c�digo da comarca, retorna Verdadeiro ou Falso conforme Sim ou N�o e sai da fun��o.
        PegaBloqueiaMatricula = IIf(rngCont.Offset(0, 5).Formula = "Sim", True, False)
    End If

End Function

Function PegaNatureza(plRelacaoCausaDePedirNatureza As Worksheet, strCausaPedir As String) As String

    Dim rngCont As Excel.Range

    Set rngCont = plRelacaoCausaDePedirNatureza.Cells().Find(What:=strCausaPedir, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu n�o conhe�o a natureza da causa de pedir deste processo (""" & strCausaPedir & """). " & _
        "Rogo que a cadastre em minha mem�ria e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Natureza da causa de pedir desconhecida"
        PegaNatureza = "Natureza da causa de pedir n�o cadastrada"
    Else
        PegaNatureza = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaGerencia(plRelacaoCausaDePedirGerencia As Worksheet, strCausaPedir As String) As String

    Dim rngCont As Excel.Range

    Set rngCont = plRelacaoCausaDePedirGerencia.Cells().Find(What:=strCausaPedir, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu n�o conhe�o a ger�ncia relacionada � causa de pedir deste processo (""" & strCausaPedir & """). " & _
        "Rogo que a cadastre em minha mem�ria e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Ger�ncia da causa de pedir desconhecida"
        PegaGerencia = "Ger�ncia da causa de pedir n�o cadastrada"
    Else
        PegaGerencia = rngCont.Offset(0, 3).Formula
    End If
    
End Function

Function PegaComarca(ByRef strJuizo As String) As String
    
    Dim rngCont As Excel.Range
    
    Set rngCont = cfJuizos.Range("B:B").Find(What:=strJuizo, LookAt:=xlWhole)
    
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", n�o h� registro da comarca dos processo que tramitam no ju�zo (""" & strJuizo & """) na minha mem�ria. " & _
            "Rogo que o cadastre em minha mem�ria e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Ju�zo n�o cadastrado"
        PegaComarca = "Ju�zo n�o cadastrado, imposs�vel encontrar comarca"
    Else
        PegaComarca = Trim(rngCont.Offset(0, 1).Formula)
    End If
    
End Function

Function PegaOrgao(ByRef strJuizo As String) As String
    
    Dim rngCont As Excel.Range
    
    Set rngCont = cfJuizos.Range("B:B").Find(What:=strJuizo, LookAt:=xlWhole)
    
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", n�o h� registro da comarca dos processo que tramitam no ju�zo (""" & strJuizo & """) na minha mem�ria. " & _
            "Rogo que o cadastre em minha mem�ria e tente novamente.", vbCritical + vbOKOnly, "S�sifo - Ju�zo n�o cadastrado"
        PegaOrgao = "Ju�zo n�o cadastrado, imposs�vel encontrar comarca"
    Else
        PegaOrgao = Trim(rngCont.Offset(0, 2).Formula)
    End If
    
End Function

Function PegaMatricula(ByRef strSci() As String, ByRef strObservacao As String) As String
    Dim strMatricula As String, strProcessosMatricula() As String, strCpjCnpjPje As String
    Dim intResposta As Integer
    Dim bolAutorEhTitular As Boolean
    
PerguntaMatricula:
    strMatricula = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", favor informar a matr�cula relacionada ao processo (se n�o for o caso, deixe o campo em branco)", "S�sifo - Matr�cula de processo", ""))
    
    ' Se tiver matr�cula, resgata c�digo da localidade e nome da parte pelo web service do SCI
    If strMatricula <> "" Then
        strSci(1, 1) = "erroCodigo"
        strSci(2, 1) = "erroMensagem"
        strSci(3, 1) = "inscricao"
        strSci(4, 1) = "numCPFCNPJ"
        strSci(5, 1) = "nomeCliente"
        ConsultaMatriculaSCI strMatricula, strSci
        
        If strSci(1, 2) <> 0 Then ' Se deu erro na consulta ao web service do SCI, pergunta se o usu�rio quer continuar assim mesmo.
            intResposta = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o SCI retornou o seguinte erro: """ & strSci(2, 2) & """. Deseja continuar mesmo sem a integra��o com o SCI?", _
                vbYesNo + vbQuestion + vbDefaultButton2, "S�sifo - Erro de integra��o com o SCI")
            
            If intResposta = vbYes Then
                strObservacao = "ATEN��O: Este cadastro foi feito sem integra��o autom�tica com o SCI." & vbCr
                PegaMatricula = ""
                Exit Function
            Else
                Set prProcesso = Nothing
                Exit Function
            End If
        End If
        
        ' Confere se o Autor � o titular da matr�cula
        strCpjCnpjPje = IIf(prProcesso.OutrosParticipantes(1).TipoParte = 1, prProcesso.OutrosParticipantes(1).CPFParte, prProcesso.OutrosParticipantes(1).CNPJParte)
        strCpjCnpjPje = Replace(strCpjCnpjPje, ".", "")
        strCpjCnpjPje = Replace(strCpjCnpjPje, "-", "")
        strCpjCnpjPje = Replace(strCpjCnpjPje, "/", "")
        
        ' Se o CPF/CNPJ n�o estiver cadastrado no processo, compara pelo nome
        If InStr(1, LCase(strCpjCnpjPje), "n�o cadastrado") <> 0 Then
            bolAutorEhTitular = IIf(Trim(LCase(strSci(5, 2))) = Trim(LCase(prProcesso.OutrosParticipantes(1).NomeParte)), True, False)
            If bolAutorEhTitular = False Then
                If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", aparentente, o nome do Autor do processo (" & Trim(prProcesso.OutrosParticipantes(1).NomeParte) & _
                ") � diferente do nome do titular da matr�cula (" & strSci(5, 2) & "). Favor conferir se a matr�cula do processo � realmente a que " & _
                "foi digitada -- " & strMatricula & ". Para prosseguir com a matr�cula informada, clique em ""OK"". Para informar uma nova matr�cula, " & _
                "clique em ""Cancelar"".", vbQuestion + vbOKCancel, "S�sifo - Autor diferente do titular da matr�cula") = vbCancel Then _
                GoTo PerguntaMatricula
            End If
        Else
        'Se o CPF/CNPJ estiver cadastrado no processo, compara pelo CPF/CNPJ mesmo
            bolAutorEhTitular = IIf(Trim(LCase(strSci(4, 2))) = strCpjCnpjPje, True, False)
            If bolAutorEhTitular = False Then
                If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", aparentente, o CPF/CNPJ do Autor do processo (" & Trim(strCpjCnpjPje) & ") " & _
                "� diferente do CPF/CNPJ do titular da matr�cula (" & strSci(4, 2) & "). Favor conferir se a matr�cula do processo � realmente a que " & _
                "foi digitada -- " & strMatricula & ". Para prosseguir com a matr�cula informada, clique em ""OK"". Para informar uma nova matr�cula, " & _
                "clique em ""Cancelar"".", vbQuestion + vbOKCancel, "S�sifo - Autor diferente do titular da matr�cula") = vbCancel Then _
                GoTo PerguntaMatricula
            End If
        End If
        
        ' Ajusta o c�digo da localidade
        strSci(3, 2) = Left(strSci(3, 2), InStr(1, strSci(3, 2), ".") - 1)
        strSci(3, 2) = Format(strSci(3, 2), "0000") ' Pega o c�digo da localidade
        
        ' Ajusta observa��o das provid�ncias
        If bolAutorEhTitular = False Then strObservacao = "O Autor n�o � titular da matr�cula" & vbCr
        
        ' Se tiver matr�cula, resgata processos em COPJ pelo web service do SCI
        strProcessosMatricula = ConsultaProcessosPorMatriculaSCI(strMatricula)
        
        ' Resgata informa��o sobre outros processos cadastrados
        If UBound(strProcessosMatricula) <> 0 Then
            strObservacao = strObservacao & "Existem outros processos da matr�cula:" & vbCr
            For Cont = 1 To UBound(strProcessosMatricula, 1)
                strObservacao = strObservacao & "   - N�mero: " & strProcessosMatricula(Cont, 1) & "; Autor: " & strProcessosMatricula(Cont, 2) & "; Causa de pedir: " & strProcessosMatricula(Cont, 3) & vbCr
            Next Cont
        Else
            strObservacao = strObservacao & "N�o existem outros processos cadastrados na matr�cula" & vbCr
        End If
    End If
    
    PegaMatricula = strMatricula
    
End Function

Function PegaUnidadeDivisao(plan As Worksheet, strCodLocal As String) As String

    Dim rngCont As Range
    
    Set rngCont = plan.Cells().Find(strCodLocal)
    If rngCont Is Nothing Then
        PegaUnidadeDivisao = ""
        Exit Function
    End If

    PegaUnidadeDivisao = rngCont.Offset(0, 2).Formula & "///" & rngCont.Offset(0, 4).Formula '/// � o separador

End Function

Function PegaNomePedido(ByVal strCodPedido As String) As String

    Dim rngCont As Excel.Range
    
    Set rngCont = cfPedidos.Cells().Find(What:=strCodPedido, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        PegaNomePedido = "Pedido n�o cadastrado"
    Else
        PegaNomePedido = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaPedido(strCausaPedir As String, strCodPedido As String, strPedido As String, curValorPedido As Currency, strRisco As String, curValorProvisionar As Currency) As Pedido
''
'' Busca o progn�stico e valores padr�o para combina��es de causa de pedir e pedido. N�o encontrando,
'' retorna sempre progn�stico remoto e valores zerados.
''
    Dim rngCont As Range
    Dim strPrimeiroEndereco As String
    Dim pePedido As Pedido
    
    Set pePedido = New Pedido
    
    Set rngCont = cfCausasPedirPedidos.Cells().Find(strCausaPedir)
    strPrimeiroEndereco = rngCont.Address
    
    Do
        If rngCont.Formula = strCausaPedir And rngCont.Offset(0, 1).Formula = strCodPedido Then
            'Se achar a combina��o de Causa de Pedir e pedido, retorna e sai da fun��o.
            pePedido.Nome = strPedido  'Pedido
            pePedido.codigoPedido = PegarCodigoPedido(strCodPedido)
            pePedido.ValorPedido = curValorPedido ' Valor pedido
            pePedido.Prognostico = strRisco  'Risco
            pePedido.ValorRisco = curValorProvisionar 'Valor a provisionar
            pePedido.DataRealizacao = DateAdd("yyyy", 1, Date) ' Data estimada do pagamento (1 ano)
            Set PegaPedido = pePedido
            Exit Function
        End If
        
        Set rngCont = cfCausasPedirPedidos.Cells().FindNext(rngCont)
    Loop Until rngCont.Address = strPrimeiroEndereco
    
    ' Se n�o achar, retorna "Remoto" com valores zerados e sai da fun��o.
    pePedido.Nome = strPedido  'Nome do Pedido
    pePedido.codigoPedido = PegarCodigoPedido(strCodPedido)
    pePedido.ValorPedido = curValorPedido ' Valor pedido
    pePedido.Prognostico = "Remoto"  'Risco
    pePedido.ValorRisco = 0  'Valor a provisionar
    pePedido.DataRealizacao = DateAdd("yyyy", 1, Date) ' Data estimada do pagamento (1 ano)
    Set PegaPedido = pePedido
    
End Function

Function PegarCodigoPedido(strCodPedido As String) As String
    Dim rngCont As Range
    
    Set rngCont = cfPedidos.Cells().Find(strCodPedido, LookAt:=xlWhole, searchorder:=xlByColumns)
    If Not rngCont Is Nothing Then
        PegarCodigoPedido = CStr(rngCont.Offset(0, 2).text)
    Else
        PegarCodigoPedido = "20"
    End If
    
End Function

Function CriarProvidenciaAdicional(ByRef oForm As SisifoEmbasaCitacoes.frmProcesso) As SisifoEmbasaCitacoes.providencia
    
    Dim contProvidencia As SisifoEmbasaCitacoes.providencia
    Dim dtData As Date
    Dim arrNucleoPrep() As String
    
    If SisifoEmbasaFuncoes.ControleExiste(oForm, "cmbProvidencia2") Then
        If Trim(oForm.Controls("cmbProvidencia2").Value) <> "" Then
            dtData = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
            
            If prProcesso.Preposto = "" Then
                ' Procura preposto
                arrNucleoPrep() = PegaNucleoAdvPrep(Preposto, prProcesso.NumeroProcesso, prProcesso.CausaPedir, dtData, False)
                
                ' Erros/resultado
                If UBound(arrNucleoPrep) = 1 Then
                    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(1) & ". A inclus�o foi cancelada e descartados os dados." & vbCrLf & _
                    "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "S�sifo - Erro na busca de n�cleo e advogado"
                    Set prProcesso = Nothing
                    Exit Function
                ElseIf Left(LCase(arrNucleoPrep(2)), 4) = "erro" Then
                    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(2) & ". A inclus�o foi cancelada e descartados os dados." & vbCrLf & _
                    "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "S�sifo - Erro na busca de n�cleo e advogado"
                    Set prProcesso = Nothing
                    Exit Function
                Else
                    prProcesso.Preposto = Trim(arrNucleoPrep(2))
                End If
            End If
            
            Set contProvidencia = New providencia
            With contProvidencia
                .Nome = Trim(oForm.Controls("cmbProvidencia2").Value)
                .Alertar = False
                .DataAlerta = Application.WorksheetFunction.WorkDay(dtData, -1, cfFeriados.Range("SisifoFeriados"))
                .DataFinal = dtData
                .Nucleo = prProcesso.Nucleo
                .Responsavel = prProcesso.Preposto
                .Observacao = Trim(oForm.Controls("txtObsProvidencia2").text)
                .GerarAndamento = True
            End With
        Else
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", a provid�ncia adicional n�o tem um nome ou uma data v�lida, e por isso seu humilde " & _
            "e leal servo n�o poder� cadastr�-la. Caso deseje cadastrar o processo normalmente sem provid�ncia adicional, clique em ""OK"". Para cancelar " & _
            "o cadastro, clique em ""Cancelar"".", vbQuestion + vbOKCancel, "S�sifo - Erro no andamento adicional") = vbCancel Then
                Set prProcesso = Nothing
                Exit Function
            End If
        End If
    End If
    
    Set CriarProvidenciaAdicional = contProvidencia
        
End Function

Function CriarProvidenciaContestar(dtData As Date, strObservacao As String) As SisifoEmbasaCitacoes.providencia
    
    Dim contProvidencia As SisifoEmbasaCitacoes.providencia
    
    Select Case prProcesso.gerencia
    Case "PPJCM"
        If dtData <= Date Then
            dtData = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
        Else
            dtData = Application.WorksheetFunction.WorkDay(CDate(dtData), -2, cfFeriados.Range("SisifoFeriados"))
        End If
        
        Set contProvidencia = New providencia
        With contProvidencia
            .Nome = "Contestar - Virtual"
            .Alertar = False
            .DataAlerta = Application.WorksheetFunction.WorkDay(prProcesso.DataAndamento, -1, cfFeriados.Range("SisifoFeriados"))
            .DataFinal = dtData
            .Nucleo = prProcesso.Nucleo
            .Responsavel = prProcesso.Advogado
            .Observacao = strObservacao
            .GerarAndamento = True
        End With
        
    Case "PPJCT"
        If dtData <= Date Then
            dtData = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
        Else
            dtData = Application.WorksheetFunction.WorkDay(CDate(dtData), -2, cfFeriados.Range("SisifoFeriados"))
        End If
        
        Set contProvidencia = New providencia
        With contProvidencia
            .Nome = "Contestar"
            .Alertar = False
            .DataAlerta = Application.WorksheetFunction.WorkDay(prProcesso.DataAndamento, -1, cfFeriados.Range("SisifoFeriados"))
            .DataFinal = dtData
            .Nucleo = prProcesso.Nucleo
            .Responsavel = prProcesso.Advogado
            .Observacao = strObservacao
            .GerarAndamento = True
        End With
        
    Case Else ' PPJCE pediu para n�o criar Contestar, e sim Analisar processo novo.
        Set contProvidencia = Nothing
        
    End Select
    
    Set CriarProvidenciaContestar = contProvidencia
        
End Function

Function CriarProvidenciaAnalisarProcesso(strObservacao As String, gerarProvidencia As Boolean) As SisifoEmbasaCitacoes.providencia
    Dim contProvidencia As SisifoEmbasaCitacoes.providencia
    
    If gerarProvidencia Then
        Set contProvidencia = New providencia
        With contProvidencia
            .Nome = "Analisar processo novo"
            .Alertar = False
            .DataAlerta = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
            .DataFinal = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
            .Nucleo = prProcesso.Nucleo
            .Responsavel = prProcesso.Advogado
            .Observacao = strObservacao
            .GerarAndamento = False
        End With
    Else
        Set contProvidencia = Nothing
    End If
    
    Set CriarProvidenciaAnalisarProcesso = contProvidencia
        
End Function

Function CriarProvidenciaLevantarSubsidios(ByRef dtDataProvContestar As Date, strObservacao As String, ByRef arrNucleoPrep() As String) As SisifoEmbasaCitacoes.providencia
    
    Dim dtData As Date
    Dim contProvidencia As SisifoEmbasaCitacoes.providencia
    
    ' Ajusta a data
    dtData = Application.WorksheetFunction.WorkDay(dtDataProvContestar, -20, cfFeriados.Range("SisifoFeriados"))
    If dtData <= Date Then dtData = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
    
    ' Pega preposto, se for PPJCE ou PPJCM
    If prProcesso.gerencia = "PPJCM" Or prProcesso.gerencia = "PPJCE" Then
        ' Procura preposto
        arrNucleoPrep() = PegaNucleoAdvPrep(Preposto, prProcesso.NumeroProcesso, prProcesso.CausaPedir, dtData, True)
        
        ' Erros/resultado
        If UBound(arrNucleoPrep) = 1 Then
            MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(1) & ". A inclus�o foi cancelada e descartados os dados." & vbCrLf & _
            "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "S�sifo - Erro na busca de n�cleo e advogado"
            Set prProcesso = Nothing
            Exit Function
        ElseIf Left(LCase(arrNucleoPrep(2)), 4) = "erro" Then
            MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(2) & ". A inclus�o foi cancelada e descartados os dados." & vbCrLf & _
            "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "S�sifo - Erro na busca de n�cleo e advogado"
            Set prProcesso = Nothing
            Exit Function
        Else
            prProcesso.Preposto = Trim(arrNucleoPrep(2))
        End If
    End If
    
    ' Ajusta as observa��es da provid�ncia
    Select Case prProcesso.gerencia
    Case "PPJCT"
        ReDim arrNucleoPrep(1 To 3)
        arrNucleoPrep(1) = ""
        arrNucleoPrep(2) = ""
        arrNucleoPrep(3) = ""
        
        strObservacao = strObservacao & "Processo: " & prProcesso.NumeroProcesso & vbCr
        strObservacao = strObservacao & "Reclamante: " & prProcesso.OutrosParticipantes(1).NomeParte & vbCr
        strObservacao = strObservacao & "Respons�vel: " & prProcesso.Advogado & vbCr
        strObservacao = strObservacao & "Ju�zo: " & prProcesso.Juizo & vbCr
        strObservacao = strObservacao & "Data da solicita��o: " & Format(Date, "dd/mm/yyyy") & vbCr
        strObservacao = strObservacao & vbCr
        strObservacao = strObservacao & "SUBS�DIOS RELEVANTES:" & vbCr
        strObservacao = strObservacao & "- Contrato de trabalho" & vbCr
        strObservacao = strObservacao & "- Ficha de Registro" & vbCr
        strObservacao = strObservacao & "- Ficha funcional e financeira" & vbCr
        strObservacao = strObservacao & "- Folha de ponto �ltimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Contracheques �ltimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante de f�rias �ltimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante do pagamento do 13�: [ ] 1� parcela / [ ] 2� parcela" & vbCr
        strObservacao = strObservacao & "- Termo de Compensa��o de Jornada" & vbCr
        strObservacao = strObservacao & "- C�pia de atestados m�dicos e/ou Concess�o de Afastamento Previdenci�rio" & vbCr
        strObservacao = strObservacao & "- C�pia de advert�ncia e suspens�es aplicadas ao Reclamante" & vbCr
        strObservacao = strObservacao & "- RIP" & vbCr
        strObservacao = strObservacao & "- Gratifica��o Motorista Usu�rio" & vbCr
        strObservacao = strObservacao & "- Avalia��es PCCS 2009" & vbCr
        strObservacao = strObservacao & "- Acordo Coletivo" & vbCr
        strObservacao = strObservacao & "- Plano de Demiss�o Volunt�ria" & vbCr
        strObservacao = strObservacao & "- PCCS/2009 e a homologa��o publicada no DOU" & vbCr
        strObservacao = strObservacao & "- PCCS/1986" & vbCr
        strObservacao = strObservacao & "- PCCS/1998" & vbCr
        strObservacao = strObservacao & "- Comprovante de pagamento do aux�lio-creche ou da justificativa para o indeferimento do beneficio" & vbCr
        strObservacao = strObservacao & "- C�pia da reclama��o inicial e da senten�a de OUTRO PROCESSO SE EXISTIR" & vbCr
        strObservacao = strObservacao & "- C�pia da homologa��o do acordo" & vbCr
        strObservacao = strObservacao & "- C�pia da peti��o de acordo" & vbCr
        strObservacao = strObservacao & "- Processo Administrativo da demiss�o, Termo de Rescis�o do contrato de trabalho, e demais comprovantes de pagamento da rescis�o" & vbCr
        strObservacao = strObservacao & "- Comprovante de entrega das guias de seguro desemprego" & vbCr
        strObservacao = strObservacao & "- Recibo de recolhimento de FGTS dos �ltimos 30 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante ades�o ao PAT" & vbCr
        strObservacao = strObservacao & "- BDV's (Boletins di�rios de ve�culos)" & vbCr
        strObservacao = strObservacao & "- C�pia do RMU (Regulamento de Motorista Usu�rio)" & vbCr
        strObservacao = strObservacao & "- C�pia de pagamento de di�rias para viagens" & vbCr
        strObservacao = strObservacao & "- Comprovante de pagamento da indeniza��o por acidente de trabalho" & vbCr
        strObservacao = strObservacao & "- Instrumento que traz as normas do PPR" & vbCr
        strObservacao = strObservacao & "OBS.: "
        
    Case "PPJCM"
        If InStr(1, LCase(prProcesso.CausaPedir), "desabastecimento") > 0 Then
            strObservacao = strObservacao & "Nota t�cnica esclarecendo APENAS os pontos favor�veis � Embasa:" & vbCr & _
                    "1. Se houve interven��o na rede no per�odo capaz de gerar desabastecimento;" & vbCr & _
                    "2. Se esse per�odo estava abrangido nos casos de manuten��o e reparo de emerg�ncia;" & vbCr & _
                    "3. Se o abastecimento poderia ser suprido por reservat�rio adequado;" & vbCr & _
                    "3. Caso o im�vel precise de reservat�rio inferior com bomba, e n�o o possui (anexar fotos, se poss�vel; se n�o, a tela do SCI);" & vbCr & _
                    "4. Se n�o houve altera��o de consumo (anexar HCON);" & vbCr & _
                    "5. Se a �gua estava cortada por outro motivo (anexar SS do corte);" & vbCr & _
                    "6. Se houve reclama��o de desabastecimento (anexar Hist�rico de SS no per�odo);"
        Else
            Select Case prProcesso.CausaPedir
            Case "Negativa��o no SPC"
                strObservacao = strObservacao & "1. Em caso de negativa de v�nculo com a Embasa, desconhecimento do contrato, etc: " & _
                        "Documentos ASSINADOS pelo autor do processo, comprovando rela��o com a Embasa, como solicita��o de liga��o " & _
                        "assinada, parcelamento assinado, notifica��o de d�bito assinada. N�o havendo estes, fazer relat�rio ASSINADO " & _
                        "explicando o resultado de visita ao im�vel para verificar se est� habitado e quem habita, comprovando atrav�s " & _
                        "v�nculo. Informa��o em OS sem assinatura geralmente � desconsiderada pelo juiz." & vbCr & _
                        "2. Quando reconhece o v�nculo, negando apenas d�bitos espec�ficos, relat�rio assinado descrevendo o ocorrido " & _
                        "trazer segundas vias das contas negativadas e, se poss�vel, segundas vias de contas posteriores com aviso de " & _
                        "exist�ncia de d�bito. Sendo indevida a negativa��o, inserir proposta de acordo."
                
            Case "Revis�o de consumo elevado"
                strObservacao = strObservacao & "Relat�rio assinado informando normalidade do hist�rico de consumo, SSs de revis�o de consumo, " & _
                "substitui��o de hidr�metro, aferi��o (se houver, anexar LAUDO do Ibametro, pois ju�zes n�o acatam SS de aferi��o) e " & _
                "tudo com fotos (se houver) e anexando as telas separadas."
                
            Case "Consumo elevado com corte"
                strObservacao = strObservacao & "Relat�rio assinado informando normalidade do hist�rico de consumo, SSs de revis�o de consumo, " & _
                "substitui��o de hidr�metro, aferi��o (se houver, anexar LAUDO do Ibametro, pois ju�zes n�o acatam SS de aferi��o) e " & _
                "tudo com fotos (se houver) e anexando as telas separadas. Notifica��o de d�bito assinada (sem assinatura, as chances s�o MUITO menores)."
                
            Case "Corte no fornecimento"
                strObservacao = strObservacao & "Relat�rio assinado explicando o ocorrido. Segundas vias de faturas cujo inadimplemento motivaram o corte. Notifica��o de d�bito " & _
                "assinada (sem assinatura, as chances s�o MUITO menores; se for o caso, esclarecer no relat�rio o porqu� de n�o estar assinada)."
                
            Case "Corte por iniciativa do cliente"
                strObservacao = strObservacao & "Relat�rio assinado explicando o ocorrido. Solicita��o de servi�o assinada ou, se foi por telefone, grava��o da liga��o (sem assinatura/grava��o, as chances s�o MUITO menores)"
                
            Case "Consumo Rateado - Medi��o Individualizada"
                strObservacao = strObservacao & "1. Relat�rio assinado explicando o ocorrido." & vbCr & _
                "2. Solicita��o de individualiza��o, ata de assembleia ou contrato espec�fico assinados pelo autor do processo;" & vbCr & _
                "3. Fotos do hidr�metro geral e das fileiras dos hidr�metros individuais;" & vbCr & _
                "4. HCON do hidr�metro geral e do hidr�metro individual da parte Autora."
                
            Case "Cobran�a de esgoto em im�vel n�o ligado � rede"
                strObservacao = strObservacao & "1. Nota t�cnica com geoweb e fotos mostrando exist�ncia de rede no local e a possibillidade de " & _
                "liga��o da resid�ncia (verificar se a resid�ncia do consumidor tem cota topogr�fica, e quaisquer outras quest�es relevantes);" & vbCr & _
                "2. Notifica��o assinada para interligar o im�vel � rede quando da implanta��o (procurar na DE, DM ou DI, se for o caso);" * vbCr & _
                "3. N�o � essencial, mas, caso esteja dispon�vel, Ordem de Servi�o da obra de implanta��o da rede no local."
                
            Case "Cobran�a de esgoto com �gua cortada"
                strObservacao = strObservacao & "Relat�rio assinado explicando o ocorrido. Anexar prova fotogr�fica, ou algum outro tipo de prova, de " & _
                "que o im�vel estava habitado no per�odo (mera observa��o nas OSs geralmente � desconsiderada pelo juiz)."
                
            Case "Fixo de esgoto"
                strObservacao = strObservacao & "Nota t�cnica assinada, com fotos demonstrando a exist�ncia de abastecimento alternativo e tabela das datas e leituras da " & _
                "medi��o do abastecimento alternativo."
                
            Case "Realizar liga��o de �gua", "Desmembramento de liga��es"
                strObservacao = strObservacao & "Nota t�cnica assinada, esclarecendo o motivo t�cnico para a liga��o n�o ter sido realizada, com PROVAS deste motivo " & _
                "(por exemplo, CONFORME O CASO, fotos do im�vel ou relat�rio de vistoria mostrando que n�o h� reserva��o adequada, ou medi��es de press�o " & _
                "com explica��o do porqu� da necessidade de reservat�rio inferior e bomba, ou relat�rio fotogr�fico de que as instala��es internas n�o s�o " & _
                "desmembradas, etc."
                
            Case "Vaz. �gua ou extravas. esgoto com danos a patrim�nio/morais", "Obra da Embasa com danos a patrim�nio/morais", _
                "Acidente com pessoa/ve�culo em buraco", "Acidente com ve�culo (colis�o ou atropelamento)"
                strObservacao = strObservacao & "Nota t�cnica e relat�rio fotogr�fico esclarecendo a verdade do incidente e eventuais medidas tomadas pela Embasa " & _
                "(sempre acompanhada das provas poss�veis)."
                
            Case "Multa por infra��o", "Suspeita de by-pass"
                strObservacao = strObservacao & "Nota t�cnica assinada descrevendo o tipo de gato realizado (bypass, furo, palheta, etc), com fotos e medidas tomadas pela Embasa."
                
            End Select
        End If
    End Select
        
    ' Cria a provid�ncia, conforme ger�ncia
    Select Case prProcesso.gerencia
    Case "PPJCM"
        If prProcesso.Preposto = "" Then
            Set contProvidencia = Nothing
        Else
            Set contProvidencia = New providencia
            With contProvidencia
                .Nome = "Levantar subs�dios"
                .Alertar = False
                .DataAlerta = Application.WorksheetFunction.WorkDay(dtData, 0, cfFeriados.Range("SisifoFeriados"))
                .DataFinal = dtData
                .Nucleo = prProcesso.Nucleo
                .Responsavel = prProcesso.Preposto
                .Observacao = strObservacao
                .GerarAndamento = False
            End With
        End If
        
    Case "PPJCE"
        If prProcesso.Preposto = "" Then
            Set contProvidencia = Nothing
        Else
            Set contProvidencia = New providencia
            With contProvidencia
                .Nome = "Levantar subs�dios"
                .Alertar = False
                .DataAlerta = 0
                .DataFinal = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
                .Nucleo = prProcesso.Nucleo
                .Responsavel = prProcesso.Preposto
                .Observacao = strObservacao
                .GerarAndamento = False
            End With
        End If
        
    Case "PPJCT"
        Set contProvidencia = New providencia
        With contProvidencia
            .Nome = "Levantar subs�dios"
            .Alertar = False
            .DataAlerta = Application.WorksheetFunction.WorkDay(dtData, 0, cfFeriados.Range("SisifoFeriados"))
            .DataFinal = dtData
            .Nucleo = prProcesso.Nucleo
            .Responsavel = prProcesso.Advogado
            .Observacao = strObservacao
            .GerarAndamento = False
        End With
        
    Case Else
        Set contProvidencia = Nothing
        
    End Select
    
    Set CriarProvidenciaLevantarSubsidios = contProvidencia
    
End Function

Function PegaNucleoAdvPrep(TipoDeResponsavel As SisifoEmbasaCitacoes.sfTipoAgente, strNumProc As String, strCausaPedir As String, dtDataPrazo As Date, bolAceitaCausaPedirSemResponsavel As Boolean) As String()
''
'' Procura a combina��o de c�digo de comarca/localidade e causa de pedir passadas como par�metro na planilha passada como par�metro.
'' Se n�o achar a combina��o exata, procura o c�digo da comarca/localidade com causa de pedir "Geral" ou em branco.
'' Retorna o nome do n�cleo, seguido da string "-,-,-" e do nome do advogado. Se n�o encontrar o c�digo da comarca (ou a linha geral
'' do c�digo da comarca, se for o caso), retorna uma mensagem de erro iniciada por "Erro em PegaNucleoAdvPrep - ".
''

    Dim plan As Worksheet
    Dim strCodBusca As String, strNucleo As String, strAdvPrep As String, strNomeGrupo As String, strPrimeiroEndereco As String
    Dim strEnderecoCodBusca As String
    Dim arrCont() As String
    Dim rngCont As Excel.Range
    
    strEnderecoCodBusca = "A:A"
    
    ' Define a planilha que cont�m os respons�veis e o c�digo de busca a ser utilizado
    Select Case TipoDeResponsavel
    Case Advogado
        Select Case LCase(prProcesso.Sistema)
        Case "projudi"
            Set plan = cfComarcasAdvsPPJCMProjudi
            strCodBusca = Right(strNumProc, 4) ' O advogado � definido pela comarca do processo (�ltimos 4 n�meros, no TJ/BA)
        Case "pje1g"
            Select Case LCase(prProcesso.gerencia)
            Case "ppjcm"
                Set plan = cfComarcasAdvsPPJCMPJe
                strCodBusca = prProcesso.Comarca ' O advogado � definido pela comarca do processo (nome da comarca)
            Case "ppjce"
                Set plan = cfComarcasAdvsPPJCEPJe
            Case "ppjct"
                Set plan = cfComarcasAdvsPPJCTPJe
            End Select
        End Select
        
    Case Preposto
        Select Case LCase(prProcesso.Sistema)
        Case "projudi"
            Set plan = cfComarcasPrepPPJCMProjudi
            strCodBusca = Trim(Replace(Left(prProcesso.Unidade, 5), "-", "")) ' O c�digo de uma unidade pode ter at� 5 letras, por isso pegamos as 5, removemos o h�fen eventual e os espa�os (o que transformar� as unidades de 3 letras de "USU -" em "USU", e tamb�m funcionar� pra as de 4 e 5 letras.)
            strCodBusca = strCodBusca & "-" & Replace(prProcesso.Divisao, "Escr. Local de", "EL") ' O formato final �, por exemplo, "USU-EL Eun�polis"
        Case "pje1g"
            Select Case LCase(prProcesso.gerencia)
            Case "ppjcm"
                Set plan = cfComarcasPrepPPJCMPJe
                strCodBusca = Trim(Replace(Left(prProcesso.Unidade, 5), "-", "")) ' O c�digo de uma unidade pode ter at� 5 letras, por isso pegamos as 5, removemos o h�fen eventual e os espa�os (o que transformar� as unidades de 3 letras de "USU -" em "USU", e tamb�m funcionar� pra as de 4 e 5 letras.)
                strCodBusca = strCodBusca & "-" & Replace(prProcesso.Divisao, "Escr. Local de", "EL") ' O formato final �, por exemplo, "USU-EL Eun�polis"
            Case "ppjce"
                Set plan = cfComarcasPrepPPJCEPJe
            'Case "ppjct"
            '    Set plan = strNomePlanilha
            End Select
        End Select
    End Select
    
InicioBusca:
    ' Busca o respons�vel pela causa de pedir E CodBusca
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, strCausaPedir, False, strCodBusca, False, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ' Busca o respons�vel s� pelo CodBusca, aceitando Causa de Pedir "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, "", True, strCodBusca, False, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    Else
        If TipoDeResponsavel = Preposto And InStr(1, strCodBusca, "-") <> 0 Then
        ' Se n�o achou uma linha geral da localidade do preposto e j� n�o est� buscando pela Unidade Regional, recome�a a buscar pela Unidade Regional
            strCodBusca = Left(strCodBusca, InStr(1, strCodBusca, "-") - 1)
            GoTo InicioBusca
        End If
    End If
    
    ' Busca o respons�vel s� pela causa de pedir, aceitando CodBusca "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, strCausaPedir, False, "", True, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ' Busca o respons�vel aceitando tudo "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, "", True, "", True, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ReDim arrCont(1 To 1)
    arrCont(1) = "Erro em PegaNucleoAdvPrep - Respons�vel n�o encontrado para o c�digo """ & strCodBusca & """ na planilha " & plan.Name & " com booleana de aceitar causa de pedir """ & bolAceitaCausaPedirSemResponsavel & """."
    PegaNucleoAdvPrep = arrCont
    Exit Function
    
Encontrado:
    
    ' N�O EST� LIDANDO COM GRUPOS DE N�CLEOS. CASO NECESS�RIO, REPROGRAMAR O C�DIGO DE ACORDO (5 LINHAS COMENTARIZADAS ABAIXO).
    ' Se for um grupo de n�cleos (nomes de grupos de n�cleos devem come�ar com o c�digo "grpnuc", dois n�meros e um h�fen),
    ' chama a fun��o que itera isonomicamente pelos membros do grupo.
    'If Left(strNucleo, 6) = "grpnuc" And IsNumeric(Mid(strNucleo, 7, 2)) And Mid(strNucleo, 9, 1) = "-" Then
    '    strNucleo = IteraNucleoAdv(strNucleo, "Nucleo")
    'End If
    
    ' Se for um grupo de advogados ou prepostos (nomes de grupos de advogados ou prepostos devem come�ar com o c�digo "grpadv" ou "grpprep", dois n�meros e um h�fen),
    ' chama a fun��o que itera isonomicamente pelos membros do grupo.
    If (Left(strAdvPrep, 6) = "grpadv" Or Left(strAdvPrep, 6) = "grprep") Then
        strNomeGrupo = strAdvPrep
        strAdvPrep = PegaResponsavelProvidenciaGrupo(ThisWorkbook.Sheets("cf" & strAdvPrep), dtDataPrazo)
    End If
    
    ReDim arrCont(1 To 3)
    arrCont(1) = strNucleo
    arrCont(2) = strAdvPrep
    arrCont(3) = strNomeGrupo
    
    PegaNucleoAdvPrep = arrCont
    
End Function

Function IteraPlanilhaNucleoAdvPrep(plan As Worksheet, strCausaPedir As String, bolAceitaCausaPedirGeral As Boolean, strCodBusca As String, bolAceitaCodBuscaGeral As Boolean, bolAceitaCausaPedirSemResponsavel As Boolean) As Excel.Range
''
'' Itera as c�lulas da coluna de Causa de Pedir da planilha passada como par�metro, buscando a combina��o de c�digo de comarca/localidade e
''  causa de pedir passadas como par�metro. Se n�o achar a combina��o exata, retorna Nothing.
''
    Dim rngColunaCausaPedir As Excel.Range, rngCel As Excel.Range
    
    Set rngColunaCausaPedir = plan.Range("C1:C" & plan.UsedRange.Rows(plan.UsedRange.Rows.Count).Row)
    
    If strCausaPedir = "" Then strCausaPedir = "/./"
    If strCodBusca = "" Then strCodBusca = "/./"
    
    For Each rngCel In rngColunaCausaPedir.Cells
        If (LCase(rngCel.Formula) = LCase(strCausaPedir) Or (bolAceitaCausaPedirGeral = True And LCase(rngCel.Formula) = "geral")) And _
            (LCase(rngCel.Offset(0, -2).Formula) = LCase(strCodBusca) Or (bolAceitaCodBuscaGeral = True And LCase(rngCel.Offset(0, -2).Formula) = "geral")) Then
            If Trim(rngCel.Offset(0, 2).Formula = "") And bolAceitaCausaPedirSemResponsavel = False Then 'Se o respons�vel estiver em branco, mas n�o for para aceitar, volta Nothing
                Set IteraPlanilhaNucleoAdvPrep = Nothing
                Exit Function
            Else
                Set IteraPlanilhaNucleoAdvPrep = rngCel.Offset(0, 1)
                Exit Function
            End If
        End If
    Next rngCel
    
End Function

Function PegaResponsavelProvidenciaGrupo(ByRef plan As Excel.Worksheet, dtDataPrazo As Date) As String
''
'' Itera a planilha do grupo e retorna o respons�vel pela provid�ncia no dia do prazo especificado
''
    Dim intTotalGeralDia As Integer
    Dim lngLinhaData As Long, lngColunaCont As Long
    Dim intCont As Integer
    Dim strAdvCont As String
    Dim dicResponsaveis As Dictionary
    Dim rngCont As Excel.Range
    
    Set dicResponsaveis = New Dictionary
    
    Set rngCont = plan.Cells().Find(dtDataPrazo, after:=Cells(6, 1), LookAt:=xlWhole, searchorder:=xlByColumns)
    If rngCont Is Nothing Then
        PegaResponsavelProvidenciaGrupo = "Erro - a data final do prazo n�o est� cadastrada na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    
    lngLinhaData = rngCont.Row
    lngColunaCont = plan.UsedRange.Columns.Count
    
    lngLinhaData = rngCont.Row
    Set rngCont = Nothing
    
    lngColunaCont = plan.Columns.Count
    lngColunaCont = plan.Cells(10, lngColunaCont).End(xlToLeft).Column
    
    ' Pegar quantidade j� distribu�da para a data final em quest�o
    intTotalGeralDia = plan.Cells(lngLinhaData, lngColunaCont).Value2
    If intTotalGeralDia = 0 Then intTotalGeralDia = 1
    
    ' Coletar advogados, calcular a Carga Di�ria Percentual de cada um e transform�-la na varia��o entre CDP e CDP pretendida
    For intCont = 1 To lngColunaCont - 2
        strAdvCont = plan.Cells(2, intCont + 1).Value2
        dicResponsaveis.Add strAdvCont, plan.Cells(lngLinhaData, intCont + 1).Value2
        dicResponsaveis(strAdvCont) = dicResponsaveis(strAdvCont) / intTotalGeralDia ' Carga di�ria percentual
        dicResponsaveis(strAdvCont) = dicResponsaveis(strAdvCont) - plan.Cells(4, intCont + 1).Value2 ' Diferen�a em rela��o � carga di�ria percentual pretendida = CDP - CDP pretendida
    Next intCont
    
    ' Remove os valores que n�o forem os menores
    Set dicResponsaveis = PegaValoresMenores(dicResponsaveis)
    
    ' Em caso de empate, pegar o total geral de provid�ncias de todos os tempos dos empatados (ponderado pela carga de trabalho da pessoa).
    If dicResponsaveis.Count > 1 Then
        For intCont = 0 To dicResponsaveis.Count - 1
            strAdvCont = dicResponsaveis.Keys(intCont)
            lngColunaCont = plan.Cells().Find(strAdvCont, LookAt:=xlWhole, searchorder:=xlByRows).Column
            dicResponsaveis(strAdvCont) = plan.Cells(5, lngColunaCont).Value / plan.Cells(3, lngColunaCont).Value
        Next intCont
        
        Set dicResponsaveis = PegaValoresMenores(dicResponsaveis)
    End If
    
    ' Em caso de novo empate, aleat�rio
    If dicResponsaveis.Count > 1 Then
        Randomize
        intCont = CInt((dicResponsaveis.Count - 1) * Rnd)
    Else
        intCont = 0
    End If
    
    ' Registra a provid�ncia para o respons�vel e retorna o valor procurado
    lngColunaCont = plan.Cells().Find(dicResponsaveis.Keys(intCont), LookAt:=xlWhole, searchorder:=xlByRows).Column
    'plan.Cells(lngLinhaData, lngColunaCont).Value = plan.Cells(lngLinhaData, lngColunaCont).Value + 1
    PegaResponsavelProvidenciaGrupo = dicResponsaveis.Keys(intCont)
    
End Function

Function RegistraResponsavelPorProvidenciaNoGrupo(ByRef plan As Excel.Worksheet, dtDataPrazo As Date, strResponsavel As String) As String
''
'' Itera a planilha do grupo e retorna o respons�vel pela provid�ncia no dia do prazo especificado
''
    Dim lngLinhaData As Long, lngColunaResponsavel As Long
    Dim rngCont As Excel.Range
    
    ' Pega a linha da data
    Set rngCont = plan.Cells().Find(dtDataPrazo, after:=Cells(6, 1), LookAt:=xlWhole, searchorder:=xlByColumns)
    If rngCont Is Nothing Then
        RegistraResponsavelPorProvidenciaNoGrupo = "A data final do prazo n�o est� cadastrada na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    lngLinhaData = rngCont.Row
    
    ' Pega a coluna do respons�vel
    Set rngCont = plan.Cells().Find(strResponsavel, LookAt:=xlWhole, searchorder:=xlByRows)
    If rngCont Is Nothing Then
        RegistraResponsavelPorProvidenciaNoGrupo = "O respons�vel n�o est� cadastrado na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    lngColunaResponsavel = rngCont.Column
    
    ' Registra a provid�ncia para o respons�vel
    plan.Cells(lngLinhaData, lngColunaResponsavel).Value = plan.Cells(lngLinhaData, lngColunaResponsavel).Value + 1
    RegistraResponsavelPorProvidenciaNoGrupo = "Sucesso"
    
End Function

Function PegaValoresMenores(ByVal Dict As Dictionary) As Dictionary
''
'' Itera uma cole��o, excluindo todos os valores que n�o sejam os menores.
''
    Dim dblMenorCargaTrabalho As Double
    Dim intCont As Integer
    
    With Dict
        'intCont = .Count
        dblMenorCargaTrabalho = .Items(.Count - 1)
        
        For intCont = .Count - 1 To 1 Step -1
            If .Items(intCont) > .Items(intCont - 1) Then
                dblMenorCargaTrabalho = .Items(intCont - 1)
                Dict.Remove .Keys(intCont)
            ElseIf .Items(intCont) < .Items(intCont - 1) Then
                dblMenorCargaTrabalho = .Items(intCont)
                Dict.Remove .Keys(intCont - 1)
            End If
        Next intCont
        
        For intCont = .Count - 1 To 1 Step -1
            If .Items(intCont) > dblMenorCargaTrabalho Then
                Dict.Remove .Keys(intCont)
            End If
        Next intCont
    End With
    
    Set PegaValoresMenores = Dict

End Function

Sub FecharExplorerZerarVariaveis(ByRef IE As InternetExplorer)
    On Error Resume Next
    IE.Quit
    Set IE = Nothing
    Set prProcesso = Nothing
    On Error GoTo 0
End Sub

Sub FecharChromeZerarVariaveis(ByRef oNavegador As Selenium.ChromeDriver)
    On Error Resume Next
    oNavegador.Quit
    Set oNavegador = Nothing
    Set prProcesso = Nothing
    On Error GoTo 0
End Sub

Function ValidaNumeros(ChaveAscii As MSForms.ReturnInteger, Optional strPermitir1 As String, Optional strPermitir2 As String, Optional strPermitir3 As String) As Boolean
''
'' Faz uma valida��o front end, s� permitindo n�meros e, caso informados, os dois caracteres passados como par�metros.
''
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'N�meros s�o sempre permitidos
        ValidaNumeros = True
    Case Else
        ValidaNumeros = False
    End Select
    
    If strPermitir1 <> "" Then
        If ChaveAscii = Asc(strPermitir1) Then ValidaNumeros = True
    End If
    
    If strPermitir2 <> "" Then
        If ChaveAscii = Asc(strPermitir2) Then ValidaNumeros = True
    End If
    
    If strPermitir3 <> "" Then
        If ChaveAscii = Asc(strPermitir3) Then ValidaNumeros = True
    End If
    
End Function

Function PerguntarArquivo(strTitulo As String, strCaminhoInicial As String, bolMultiselecao As Boolean) As String
    Dim objPopup As FileDialog
    Dim strArquivo As String
    
    Set objPopup = Application.FileDialog(msoFileDialogFilePicker)
    With objPopup
        .Title = strTitulo
        .AllowMultiSelect = bolMultiselecao
        .InitialFileName = strCaminhoInicial
        If .Show <> -1 Then GoTo AtribuirValor
        strArquivo = .SelectedItems(1)
    End With
    
AtribuirValor:
    PerguntarArquivo = strArquivo
    Set objPopup = Nothing
    
End Function




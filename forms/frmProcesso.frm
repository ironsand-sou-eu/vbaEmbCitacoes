VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProcesso 
   Caption         =   "Sísifo - Insira/confirme os dados do processo"
   ClientHeight    =   8445
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10140
   OleObjectBlob   =   "frmProcesso.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colGerenciadoresDeEvento As Collection

'Declarações de APIs
Private Declare PtrSafe Function GetCursorInfo Lib "user32" (ByRef pci As CursorInfo) As Boolean
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Cursores padrão do windows
Public Enum CursorTypes
    IDC_ARROW = 32512
    IDC_IBEAM = 32513
    IDC_WAIT = 32514
    IDC_CROSS = 32515
    IDC_UPARROW = 32516
    IDC_SIZE = 32640
    IDC_ICON = 32641
    IDC_SIZENWSE = 32642
    IDC_SIZENESW = 32643
    IDC_SIZEWE = 32644
    IDC_SIZENS = 32645
    IDC_SIZEALL = 32646
    IDC_NO = 32648
    IDC_HAND = 32649
    IDC_APPSTARTING = 32650
End Enum

'Necessário para a função GetCursorInfo
Private Type POINT
    X As Long
    Y As Long
End Type

'Necessário para a função GetCursorInfo
Private Type CursorInfo
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINT
End Type

'Configura um cursor
Private Function AddCursor(CursorType As CursorTypes)
    If Not TipoCursorEh(CursorType) Then
        SetCursor LoadCursor(0, CursorType)
        Sleep 200 ' wait a bit, needed for rendering
    End If
End Function

'Verifica se o cursor já está
Private Function TipoCursorEh(CursorType As CursorTypes) As Boolean
    Dim CursorHandle As Long: CursorHandle = LoadCursor(ByVal 0&, CursorType)
    Dim Cursor As CursorInfo: Cursor.cbSize = Len(Cursor)
    Dim CursorInfo As Boolean: CursorInfo = GetCursorInfo(Cursor)

    If Not CursorInfo Then
        TipoCursorEh = False
        Exit Function
    End If

    TipoCursorEh = (Cursor.hCursor = CursorHandle)
End Function

Private Sub LabelAutor_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim oGereEventoExitBotao As GereEventoExitBotao
    Dim oGereEventoExitCxTexto As GereEventoExitCxTexto
    Dim oGereEventoExitCombo As GereEventoExitCombo
    Dim strgerencia As String
    
    Set colGerenciadoresDeEvento = New Collection
    cmbCausaPedir.SetFocus
    AjustarLegendaSemTransicao LabelCausaPedir
    cmbAndamento.List = cfConfigurações.Range("AndamentosReferencia").Value
    cmbTercProprio.List = cfConfigurações.Range("TercProprio").Value
    Me.chbAnalisarProcNovo.Value = ConferirPadraoParaGerarProvidenciaAnalisarProcessoNovo
    AjustarVisibilidadeCamposMatriculaOuTerceiro
    CriarGerenciadoresDeEventoParaAnimacaoDosControles oGereEventoExitBotao, oGereEventoExitCxTexto, oGereEventoExitCombo
End Sub

Private Sub CriarGerenciadoresDeEventoParaAnimacaoDosControles(ByRef oGereEventoExitBotao As GereEventoExitBotao, _
            ByRef oGereEventoExitCxTexto As GereEventoExitCxTexto, ByRef oGereEventoExitCombo As GereEventoExitCombo)
    Dim oControle As MSForms.Control
    Dim retorno As Collection
    
    For Each oControle In Me.Controls
        If oControle.Visible = True Then
            Select Case TypeName(oControle)
            Case "TextBox"
                If oControle.TabStop = True Then
                    Set oGereEventoExitCxTexto = New GereEventoExitCxTexto
                    Set oGereEventoExitCxTexto.CxTexto = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitCxTexto
                End If
                
            Case "CommandButton"
                If oControle.TabStop = True Then
                    Set oGereEventoExitBotao = New GereEventoExitBotao
                    Set oGereEventoExitBotao.Botao = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitBotao
                End If
                
            Case "ComboBox"
                If oControle.TabStop = True Then
                    Set oGereEventoExitCombo = New GereEventoExitCombo
                    Set oGereEventoExitCombo.Combo = oControle
                    colGerenciadoresDeEvento.Add oGereEventoExitCombo
                End If
                
            End Select
        End If
    Next oControle
End Sub

Private Sub AjustarVisibilidadeCamposMatriculaOuTerceiro()
    If prProcesso.Tribunal = sfTRT5 Then
        Me.txtMatricula.Visible = False
        Me.LabelMatricula.Visible = False
        Me.LinhaBaseMatricula.Visible = False
        Me.LinhaMatricula.Visible = False
        
        Me.txtCodLocal.Visible = False
        Me.LabelCodLocal.Visible = False
        Me.LinhaBaseCodLocal.Visible = False
        Me.LinhaCodLocal.Visible = False
        
        Me.cmbTercProprio.Visible = True
        Me.LabelTercProprio.Visible = True
        Me.LinhaBaseTercProprio.Visible = True
        Me.LinhaTercProprio.Visible = True
    Else
        Me.txtMatricula.Visible = True
        Me.LabelMatricula.Visible = True
        Me.LinhaBaseMatricula.Visible = True
        Me.LinhaMatricula.Visible = True
        
        Me.txtCodLocal.Visible = True
        Me.LabelCodLocal.Visible = True
        Me.LinhaBaseCodLocal.Visible = True
        Me.LinhaCodLocal.Visible = True
        
        Me.cmbTercProprio.Visible = False
        Me.LabelTercProprio.Visible = False
        Me.LinhaBaseTercProprio.Visible = False
        Me.LinhaTercProprio.Visible = False
    End If
End Sub

Private Function ConferirPadraoParaGerarProvidenciaAnalisarProcessoNovo() As Boolean
    Dim gerencia As String
    Dim resposta As Boolean
    
    On Error Resume Next
    gerencia = LCase(prProcesso.gerencia)
    On Error GoTo 0
    
    Select Case UCase(gerencia)
    Case "PPJCM"
        If prProcesso.Sistema = sfPJe1g Then
            resposta = True
        Else
            resposta = False
        End If
    
    Case "PPJCE"
        resposta = True
        
    Case Else
        resposta = False
        
    End Select
    
    ConferirPadraoParaGerarProvidenciaAnalisarProcessoNovo = resposta

End Function

Private Sub UserForm_AddControl(ByVal Control As MSForms.Control)
    
    Select Case TypeName(Control)
    Case "TextBox"
        If Control.TabStop = True Then
            Dim oGereEventoExitCxTexto As GereEventoExitCxTexto
            Set oGereEventoExitCxTexto = New GereEventoExitCxTexto
            Set oGereEventoExitCxTexto.CxTexto = Control
            colGerenciadoresDeEvento.Add oGereEventoExitCxTexto
        End If
        
    Case "CommandButton"
        If Control.TabStop = True Then
            Dim oGereEventoExitBotao As GereEventoExitBotao
            Set oGereEventoExitBotao = New GereEventoExitBotao
            Set oGereEventoExitBotao.Botao = Control
            colGerenciadoresDeEvento.Add oGereEventoExitBotao
        End If
        
    Case "ComboBox"
        If Control.TabStop = True Then
            Dim oGereEventoExitCombo As GereEventoExitCombo
            Set oGereEventoExitCombo = New GereEventoExitCombo
            Set oGereEventoExitCombo.Combo = Control
            colGerenciadoresDeEvento.Add oGereEventoExitCombo
        End If
        
    End Select
End Sub

Private Sub UserForm_Terminate()
    Set oControleAtual = Nothing
    Set oControleAnterior = Nothing
End Sub

Private Sub LabelAdLiminar_Click()
    If LabelAdLiminar.Caption = "+" Then
        LabelAdLiminar.Visible = False
        Me.Width = Me.Width + 292
        ' Criar Controles: para nome do andamento, data e observação (para cada um a linha, a linha base, o label e a caixa de texto)
        AdicionarComboBox "Andamento2", "Andamento adicional", 17, 185, 24, 508, cfConfigurações.Range("AndamentosAdicionais")
        AdicionarCxTexto "DataAndamento2", "Data", 17, 73, 24, 712, False, vbDate
        AdicionarCxTexto "ObsAndamento2", "Texto andamento", 101, 274, 66, 508, True
        AdicionarComboBox "Providencia2", "Providência", 17, 274, 192, 508, cfConfigurações.Range("ProvidenciasAdicionais")
        AdicionarCxTexto "ObsProvidencia2", "Texto providência", 89, 274, 234, 508, True
    End If
    
    ' Centraliza o botão de inserir o processo
    Me.Controls("cmdIr").Left = Me.Width / 2 - (Me.Controls("cmdIr").Width / 2)
    
    MsgBox SisifoEmbasaFuncoes.ControleExiste(Me, "Joaquinho") & " " & SisifoEmbasaFuncoes.ControleExiste(Me, "cmbAndamento2")

End Sub

Private Sub LabelAdLiminar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    AddCursor IDC_HAND
End Sub

Private Sub txtDataAndamento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ValidaNumeros(KeyAscii, "/", ":", " ") = False Then KeyAscii = 0
End Sub

Private Sub txtDataAndamento_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    txtDataAndamento.text = SisifoEmbasaFuncoes.ValidaData(txtDataAndamento.text)
End Sub

Private Sub cmbCausaPedir_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'Pega pedidos e valores na planilha cfCausasPedirPedidos e preenche os checkboxes

    Dim rngCont As Range
    Dim arrPedidos() As String, strContNomePedido As String, strContRisco As String, strPrimeiroEndereco As String, strNucleoAdv() As String
    Dim btCont As Byte, btCont2 As Byte, btQtdLinhas As Byte, btControlesPorLinha As Byte, btControlesPreexistentes As Byte
    Dim intCont2 As Integer
    Dim curContValorPedido As Currency, curContProvisionar As Currency
    Dim dtCont As Date
    Dim varCont As Variant
    Dim bolExisteControle As Boolean
    
    cmbCausaPedir.Value = Trim(cmbCausaPedir.Value)
    
    If Trim(cmbCausaPedir.text) = "" Then Exit Sub
    
    Set rngCont = cfCausasPedirPedidos.Cells().Find(cmbCausaPedir.text)
    If rngCont Is Nothing Then Exit Sub
    strPrimeiroEndereco = rngCont.Address
    btCont = 0
    
    '''''''''''''''''''''''''''''''''''''''''
    ''' Ajustar causa de pedir e gerência '''
    '''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.CausaPedir = Trim(Me.cmbCausaPedir.Value)
    prProcesso.gerencia = PegaGerencia(cfCausasPedir, prProcesso.CausaPedir)
    
    
    '''''''''''''''''''''''
    ''' Ajustar pedidos '''
    '''''''''''''''''''''''
    
    ' Pega os nomes dos pedidos da causa de pedir em questão e armazena em arrPedidos()
    Do
        btCont = btCont + 1
        ReDim Preserve arrPedidos(1 To 3, 1 To btCont)
        arrPedidos(1, btCont) = rngCont.Offset(0, 1).Formula ' Nome do pedido
        arrPedidos(2, btCont) = rngCont.Offset(0, 2).Formula ' Risco
        arrPedidos(3, btCont) = CCur(rngCont.Offset(0, 3).Formula) ' Valor a provisionar
        
        Set rngCont = cfCausasPedirPedidos.Cells().FindNext(rngCont)
    Loop Until rngCont.Address = strPrimeiroEndereco
    
    ' Descobre quantas linhas de pedidos existem
    btControlesPorLinha = 5
    btControlesPreexistentes = 4
    btQtdLinhas = (Me.fraPedidos.Controls.Count - btControlesPreexistentes) / btControlesPorLinha
    
    ' Exclui as linhas de pedido existentes
    If btQtdLinhas > 0 Then
        For intCont2 = btQtdLinhas To 1 Step -1
            RemoverLinhaPedidoFormulario Me.Controls("cmdExcluir" & intCont2)
        Next intCont2
    End If
    'Pega as informações respectivas e insere uma linha para cada pedido
    For btCont2 = 1 To btCont Step 1
        strContNomePedido = PegaNomePedido(arrPedidos(1, btCont2))
        curContValorPedido = 0
        strContRisco = arrPedidos(2, btCont2)
        curContProvisionar = CCur(arrPedidos(3, btCont2))
        AdicionarLinhaPedido Me.fraPedidos, strContNomePedido, curContValorPedido, strContRisco, curContProvisionar
    Next btCont2
    
    
    '''''''''''''''''''''''''''
    ''' Ajustar unidade CCR '''
    '''''''''''''''''''''''''''
    
    ' Preenche as linhas com as informações respectivas
    ' Se for CCR, coloca unidade PPJ.
    If cmbCausaPedir.Value = "Desabastecimento CCR 04/2015" Or _
        cmbCausaPedir.Value = "Desabastecimento CCR 04/2016" Then txtCodLocal.text = "PPJ"
    
    
    ''''''''''''''''''''''''
    ''' Ajustar advogado '''
    ''''''''''''''''''''''''
    
    ' Povoa a combobox de advogados
    Me.cmbAdvogado.List = cfListaAdvogados.Range("ListaAdvogados").Value
    
    ' Busca e preenche o Núcleo e advogado
    If CDate(Trim(Me.txtDataAndamento)) = Date Or Trim(Me.txtDataAndamento) = 0 Or Trim(Me.txtDataAndamento) = "" Then
        dtCont = Application.WorksheetFunction.WorkDay(Date, 1, cfFeriados.Range("SisifoFeriados"))
    Else
        dtCont = Application.WorksheetFunction.WorkDay(CDate(Me.txtDataAndamento), -2, cfFeriados.Range("SisifoFeriados"))
    End If
    strNucleoAdv() = PegaNucleoAdvPrep(Advogado, prProcesso.NumeroProcesso, cmbCausaPedir.Value, dtCont, False)
    
    ' Erros/resultado
    If UBound(strNucleoAdv) = 1 Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & strNucleoAdv(1) & ". A inclusão foi cancelada e descartados os dados" & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
        Set prProcesso = Nothing
        Me.Hide
        Exit Sub
    ElseIf Left(LCase(strNucleoAdv(2)), 4) = "erro" Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & strNucleoAdv(2) & ". A inclusão foi cancelada e descartados os dados" & vbCrLf & _
        "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
        Set prProcesso = Nothing
        Me.Hide
        Exit Sub
    Else
        AdicionarLinhaAjustarLegenda cmbAdvogado, 4
        prProcesso.Nucleo = Trim(strNucleoAdv(1))
        cmbAdvogado.Value = Trim(strNucleoAdv(2))
        cmbAdvogado.Tag = Trim(strNucleoAdv(3)) ' A tag da textbox txtAdvogado terá o nome do grupo de onde ele saiu
    End If
    
    
    ''''''''''''''''''''''''''''
    ''' Ajustar tipo de ação '''
    ''''''''''''''''''''''''''''
    
    ' Inserir a opção de Tipo de Ação secundária (conforme causa de pedir) logo após a opção primária na Combobox
    Set rngCont = cfCausasPedir.Cells().Find(What:=cmbCausaPedir.Value, LookAt:=xlWhole)
    If Not rngCont Is Nothing Then
        cmbTipoAcao.AddItem rngCont.Offset(0, 2).Formula, Me.cmbTipoAcao.ListIndex + 1
    End If
    
End Sub

Private Sub cmdAdicionarPedido_Click()
    AdicionarLinhaPedido Me.fraPedidos
End Sub

Private Sub cmdIr_Click()
    Me.chbDeveGerar.Value = True
        
                Me.Hide
End Sub

Private Sub AdicionarLinhaPedido(ByRef frame As MSForms.frame, Optional ByVal strNomePedido As String, Optional ByVal strValorPedido As String, Optional ByVal strRisco As String, Optional ByVal strValorProvisionar As String)

    Dim btCont As Byte, btQtdLinhas As Byte, btControlesPorLinha As Byte, btControlesPreexistentes As Byte
    Dim Combo As MSForms.ComboBox
    Dim text As MSForms.TextBox
    Dim Botao As MSForms.CommandButton
    Dim oContGerenciadorEventoBotao As GereEventosBotaoAdPedido
    Dim oContGerenciadorEventoTexto As GereEventosCxTextoValor
    Dim varCont As Variant
    
    'Descobre a quantidad de linhas da Frame fraPedidos
    btControlesPorLinha = 5
    btControlesPreexistentes = 4
    btQtdLinhas = (Me.fraPedidos.Controls.Count - btControlesPreexistentes) / btControlesPorLinha
    
    ' Inclui os 5 controles desta linha
    Set Combo = frame.Controls.Add("Forms.ComboBox.1", "cmbPed" & btQtdLinhas + 1)
    Combo.Top = btQtdLinhas * 24 + 18
    Combo.Left = 6
    Combo.Height = 18
    Combo.Width = 204
    Combo.BorderStyle = fmBorderStyleSingle
    Combo.List = cfPedidos.Range("Pedidos").Value
    Combo.MatchRequired = True
    If strNomePedido <> "" Then Combo.Value = strNomePedido
    
    Set text = frame.Controls.Add("Forms.TextBox.1", "txtPed" & btQtdLinhas + 1)
    text.Top = btQtdLinhas * 24 + 18
    text.Left = 216
    text.Height = 17
    text.Width = 78
    text.BorderStyle = fmBorderStyleSingle
    If strValorPedido = "" Then strValorPedido = "0"
    text.Value = Format(strValorPedido, "#,##0.00")
    Set oContGerenciadorEventoTexto = New GereEventosCxTextoValor
    Set oContGerenciadorEventoTexto.CaixaDeTextoValor = text
    colGerenciadoresDeEvento.Add oContGerenciadorEventoTexto
    
    Set Combo = frame.Controls.Add("Forms.ComboBox.1", "cmbRisco" & btQtdLinhas + 1)
    Combo.Top = btQtdLinhas * 24 + 18
    Combo.Left = 300
    Combo.Height = 18
    Combo.Width = 52
    Combo.BorderStyle = fmBorderStyleSingle
    Combo.List = cfPedidos.Range("Riscos").Value
    Combo.MatchRequired = True
    If strRisco = "" Then strRisco = "Remoto"
    Combo.Value = strRisco
    
    Set text = frame.Controls.Add("Forms.TextBox.1", "txtProv" & btQtdLinhas + 1)
    text.Top = btQtdLinhas * 24 + 18
    text.Left = 358
    text.Height = 17
    text.Width = 78
    text.BorderStyle = fmBorderStyleSingle
    If strValorProvisionar = "" Then strValorProvisionar = "0"
    text.Value = Format(strValorProvisionar, "#,##0.00")
    Set oContGerenciadorEventoTexto = New GereEventosCxTextoValor
    Set oContGerenciadorEventoTexto.CaixaDeTextoValor = text
    colGerenciadoresDeEvento.Add oContGerenciadorEventoTexto
    
    Set Botao = frame.Controls.Add("Forms.CommandButton.1", "cmdExcluir" & btQtdLinhas + 1)
    Botao.Top = btQtdLinhas * 24 + 18
    Botao.Left = 442
    Botao.Height = 18
    Botao.Width = 18
    Botao.Picture = Me.btImagem.Picture
    Set oContGerenciadorEventoBotao = New GereEventosBotaoAdPedido
    Set oContGerenciadorEventoBotao.Botao = Botao
    colGerenciadoresDeEvento.Add oContGerenciadorEventoBotao
    
    ' Reposiciona e renomeia o botão de adicionar e a área de rolagem da frame
    Set Botao = frame.Controls("cmdAdicionarPedido")
    Botao.Top = (btQtdLinhas + 1) * 24 + 21
    frame.ScrollHeight = Botao.Top + Botao.Height + 6
    
End Sub

Public Sub AdicionarLinhaAjustarLegenda(ByRef Controle As MSForms.Control, sngVelocidade As Single)

    Dim lblLabel As MSForms.Label, lblLinha As MSForms.Label, ctrLabel As MSForms.Control
    Dim sngTimerInicio As Single, sngCont As Single, sngContCor As Single
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    Dim strCategoriaControle As String
    
    strCategoriaControle = Right(Controle.Name, Len(Controle.Name) - 3)
    Set lblLinha = Controle.Parent.Controls("Linha" & strCategoriaControle)
    Set ctrLabel = Controle.Parent.Controls("Label" & strCategoriaControle)
    
    ' Configura a linha azul
    lblLinha.Visible = True
    lblLinha.Width = 1
    
    If Controle.text = "" Then
        ' Configurações do label (se estiver sem texto)
        Set lblLabel = ctrLabel
        sngTopDestino = ctrLabel.Top - 12
        sngLeftDestino = ctrLabel.Left - 6
        btTamanhoFonteDestino = 8
        sngCorDestino = 100
    End If
    
    ' Transição
    On Error Resume Next
    sngTimerInicio = Timer
    While Timer <= sngTimerInicio + (1 / sngVelocidade)
        sngCont = IIf(Timer - sngTimerInicio = 0, 0.01, Timer - sngTimerInicio)
        lblLinha.Width = Controle.Width * sngCont * sngVelocidade
        If Controle.text = "" Then
            ctrLabel.Top = sngTopDestino + ctrLabel.Top * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            ctrLabel.Left = sngLeftDestino + ctrLabel.Left * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            lblLabel.Font.Size = btTamanhoFonteDestino + lblLabel.Font.Size * (0.01 / (sngCont * sngVelocidade / 2)) ' Expressão decrescente que tende a 0
            sngContCor = sngCorDestino * sngCont * sngVelocidade
            lblLabel.ForeColor = RGB(sngContCor, sngContCor, sngContCor)
        End If
        DoEvents
    Wend
    On Error GoTo 0
    
    ' Valores finais
    lblLinha.Width = Controle.Width
    If Controle.text = "" Then
        ctrLabel.Top = sngTopDestino
        ctrLabel.Left = sngLeftDestino
        lblLabel.Font.Size = btTamanhoFonteDestino
        lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    End If
    
End Sub

Sub RemoverLinhaPedidoFormulario(ByRef cmdExcluir As MSForms.Control)
    
    Dim btLinhaExcluir As Byte, btCont As Byte, btQtdLinhas As Byte, btControlesPorLinha As Byte, btControlesPreexistentes As Byte
    Dim Controle As MSForms.Control
    
    ' Descobre quantas linhas de pedidos existem
    btControlesPorLinha = 5
    btControlesPreexistentes = 4
    btQtdLinhas = (Me.fraPedidos.Controls.Count - btControlesPreexistentes) / btControlesPorLinha
    
    ' Descobre linha a excluir
    btLinhaExcluir = CByte(Replace(cmdExcluir.Name, "cmdExcluir", ""))
    
    ' Exclui os 5 controles desta linha
    Me.Controls.Remove "cmbPed" & btLinhaExcluir
    Me.Controls.Remove "txtPed" & btLinhaExcluir
    Me.Controls.Remove "cmbRisco" & btLinhaExcluir
    Me.Controls.Remove "txtProv" & btLinhaExcluir
    Me.Controls.Remove "cmdExcluir" & btLinhaExcluir
    
    ' Reposiciona e renomeia os que vêm abaixo
    For btCont = btLinhaExcluir + 1 To btQtdLinhas Step 1
        Set Controle = Me.Controls("cmbPed" & btCont)
        Controle.Top = Controle.Top - 24
        Controle.Name = "cmbPed" & btCont - 1
        
        Set Controle = Me.fraPedidos.Controls("txtPed" & btCont)
        Controle.Top = Controle.Top - 24
        Controle.Name = "txtPed" & btCont - 1
        
        Set Controle = Me.fraPedidos.Controls("txtProv" & btCont)
        Controle.Top = Controle.Top - 24
        Controle.Name = "txtProv" & btCont - 1
        
        Set Controle = Me.Controls("cmbRisco" & btCont)
        Controle.Top = Controle.Top - 24
        Controle.Name = "cmbRisco" & btCont - 1
        
        Set Controle = Me.fraPedidos.Controls("cmdExcluir" & btCont)
        Controle.Top = Controle.Top - 24
        Controle.Name = "cmdExcluir" & btCont - 1
    Next btCont
    
    Set Controle = Me.fraPedidos.Controls("cmdAdicionarPedido")
    Controle.Top = Controle.Top - 24
    Controle.Parent.ScrollHeight = Controle.Top + Controle.Height + 6

End Sub

Public Sub AjustarLegendaSemTransicao(ByRef ctrLabel As MSForms.Control)

    Dim lblLabel As MSForms.Label
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    
    ' Configurações do label (se estiver sem texto)
    Set lblLabel = ctrLabel
    sngTopDestino = ctrLabel.Top - 12
    sngLeftDestino = ctrLabel.Left - 6
    btTamanhoFonteDestino = 8
    sngCorDestino = 100
    
    ctrLabel.Top = sngTopDestino
    ctrLabel.Left = sngLeftDestino
    lblLabel.Font.Size = btTamanhoFonteDestino
    lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    
End Sub

Public Sub RetornarFormato(ByRef Controle As MSForms.Control, sngVelocidade As Single)

    Dim lblLabel As MSForms.Label, ctrLabel As MSForms.Control, lblLinha As MSForms.Label
    Dim sngTimerInicio As Single, sngCont As Single, sngContCor As Single
    Dim sngTopDestino As Single, sngLeftDestino As Single, sngCorDestino As Single, btTamanhoFonteDestino As Byte
    Dim strCategoriaControle As String
    
    strCategoriaControle = Right(Controle.Name, Len(Controle.Name) - 3)
    Set lblLinha = Controle.Parent.Controls("Linha" & strCategoriaControle)
    Set ctrLabel = Controle.Parent.Controls("Label" & strCategoriaControle)
    
    ' Configura a linha azul
    lblLinha.Visible = False
    
    ' Configurações
    If Controle.text = "" Then
        Set lblLabel = ctrLabel
        sngTopDestino = ctrLabel.Top + 12
        sngLeftDestino = ctrLabel.Left + 6
        btTamanhoFonteDestino = 11.25
        sngCorDestino = 0
        
        ' Transição
        On Error Resume Next
        sngTimerInicio = Timer
        While Timer <= sngTimerInicio + (1 / sngVelocidade)
            sngCont = IIf(Timer - sngTimerInicio = 0, 0.01, Timer - sngTimerInicio)
            ctrLabel.Top = sngTopDestino * sngCont * sngVelocidade
            ctrLabel.Left = sngLeftDestino * sngCont * sngVelocidade
            lblLabel.Font.Size = btTamanhoFonteDestino * sngCont * sngVelocidade
            DoEvents
        Wend
        On Error GoTo 0
        
        ' Valores finais
        ctrLabel.Top = sngTopDestino
        ctrLabel.Left = sngLeftDestino
        lblLabel.Font.Size = btTamanhoFonteDestino
        lblLabel.ForeColor = RGB(sngCorDestino, sngCorDestino, sngCorDestino)
    End If
    
End Sub

Private Sub AdicionarCxTexto(ByVal strNomeControle As String, strLegenda As String, lngHeight As Long, lngWidth As Long, lngTop As Long, lngLeft As Long, Optional bolMultilinha As Boolean, Optional vbValidacao As VbVarType)
    Dim oCxTexto As MSForms.TextBox
    
    ' Inclui os 3 controles que compõem as bordas do macrocontrole (LinhaBase, Linha e Label)
    AdicionarBordasControle strNomeControle, strLegenda, lngHeight, lngWidth, lngTop, lngLeft
    
    ' Inclui a caixa de texto
    Set oCxTexto = Me.Controls.Add("Forms.TextBox.1", "txt" & strNomeControle)
    With oCxTexto
        .Top = lngTop
        .Left = lngLeft
        .Height = lngHeight
        .Width = lngWidth
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Size = 12
        .ForeColor = &HFD0C28
        .MultiLine = bolMultilinha
        .EnterKeyBehavior = bolMultilinha
    End With
    
    Select Case vbValidacao
    Case vbDate
        oCxTexto.Tag = "data"
        Dim oContGereEventosCxTextoData As GereEventosCxTextoData
        Set oContGereEventosCxTextoData = New GereEventosCxTextoData
        Set oContGereEventosCxTextoData.CaixaDeTextoData = oCxTexto
        colGerenciadoresDeEvento.Add oContGereEventosCxTextoData
    Case vbCurrency
        Dim oContGereEventosCxTextoValor As GereEventosCxTextoValor
        Set oContGereEventosCxTextoValor = New GereEventosCxTextoValor
        Set oContGereEventosCxTextoValor.CaixaDeTextoValor = oCxTexto
        colGerenciadoresDeEvento.Add oContGereEventosCxTextoValor
    End Select
    
End Sub

Private Sub AdicionarComboBox(ByVal strNomeControle As String, strLegenda As String, lngHeight As Long, lngWidth As Long, lngTop As Long, lngLeft As Long, Optional rngValoresValidos As Excel.Range)
    Dim oCombo As MSForms.ComboBox
    
    ' Inclui os 3 controles que compõem as bordas do macrocontrole (LinhaBase, Linha e Label)
    AdicionarBordasControle strNomeControle, strLegenda, lngHeight, lngWidth, lngTop, lngLeft
    
    ' Inclui a caixa de texto
    Set oCombo = Me.Controls.Add("Forms.ComboBox.1", "cmb" & strNomeControle)
    With oCombo
        .Top = lngTop
        .Left = lngLeft
        .Height = lngHeight
        .Width = lngWidth
        .BorderStyle = fmBorderStyleNone
        .BackStyle = fmBackStyleTransparent
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Size = 12
        .ForeColor = &HFD0C28
        .MatchRequired = True
        If Not rngValoresValidos Is Nothing Then .List = rngValoresValidos.Value
    End With
    
End Sub

Private Sub AdicionarBordasControle(ByVal strNomeControle As String, strLegenda As String, lngHeight As Long, lngWidth As Long, lngTop As Long, lngLeft As Long)
    Dim lbLabel As MSForms.Label
    
    ' Inclui os 4 controles que compõem o macrocontrole (LinhaBase, Linha, Caixa de texto e Label)
    Set lbLabel = Me.Controls.Add("Forms.Label.1", "LinhaBase" & strNomeControle)
    With lbLabel
        .Top = lngTop + lngHeight + 1
        .Left = lngLeft
        .Height = 2
        .Width = lngWidth
        .BorderStyle = fmBorderStyleNone
        .BorderColor = &H80000006
        .SpecialEffect = fmSpecialEffectEtched
    End With
    
    Set lbLabel = Me.Controls.Add("Forms.Label.1", "Linha" & strNomeControle)
    With lbLabel
        .Top = lngTop + lngHeight + 1
        .Left = lngLeft
        .Height = 2
        .Width = lngWidth / 3
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HFD0C28
        .Visible = False
    End With
    
    Set lbLabel = Me.Controls.Add("Forms.Label.1", "Label" & strNomeControle)
    With lbLabel
        .Top = lngTop
        .Left = lngLeft + 6
        .Height = lngHeight
        .Width = lngWidth
        .Caption = strLegenda
        .Font.Size = 11
        .BackStyle = fmBackStyleTransparent
    End With
End Sub


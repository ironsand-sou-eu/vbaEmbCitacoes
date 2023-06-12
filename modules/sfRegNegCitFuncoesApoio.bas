Attribute VB_Name = "sfRegNegCitFuncoesApoio"
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
 
Sub FechaConfigCitacoesVisivel(ByVal Controle As IRibbonControl, Optional ByRef returnedVal)
    SisifoEmbasaFuncoes.FechaConfigVisivel ThisWorkbook, cfConfigurações, Controle, returnedVal
End Sub

Private Sub AoCarregarRibbonCitacoes(Ribbon As IRibbonUI)
' Chama a função geral AoCarregarRibbon com os parâmetros corretos.
    SisifoEmbasaFuncoes.AoCarregarRibbon cfConfigurações, Ribbon
End Sub

Sub LiberarEdicaoCitacoes(ByVal Controle As IRibbonControl)
' Chama a função geral LiberarEdicao
    SisifoEmbasaFuncoes.LiberarEdicao ThisWorkbook, cfConfigurações
    
End Sub

Sub RestringirEdicaoRibbonCitacoes(ByVal Controle As IRibbonControl)
' Chama a função geral RestringirEdicaoRibbon
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfConfigurações, Controle
End Sub

Sub ssfCitCmbSistemaMudou(ByVal Controle As IRibbonControl, ByVal text As String)
    cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula = Trim(text)
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook ' Salva as alterações
End Sub

Sub ssfCitCmbSistemaTexto(ByVal Control As IRibbonControl, ByRef returnedVal)
    If Control.ID = "cmbSistema" Then
        If Trim(stcitgerencia) <> "" Then
            returnedVal = cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
        Else
            'Pega o valor do sistema na planilha de configurações.
            returnedVal = cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
        End If
    End If
End Sub

Sub CadastrarProcessoIndividual(ByVal Controle As IRibbonControl)
''
'' Com o PJe aberto e logado no Internet Explorer, busca um processo e o cadastra
''
    Dim strNumeroProcesso As String
    Dim rngCelula As Excel.Range
    
    ' Zerar a variável prProcesso
    Set prProcesso = Nothing
    Set prProcesso = New Processo
    
    ' Pegar o número
    Set rngCelula = ActiveCell
    strNumeroProcesso = PegaNumeroProcessoDeCelula(rngCelula)
    If strNumeroProcesso = "Número não é CNJ" Then Exit Sub
    
    ' Descobrir o sistema
    prProcesso.Tribunal = DescobrirTribunal(strNumeroProcesso)
    prProcesso.Sistema = DescobrirSistema(strNumeroProcesso)
    
    ' Chama a função de cadastro correspondente ao sistema e Tribunal
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
'' Retorna o número do processo contido na primeira célula da range passada como parâmetro -- ou, se não for padrão CNJ, pergunta.
'' Em caso de erro, retorna a mensagem de erro.
''
    Dim strNumeroProcesso As String
    Dim intTentarDeNovo As Integer
    Dim rngCelula As Range
    
    Set rngCelula = rngRange(1, 1)
    strNumeroProcesso = rngCelula.text
    
    ' Se não houver célula no espaço enviado, ou se estiver vazia, ou se contiver algo em formato não CNJ, pergunta o número do processo.
    If rngCelula Is Nothing Or rngCelula.text = "" Or Not EhCNJ(ActiveCell.text) Then
PerguntaNumero:
        strNumeroProcesso = InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", informe o número do processo do Projudi a cadastrar no formato CNJ " & _
                "(""0000000-00.0000.0.00.0000""):", "Sísifo - Cadastrar processo")
            
AvisoNaoCNJ:
        If Not EhCNJ(strNumeroProcesso) Then
            intTentarDeNovo = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o número informado (" & strNumeroProcesso & ") não está no padrão do CNJ. " & _
                "Deseja tentar novamente com um número no padrão ""0000000-00.0000.0.00.0000""?", vbYesNo + vbCritical + vbDefaultButton1, _
                "Sísifo - Erro no cadastro")
                
            If intTentarDeNovo = vbYes Then
                GoTo PerguntaNumero
            Else
                PegarNumeroProcesso = "Número não é CNJ"
                Exit Function
            End If
        End If
    End If
    
    ' Se o conteúdo de strNumeroProcesso for um número de processo CNJ, aceita.
    PegaNumeroProcessoDeCelula = strNumeroProcesso
    
End Function

Function DescobrirTribunal(strNumero As String) As String
''
'' Confere a que tribunal pertence o número strNumero. Retorna uma string com o sistema, ou a mensagem de erro se não for CNJ.
''

    Dim strNumJustica As String, strNumTibunal As String, strCmbTribunal As String
    
    ' Se o número não for CNJ, retorna erro
    If Not EhCNJ(strNumero) Then
        DescobrirTribunal = "Número passado não é CNJ"
        Exit Function
    End If
    
    ''' Pega o tribunal de acordo com a combobox do Ribbon do Sísifo
    Select Case cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
    Case "Projudi"
        strCmbTribunal = sfTjba
    Case "PJe 1g TJ/BA"
        strCmbTribunal = sfTjba
    Case "PJe 1g TRT5"
        strCmbTribunal = sfTRT5
    End Select
    
    
    ''' Pega o tribunal escolhido de acordo com o número
    strNumJustica = Mid(strNumero, 17, 1)
    strNumTibunal = Mid(strNumero, 19, 2)
    
    ' Se não for TJ/BA, retorna erro.
    If strNumJustica <> "8" And strNumJustica <> "5" Then
        DescobrirTribunal = "o processo não pertence a TJ nem a TRT, e ainda não é cadastrado pelo Sísifo."
        Exit Function
    
    ' Justiça estadual de outros estados
    ElseIf strNumJustica = "8" And strNumTibunal <> "05" Then
        DescobrirTribunal = "o processo pertence ao Tribunal de Justiça de outro estado, que ainda não é cadastrado pelo Sísifo."
        Exit Function
    
    ' Justiça do Trabalho de outros regionais
    ElseIf strNumJustica = "5" And strNumTibunal <> "05" Then
        DescobrirTribunal = "o processo pertence ao TRT de outra região, que ainda não é cadastrado pelo Sísifo."
        Exit Function
    
    ' TJ/BA
    ElseIf strNumJustica = "8" And strNumTibunal = "05" Then
        If strCmbTribunal = sfTjba Then
            DescobrirTribunal = sfTjba
        Else
            DescobrirTribunal = "embora tenhais selecionado um sistema do tribunal " & strCmbTribunal & ", o número parece indicar que " & _
                "o processo pertence, em verdade, ao TJ/BA. Imploro que verifiqueis a inconsistência."
            Exit Function
        End If

    ' TRT 5ª Região
    ElseIf strNumJustica = "5" And strNumTibunal = "05" Then
        If strCmbTribunal = sfTRT5 Then
            DescobrirTribunal = sfTRT5
        Else
            DescobrirTribunal = "embora tenhais selecionado um sistema do tribunal " & strCmbTribunal & ", o número parece indicar que " & _
                "o processo pertence, em verdade, ao TRT da 5ª região. Imploro que verifiqueis a inconsistência."
            Exit Function
        End If

    End If
    
End Function

Function DescobrirSistema(strNumero As String) As String
''
'' Confere se o número passado corresponde a um número de processo no padrão CNJ, e depois tenta descobrir o sistema em que tramita
''   (baseia-se apenas nos padrões rotineiros de numeração do TJ/BA).
''    Retorna uma string com o sistema, ou a mensagem de erro correspondente.
''

    Dim strCont As String, strCmbSistema As String
    
    ''' Pega o tribunal de acordo com a combobox do Ribbon do Sísifo
    Select Case cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula
    Case "Projudi"
        strCmbSistema = sfProjudi
    Case "PJe 1g TJ/BA"
        strCmbSistema = sfPJe1g
    Case "PJe 1g TRT5"
        strCmbSistema = sfPJe1g
    End Select
    
    ' Se não for CNJ, retorna erro
    If Not EhCNJ(strNumero) Then
        DescobrirSistema = "Número de processo indicado não está no padrão CNJ"
    
    ' Se não for dos tribunais abrangidos, retorna o erro respectivo.
    ElseIf prProcesso.Tribunal <> sfTjba And prProcesso.Tribunal <> sfTRT5 Then
        DescobrirSistema = prProcesso.Tribunal
    
    Else
        Select Case prProcesso.Tribunal
        Case sfTjba 'TJ/BA
            ' Se não começar com 0, é PJe
            If Left(strNumero, 1) <> "0" Then
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o número parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema do tribunal. Imploro que verifiqueis a inconsistência."
                    Exit Function
                End If
            
            ' Se começar com 03 ou 05, é eSaj (o Sísifo vai tratar como PJe, pois o TJ/BA descontinuará o eSaj e já migrou alguns processos)
            ElseIf Left(strNumero, 2) = "03" Or Left(strNumero, 2) = "05" Then
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o número parece indicar que " & _
                        "o processo pertence, em verdade, ao eSaj. Imploro que verifiqueis a inconsistência."
                    Exit Function
                End If
            
            ' Nos demais casos, é Projudi
            Else
                If strCmbSistema = sfProjudi Then
                    DescobrirSistema = sfProjudi
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o número parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema. Imploro que verifiqueis a inconsistência."
                    Exit Function
                End If
            End If
            
        Case sfTRT5 'TRT5
                If strCmbSistema = sfPJe1g Then
                    DescobrirSistema = sfPJe1g
                Else
                    DescobrirSistema = "embora tenhais selecionado um processo do sistema" & strCmbSistema & ", o número parece indicar que " & _
                        "o processo pertence, em verdade, a outro sistema do tribunal. Imploro que verifiqueis a inconsistência."
                    Exit Function
                End If
        
        End Select
                
    End If
    
End Function

Function EhCNJ(strNumero As String) As Boolean
''
'' Confere se o número passado corresponde a um número de processo no padrão CNJ.
'' Só retorna VERDADEIRO se for padrão CNJ com pontos e traços.
'' O padrão é 0000000-00.0000.0.00.0000 - zeros significam qualquer número; hífens e pontos são hífens e pontos mesmo.
''

    Dim strCont As String
    Dim btCont As Byte
    Dim bolEhCNJ As Boolean
    
    bolEhCNJ = True
    
    ' Se não tiver 25 caracteres, não é CNJ
    If Len(strNumero) <> 25 Then bolEhCNJ = False
    
    For btCont = 1 To 25 Step 1 ' Itera caractere a caractere, verificando:
    
        strCont = Mid(strNumero, btCont, 1)
        
        Select Case btCont
        Case 8
            If strCont <> "-" Then bolEhCNJ = False ' Se o hífen está no lugar
            
        Case 11, 16, 18, 21
            If strCont <> "." Then bolEhCNJ = False ' Se os pontos estão no lugar
        
        Case Else
            If Not IsNumeric(strCont) Then bolEhCNJ = False ' Se os demais são números
            
        End Select
        
    Next btCont
    
    EhCNJ = bolEhCNJ
    
End Function

Function ValidarCPFCNPJ(strStringAValidar As String) As String
''
'' Com uma string qualquer passada como parâmetro, retorna um CPF/CNPJ válido ou retorna uma mensagem de erro correspondente.
''
    Dim strCont As String, strCont2
    Dim btCont As Integer
    Dim intCont As Integer
    
    
    ' Remove o que não for numérico
    For intCont = Len(strStringAValidar) To 1 Step -1
        If IsNumeric(Mid(strStringAValidar, intCont, 1)) Then strCont = Mid(strStringAValidar, intCont, 1) & strCont
    Next intCont
    
    ' Verifica se o tamanho é compatível com CPF ou CNPJ
    If Len(strCont) <> 11 And Len(strCont) <> 14 Then
        ValidarCPFCNPJ = "o texto informado tem tamanho incompatível com CPF ou CNPJ"
        Exit Function
    End If
    
    ' Valida dígito do CPF
    If Len(strCont) = 11 Then
        strCont2 = Left(strCont, 9)

        ' Descobre primeiro dígito
        For btCont = 1 To 9 Step 1
            intCont = (CInt(Mid(strCont2, btCont, 1)) * (11 - btCont)) + intCont
        Next btCont
        intCont = ((intCont \ 11 + 1) * 11) - intCont
        If intCont = 10 Then intCont = 0
        strCont2 = strCont2 & CStr(intCont)
        
        ' Descobre segundo dígito
        intCont = 0
        For btCont = 1 To 10 Step 1
            intCont = (CInt(Mid(strCont2, btCont, 1)) * (12 - btCont)) + intCont
        Next btCont
        intCont = ((intCont \ 11 + 1) * 11) - intCont
        If intCont = 10 Or intCont = 11 Then intCont = 0
        strCont2 = strCont2 & CStr(intCont)
        
        If strCont <> strCont2 Then
            ValidarCPFCNPJ = "o dígito verificador do CPF é inválido"
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
'' Pega um código único de parte sem CPF. Os códigos são compostos por dia, mês, ano e quatro dígitos sequenciais
''   para a quantidade de partes sem CPF cadastradas hoje.
''
    Dim rngRange As Range
    Dim strCodigo As String
    
    ' Determina a célula que contém o código atual
    Set rngRange = cfConfigurações.Cells().Find(What:="Partes sem CPF cadastrado", LookAt:=xlWhole).Offset(0, 1)
    
    If Left(rngRange.text, 8) = Format(Date, "ddmmyyyy") Then
    ' Se o código atual está no mesmo dia, soma um e transforma-o no código atual
        rngRange.Formula = "'" & Trim(rngRange.text + 1)
        
    Else
    ' Se código atual está em outro dia, cria o código de um novo dia no número "0001"
        rngRange.Formula = "'" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "0001"
        
    End If
    
    PegaCodParteSemCPF = rngRange.text
    
End Function

Sub PegaInformacoesProcessoGeral(ByRef dtDataProvContestar As Date, ByRef dtDataProvSubsidios As Date, ByRef planGrupoProvContestar As Excel.Worksheet, ByRef planGrupoProvSubsidios As Excel.Worksheet)
''
'' Faz a coleta dos dados do processo que não dependem de sistema (PJe ou Projudi), armazena num objeto da classe Processo e devolve este objeto.
'' Deve ser rodada após a função específica do sistema (PJe ou Projudi)
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
    'Dim bolMaior20SM As Boolean, bolAgendaPautista As Boolean ''Eram apenas para a providência de agendar pautista
    
    '''''''''''''''''''''''
    ''' Comarca e Órgão '''
    '''''''''''''''''''''''
    
    prProcesso.Comarca = PegaComarca(prProcesso.Juizo)
    prProcesso.Orgao = PegaOrgao(prProcesso.Juizo)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Exceto para PJe TRT5, consulta SCI: Matrícula, '''
    ''' titularidade, EL vinculado, outros processos   '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If LCase(cfConfigurações.Cells().Find(What:="Sistema no combobox", LookAt:=xlWhole).Offset(0, 1).Formula) <> "pje 1g trt5" Then
        strMatricula = PegaMatricula(strSci(), strObservacao)
        If prProcesso Is Nothing Then Exit Sub
    End If
    
    
    '''''''''''''''''''''''''
    ''' Mostra formulário '''
    '''''''''''''''''''''''''
    
    ' Caso seja trabalhista, verifica se tem outros réus
    If prProcesso.Tribunal = sfTRT5 Then
        For Cont = 1 To prProcesso.OutrosParticipantes.Count Step 1
            If prProcesso.OutrosParticipantes(Cont).CondicaoParte = "Réu" Then
                bolTemOutrosReus = True
                Exit For
            End If
        Next Cont
    End If
    
    ' Mostra formulário. Exibe número do processo e nome do autor para conferência.
    ' IMPORTANTE: O núcleo da variável global prProcesso é definido dentro do evento 'Change' do controle cmbCausaPedir do formulário.
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
            .cmbTercProprio.Value = IIf(bolTemOutrosReus = True, "Terceirizado", "Próprio")
            .AjustarLegendaSemTransicao .LabelTercProprio
        End If
        .txtDataAndamento.text = IIf(prProcesso.DataAndamento <> 0, prProcesso.DataAndamento, Date)
        If Len(.txtDataAndamento.text) = 19 And Right(.txtDataAndamento.text, 3) = ":00" Then .txtDataAndamento.text = Left(.txtDataAndamento.text, 16) ' Não exibe o ":00" dos segundos
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
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o andamento adicional não tem um nome ou uma data válida, e por isso seu humilde " & _
            "e leal servo não poderá cadastrá-lo. Caso deseje cadastrar o processo normalmente sem andamento adicional, clique em ""OK"". Para cancelar " & _
            "o cadastro, clique em ""Cancelar"".", vbQuestion + vbOKCancel, "Sísifo - Erro no andamento adicional") = vbCancel Then GoTo EncerrarCadastro
        End If
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''' Causa de pedir, gerência, natureza (ramo do direito) '''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    prProcesso.CausaPedir = form.cmbCausaPedir.Value
    prProcesso.gerencia = "PPJCM"
    prProcesso.Natureza = PegaNatureza(cfCausasPedir, prProcesso.CausaPedir)
    
    
    ''''''''''''''''''''''''''''
    ''' Advogado responsável '''
    ''''''''''''''''''''''''''''
    prProcesso.Advogado = Trim(form.cmbAdvogado.Value)
    ' Se tiver mudado advogado para escritório externo, muda o núcleo
    If prProcesso.gerencia = "PPJCM" And prProcesso.Advogado = "ESCRITORIO: SANTOS NETO E BOA SORTE" Then
        prProcesso.Nucleo = "PPJCE - Cont. Especial - Santos Neto"
        prProcesso.Advogado = "JORGE KIDELMIR NASCIMENTO DE OLIVEIRA FILHO"
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''
    ''' Matrícula e bloquear matrícula   '''
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
    
    ' Coloca Unidade. Se não houver o número de localidade na planilha, alerta.
    If form.txtCodLocal.Visible = True And form.txtCodLocal.text <> "" Then
        strUnidadeDiv = PegaUnidadeDivisao(cfUnidadesELs, form.txtCodLocal.text)
        If strUnidadeDiv = "" Then
            MsgBox "O número da localidade digitado não foi encontrado em nossa base. Favor inserir o nome da " & _
                    "Unidade manualmente na planilha antes de importar para o Espaider. ATENÇÃO: O nome da Unidade " & _
                    "precisa ser copiado e colado do Espaider, pois caso seja diferente, a importação deste processo " & _
                    "resultará em erro.", vbCritical + vbOKOnly, "Sísifo - Código de localidade não encontrado"
            
            GoTo EncerrarCadastro
        Else
            prProcesso.Unidade = Left(strUnidadeDiv, InStr(1, strUnidadeDiv, "///") - 1)
        End If
        
        ' Coloca escritório local. Se o EL estiver em branco mas a unidade não (e a unidade digitada for numérica, isto é:
        ' não for um dos códigos diretos, como "UMC", "UNA", etc), coloca no log pra adicionar à lista depois.
        prProcesso.Divisao = Right(strUnidadeDiv, Len(strUnidadeDiv) - InStr(1, strUnidadeDiv, "///") - 2)
        If IsNumeric(form.txtCodLocal.text) And prProcesso.Unidade <> "" And prProcesso.Divisao = "" Then
            cfLogDivisoes.Cells(1, 1).End(xlDown).Offset(1, 0).Formula = form.txtCodLocal
        End If
    End If
    
    ''''''''''''''''''''
    ''' Providências '''
    ''''''''''''''''''''
    ' Insere a observação do cadastro
    If form.txtObs.text <> "" Then strObservacao = "Obs. do cadastro: " & form.txtObs.text & vbCr & strObservacao
    strObservacao = "Sísifo: " & prProcesso.NomeAndamento & " - " & prProcesso.DataAndamento & vbCr & strObservacao
    
    ' Adiciona providência de Contestar
    Set contProvidencias = CriarProvidenciaContestar(CDate(Trim(form.txtDataAndamento.text)), strObservacao)
    If Not contProvidencias Is Nothing Then
        prProcesso.Providencias.Add contProvidencias
        dtDataProvContestar = contProvidencias.DataFinal
        If form.cmbAdvogado.Tag <> "" Then Set planGrupoProvContestar = ThisWorkbook.Sheets("cf" & form.cmbAdvogado.Tag)
    End If
    
    ' Adiciona providência de Analisar processo novo
    Set contProvidencias = CriarProvidenciaAnalisarProcesso(strObservacao, form.chbAnalisarProcNovo.Value)
    If Not contProvidencias Is Nothing Then prProcesso.Providencias.Add contProvidencias
    
    ' Adiciona providência de Levantar subsídios
    ' Insere Levantar Subsídios
    Set contProvidencias = CriarProvidenciaLevantarSubsidios(dtDataProvContestar, strObservacao, arrNucleoPrep)
    If Not contProvidencias Is Nothing Then
        prProcesso.Providencias.Add contProvidencias
        dtDataProvSubsidios = contProvidencias.DataFinal
        If arrNucleoPrep(3) <> "" Then Set planGrupoProvSubsidios = ThisWorkbook.Sheets("cf" & Trim(arrNucleoPrep(3)))
    End If
    
    ' Se houver providência adicional, cadastra como a primeira
    Set contProvidencias = CriarProvidenciaAdicional(form)
    If Not contProvidencias Is Nothing Then prProcesso.Providencias.Add contProvidencias, Before:=1
    
    If prProcesso Is Nothing Then GoTo EncerrarCadastro
    
    ' Se for Projudi de Lauro de Freitas/BA, adiciona providência de Triar Sentença para 20 dias depois da audiência
    If prProcesso.DataAndamento <> 0 And prProcesso.Comarca = "Lauro de Freitas" And prProcesso.Sistema = sfProjudi Then
        Set contProvidencias = New providencia
        contProvidencias.Nome = "Triar Sentença / Acórdão e atualizar Provisionamento"
        contProvidencias.Alertar = False
        contProvidencias.DataAlerta = CDate(prProcesso.DataAndamento + 19)
        contProvidencias.DataFinal = CDate(prProcesso.DataAndamento + 20)
        contProvidencias.Nucleo = prProcesso.Nucleo
        contProvidencias.Responsavel = prProcesso.Advogado
        contProvidencias.Observacao = "Providência criada pelo Sísifo, 20 dias após audiência - intimações pré-datadas de Lauro de Freitas"
        contProvidencias.GerarAndamento = False
        prProcesso.Providencias.Add contProvidencias
    End If
    
    '' Se for > 20 salários de comarcas configuradas para autoagendar pautista, adiciona providência de Agendar advogado
    'bolMaior20SM = IIf(prProcesso.ValorCausa > 20 * cfConfigurações.Cells().Find(what:="Valor do salário mínimo", lookat:=xlWhole).Offset(0, 1).Formula, True, False)
    'bolAgendaPautista = AutoAgendaPautista(Right(prProcesso.NumeroProcesso, 4), Left(prProcesso.Nucleo, 5))
    'If bolMaior20SM = True And bolAgendaPautista = True And prProcesso.DataAndamento <> 0 Then
    '    Set contProvidencias = New Providencia
    '    contProvidencias.Nome = "Agendar advogado pautista"
    '    contProvidencias.Alertar = False
    '    contProvidencias.DataAlerta = Application.WorksheetFunction.WorkDay(Date, 2)
    '    contProvidencias.DataFinal = Application.WorksheetFunction.WorkDay(prProcesso.DataAndamento, -2)
    '    contProvidencias.Nucleo = prProcesso.Nucleo
    '    contProvidencias.Responsavel = cfConfigurações.Cells().Find(what:="Responsável pelo agendamento de pautista", lookat:=xlWhole).Offset(0, 1).Formula
    '    contProvidencias.Observacao = "Audiência Una - " & prProcesso.DataAndamento
    '    contProvidencias.GerarAndamento = False
    '    prProcesso.Providencias.Add contProvidencias
    'End If
    
    '''''''''''''''
    ''' Pedidos '''
    '''''''''''''''
    
    ' Descobre quantas linhas de pedidos existem no formulário
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
        
        ' Se já tiver sido adicionado pedido com esse nome, pula (pedidos com mesmo no mesmo processo podem causar bug na interpretação da planilha pelo Espaider)
        For Cont2 = prProcesso.Pedidos.Count To 1 Step -1
            If strNomePedido = prProcesso.Pedidos(Cont2).Nome Then bolPularPedido = True
        Next Cont2
            
        If bolPularPedido = False Then
            ReDim varCont(1 To 5)
            varCont(1) = cfPedidos.Cells().Find(What:=form.Controls("cmbPed" & Cont).Value, LookAt:=xlWhole).Offset(0, -1).Formula ' Código do pedido
            varCont(2) = form.Controls("cmbPed" & Cont).Value  ' Nome do pedido
            varCont(3) = CCur(IIf(Trim(form.Controls("txtPed" & Cont).text) = "", 0, form.Controls("txtPed" & Cont).text)) ' Valor pedido; se for vazio, coloca 0
            varCont(4) = IIf(Trim(form.Controls("cmbRisco" & Cont).Value) = "", "Remoto", Trim(form.Controls("cmbRisco" & Cont).Value)) ' Risco; se for vazio, coloca Remoto
            
            If prProcesso.gerencia = "PPJCT" Then
                If Trim(form.Controls("txtProv" & Cont).text) = "" Or Trim(form.Controls("txtProv" & Cont).text) = 0 Then
                    varCont(5) = CCur(prProcesso.ValorCausa / btQtdLinhas) ' Valor do risco a provisionar; se for trabalhista e estiver zerado, divide valor da causa por igual
                Else
                    varCont(5) = CCur(form.Controls("txtProv" & Cont).text) ' Valor do risco a provisionar; se for trabalhista e não estiver zerado, usa o valor que está no formulário
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
'' Confere se as matrículas de uma comarca devem ser bloqueadas para cobrança ou negativação.
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
        ' Se não achar o Código da comarca na planilha, retorna Falso e sai da função.
        PegaBloqueiaMatricula = False
    Else
        ' Se achar o código da comarca, retorna Verdadeiro ou Falso conforme Sim ou Não e sai da função.
        PegaBloqueiaMatricula = IIf(rngCont.Offset(0, 5).Formula = "Sim", True, False)
    End If

End Function

Function PegaNatureza(plRelacaoCausaDePedirNatureza As Worksheet, strCausaPedir As String) As String

    Dim rngCont As Excel.Range

    Set rngCont = plRelacaoCausaDePedirNatureza.Cells().Find(What:=strCausaPedir, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu não conheço a natureza da causa de pedir deste processo (""" & strCausaPedir & """). " & _
        "Rogo que a cadastre em minha memória e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Natureza da causa de pedir desconhecida"
        PegaNatureza = "Natureza da causa de pedir não cadastrada"
    Else
        PegaNatureza = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaGerencia(plRelacaoCausaDePedirGerencia As Worksheet, strCausaPedir As String) As String

    Dim rngCont As Excel.Range

    Set rngCont = plRelacaoCausaDePedirGerencia.Cells().Find(What:=strCausaPedir, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", eu não conheço a gerência relacionada à causa de pedir deste processo (""" & strCausaPedir & """). " & _
        "Rogo que a cadastre em minha memória e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Gerência da causa de pedir desconhecida"
        PegaGerencia = "Gerência da causa de pedir não cadastrada"
    Else
        PegaGerencia = rngCont.Offset(0, 3).Formula
    End If
    
End Function

Function PegaComarca(ByRef strJuizo As String) As String
    
    Dim rngCont As Excel.Range
    
    Set rngCont = cfJuizos.Range("B:B").Find(What:=strJuizo, LookAt:=xlWhole)
    
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", não há registro da comarca dos processo que tramitam no juízo (""" & strJuizo & """) na minha memória. " & _
            "Rogo que o cadastre em minha memória e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Juízo não cadastrado"
        PegaComarca = "Juízo não cadastrado, impossível encontrar comarca"
    Else
        PegaComarca = Trim(rngCont.Offset(0, 1).Formula)
    End If
    
End Function

Function PegaOrgao(ByRef strJuizo As String) As String
    
    Dim rngCont As Excel.Range
    
    Set rngCont = cfJuizos.Range("B:B").Find(What:=strJuizo, LookAt:=xlWhole)
    
    If rngCont Is Nothing Then
        MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", não há registro da comarca dos processo que tramitam no juízo (""" & strJuizo & """) na minha memória. " & _
            "Rogo que o cadastre em minha memória e tente novamente.", vbCritical + vbOKOnly, "Sísifo - Juízo não cadastrado"
        PegaOrgao = "Juízo não cadastrado, impossível encontrar comarca"
    Else
        PegaOrgao = Trim(rngCont.Offset(0, 2).Formula)
    End If
    
End Function

Function PegaMatricula(ByRef strSci() As String, ByRef strObservacao As String) As String
    Dim strMatricula As String, strProcessosMatricula() As String, strCpjCnpjPje As String
    Dim intResposta As Integer
    Dim bolAutorEhTitular As Boolean
    
PerguntaMatricula:
    strMatricula = Trim(InputBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", favor informar a matrícula relacionada ao processo (se não for o caso, deixe o campo em branco)", "Sísifo - Matrícula de processo", ""))
    
    ' Se tiver matrícula, resgata código da localidade e nome da parte pelo web service do SCI
    If strMatricula <> "" Then
        strSci(1, 1) = "erroCodigo"
        strSci(2, 1) = "erroMensagem"
        strSci(3, 1) = "inscricao"
        strSci(4, 1) = "numCPFCNPJ"
        strSci(5, 1) = "nomeCliente"
        ConsultaMatriculaSCI strMatricula, strSci
        
        If strSci(1, 2) <> 0 Then ' Se deu erro na consulta ao web service do SCI, pergunta se o usuário quer continuar assim mesmo.
            intResposta = MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", o SCI retornou o seguinte erro: """ & strSci(2, 2) & """. Deseja continuar mesmo sem a integração com o SCI?", _
                vbYesNo + vbQuestion + vbDefaultButton2, "Sísifo - Erro de integração com o SCI")
            
            If intResposta = vbYes Then
                strObservacao = "ATENÇÃO: Este cadastro foi feito sem integração automática com o SCI." & vbCr
                PegaMatricula = ""
                Exit Function
            Else
                Set prProcesso = Nothing
                Exit Function
            End If
        End If
        
        ' Confere se o Autor é o titular da matrícula
        strCpjCnpjPje = IIf(prProcesso.OutrosParticipantes(1).TipoParte = 1, prProcesso.OutrosParticipantes(1).CPFParte, prProcesso.OutrosParticipantes(1).CNPJParte)
        strCpjCnpjPje = Replace(strCpjCnpjPje, ".", "")
        strCpjCnpjPje = Replace(strCpjCnpjPje, "-", "")
        strCpjCnpjPje = Replace(strCpjCnpjPje, "/", "")
        
        ' Se o CPF/CNPJ não estiver cadastrado no processo, compara pelo nome
        If InStr(1, LCase(strCpjCnpjPje), "não cadastrado") <> 0 Then
            bolAutorEhTitular = IIf(Trim(LCase(strSci(5, 2))) = Trim(LCase(prProcesso.OutrosParticipantes(1).NomeParte)), True, False)
            If bolAutorEhTitular = False Then
                If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", aparentente, o nome do Autor do processo (" & Trim(prProcesso.OutrosParticipantes(1).NomeParte) & _
                ") é diferente do nome do titular da matrícula (" & strSci(5, 2) & "). Favor conferir se a matrícula do processo é realmente a que " & _
                "foi digitada -- " & strMatricula & ". Para prosseguir com a matrícula informada, clique em ""OK"". Para informar uma nova matrícula, " & _
                "clique em ""Cancelar"".", vbQuestion + vbOKCancel, "Sísifo - Autor diferente do titular da matrícula") = vbCancel Then _
                GoTo PerguntaMatricula
            End If
        Else
        'Se o CPF/CNPJ estiver cadastrado no processo, compara pelo CPF/CNPJ mesmo
            bolAutorEhTitular = IIf(Trim(LCase(strSci(4, 2))) = strCpjCnpjPje, True, False)
            If bolAutorEhTitular = False Then
                If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", aparentente, o CPF/CNPJ do Autor do processo (" & Trim(strCpjCnpjPje) & ") " & _
                "é diferente do CPF/CNPJ do titular da matrícula (" & strSci(4, 2) & "). Favor conferir se a matrícula do processo é realmente a que " & _
                "foi digitada -- " & strMatricula & ". Para prosseguir com a matrícula informada, clique em ""OK"". Para informar uma nova matrícula, " & _
                "clique em ""Cancelar"".", vbQuestion + vbOKCancel, "Sísifo - Autor diferente do titular da matrícula") = vbCancel Then _
                GoTo PerguntaMatricula
            End If
        End If
        
        ' Ajusta o código da localidade
        strSci(3, 2) = Left(strSci(3, 2), InStr(1, strSci(3, 2), ".") - 1)
        strSci(3, 2) = Format(strSci(3, 2), "0000") ' Pega o código da localidade
        
        ' Ajusta observação das providências
        If bolAutorEhTitular = False Then strObservacao = "O Autor não é titular da matrícula" & vbCr
        
        ' Se tiver matrícula, resgata processos em COPJ pelo web service do SCI
        strProcessosMatricula = ConsultaProcessosPorMatriculaSCI(strMatricula)
        
        ' Resgata informação sobre outros processos cadastrados
        If UBound(strProcessosMatricula) <> 0 Then
            strObservacao = strObservacao & "Existem outros processos da matrícula:" & vbCr
            For Cont = 1 To UBound(strProcessosMatricula, 1)
                strObservacao = strObservacao & "   - Número: " & strProcessosMatricula(Cont, 1) & "; Autor: " & strProcessosMatricula(Cont, 2) & "; Causa de pedir: " & strProcessosMatricula(Cont, 3) & vbCr
            Next Cont
        Else
            strObservacao = strObservacao & "Não existem outros processos cadastrados na matrícula" & vbCr
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

    PegaUnidadeDivisao = rngCont.Offset(0, 2).Formula & "///" & rngCont.Offset(0, 4).Formula '/// é o separador

End Function

Function PegaNomePedido(ByVal strCodPedido As String) As String

    Dim rngCont As Excel.Range
    
    Set rngCont = cfPedidos.Cells().Find(What:=strCodPedido, LookAt:=xlWhole)
    If rngCont Is Nothing Then
        PegaNomePedido = "Pedido não cadastrado"
    Else
        PegaNomePedido = rngCont.Offset(0, 1).Formula
    End If
    
End Function

Function PegaPedido(strCausaPedir As String, strCodPedido As String, strPedido As String, curValorPedido As Currency, strRisco As String, curValorProvisionar As Currency) As Pedido
''
'' Busca o prognóstico e valores padrão para combinações de causa de pedir e pedido. Não encontrando,
'' retorna sempre prognóstico remoto e valores zerados.
''
    Dim rngCont As Range
    Dim strPrimeiroEndereco As String
    Dim pePedido As Pedido
    
    Set pePedido = New Pedido
    
    Set rngCont = cfCausasPedirPedidos.Cells().Find(strCausaPedir)
    strPrimeiroEndereco = rngCont.Address
    
    Do
        If rngCont.Formula = strCausaPedir And rngCont.Offset(0, 1).Formula = strCodPedido Then
            'Se achar a combinação de Causa de Pedir e pedido, retorna e sai da função.
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
    
    ' Se não achar, retorna "Remoto" com valores zerados e sai da função.
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
                    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(1) & ". A inclusão foi cancelada e descartados os dados." & vbCrLf & _
                    "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
                    Set prProcesso = Nothing
                    Exit Function
                ElseIf Left(LCase(arrNucleoPrep(2)), 4) = "erro" Then
                    MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(2) & ". A inclusão foi cancelada e descartados os dados." & vbCrLf & _
                    "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
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
            If MsgBox(SisifoEmbasaFuncoes.DeterminarTratamento & ", a providência adicional não tem um nome ou uma data válida, e por isso seu humilde " & _
            "e leal servo não poderá cadastrá-la. Caso deseje cadastrar o processo normalmente sem providência adicional, clique em ""OK"". Para cancelar " & _
            "o cadastro, clique em ""Cancelar"".", vbQuestion + vbOKCancel, "Sísifo - Erro no andamento adicional") = vbCancel Then
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
        
    Case Else ' PPJCE pediu para não criar Contestar, e sim Analisar processo novo.
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
            MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(1) & ". A inclusão foi cancelada e descartados os dados." & vbCrLf & _
            "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
            Set prProcesso = Nothing
            Exit Function
        ElseIf Left(LCase(arrNucleoPrep(2)), 4) = "erro" Then
            MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve o seguinte erro: """ & arrNucleoPrep(2) & ". A inclusão foi cancelada e descartados os dados." & vbCrLf & _
            "Processo: " & prProcesso.NumeroProcesso & vbCrLf, vbCritical + vbOKOnly, "Sísifo - Erro na busca de núcleo e advogado"
            Set prProcesso = Nothing
            Exit Function
        Else
            prProcesso.Preposto = Trim(arrNucleoPrep(2))
        End If
    End If
    
    ' Ajusta as observações da providência
    Select Case prProcesso.gerencia
    Case "PPJCT"
        ReDim arrNucleoPrep(1 To 3)
        arrNucleoPrep(1) = ""
        arrNucleoPrep(2) = ""
        arrNucleoPrep(3) = ""
        
        strObservacao = strObservacao & "Processo: " & prProcesso.NumeroProcesso & vbCr
        strObservacao = strObservacao & "Reclamante: " & prProcesso.OutrosParticipantes(1).NomeParte & vbCr
        strObservacao = strObservacao & "Responsável: " & prProcesso.Advogado & vbCr
        strObservacao = strObservacao & "Juízo: " & prProcesso.Juizo & vbCr
        strObservacao = strObservacao & "Data da solicitação: " & Format(Date, "dd/mm/yyyy") & vbCr
        strObservacao = strObservacao & vbCr
        strObservacao = strObservacao & "SUBSÍDIOS RELEVANTES:" & vbCr
        strObservacao = strObservacao & "- Contrato de trabalho" & vbCr
        strObservacao = strObservacao & "- Ficha de Registro" & vbCr
        strObservacao = strObservacao & "- Ficha funcional e financeira" & vbCr
        strObservacao = strObservacao & "- Folha de ponto últimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Contracheques últimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante de férias últimos 5 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante do pagamento do 13º: [ ] 1ª parcela / [ ] 2ª parcela" & vbCr
        strObservacao = strObservacao & "- Termo de Compensação de Jornada" & vbCr
        strObservacao = strObservacao & "- Cópia de atestados médicos e/ou Concessão de Afastamento Previdenciário" & vbCr
        strObservacao = strObservacao & "- Cópia de advertência e suspensões aplicadas ao Reclamante" & vbCr
        strObservacao = strObservacao & "- RIP" & vbCr
        strObservacao = strObservacao & "- Gratificação Motorista Usuário" & vbCr
        strObservacao = strObservacao & "- Avaliações PCCS 2009" & vbCr
        strObservacao = strObservacao & "- Acordo Coletivo" & vbCr
        strObservacao = strObservacao & "- Plano de Demissão Voluntária" & vbCr
        strObservacao = strObservacao & "- PCCS/2009 e a homologação publicada no DOU" & vbCr
        strObservacao = strObservacao & "- PCCS/1986" & vbCr
        strObservacao = strObservacao & "- PCCS/1998" & vbCr
        strObservacao = strObservacao & "- Comprovante de pagamento do auxílio-creche ou da justificativa para o indeferimento do beneficio" & vbCr
        strObservacao = strObservacao & "- Cópia da reclamação inicial e da sentença de OUTRO PROCESSO SE EXISTIR" & vbCr
        strObservacao = strObservacao & "- Cópia da homologação do acordo" & vbCr
        strObservacao = strObservacao & "- Cópia da petição de acordo" & vbCr
        strObservacao = strObservacao & "- Processo Administrativo da demissão, Termo de Rescisão do contrato de trabalho, e demais comprovantes de pagamento da rescisão" & vbCr
        strObservacao = strObservacao & "- Comprovante de entrega das guias de seguro desemprego" & vbCr
        strObservacao = strObservacao & "- Recibo de recolhimento de FGTS dos últimos 30 anos" & vbCr
        strObservacao = strObservacao & "- Comprovante adesão ao PAT" & vbCr
        strObservacao = strObservacao & "- BDV's (Boletins diários de veículos)" & vbCr
        strObservacao = strObservacao & "- Cópia do RMU (Regulamento de Motorista Usuário)" & vbCr
        strObservacao = strObservacao & "- Cópia de pagamento de diárias para viagens" & vbCr
        strObservacao = strObservacao & "- Comprovante de pagamento da indenização por acidente de trabalho" & vbCr
        strObservacao = strObservacao & "- Instrumento que traz as normas do PPR" & vbCr
        strObservacao = strObservacao & "OBS.: "
        
    Case "PPJCM"
        If InStr(1, LCase(prProcesso.CausaPedir), "desabastecimento") > 0 Then
            strObservacao = strObservacao & "Nota técnica esclarecendo APENAS os pontos favoráveis à Embasa:" & vbCr & _
                    "1. Se houve intervenção na rede no período capaz de gerar desabastecimento;" & vbCr & _
                    "2. Se esse período estava abrangido nos casos de manutenção e reparo de emergência;" & vbCr & _
                    "3. Se o abastecimento poderia ser suprido por reservatório adequado;" & vbCr & _
                    "3. Caso o imóvel precise de reservatório inferior com bomba, e não o possui (anexar fotos, se possível; se não, a tela do SCI);" & vbCr & _
                    "4. Se não houve alteração de consumo (anexar HCON);" & vbCr & _
                    "5. Se a água estava cortada por outro motivo (anexar SS do corte);" & vbCr & _
                    "6. Se houve reclamação de desabastecimento (anexar Histórico de SS no período);"
        Else
            Select Case prProcesso.CausaPedir
            Case "Negativação no SPC"
                strObservacao = strObservacao & "1. Em caso de negativa de vínculo com a Embasa, desconhecimento do contrato, etc: " & _
                        "Documentos ASSINADOS pelo autor do processo, comprovando relação com a Embasa, como solicitação de ligação " & _
                        "assinada, parcelamento assinado, notificação de débito assinada. Não havendo estes, fazer relatório ASSINADO " & _
                        "explicando o resultado de visita ao imóvel para verificar se está habitado e quem habita, comprovando através " & _
                        "vínculo. Informação em OS sem assinatura geralmente é desconsiderada pelo juiz." & vbCr & _
                        "2. Quando reconhece o vínculo, negando apenas débitos específicos, relatório assinado descrevendo o ocorrido " & _
                        "trazer segundas vias das contas negativadas e, se possível, segundas vias de contas posteriores com aviso de " & _
                        "existência de débito. Sendo indevida a negativação, inserir proposta de acordo."
                
            Case "Revisão de consumo elevado"
                strObservacao = strObservacao & "Relatório assinado informando normalidade do histórico de consumo, SSs de revisão de consumo, " & _
                "substituição de hidrômetro, aferição (se houver, anexar LAUDO do Ibametro, pois juízes não acatam SS de aferição) e " & _
                "tudo com fotos (se houver) e anexando as telas separadas."
                
            Case "Consumo elevado com corte"
                strObservacao = strObservacao & "Relatório assinado informando normalidade do histórico de consumo, SSs de revisão de consumo, " & _
                "substituição de hidrômetro, aferição (se houver, anexar LAUDO do Ibametro, pois juízes não acatam SS de aferição) e " & _
                "tudo com fotos (se houver) e anexando as telas separadas. Notificação de débito assinada (sem assinatura, as chances são MUITO menores)."
                
            Case "Corte no fornecimento"
                strObservacao = strObservacao & "Relatório assinado explicando o ocorrido. Segundas vias de faturas cujo inadimplemento motivaram o corte. Notificação de débito " & _
                "assinada (sem assinatura, as chances são MUITO menores; se for o caso, esclarecer no relatório o porquê de não estar assinada)."
                
            Case "Corte por iniciativa do cliente"
                strObservacao = strObservacao & "Relatório assinado explicando o ocorrido. Solicitação de serviço assinada ou, se foi por telefone, gravação da ligação (sem assinatura/gravação, as chances são MUITO menores)"
                
            Case "Consumo Rateado - Medição Individualizada"
                strObservacao = strObservacao & "1. Relatório assinado explicando o ocorrido." & vbCr & _
                "2. Solicitação de individualização, ata de assembleia ou contrato específico assinados pelo autor do processo;" & vbCr & _
                "3. Fotos do hidrômetro geral e das fileiras dos hidrômetros individuais;" & vbCr & _
                "4. HCON do hidrômetro geral e do hidrômetro individual da parte Autora."
                
            Case "Cobrança de esgoto em imóvel não ligado à rede"
                strObservacao = strObservacao & "1. Nota técnica com geoweb e fotos mostrando existência de rede no local e a possibillidade de " & _
                "ligação da residência (verificar se a residência do consumidor tem cota topográfica, e quaisquer outras questões relevantes);" & vbCr & _
                "2. Notificação assinada para interligar o imóvel à rede quando da implantação (procurar na DE, DM ou DI, se for o caso);" * vbCr & _
                "3. Não é essencial, mas, caso esteja disponível, Ordem de Serviço da obra de implantação da rede no local."
                
            Case "Cobrança de esgoto com água cortada"
                strObservacao = strObservacao & "Relatório assinado explicando o ocorrido. Anexar prova fotográfica, ou algum outro tipo de prova, de " & _
                "que o imóvel estava habitado no período (mera observação nas OSs geralmente é desconsiderada pelo juiz)."
                
            Case "Fixo de esgoto"
                strObservacao = strObservacao & "Nota técnica assinada, com fotos demonstrando a existência de abastecimento alternativo e tabela das datas e leituras da " & _
                "medição do abastecimento alternativo."
                
            Case "Realizar ligação de água", "Desmembramento de ligações"
                strObservacao = strObservacao & "Nota técnica assinada, esclarecendo o motivo técnico para a ligação não ter sido realizada, com PROVAS deste motivo " & _
                "(por exemplo, CONFORME O CASO, fotos do imóvel ou relatório de vistoria mostrando que não há reservação adequada, ou medições de pressão " & _
                "com explicação do porquê da necessidade de reservatório inferior e bomba, ou relatório fotográfico de que as instalações internas não são " & _
                "desmembradas, etc."
                
            Case "Vaz. água ou extravas. esgoto com danos a patrimônio/morais", "Obra da Embasa com danos a patrimônio/morais", _
                "Acidente com pessoa/veículo em buraco", "Acidente com veículo (colisão ou atropelamento)"
                strObservacao = strObservacao & "Nota técnica e relatório fotográfico esclarecendo a verdade do incidente e eventuais medidas tomadas pela Embasa " & _
                "(sempre acompanhada das provas possíveis)."
                
            Case "Multa por infração", "Suspeita de by-pass"
                strObservacao = strObservacao & "Nota técnica assinada descrevendo o tipo de gato realizado (bypass, furo, palheta, etc), com fotos e medidas tomadas pela Embasa."
                
            End Select
        End If
    End Select
        
    ' Cria a providência, conforme gerência
    Select Case prProcesso.gerencia
    Case "PPJCM"
        If prProcesso.Preposto = "" Then
            Set contProvidencia = Nothing
        Else
            Set contProvidencia = New providencia
            With contProvidencia
                .Nome = "Levantar subsídios"
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
                .Nome = "Levantar subsídios"
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
            .Nome = "Levantar subsídios"
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
'' Procura a combinação de código de comarca/localidade e causa de pedir passadas como parâmetro na planilha passada como parâmetro.
'' Se não achar a combinação exata, procura o código da comarca/localidade com causa de pedir "Geral" ou em branco.
'' Retorna o nome do núcleo, seguido da string "-,-,-" e do nome do advogado. Se não encontrar o código da comarca (ou a linha geral
'' do código da comarca, se for o caso), retorna uma mensagem de erro iniciada por "Erro em PegaNucleoAdvPrep - ".
''

    Dim plan As Worksheet
    Dim strCodBusca As String, strNucleo As String, strAdvPrep As String, strNomeGrupo As String, strPrimeiroEndereco As String
    Dim strEnderecoCodBusca As String
    Dim arrCont() As String
    Dim rngCont As Excel.Range
    
    strEnderecoCodBusca = "A:A"
    
    ' Define a planilha que contém os responsáveis e o código de busca a ser utilizado
    Select Case TipoDeResponsavel
    Case Advogado
        Select Case LCase(prProcesso.Sistema)
        Case "projudi"
            Set plan = cfComarcasAdvsPPJCMProjudi
            strCodBusca = Right(strNumProc, 4) ' O advogado é definido pela comarca do processo (últimos 4 números, no TJ/BA)
        Case "pje1g"
            Select Case LCase(prProcesso.gerencia)
            Case "ppjcm"
                Set plan = cfComarcasAdvsPPJCMPJe
                strCodBusca = prProcesso.Comarca ' O advogado é definido pela comarca do processo (nome da comarca)
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
            strCodBusca = Trim(Replace(Left(prProcesso.Unidade, 5), "-", "")) ' O código de uma unidade pode ter até 5 letras, por isso pegamos as 5, removemos o hífen eventual e os espaços (o que transformará as unidades de 3 letras de "USU -" em "USU", e também funcionará pra as de 4 e 5 letras.)
            strCodBusca = strCodBusca & "-" & Replace(prProcesso.Divisao, "Escr. Local de", "EL") ' O formato final é, por exemplo, "USU-EL Eunápolis"
        Case "pje1g"
            Select Case LCase(prProcesso.gerencia)
            Case "ppjcm"
                Set plan = cfComarcasPrepPPJCMPJe
                strCodBusca = Trim(Replace(Left(prProcesso.Unidade, 5), "-", "")) ' O código de uma unidade pode ter até 5 letras, por isso pegamos as 5, removemos o hífen eventual e os espaços (o que transformará as unidades de 3 letras de "USU -" em "USU", e também funcionará pra as de 4 e 5 letras.)
                strCodBusca = strCodBusca & "-" & Replace(prProcesso.Divisao, "Escr. Local de", "EL") ' O formato final é, por exemplo, "USU-EL Eunápolis"
            Case "ppjce"
                Set plan = cfComarcasPrepPPJCEPJe
            'Case "ppjct"
            '    Set plan = strNomePlanilha
            End Select
        End Select
    End Select
    
InicioBusca:
    ' Busca o responsável pela causa de pedir E CodBusca
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, strCausaPedir, False, strCodBusca, False, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ' Busca o responsável só pelo CodBusca, aceitando Causa de Pedir "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, "", True, strCodBusca, False, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    Else
        If TipoDeResponsavel = Preposto And InStr(1, strCodBusca, "-") <> 0 Then
        ' Se não achou uma linha geral da localidade do preposto e já não está buscando pela Unidade Regional, recomeça a buscar pela Unidade Regional
            strCodBusca = Left(strCodBusca, InStr(1, strCodBusca, "-") - 1)
            GoTo InicioBusca
        End If
    End If
    
    ' Busca o responsável só pela causa de pedir, aceitando CodBusca "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, strCausaPedir, False, "", True, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ' Busca o responsável aceitando tudo "Geral"
    Set rngCont = IteraPlanilhaNucleoAdvPrep(plan, "", True, "", True, bolAceitaCausaPedirSemResponsavel)
    If Not rngCont Is Nothing Then
        strNucleo = rngCont.Formula
        strAdvPrep = rngCont.Offset(0, 1).Formula
        GoTo Encontrado
    End If
    
    ReDim arrCont(1 To 1)
    arrCont(1) = "Erro em PegaNucleoAdvPrep - Responsável não encontrado para o código """ & strCodBusca & """ na planilha " & plan.Name & " com booleana de aceitar causa de pedir """ & bolAceitaCausaPedirSemResponsavel & """."
    PegaNucleoAdvPrep = arrCont
    Exit Function
    
Encontrado:
    
    ' NÃO ESTÁ LIDANDO COM GRUPOS DE NÚCLEOS. CASO NECESSÁRIO, REPROGRAMAR O CÓDIGO DE ACORDO (5 LINHAS COMENTARIZADAS ABAIXO).
    ' Se for um grupo de núcleos (nomes de grupos de núcleos devem começar com o código "grpnuc", dois números e um hífen),
    ' chama a função que itera isonomicamente pelos membros do grupo.
    'If Left(strNucleo, 6) = "grpnuc" And IsNumeric(Mid(strNucleo, 7, 2)) And Mid(strNucleo, 9, 1) = "-" Then
    '    strNucleo = IteraNucleoAdv(strNucleo, "Nucleo")
    'End If
    
    ' Se for um grupo de advogados ou prepostos (nomes de grupos de advogados ou prepostos devem começar com o código "grpadv" ou "grpprep", dois números e um hífen),
    ' chama a função que itera isonomicamente pelos membros do grupo.
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
'' Itera as células da coluna de Causa de Pedir da planilha passada como parâmetro, buscando a combinação de código de comarca/localidade e
''  causa de pedir passadas como parâmetro. Se não achar a combinação exata, retorna Nothing.
''
    Dim rngColunaCausaPedir As Excel.Range, rngCel As Excel.Range
    
    Set rngColunaCausaPedir = plan.Range("C1:C" & plan.UsedRange.Rows(plan.UsedRange.Rows.Count).Row)
    
    If strCausaPedir = "" Then strCausaPedir = "/./"
    If strCodBusca = "" Then strCodBusca = "/./"
    
    For Each rngCel In rngColunaCausaPedir.Cells
        If (LCase(rngCel.Formula) = LCase(strCausaPedir) Or (bolAceitaCausaPedirGeral = True And LCase(rngCel.Formula) = "geral")) And _
            (LCase(rngCel.Offset(0, -2).Formula) = LCase(strCodBusca) Or (bolAceitaCodBuscaGeral = True And LCase(rngCel.Offset(0, -2).Formula) = "geral")) Then
            If Trim(rngCel.Offset(0, 2).Formula = "") And bolAceitaCausaPedirSemResponsavel = False Then 'Se o responsável estiver em branco, mas não for para aceitar, volta Nothing
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
'' Itera a planilha do grupo e retorna o responsável pela providência no dia do prazo especificado
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
        PegaResponsavelProvidenciaGrupo = "Erro - a data final do prazo não está cadastrada na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    
    lngLinhaData = rngCont.Row
    lngColunaCont = plan.UsedRange.Columns.Count
    
    lngLinhaData = rngCont.Row
    Set rngCont = Nothing
    
    lngColunaCont = plan.Columns.Count
    lngColunaCont = plan.Cells(10, lngColunaCont).End(xlToLeft).Column
    
    ' Pegar quantidade já distribuída para a data final em questão
    intTotalGeralDia = plan.Cells(lngLinhaData, lngColunaCont).Value2
    If intTotalGeralDia = 0 Then intTotalGeralDia = 1
    
    ' Coletar advogados, calcular a Carga Diária Percentual de cada um e transformá-la na variação entre CDP e CDP pretendida
    For intCont = 1 To lngColunaCont - 2
        strAdvCont = plan.Cells(2, intCont + 1).Value2
        dicResponsaveis.Add strAdvCont, plan.Cells(lngLinhaData, intCont + 1).Value2
        dicResponsaveis(strAdvCont) = dicResponsaveis(strAdvCont) / intTotalGeralDia ' Carga diária percentual
        dicResponsaveis(strAdvCont) = dicResponsaveis(strAdvCont) - plan.Cells(4, intCont + 1).Value2 ' Diferença em relação à carga diária percentual pretendida = CDP - CDP pretendida
    Next intCont
    
    ' Remove os valores que não forem os menores
    Set dicResponsaveis = PegaValoresMenores(dicResponsaveis)
    
    ' Em caso de empate, pegar o total geral de providências de todos os tempos dos empatados (ponderado pela carga de trabalho da pessoa).
    If dicResponsaveis.Count > 1 Then
        For intCont = 0 To dicResponsaveis.Count - 1
            strAdvCont = dicResponsaveis.Keys(intCont)
            lngColunaCont = plan.Cells().Find(strAdvCont, LookAt:=xlWhole, searchorder:=xlByRows).Column
            dicResponsaveis(strAdvCont) = plan.Cells(5, lngColunaCont).Value / plan.Cells(3, lngColunaCont).Value
        Next intCont
        
        Set dicResponsaveis = PegaValoresMenores(dicResponsaveis)
    End If
    
    ' Em caso de novo empate, aleatório
    If dicResponsaveis.Count > 1 Then
        Randomize
        intCont = CInt((dicResponsaveis.Count - 1) * Rnd)
    Else
        intCont = 0
    End If
    
    ' Registra a providência para o responsável e retorna o valor procurado
    lngColunaCont = plan.Cells().Find(dicResponsaveis.Keys(intCont), LookAt:=xlWhole, searchorder:=xlByRows).Column
    'plan.Cells(lngLinhaData, lngColunaCont).Value = plan.Cells(lngLinhaData, lngColunaCont).Value + 1
    PegaResponsavelProvidenciaGrupo = dicResponsaveis.Keys(intCont)
    
End Function

Function RegistraResponsavelPorProvidenciaNoGrupo(ByRef plan As Excel.Worksheet, dtDataPrazo As Date, strResponsavel As String) As String
''
'' Itera a planilha do grupo e retorna o responsável pela providência no dia do prazo especificado
''
    Dim lngLinhaData As Long, lngColunaResponsavel As Long
    Dim rngCont As Excel.Range
    
    ' Pega a linha da data
    Set rngCont = plan.Cells().Find(dtDataPrazo, after:=Cells(6, 1), LookAt:=xlWhole, searchorder:=xlByColumns)
    If rngCont Is Nothing Then
        RegistraResponsavelPorProvidenciaNoGrupo = "A data final do prazo não está cadastrada na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    lngLinhaData = rngCont.Row
    
    ' Pega a coluna do responsável
    Set rngCont = plan.Cells().Find(strResponsavel, LookAt:=xlWhole, searchorder:=xlByRows)
    If rngCont Is Nothing Then
        RegistraResponsavelPorProvidenciaNoGrupo = "O responsável não está cadastrado na planilha do grupo " & Replace(plan.Name, "cf", "", 1, 1) & "."
        Exit Function
    End If
    lngColunaResponsavel = rngCont.Column
    
    ' Registra a providência para o responsável
    plan.Cells(lngLinhaData, lngColunaResponsavel).Value = plan.Cells(lngLinhaData, lngColunaResponsavel).Value + 1
    RegistraResponsavelPorProvidenciaNoGrupo = "Sucesso"
    
End Function

Function PegaValoresMenores(ByVal Dict As Dictionary) As Dictionary
''
'' Itera uma coleção, excluindo todos os valores que não sejam os menores.
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
'' Faz uma validação front end, só permitindo números e, caso informados, os dois caracteres passados como parâmetros.
''
    Select Case ChaveAscii
    Case Asc("0") To Asc("9") 'Números são sempre permitidos
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




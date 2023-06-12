Attribute VB_Name = "sfAprDadGerenciaPlanilhas"
Sub RelatarProcessosArmazenadosCitacoes(ByVal Controle As IRibbonControl)
''
'' Contar processos e registros existentes nas planilhas da mem�ria
''
    Dim arrPlans(1 To 7) As Excel.Worksheet
    
    Set arrPlans(1) = sfCadProcessos
    Set arrPlans(2) = sfCadMatricula
    Set arrPlans(3) = sfCadAndamentos
    Set arrPlans(4) = sfCadProvidencias
    Set arrPlans(5) = sfCadPedidos
    Set arrPlans(6) = sfCadSemCPF
    Set arrPlans(7) = sfCadLitisc
    
    RelatarProcessosArmazenados arrPlans
    
End Sub

' Gerar planilha para Espaider (S� exclui ap�s usu�rio confimar que fez o upload no Espaider)
Sub ExportarPlanilhasCitacoesEspaider(ByVal Controle As IRibbonControl)
    Dim arq As Workbook
    Dim plan As Worksheet
    Dim bolProcessosVazia As Boolean, bolMatriculasVazia As Boolean, bolAndamentosVazia As Boolean, bolProvidenciasVazia As Boolean
    Dim bolPedidosVazia As Boolean, bolLitisVazia As Boolean, bolSemCPFVazia As Boolean
    Dim lnUltimaLinhaProcessos As Long, lnUltimaLinhaMatriculas As Long, lnUltimaLinhaAndamentos As Long
    Dim lnUltimaLinhaProvidencias As Long, lnUltimaLinhaPedidos As Long, lnUltimaLinhaLitis As Long
    Dim lnUltimaLinhasemCPF As Long
    Dim strDesktop As String
    Dim contX As Byte
    
    bolProcessosVazia = False
    bolMatriculasVazia = False
    bolAndamentosVazia = False
    bolProvidenciasVazia = False
    bolPedidosVazia = False
    bolLitisVazia = False
    bolSemCPFVazia = False
    
    ' Pergunta se deseja continuar, para o caso de ser apertado sem querer (percebi que � f�cil errar, bot�es
    ' pr�ximos, e essa tarefa s� roda uma vez por dia, portanto a confirma��o n�o sobrecarrega o usu�rio.
    If MsgBox("Deseja gerar a planilha de exporta��o no formato do Espaider?", vbQuestion + vbYesNo, _
    "S�sifo - Exportar processos?") = vbNo Then Exit Sub
    
    ' Testa a planilha CadProcessos
    lnUltimaLinhaProcessos = sfCadProcessos.UsedRange.Rows(sfCadProcessos.UsedRange.Rows.Count).Row
    If lnUltimaLinhaProcessos = 4 Then bolProcessosVazia = True
    
    ' Testa a planilha CadMatricula
    lnUltimaLinhaMatriculas = sfCadMatricula.UsedRange.Rows(sfCadMatricula.UsedRange.Rows.Count).Row
    If lnUltimaLinhaMatriculas = 4 Then bolMatriculasVazia = True
    
    ' Testa a planilha CadAndamentos
    lnUltimaLinhaAndamentos = sfCadAndamentos.UsedRange.Rows(sfCadAndamentos.UsedRange.Rows.Count).Row
    If lnUltimaLinhaAndamentos = 4 Then bolAndamentosVazia = True
    
    ' Testa a planilha CadProvidencias
    lnUltimaLinhaProvidencias = sfCadProvidencias.UsedRange.Rows(sfCadProvidencias.UsedRange.Rows.Count).Row
    If lnUltimaLinhaProvidencias = 4 Then bolProvidenciasVazia = True
    
    ' Testa a planilha CadPedidos
    lnUltimaLinhaPedidos = sfCadPedidos.UsedRange.Rows(sfCadPedidos.UsedRange.Rows.Count).Row
    If lnUltimaLinhaPedidos = 4 Then bolPedidosVazia = True
    
    ' Testa a planilha CadLitis
    lnUltimaLinhaLitis = sfCadLitisc.UsedRange.Rows(sfCadLitisc.UsedRange.Rows.Count).Row
    If lnUltimaLinhaLitis = 4 Then bolLitisVazia = True
    
    ' Testa a planilha CadSemCPF
    lnUltimaLinhasemCPF = sfCadSemCPF.UsedRange.Rows(sfCadSemCPF.UsedRange.Rows.Count).Row
    If lnUltimaLinhasemCPF = 4 Then bolSemCPFVazia = True
    
    ' Se estiverem todas vazias, avisa e para o procedimento
    If bolProcessosVazia = True And bolMatriculasVazia = True And bolAndamentosVazia = True And _
        bolProvidenciasVazia = True And bolPedidosVazia = True And bolLitisVazia And bolSemCPFVazia = True Then
        MsgBox "As planilhas de processos est�o vazias. N�o h� processos para exportar.", _
         vbInformation + vbOKOnly, "S�sifo - Planilhas vazias"
        Exit Sub
    End If

    ' Se as planilhas n�o estiverem vazias, exporta-as
    Set arq = Workbooks.Add
    
    ' Se houver mais de uma planilha na pasta de trabalho, exclui as demais
    Application.DisplayAlerts = False
    If arq.Sheets.Count > 1 Then
        For contX = arq.Sheets.Count To 2
            arq.Sheets(contX).Delete
        Next contX
    End If
    Application.DisplayAlerts = True
    
    If bolProcessosVazia = False Then _
        sfCadProcessos.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolMatriculasVazia = False Then _
        sfCadMatricula.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolAndamentosVazia = False Then _
        sfCadAndamentos.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolProvidenciasVazia = False Then _
        sfCadProvidencias.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolPedidosVazia = False Then _
        sfCadPedidos.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolLitisVazia = False Then _
        sfCadLitisc.Copy after:=arq.Sheets(arq.Sheets.Count)
        
    If bolSemCPFVazia = False Then _
        sfCadSemCPF.Copy after:=arq.Sheets(arq.Sheets.Count)
    
    Application.DisplayAlerts = False
    arq.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    strDesktop = SisifoEmbasaFuncoes.CaminhoDesktop & IIf(Right(SisifoEmbasaFuncoes.CaminhoDesktop, 1) = "\", "", "\")
    arq.SaveAs strDesktop & "Sisifo - Processos - " & Format(Year(Now), "0000") & "." & Format(Month(Now), "00") & "." & Format(Day(Now), "00") & " " & Format(Hour(Time), "00") & "." & Format(Minute(Time), "00") & ".xlsx"
        
    'Confirmar a inser��o no Espaider
    
    If arq.Saved = False Then
        MsgBox "N�o foi poss�vel salvar o arquivo para exporta��o de processos. Ele ser� fechado." & Chr(13) & _
        "Caso o arquivo para exporta��o n�o seja fechado automaticamente, descarte-o e tente exportar " & _
        "novamente, at� obter a confirma��o da exporta��o dos processos.", vbCritical + vbOKOnly, _
        "S�sifo - Erro ao salvar o arquivo"
        arq.Close False
    Else
        If MsgBox("Confira se a planilha de processos foi salva na �rea de trabalho e clique em OK. " & _
            "N�o esque�a de importar no Espaider. Caso n�o consiga fazer o upload no Espaider, " & _
            "tente novamente mais tarde.", vbExclamation + vbOKCancel + vbApplicationModal, _
            "S�sifo - Confirma exporta��o") = vbOK Then
            ' Usu�rio confirmou salvamento. Limpa a planilha de processos e salva.
            If bolProcessosVazia = False Then sfCadProcessos.Rows("5:" & lnUltimaLinhaProcessos).Delete
            If bolMatriculasVazia = False Then sfCadMatricula.Rows("5:" & lnUltimaLinhaMatriculas).Delete
            If bolAndamentosVazia = False Then sfCadAndamentos.Rows("5:" & lnUltimaLinhaAndamentos).Delete
            If bolProvidenciasVazia = False Then sfCadProvidencias.Rows("5:" & lnUltimaLinhaProvidencias).Delete
            If bolPedidosVazia = False Then sfCadPedidos.Rows("5:" & lnUltimaLinhaPedidos).Delete
            If bolLitisVazia = False Then sfCadLitisc.Rows("5:" & lnUltimaLinhaLitis).Delete
            If bolSemCPFVazia = False Then sfCadSemCPF.Rows("5:" & lnUltimaLinhasemCPF).Delete
            
            Application.DisplayAlerts = False
            ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
            Application.DisplayAlerts = True
        End If
    End If
End Sub

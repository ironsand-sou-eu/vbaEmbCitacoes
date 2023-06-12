Attribute VB_Name = "sfAprDadGerenciaPlanilhas"
Sub RelatarProcessosArmazenadosCitacoes(ByVal Controle As IRibbonControl)
''
'' Contar processos e registros existentes nas planilhas da memória
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

' Gerar planilha para Espaider (Só exclui após usuário confimar que fez o upload no Espaider)
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
    
    ' Pergunta se deseja continuar, para o caso de ser apertado sem querer (percebi que é fácil errar, botões
    ' próximos, e essa tarefa só roda uma vez por dia, portanto a confirmação não sobrecarrega o usuário.
    If MsgBox("Deseja gerar a planilha de exportação no formato do Espaider?", vbQuestion + vbYesNo, _
    "Sísifo - Exportar processos?") = vbNo Then Exit Sub
    
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
        MsgBox "As planilhas de processos estão vazias. Não há processos para exportar.", _
         vbInformation + vbOKOnly, "Sísifo - Planilhas vazias"
        Exit Sub
    End If

    ' Se as planilhas não estiverem vazias, exporta-as
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
        
    'Confirmar a inserção no Espaider
    
    If arq.Saved = False Then
        MsgBox "Não foi possível salvar o arquivo para exportação de processos. Ele será fechado." & Chr(13) & _
        "Caso o arquivo para exportação não seja fechado automaticamente, descarte-o e tente exportar " & _
        "novamente, até obter a confirmação da exportação dos processos.", vbCritical + vbOKOnly, _
        "Sísifo - Erro ao salvar o arquivo"
        arq.Close False
    Else
        If MsgBox("Confira se a planilha de processos foi salva na área de trabalho e clique em OK. " & _
            "Não esqueça de importar no Espaider. Caso não consiga fazer o upload no Espaider, " & _
            "tente novamente mais tarde.", vbExclamation + vbOKCancel + vbApplicationModal, _
            "Sísifo - Confirma exportação") = vbOK Then
            ' Usuário confirmou salvamento. Limpa a planilha de processos e salva.
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

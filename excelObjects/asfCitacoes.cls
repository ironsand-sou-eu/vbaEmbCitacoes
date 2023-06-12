VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "asfCitacoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Success = False Then MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro ao salvar a planilha. Recomendamos que realize novamente a operação " & _
        "que acaba de realizar, até que não receba esta mensagem.", vbInformation + vbOKOnly, "Sísifo - Erro ao salvar"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim oRef As Variant
    
    ' Confere se a Addin de funções de sistema está referenciada (se o salvamento tiver sido feito na hora de fechar, ela não está)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.Name = "SisifoEmbasaFuncoes" Then GoTo AddinEncontrada
    Next oRef
    
AddinEncontrada:
    ' Se estiver, remove a referencia à Addin das funções de sistema deste projeto
    If Not IsEmpty(oRef) Then If oRef.Name = "SisifoEmbasaFuncoes" Then SisifoEmbasaFuncoes.FechaConfiguracoesAoSalvar ThisWorkbook, cfConfigurações
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim arrPlans(1 To 7) As Excel.Worksheet
    Dim oRef As Variant
    
    Set arrPlans(1) = sfCadProcessos
    Set arrPlans(2) = sfCadMatricula
    Set arrPlans(3) = sfCadAndamentos
    Set arrPlans(4) = sfCadProvidencias
    Set arrPlans(5) = sfCadPedidos
    Set arrPlans(6) = sfCadSemCPF
    Set arrPlans(7) = sfCadLitisc
    Cancel = SisifoEmbasaFuncoes.ConferePlanilhasPerguntaSeFecha(arrPlans, cfConfigurações)
    If Cancel = True Then Exit Sub
    
    ' Confere se a Addin de funções de sistema está referenciada (deveria estar, pois foi referenciada ao abrir a planilha)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.Name = "SisifoEmbasaFuncoes" Then GoTo AddinEncontrada
    Next oRef
    
AddinEncontrada:
    ' Se estiver, remove a referencia à Addin das funções de sistema deste projeto
    If oRef.Name = "SisifoEmbasaFuncoes" Then ThisWorkbook.VBProject.References.Remove oRef
    
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfConfigurações
    
End Sub

Private Sub Workbook_Open()
    Dim strCaminhoFS As String, strArquivoFS As String, arrCont() As String
    Dim bolAddinEstahReferenciada As Boolean
    Dim oRef As Variant
    
    ' Busca o caminho do arquivo da Addin de funções de sistema
    
    strCaminhoFS = ThisWorkbook.Path & IIf(Right(ThisWorkbook.Path, 1) = "\", "", "\")
    strArquivoFS = "sfFuncoesComuns*.xlam"
    
Reinicio:
    strArquivoFS = Dir(strCaminhoFS & strArquivoFS, vbArchive)
    If strArquivoFS = "" Then ' Se não existe o arquivo, pergunta
        MsgBox DeterminarTratamento & ", não consigo encontrar o suplemento (AddIn) de funções de sistema. Suplico que me mostre onde ele está localizado " & _
        "neste computador!", vbCritical + vbOKOnly, "Sísifo - Suplemento não encontrada"
        strArquivoFS = PerguntarArquivo("Sísifo - Informar arquivo das funções de sistema do Sísifo", strCaminhoFS, False)
        arrCont = Split(strArquivoFS, "\")
        strArquivoFS = arrCont(UBound(arrCont))
        ReDim Preserve arrCont(UBound(arrCont) - 1)
        strCaminhoFS = Join(arrCont, "\") & "\"
        GoTo Reinicio
    End If
    
    ' Confere se já está referenciada (não é para estar para evitar a popup ao abrir o Excel, vide _____)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.FullPath = strCaminhoFS & strArquivoFS Then bolAddinEstahReferenciada = True
    Next oRef
    
    ' Se não estiver, referencia a Addin das funções de sistema neste projeto
    If bolAddinEstahReferenciada = False Then ThisWorkbook.VBProject.References.AddFromFile strCaminhoFS & strArquivoFS
    
End Sub

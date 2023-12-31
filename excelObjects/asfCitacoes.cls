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
    If Success = False Then MsgBox SisifoEmbasaFuncoes.DeterminarTratamento & ", houve algum erro ao salvar a planilha. Recomendamos que realize novamente a opera��o " & _
        "que acaba de realizar, at� que n�o receba esta mensagem.", vbInformation + vbOKOnly, "S�sifo - Erro ao salvar"
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim oRef As Variant
    
    ' Confere se a Addin de fun��es de sistema est� referenciada (se o salvamento tiver sido feito na hora de fechar, ela n�o est�)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.Name = "SisifoEmbasaFuncoes" Then GoTo AddinEncontrada
    Next oRef
    
AddinEncontrada:
    ' Se estiver, remove a referencia � Addin das fun��es de sistema deste projeto
    If Not IsEmpty(oRef) Then If oRef.Name = "SisifoEmbasaFuncoes" Then SisifoEmbasaFuncoes.FechaConfiguracoesAoSalvar ThisWorkbook, cfConfigura��es
    
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
    Cancel = SisifoEmbasaFuncoes.ConferePlanilhasPerguntaSeFecha(arrPlans, cfConfigura��es)
    If Cancel = True Then Exit Sub
    
    ' Confere se a Addin de fun��es de sistema est� referenciada (deveria estar, pois foi referenciada ao abrir a planilha)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.Name = "SisifoEmbasaFuncoes" Then GoTo AddinEncontrada
    Next oRef
    
AddinEncontrada:
    ' Se estiver, remove a referencia � Addin das fun��es de sistema deste projeto
    If oRef.Name = "SisifoEmbasaFuncoes" Then ThisWorkbook.VBProject.References.Remove oRef
    
    SisifoEmbasaFuncoes.RestringirEdicaoRibbon ThisWorkbook, cfConfigura��es
    
End Sub

Private Sub Workbook_Open()
    Dim strCaminhoFS As String, strArquivoFS As String, arrCont() As String
    Dim bolAddinEstahReferenciada As Boolean
    Dim oRef As Variant
    
    ' Busca o caminho do arquivo da Addin de fun��es de sistema
    
    strCaminhoFS = ThisWorkbook.Path & IIf(Right(ThisWorkbook.Path, 1) = "\", "", "\")
    strArquivoFS = "sfFuncoesComuns*.xlam"
    
Reinicio:
    strArquivoFS = Dir(strCaminhoFS & strArquivoFS, vbArchive)
    If strArquivoFS = "" Then ' Se n�o existe o arquivo, pergunta
        MsgBox DeterminarTratamento & ", n�o consigo encontrar o suplemento (AddIn) de fun��es de sistema. Suplico que me mostre onde ele est� localizado " & _
        "neste computador!", vbCritical + vbOKOnly, "S�sifo - Suplemento n�o encontrada"
        strArquivoFS = PerguntarArquivo("S�sifo - Informar arquivo das fun��es de sistema do S�sifo", strCaminhoFS, False)
        arrCont = Split(strArquivoFS, "\")
        strArquivoFS = arrCont(UBound(arrCont))
        ReDim Preserve arrCont(UBound(arrCont) - 1)
        strCaminhoFS = Join(arrCont, "\") & "\"
        GoTo Reinicio
    End If
    
    ' Confere se j� est� referenciada (n�o � para estar para evitar a popup ao abrir o Excel, vide _____)
    For Each oRef In ThisWorkbook.VBProject.References
        If oRef.FullPath = strCaminhoFS & strArquivoFS Then bolAddinEstahReferenciada = True
    Next oRef
    
    ' Se n�o estiver, referencia a Addin das fun��es de sistema neste projeto
    If bolAddinEstahReferenciada = False Then ThisWorkbook.VBProject.References.AddFromFile strCaminhoFS & strArquivoFS
    
End Sub

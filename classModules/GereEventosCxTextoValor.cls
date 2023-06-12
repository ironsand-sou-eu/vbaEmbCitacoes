VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventosCxTextoValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moCxTextoValor As MSForms.TextBox
Attribute moCxTextoValor.VB_VarHelpID = -1

Public Property Set CaixaDeTextoValor(ByVal oCaixaTexto As MSForms.TextBox)
    Set moCxTextoValor = oCaixaTexto
End Property

Private Sub moCxTextoValor_Change()
    Dim strCont As String
    
    strCont = Replace(moCxTextoValor.text, ",", "")
    strCont = Replace(strCont, ".", "")
    strCont = strCont / 100
    moCxTextoValor.text = Format(strCont, "#,##0.00")
End Sub

Private Sub moCxTextoValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ValidaNumeros(KeyAscii, ",") = False Then KeyAscii = 0
End Sub

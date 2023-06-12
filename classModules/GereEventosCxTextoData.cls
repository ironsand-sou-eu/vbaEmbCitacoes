VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventosCxTextoData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moCxTextoData As MSForms.TextBox
Attribute moCxTextoData.VB_VarHelpID = -1

Public Property Set CaixaDeTextoData(ByVal oCaixaTexto As MSForms.TextBox)
    Set moCxTextoData = oCaixaTexto
End Property

Private Sub moCxTextoData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If ValidaNumeros(KeyAscii, "/", ":", " ") = False Then KeyAscii = 0
End Sub

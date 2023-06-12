VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GereEventosBotaoAdPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents moBotao As MSForms.CommandButton
Attribute moBotao.VB_VarHelpID = -1

Public Property Set Botao(ByVal oBotao As MSForms.CommandButton)
    Set moBotao = oBotao
End Property

Private Sub moBotao_Click()
    Dim form As frmProcesso
    Set form = moBotao.Parent.Parent
    form.RemoverLinhaPedidoFormulario moBotao
End Sub

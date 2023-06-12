VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:

Private intBaseadoEm As Integer
Private dtDataRealizacao As Date
Private strNomePedido As String, strPrognostico As String, strCodigoPedido As String
Private curValorPedido As Currency, curValorRisco As Currency

Private Sub Class_Initialize()
    ' Carrega os valores padrão
    intBaseadoEm = cfConfigurações.Cells().Find(What:="Risco Baseado em", LookAt:=xlWhole).Offset(0, 1).Formula
    
End Sub

'Propriedade Nome:
Public Property Get Nome() As String
Nome = strNomePedido
End Property
Public Property Let Nome(str As String)
strNomePedido = str
End Property

'Propriedade Prognostico:
Public Property Get Prognostico() As String
Prognostico = strPrognostico
End Property
Public Property Let Prognostico(str As String)
strPrognostico = str
End Property

'Propriedade CodigoPedido:
Public Property Get codigoPedido() As String
codigoPedido = strCodigoPedido
End Property
Public Property Let codigoPedido(str As String)
strCodigoPedido = str
End Property

'Propriedade ValorPedido:
Public Property Get ValorPedido() As Currency
ValorPedido = curValorPedido
End Property
Public Property Let ValorPedido(valor As Currency)
curValorPedido = valor
End Property

'Propriedade ValorRisco:
Public Property Get ValorRisco() As Currency
ValorRisco = curValorRisco
End Property
Public Property Let ValorRisco(valor As Currency)
curValorRisco = valor
End Property

'Propriedade BaseadoEm:
Public Property Get BaseadoEm() As Integer
BaseadoEm = intBaseadoEm
End Property

'Propriedade DataRealizacao:
Public Property Get DataRealizacao() As Date
DataRealizacao = dtDataRealizacao
End Property
Public Property Let DataRealizacao(dt As Date)
dtDataRealizacao = dt
End Property



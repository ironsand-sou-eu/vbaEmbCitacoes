VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OutroParticipante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Devem ser registrados aqui todas as outras partes além da Embasa (mesmo que seja apenas um Autor).
'Esta Classe possui várias propriedades, representadas pelas seguintes variáveis:

Private strNomeParte As String, strCondicaoParte As String, strCPFParte As String, strCNPJParte As String
Private intTipoParte As Integer, intClasseParte As Integer


'Propriedade NomeParte:
Public Property Get NomeParte() As String
NomeParte = strNomeParte
End Property
Public Property Let NomeParte(str As String)
strNomeParte = str
End Property

'Propriedade CondicaoParte:
Public Property Get CondicaoParte() As String
CondicaoParte = strCondicaoParte
End Property
Public Property Let CondicaoParte(str As String)
strCondicaoParte = str
End Property

'Propriedade CPFParte:
Public Property Get CPFParte() As String
CPFParte = strCPFParte
End Property
Public Property Let CPFParte(str As String)
strCPFParte = str
End Property

'Propriedade CNPJParte:
Public Property Get CNPJParte() As String
CNPJParte = strCNPJParte
End Property
Public Property Let CNPJParte(str As String)
strCNPJParte = str
End Property

'Propriedade TipoParte:
Public Property Get TipoParte() As Integer
TipoParte = intTipoParte
End Property
Public Property Let TipoParte(num As Integer)
intTipoParte = num
End Property

'Propriedade ClasseParte:
Public Property Get ClasseParte() As Integer
ClasseParte = intClasseParte
End Property
Public Property Let ClasseParte(num As Integer)
intClasseParte = num
End Property

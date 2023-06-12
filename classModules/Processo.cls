VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Processo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Esta Classe possiu muitas propriedades, representadas pelas seguintes variáveis:
Private dtCitacao As Date, dtDataAndamento As Date, dtDataAndamentoAdicional As Date
Private intClasseProcesso As Integer, intTipoEmpresa As Integer, intPeticionamento As Integer
Private strNumProc As String, strUnidade As String, strDivisao As String, strFase As String, strNatureza As String
Private strTipoAcao As String, strCausaPedir As String, strNucleo As String, strAdvogadoNucleo As String, strPrepostoNucleo As String
Private strResponsavelRegressivo As String, strCNPJEmbasa As String, strNomeEmpresa As String, strCondicaoEmbasa As String
Private strAdvogadoAdverso As String, strNomeDesdobramento As String, strOrgao As String, strJuizo As String, strComarca As String
Private strRito As String, strNomeAndamento As String, strNomeAndamentoAdicional As String, strObsAndamentoAdicional As String
Private strgerencia As String, strSistema As String, strTribunal As String, strMensagemErro As String
Private strSequencialAndamento As String, strSequencialAndamentoAdicional As String
Private lnMatricula As Long
Private curValorCausa As Currency
Private bolBloqueiaMatricula As Boolean
Private btCont As Byte
Private listaPedidos As Collection
Private listaProvidencias As Collection
Private listaOutrosParticipantes As Collection

Private Sub Class_Initialize()
    ' Carrega os valores padrão
    intClasseProcesso = cfConfigurações.Cells().Find(What:="Classe do Processo", LookAt:=xlWhole).Offset(0, 1).Formula
    intPeticionamento = cfConfigurações.Cells().Find(What:="Peticionamento", LookAt:=xlWhole).Offset(0, 1).Formula
    strFase = cfConfigurações.Cells().Find(What:="Fase", LookAt:=xlWhole).Offset(0, 1).Formula
    strCNPJEmbasa = cfConfigurações.Cells().Find(What:="CNPJ Embasa", LookAt:=xlWhole).Offset(0, 1).Formula
    strNomeEmpresa = cfConfigurações.Cells().Find(What:="Nome da Embasa", LookAt:=xlWhole).Offset(0, 1).Formula
    intTipoEmpresa = cfConfigurações.Cells().Find(What:="Tipo da Empresa", LookAt:=xlWhole).Offset(0, 1).Formula
    strCondicaoEmbasa = cfConfigurações.Cells().Find(What:="Condição da Embasa", LookAt:=xlWhole).Offset(0, 1).Formula
    strNomeDesdobramento = cfConfigurações.Cells().Find(What:="Nome do Desdobramento", LookAt:=xlWhole).Offset(0, 1).Formula
    
    'Seta as coleções
    Set listaPedidos = New Collection
    Set listaProvidencias = New Collection
    Set listaOutrosParticipantes = New Collection
    
End Sub

'Propriedade Pedidos:
Public Property Get Pedidos() As Collection
Set Pedidos = listaPedidos
End Property

'Propriedade Providencias:
Public Property Get Providencias() As Collection
Set Providencias = listaProvidencias
End Property

'Propriedade OutrosParticipantes:
Public Property Get OutrosParticipantes() As Collection
Set OutrosParticipantes = listaOutrosParticipantes
End Property

'Propriedade NumeroProcesso:
Public Property Get NumeroProcesso() As String
NumeroProcesso = strNumProc
End Property
Public Property Let NumeroProcesso(str As String)
strNumProc = str
End Property

'Propriedade Unidade:
Public Property Get Unidade() As String
Unidade = strUnidade
End Property
Public Property Let Unidade(str As String)
strUnidade = str
End Property

'Propriedade Divisao:
Public Property Get Divisao() As String
Divisao = strDivisao
End Property
Public Property Let Divisao(str As String)
strDivisao = str
End Property

'Propriedade Fase:
Public Property Get Fase() As String
Fase = strFase
End Property
Public Property Let Fase(str As String)
strFase = str
End Property

'Propriedade Natureza:
Public Property Get Natureza() As String
Natureza = strNatureza
End Property
Public Property Let Natureza(str As String)
strNatureza = str
End Property

'Propriedade TipoAcao:
Public Property Get TipoAcao() As String
TipoAcao = strTipoAcao
End Property
Public Property Let TipoAcao(str As String)
strTipoAcao = str
End Property

'Propriedade CausaPedir:
Public Property Get CausaPedir() As String
CausaPedir = strCausaPedir
End Property
Public Property Let CausaPedir(str As String)
strCausaPedir = str
End Property

'Propriedade Nucleo:
Public Property Get Nucleo() As String
Nucleo = strNucleo
End Property
Public Property Let Nucleo(str As String)
strNucleo = str
End Property

'Propriedade Preposto:
Public Property Get Preposto() As String
Preposto = strPrepostoNucleo
End Property
Public Property Let Preposto(str As String)
strPrepostoNucleo = str
End Property

'Propriedade Advogado:
Public Property Get Advogado() As String
Advogado = strAdvogadoNucleo
End Property
Public Property Let Advogado(str As String)
strAdvogadoNucleo = str
End Property

'Propriedade ResponsavelRegressivo:
Public Property Get ResponsavelRegressivo() As String
ResponsavelRegressivo = strResponsavelRegressivo
End Property
Public Property Let ResponsavelRegressivo(str As String)
strResponsavelRegressivo = str
End Property

'Propriedade CNPJEmbasa:
Public Property Get CNPJEmbasa() As String
CNPJEmbasa = strCNPJEmbasa
End Property
Public Property Let CNPJEmbasa(str As String)
strCNPJEmbasa = str
End Property

'Propriedade NomeEmpresa:
Public Property Get NomeEmpresa() As String
NomeEmpresa = strNomeEmpresa
End Property
Public Property Let NomeEmpresa(str As String)
strNomeEmpresa = str
End Property

'Propriedade CondicaoEmbasa:
Public Property Get CondicaoEmbasa() As String
CondicaoEmbasa = strCondicaoEmbasa
End Property
Public Property Let CondicaoEmbasa(str As String)
strCondicaoEmbasa = str
End Property

'Propriedade AdvogadoAdverso:
Public Property Get AdvogadoAdverso() As String
AdvogadoAdverso = strAdvogadoAdverso
End Property
Public Property Let AdvogadoAdverso(str As String)
strAdvogadoAdverso = str
End Property

'Propriedade NomeDesdobramento:
Public Property Get NomeDesdobramento() As String
NomeDesdobramento = strNomeDesdobramento
End Property
Public Property Let NomeDesdobramento(str As String)
strNomeDesdobramento = str
End Property

'Propriedade Orgao:
Public Property Get Orgao() As String
Orgao = strOrgao
End Property
Public Property Let Orgao(str As String)
strOrgao = str
End Property

'Propriedade Juizo:
Public Property Get Juizo() As String
Juizo = strJuizo
End Property
Public Property Let Juizo(str As String)
strJuizo = str
End Property

'Propriedade Comarca:
Public Property Get Comarca() As String
Comarca = strComarca
End Property
Public Property Let Comarca(str As String)
strComarca = str
End Property

'Propriedade Rito:
Public Property Get Rito() As String
Rito = strRito
End Property
Public Property Let Rito(str As String)
strRito = str
End Property

'Propriedade NomeAndamento:
Public Property Get NomeAndamento() As String
NomeAndamento = strNomeAndamento
End Property
Public Property Let NomeAndamento(str As String)
strNomeAndamento = str
End Property

'Propriedade DataAndamento:
Public Property Get DataAndamento() As Date
DataAndamento = dtDataAndamento
End Property
Public Property Let DataAndamento(dt As Date)
dtDataAndamento = dt
End Property

'Propriedade NomeAndamentoAdicional:
Public Property Get NomeAndamentoAdicional() As String
NomeAndamentoAdicional = strNomeAndamentoAdicional
End Property
Public Property Let NomeAndamentoAdicional(str As String)
strNomeAndamentoAdicional = str
End Property

'Propriedade DataAndamentoAdicional:
Public Property Get DataAndamentoAdicional() As Date
DataAndamentoAdicional = dtDataAndamentoAdicional
End Property
Public Property Let DataAndamentoAdicional(dt As Date)
dtDataAndamentoAdicional = dt
End Property

'Propriedade ObsAndamentoAdicional:
Public Property Get ObsAndamentoAdicional() As String
ObsAndamentoAdicional = strObsAndamentoAdicional
End Property
Public Property Let ObsAndamentoAdicional(str As String)
strObsAndamentoAdicional = str
End Property

'Propriedade Citacao:
Public Property Get Citacao() As Date
Citacao = dtCitacao
End Property
Public Property Let Citacao(dt As Date)
dtCitacao = dt
End Property

'Propriedade ClasseProcesso:
Public Property Get ClasseProcesso() As Integer
ClasseProcesso = intClasseProcesso
End Property
Public Property Let ClasseProcesso(num As Integer)
intClasseProcesso = num
End Property

'Propriedade TipoEmpresa:
Public Property Get TipoEmpresa() As Integer
TipoEmpresa = intTipoEmpresa
End Property
Public Property Let TipoEmpresa(num As Integer)
intTipoEmpresa = num
End Property

'Propriedade Peticionamento:
Public Property Get Peticionamento() As Integer
Peticionamento = intPeticionamento
End Property
Public Property Let Peticionamento(num As Integer)
intPeticionamento = num
End Property

'Propriedade Matricula:
Public Property Get Matricula() As Long
Matricula = lnMatricula
End Property
Public Property Let Matricula(lon As Long)
lnMatricula = Trim(lon)
End Property

'Propriedade ValorCausa:
Public Property Get ValorCausa() As Currency
ValorCausa = curValorCausa
End Property
Public Property Let ValorCausa(valor As Currency)
curValorCausa = valor
End Property

'Propriedade BloqueiaMatricula:
Public Property Get BloqueiaMatricula() As Boolean
BloqueiaMatricula = bolBloqueiaMatricula
End Property
Public Property Let BloqueiaMatricula(bloqmat As Boolean)
bolBloqueiaMatricula = bloqmat
End Property

'Propriedade Sistema
Public Property Get Sistema() As String
Sistema = strSistema
End Property
Public Property Let Sistema(str As String)
strSistema = str
End Property

'Propriedade Tribunal
Public Property Get Tribunal() As String
Tribunal = strTribunal
End Property
Public Property Let Tribunal(str As String)
strTribunal = str
End Property

'Propriedade Gerencia:
Public Property Get gerencia() As String
gerencia = strgerencia
End Property
Public Property Let gerencia(str As String)
strgerencia = str
End Property

'Propriedade MensagemErro:
Public Property Get MensagemErro() As String
MensagemErro = strMensagemErro
End Property
Public Property Let MensagemErro(str As String)
strMensagemErro = str
End Property

Private Sub CriarSequenciaisDosAndamentos()
    Dim sequencial As String
    
    sequencial = PegarSequencialAndamento
    strSequencialAndamento = sequencial
    strSequencialAndamentoAdicional = SomarUmAoStringNumericoSemConverter(sequencial)
End Sub

Private Function PegarSequencialAndamento() As String
    Dim textoMarcador As String
    textoMarcador = Format(dtDataAndamento, "ddmmyyhhmm")
    PegarSequencialAndamento = PegarSequencial(textoMarcador)
End Function

Private Function PegarSequencialProvidencia(providencia As providencia, indice As Integer) As String
    Dim textoMarcador As String
    textoMarcador = Format(providencia.DataFinal, "ddmmyy") & indice
    PegarSequencialProvidencia = PegarSequencial(textoMarcador)
End Function

Private Function PegarSequencialPedido(codigoPedido As String) As String
    Dim textoMarcador As String
    textoMarcador = codigoPedido
    PegarSequencialPedido = PegarSequencial(textoMarcador)
End Function

Private Function PegarSequencial(textoAAdicionar As String) As String
    Dim sequencial As String

    sequencial = Replace(strNumProc, ".8.05.", "")
    sequencial = sequencial & textoAAdicionar
    sequencial = DevolverSomenteNumeros(sequencial)
    sequencial = RemoverZerosEsquerda(sequencial)
    PegarSequencial = sequencial
End Function

Private Function DevolverSomenteNumeros(textoEntrada As String) As String
    Dim sequencial() As String, caractere As String
    Dim i As Integer
    
    ReDim sequencial(Len(textoEntrada) - 1)
    For i = 1 To Len(textoEntrada)
        caractere = Mid(textoEntrada, i, 1)
        sequencial(i - 1) = IIf(IsNumeric(caractere), caractere, "")
    Next
    
    DevolverSomenteNumeros = Join(sequencial, "")
    
End Function

Private Function RemoverZerosEsquerda(textoEntrada As String) As String
    Dim texto As String
    Dim i As Integer
    Dim parar As Boolean
    
    texto = textoEntrada
    Do While Left(texto, 1) = "0"
        texto = Replace(texto, "0", "", 1, 1)
    Loop
    
    RemoverZerosEsquerda = texto
End Function

Private Function SomarUmAoStringNumericoSemConverter(numeroEmString As String) As String
    Dim inicioTexto As String, ultimosDoisCaracteres As String, resposta As String
    
    If Not IsNumeric(numeroEmString) Then
        resposta = ""
    Else
        If Len(numeroEmString) <= 3 Then
            resposta = CStr(numeroEmString + 1)
        Else
            inicioTexto = Left(numeroEmString, Len(numeroEmString) - 2)
            ultimosDoisCaracteres = Right(numeroEmString, 2)
            ultimosDoisCaracteres = CStr(ultimosDoisCaracteres + 1)
            If ultimosDoisCaracteres = "100" Then ultimosDoisCaracteres = "00"
            resposta = inicioTexto & ultimosDoisCaracteres
        End If
    End If
    SomarUmAoStringNumericoSemConverter = resposta
End Function

'Método ExportaLinhasEspaider:
Public Sub ExportaLinhasEspaider(arq As Workbook, plan As Worksheet)
''
'' Exporta as linhas. A planilha passada como parâmetro deve estar limpa!
'' Se houver mais de uma parte adversa, cria também uma outra planilha na mesma pasta de trabalho, preparada
'' para exportar para ProcessoParticipantes.
''
    Dim intCont As Integer, intNaoCadastrados As Integer, NaoCadastrados() As Integer
    Dim lnUltimaLinha As Long
    Dim strNaoCadastrados As String
    Dim planProcessos As Worksheet, planAndamentos As Worksheet
    
    Set planProcessos = arq.Worksheets.Add
    With planProcessos
        .Cells(1, 1) = strNumProc 'Número
        .Cells(1, 2) = strNumProc 'Número
        .Cells(1, 3) = strNumProc 'Número
        .Cells(1, 4) = IIf(listaOutrosParticipantes(1).TipoParte = 1, listaOutrosParticipantes(1).CPFParte, listaOutrosParticipantes(1).CNPJParte) 'CPF/CNPJ Adverso
        .Cells(1, 5) = listaOutrosParticipantes(1).NomeParte 'Nome do Adverso
        .Cells(1, 6) = listaOutrosParticipantes(1).TipoParte 'Tipo Adverso
        .Cells(1, 7) = listaOutrosParticipantes(1).CondicaoParte 'Condição Adverso
            
        ' Se CPF ou CNPJ for "Não cadastrado", anota esta informação.
        If listaOutrosParticipantes(1).CPFParte = "Não cadastrado" Or listaOutrosParticipantes(1).CNPJParte = "Não cadastrado" Then strNaoCadastrados = "1,"
        
        .Cells(1, 8) = strAdvogadoAdverso 'Advogado do Adverso
        .Cells(1, 9) = strNucleo 'Núcleo
        .Cells(1, 10) = IIf(strAdvogadoNucleo = "", "", strNucleo) 'Núcleo-adv
        .Cells(1, 11) = strAdvogadoNucleo 'Advogado do Núcleo
        .Cells(1, 12) = strUnidade 'Unidade/Depto.
        .Cells(1, 13) = IIf(strDivisao = "", "", strUnidade) 'Unidade da Divisão
        .Cells(1, 14) = strDivisao 'Divisao/EL
        .Cells(1, 15) = intClasseProcesso 'Classe
        .Cells(1, 16) = strNatureza 'Natureza
        .Cells(1, 17) = strNatureza 'Natureza Tipo
        .Cells(1, 18) = strTipoAcao 'Tipo de ação
        .Cells(1, 19) = strCausaPedir 'Causa de pedir
        .Cells(1, 20) = curValorCausa 'Valor da causa
        .Cells(1, 21) = strNomeDesdobramento 'Nome do desdobramento
        .Cells(1, 22) = strOrgao 'Órgão
        .Cells(1, 23) = strOrgao 'Órgão-Juízo
        .Cells(1, 24) = strJuizo 'Juízo
        .Cells(1, 25) = strComarca 'Comarca
        .Cells(1, 26) = strCNPJEmbasa 'CNPJ Empresa
        .Cells(1, 27) = strNomeEmpresa 'Empresa
        .Cells(1, 28) = intTipoEmpresa 'Tipo Empresa
        .Cells(1, 29) = strCondicaoEmbasa 'Condição
        .Cells(1, 30) = strFase 'Fase atual
        .Cells(1, 31) = strFase 'Fase atual do desdobramento
        .Cells(1, 32) = strResponsavelRegressivo 'Responsável Regressivo
        .Cells(1, 33) = strRito 'Rito
        .Cells(1, 34) = intPeticionamento 'Peticionamento
        .Cells(1, 35) = dtCitacao 'Citação
    End With
    
    If lnMatricula <> 0 Then
        Dim planMatriculas As Worksheet
        Set planMatriculas = arq.Worksheets.Add
        
        With planMatriculas
            .Cells(1, 1) = strNumProc 'Número do processo
            .Cells(1, 2) = lnMatricula 'Matrícula
            .Cells(1, 3) = IIf(bolBloqueiaMatricula, "CLI_N", "CLI_S") 'Negativação
            .Cells(1, 4) = IIf(bolBloqueiaMatricula, 1, 2) 'Cobrança
            .Cells(1, 5) = "Sim" 'Pegar Histórico
        End With
    End If
    
    Set planAndamentos = arq.Worksheets.Add
    CriarSequenciaisDosAndamentos
    With planAndamentos
        If strNomeAndamentoAdicional <> "" Then
            .Cells(1, 1) = "'" & strSequencialAndamentoAdicional 'Sequencial andamento
            .Cells(1, 2) = strNumProc 'Número do Processo
            .Cells(1, 3) = strNumProc 'Número do desdobramento
            .Cells(1, 4) = strNomeAndamentoAdicional 'Nome do andamento
            .Cells(1, 5) = dtDataAndamentoAdicional 'Data
            .Cells(1, 6) = strObsAndamentoAdicional 'Obs Andamento
            btCont = 1
        End If
        
        If dtDataAndamento <> 0 Then
            .Cells(1 + btCont, 1) = "'" & strSequencialAndamento 'Sequencial andamento
            .Cells(1 + btCont, 2) = strNumProc 'Número do Processo
            .Cells(1 + btCont, 3) = strNumProc 'Número do desdobramento
            .Cells(1 + btCont, 4) = strNomeAndamento 'Nome do andamento
            .Cells(1 + btCont, 5) = dtDataAndamento 'Data
            .Cells(1 + btCont, 6) = strObsAndamento 'Obs Andamento
        End If
    End With
    
    If listaProvidencias.Count > 0 Then
        Dim planProvidencias As Worksheet
        Set planProvidencias = arq.Worksheets.Add
        
        intCont = 1
            
        With planProvidencias
            For intCont = 1 To listaProvidencias.Count
                .Cells(intCont, 1) = "'" & PegarSequencialProvidencia(listaProvidencias(intCont), intCont) 'Sequencial da providência
                .Cells(intCont, 2) = strNumProc 'Número do Processo
                .Cells(intCont, 3) = strNumProc 'Número do desdobramento
                .Cells(intCont, 4) = "'" & strSequencialAndamento 'Sequencial do andamento
                .Cells(intCont, 5) = listaProvidencias(intCont).Nome 'Providência
                .Cells(intCont, 6) = listaProvidencias(intCont).DataFinal 'Data Final
                .Cells(intCont, 7) = listaProvidencias(intCont).PrazoDias 'Prazo
                .Cells(intCont, 8) = listaProvidencias(intCont).Situacao 'Situação
                .Cells(intCont, 9) = IIf(listaProvidencias(intCont).Alertar, "Sim", "Não") 'Enviar alertas
                .Cells(intCont, 10) = listaProvidencias(intCont).DiasAntecedenciaAlerta 'Dias de Antecedencia
                .Cells(intCont, 11) = listaProvidencias(intCont).PeriodicidadeAlerta 'Periodicidade
                .Cells(intCont, 12) = listaProvidencias(intCont).DataAlerta 'Data do alerta
                .Cells(intCont, 13) = listaProvidencias(intCont).Nucleo 'Núcleo
                .Cells(intCont, 14) = listaProvidencias(intCont).Responsavel 'Responsável
                .Cells(intCont, 15) = listaProvidencias(intCont).Observacao 'Observacoes
                .Cells(intCont, 16) = IIf(listaProvidencias(intCont).GerarAndamento, "Sim", "Não") 'Gerar Evento
                .Cells(intCont, 17) = IIf(listaProvidencias(intCont).GerarAndamento, listaProvidencias(intCont).DataAndamento, "") 'Data Base
                .Cells(intCont, 18) = IIf(listaProvidencias(intCont).GerarAndamento, listaProvidencias(intCont).Andamento, "") 'Evento
            Next intCont
        End With
    End If
    
    If listaPedidos.Count > 0 Then
        Dim planPedidos As Worksheet
        Set planPedidos = arq.Worksheets.Add
        
        intCont = 1
        
        With planPedidos
            For intCont = 1 To listaPedidos.Count
                .Cells(intCont, 1) = "'" & PegarSequencialPedido(listaPedidos(intCont).codigoPedido) 'sequencial Pedidos
                .Cells(intCont, 2) = strNumProc 'Número do Processo
                .Cells(intCont, 3) = listaPedidos(intCont).Nome 'Pedido
                .Cells(intCont, 4) = listaPedidos(intCont).ValorPedido 'valor Pedido
                .Cells(intCont, 5) = dtCitacao 'Data Base Atualização
                .Cells(intCont, 6) = listaPedidos(intCont).Prognostico 'risco original
                .Cells(intCont, 7) = listaPedidos(intCont).BaseadoEm 'Risco Baseado em
                .Cells(intCont, 8) = listaPedidos(intCont).ValorRisco 'Valor risco original
                .Cells(intCont, 9) = listaPedidos(intCont).DataRealizacao 'Data estim. pagto.
                .Cells(intCont, 10) = "INPC" 'Índice
                .Cells(intCont, 11) = dtCitacao 'Database Juros
                .Cells(intCont, 12) = 1 'Juros
            Next intCont
        End With
    End If
    
    ' Testa se há outras partes em listaOutrosParticipantes; caso haja, faz a outra planilha.
    If listaOutrosParticipantes.Count >= 2 Then
        Dim planLitis As Worksheet
        Set planLitis = arq.Worksheets.Add
        
        With planLitis
            
            intCont = 2
            For intCont = 2 To listaOutrosParticipantes.Count

                ' Se CPF ou CNPJ for "Não cadastrado", deixa para o próximo passo.
                If listaOutrosParticipantes(intCont).CPFParte = "Não cadastrado" Or listaOutrosParticipantes(intCont).CPFParte = "Não cadastrado" Then
                    strNaoCadastrados = strNaoCadastrados & intCont & ","
                    intNaoCadastrados = intNaoCadastrados + 1
                Else
                    .Cells(intCont - intNaoCadastrados - 1, 1) = strNumProc
                    .Cells(intCont - intNaoCadastrados - 1, 2) = strNumProc
                    .Cells(intCont - intNaoCadastrados - 1, 3) = CStr(listaOutrosParticipantes(intCont).CPFParte)
                    .Cells(intCont - intNaoCadastrados - 1, 4) = CStr(listaOutrosParticipantes(intCont).CNPJParte)
                    .Cells(intCont - intNaoCadastrados - 1, 5) = listaOutrosParticipantes(intCont).NomeParte
                    .Cells(intCont - intNaoCadastrados - 1, 6) = listaOutrosParticipantes(intCont).NomeParte
                    .Cells(intCont - intNaoCadastrados - 1, 7) = listaOutrosParticipantes(intCont).TipoParte
                    .Cells(intCont - intNaoCadastrados - 1, 8) = listaOutrosParticipantes(intCont).CondicaoParte
                    .Cells(intCont - intNaoCadastrados - 1, 9) = listaOutrosParticipantes(intCont).ClasseParte
                End If
            Next intCont
            
        End With
    End If
    
    ' Testa se há partes com CPF ou CNPJ não cadastrado; caso haja, faz a outra planilha.
    If strNaoCadastrados <> "" Then
        Dim planNaoCad As Worksheet
        Set planNaoCad = arq.Worksheets.Add
        
        With planNaoCad
            strNaoCadastrados = Left(strNaoCadastrados, Len(strNaoCadastrados) - 1)
            
            If InStr(strNaoCadastrados, ",") <> 0 Then
                NaoCadastrados = Split(strNaoCadastrados, ",")
            ElseIf strNaoCadastrados <> "" Then
                ReDim NaoCadastrados(1 To 1)
                NaoCadastrados(1) = strNaoCadastrados
            End If
            
            intCont = 1
            For intCont = 1 To UBound(NaoCadastrados)
                .Cells(intCont, 1) = strNumProc
                .Cells(intCont, 2) = strNumProc
                .Cells(intCont, 3) = ""
                .Cells(intCont, 4) = ""
                .Cells(intCont, 5) = listaOutrosParticipantes(NaoCadastrados(intCont)).NomeParte
                .Cells(intCont, 6) = listaOutrosParticipantes(NaoCadastrados(intCont)).NomeParte
                .Cells(intCont, 7) = listaOutrosParticipantes(NaoCadastrados(intCont)).TipoParte
                .Cells(intCont, 8) = listaOutrosParticipantes(NaoCadastrados(intCont)).CondicaoParte
                .Cells(intCont, 9) = listaOutrosParticipantes(NaoCadastrados(intCont)).ClasseParte
                .Cells(intCont, 10) = IIf(NaoCadastrados(intCont) = 1, "Sim", "Não")
            Next intCont

            
        End With
    End If
    
    plan.UsedRange.Columns.AutoFit
    plan.Activate
    plan.Range("U1:Y1").Select
    
    ' Confirmados os dados, confere se o processo já está na planilha
    If Not sfCadProcessos.Cells().Find(strNumProc) Is Nothing Then
        arq.Close False
        Do
        Loop Until MsgBox("O processo já existe na planilha! Inclusão cancelada. Descartados os dados." & Chr(13) & _
        "Processo: " & strNumProc & Chr(13) & _
        "Adverso: " & listaOutrosParticipantes(1).NomeParte & Chr(13) & _
        "Clique em 'Cancelar' e insira o próximo processo.", vbCritical + vbOKCancel, "Sísifo - Processo repetido") = vbCancel
        Exit Sub
    End If
    
    ' Não estando na planilha, pimba na gorduchinha!
    AdicionarLinhasAoFinal planProcessos, sfCadProcessos
    If Not planMatriculas Is Nothing Then AdicionarLinhasAoFinal planMatriculas, sfCadMatricula
    If Not planAndamentos Is Nothing Then AdicionarLinhasAoFinal planAndamentos, sfCadAndamentos
    If Not planProvidencias Is Nothing Then AdicionarLinhasAoFinal planProvidencias, sfCadProvidencias
    If Not planPedidos Is Nothing Then AdicionarLinhasAoFinal planPedidos, sfCadPedidos
    If Not planLitis Is Nothing Then AdicionarLinhasAoFinal planLitis, sfCadLitisc
    If Not planNaoCad Is Nothing Then AdicionarLinhasAoFinal planNaoCad, sfCadSemCPF
    
    'Após salvar, confirma a inclusão e fecha a pasta de trabalho de rascunho.
    Application.DisplayAlerts = False
    'ThisWorkbook.SaveAs Filename:=ThisWorkbook.FullName, FileFormat:=xlOpenXMLAddIn
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    MsgBox "Processo registrado com sucesso!" & Chr(13) & _
    "Processo: " & strNumProc & Chr(13) & _
    "Adverso: " & listaOutrosParticipantes(1).NomeParte & Chr(13) & _
    "Lembre-se de inserir no Espaider no final do dia.", vbInformation + vbOKOnly, "Sísifo - Processo salvo para inclusão"
    arq.Close False

End Sub

Private Sub AdicionarLinhasAoFinal(origem As Worksheet, destino As Worksheet)
    Dim ultimaLinha As Long
    
    origem.UsedRange.Copy
    ultimaLinha = destino.UsedRange.Rows(destino.UsedRange.Rows.Count).Row
    destino.Cells(ultimaLinha + 1, 1).PasteSpecial xlPasteAll
End Sub

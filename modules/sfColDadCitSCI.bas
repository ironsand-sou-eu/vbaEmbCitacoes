Attribute VB_Name = "sfColDadCitSCI"
Sub ConsultaMatriculaSCI(ByVal strMatricula As String, ByRef strConsulta() As String)
''
'' Executa o Web Service de consultar matrícula, utilizando a matrícula fornecida como parâmetro. A variável strConsulta deve ser uma matriz com duas
'' dimensões: a primeira dimensão contém os parâmetros de resposta do Web Service que se deseja, e a segunda dimensão deve estar em branco.
'' A função preencherá a segunda dimensão com o valor da resposta do Web Service correspondente ao nome de parâmetro contido na primeira dimensão.
'' EM CASO DE ERRO, a função retornará apenas as dimensões que tiveram resposta (provavelmente, o "erroCodigo" e o "erroMensagem").
''
'' Exemplo de Consulta que busca 5 parâmetros (erroCodigo, erroMensagem, inscricao, CPF e nomeCliente):
'' A variável passada deve ser uma array (4,2), com os valores seguintes:
''      strConsulta(1, 1) = "erroCodigo";   (1, 2) = em branco
''      strConsulta(2, 1) = "erroMensagem";   (2, 2) = em branco
''      strConsulta(3, 1) = "inscricao";    (3, 2) = em branco
''      strConsulta(4, 1) = "numCPFCNPJ";  (4, 2) = em branco
''      strConsulta(5, 1) = "nomeCliente";  (5, 2) = em branco
''

    Dim reqSCI As MSXML2.XMLHTTP60
    Dim strCorpoXML As String, strRetornoReq As String
    Dim btCont As Byte
    Dim lngInicio As Long, lngFinal As Long
    'Dim XMLResposta As MSXML2.DOMDocument60
    ' Originalmente, eu buscava a resposta como XML, mas algumas matrículas (coisa de 1 ou 2%) tinham resposta da requisição com XML vazio,
    '   mas texto normal. Por isso, passei a utilizar a resposta de texto. As linhas de quando eu usava XML continuam abaixo, comentarizadas,
    '   para o caso de ser necessário reverter.
    
    strCorpoXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:jur=""http://juridico.webservices.web.sci.embasa.ba.gov.br/"">" & vbCrLf & _
                   "<soapenv:Header/>" & vbCrLf & _
                   "<soapenv:Body>" & vbCrLf & _
                      "<jur:consultarMatricula>" & vbCrLf & _
                         "<matricula>" & strMatricula & "</matricula>" & vbCrLf & _
                      "</jur:consultarMatricula>" & vbCrLf & _
                   "</soapenv:Body>" & vbCrLf & _
                "</soapenv:Envelope>"

    Set reqSCI = New MSXML2.XMLHTTP60
    reqSCI.Open "POST", sfWsSciUrl, bstrUser:=sfWsSciUsuario, bstrPassword:=sfWsSciSenha
    reqSCI.send strCorpoXML
    strRetornoReq = reqSCI.responseText
    
    For btCont = 1 To UBound(strConsulta, 1) Step 1
        On Error Resume Next
        lngInicio = InStr(1, strRetornoReq, "<" & strConsulta(btCont, 1) & ">") + Len("<" & strConsulta(btCont, 1) & ">")
        lngFinal = InStr(1, strRetornoReq, "</" & strConsulta(btCont, 1) & ">")
        strConsulta(btCont, 2) = Trim(Mid(strRetornoReq, lngInicio, lngFinal - lngInicio))
        'strConsulta(btCont, 2) = XMLResposta.getElementsByTagName(strConsulta(btCont, 1))(0).Text
        On Error GoTo 0
    Next btCont

End Sub

Function ConsultaProcessosPorMatriculaSCI(ByVal strMatricula As String)
''
'' Executa o Web Service de consultar processos cadastrados na matrícula, utilizando a matrícula fornecida como parâmetro de busca no web service.
'' Retorna uma matriz com os processos cadastrados, na qual:
'' A primeira dimensão é um cardinal referente à quantidade de processos retornada
'' A segunda dimensão é o retorno da busca, sendo que 1 = numeroProcesso, 2 = nomeAutor, 3 = objetoAcao
''

    Dim reqSCI As MSXML2.XMLHTTP60
    Dim strCorpoXML As String, strResposta() As String, strRetornoReq As String, strCont As String
    Dim intCont As Integer, intQtdProcessos As Integer
    Dim lngInicio As Long, lngFinal As Long, lngSubInicio As Long, lngSubFinal As Long
    'Dim XMLResposta As MSXML2.DOMDocument60
    ' Originalmente, eu buscava a resposta como XML, mas algumas matrículas (coisa de 1 ou 2%) tinham resposta da requisição com XML vazio,
    '   mas texto normal. Por isso, passei a utilizar a resposta de texto. As linhas de quando eu usava XML continuam abaixo, comentarizadas,
    '   para o caso de ser necessário reverter.
    
    strCorpoXML = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:jur=""http://juridico.webservices.web.sci.embasa.ba.gov.br/"">" & vbCrLf & _
                   "<soapenv:Header/>" & vbCrLf & _
                   "<soapenv:Body>" & vbCrLf & _
                      "<jur:consultarProcessosPorMatricula>" & vbCrLf & _
                         "<matricula>" & strMatricula & "</matricula>" & vbCrLf & _
                      "</jur:consultarProcessosPorMatricula>" & vbCrLf & _
                   "</soapenv:Body>" & vbCrLf & _
                "</soapenv:Envelope>"

    Set reqSCI = New MSXML2.XMLHTTP60
    reqSCI.Open "POST", sfWsSciUrl, bstrUser:=sfWsSciUsuario, bstrPassword:=sfWsSciSenha
    reqSCI.send strCorpoXML
    
    'Set XMLResposta = New MSXML2.DOMDocument60
    'XMLResposta.Load reqSCI.responseXML
    strRetornoReq = reqSCI.responseText
    
    ' Pega o número de tags de processo
    'intCont = XMLResposta.getElementsByTagName("processosDaMatriculaNoSCI").length
    intCont = (Len(strRetornoReq) - Len(Replace(strRetornoReq, "<processosDaMatriculaNoSCI>", ""))) / Len("<processosDaMatriculaNoSCI>")
    intQtdProcessos = intCont
    
    If intCont > 0 Then
        ReDim strResposta(1 To intCont, 1 To 3)
        lngInicio = 1
        lngFinal = 1
        
        For intCont = 1 To intQtdProcessos Step 1
            ' Para cada processo, separa na substring strCont as tags daquele processo...
            lngInicio = InStr(lngFinal, strRetornoReq, "<processosDaMatriculaNoSCI>")
            lngFinal = InStr(lngFinal, strRetornoReq, "</processosDaMatriculaNoSCI>") + Len("<processosDaMatriculaNoSCI>")
            strCont = Trim(Mid(strRetornoReq, lngInicio, lngFinal - lngInicio + 1))
            
            ' E pega os dados respectivos.
            lngSubInicio = InStr(1, strCont, "<numeroProcesso>") + Len("<numeroProcesso>")
            lngSubFinal = InStr(1, strCont, "</numeroProcesso>")
            strResposta(intCont, 1) = Trim(Mid(strCont, lngSubInicio, lngSubFinal - lngSubInicio))
            
            lngSubInicio = InStr(1, strCont, "<nomeAutor>") + Len("<nomeAutor>")
            lngSubFinal = InStr(1, strCont, "</nomeAutor>")
            strResposta(intCont, 2) = Trim(Mid(strCont, lngSubInicio, lngSubFinal - lngSubInicio))
            
            lngSubInicio = InStr(1, strCont, "<objetoAcao>") + Len("<objetoAcao>")
            lngSubFinal = InStr(1, strCont, "</objetoAcao>")
            strResposta(intCont, 3) = Trim(Mid(strCont, lngSubInicio, lngSubFinal - lngSubInicio))
            
            'strResposta(intCont, 1) = XMLResposta.getElementsByTagName("processosDaMatriculaNoSCI")(intCont - 1).SelectSingleNode("numeroProcesso").Text
            'strResposta(intCont, 2) = XMLResposta.getElementsByTagName("processosDaMatriculaNoSCI")(intCont - 1).SelectSingleNode("nomeAutor").Text
            'strResposta(intCont, 3) = XMLResposta.getElementsByTagName("processosDaMatriculaNoSCI")(intCont - 1).SelectSingleNode("objetoAcao").Text
        Next intCont
    
    Else
        ReDim strResposta(0)
        strResposta(0) = "Não há processos cadastrados para a matrícula"
    
    End If
    
    ConsultaProcessosPorMatriculaSCI = strResposta
    
End Function



Attribute VB_Name = "modFuncoesDesuso"
Option Explicit

Function xxIteraNucleoAdv(strGrupo As String, strNucleoOuAdv As String) As String ' Coloquei "xx" na frente para que, caso seja chamada, eu veja a tela de erro.
''
'' A variável strNucleoouAdv indica se estamos buscando um Núcleo ou um Advogado. Se for um Núcleo, retornará uma string
''  com o nome do núcleo e o nome do advogado responsável, separados por "-,-,-". Se for um Advogado, retornará apenas o
''  nome do Advogado.
''
    Dim plan As Worksheet
    Dim rngProximo As Range
    Dim strAdv As String, strProximoAdv As String
    Dim Cont As Integer
    Select Case strNucleoOuAdv
    Case "Nucleo"
        Set plan = ThisWorkbook.Sheets("cfGruposNucleo")
    Case "Adv"
        'Set plan = ThisWorkbook.Sheets("cfGruposAdv")
    End Select
    
    ' Aponta a variável rngProximo para a célula que contém o núcleo/advogado da vez. Coloca o nome do
    ' núcleo/advogado da vez em strAdv.
    Set rngProximo = plan.Cells().Find(strGrupo).Offset(1, 0)
    strAdv = Replace(rngProximo.Formula, "Próximo: ", "")
    
    ' Se for um núcleo, adiciona o Advogado do Núcleo, separado por "-,-,-"
    If strNucleoOuAdv = "Nucleo" Then
        strAdv = strAdv & "-,-,-" & rngProximo.Offset(0, 1).Formula
    End If
    
    ' Itera as pessoas abaixo do grupo até achar a que foi utilizada em strAdv; coloca a próxima
    ' na célula "próximo".
    Cont = 0
    Do
        Cont = Cont + 1
        If InStr(1, strAdv, rngProximo.Offset(Cont, 0).Formula) = 1 Then
            strProximoAdv = IIf(rngProximo.Offset(Cont + 1, 0).Formula <> "", rngProximo.Offset(Cont + 1, 0).Formula, rngProximo.Offset(1, 0).Formula)
            rngProximo.Formula = "Próximo: " & strProximoAdv
            If strNucleoOuAdv = "Nucleo" Then
                rngProximo.Offset(0, 1).Formula = IIf(rngProximo.Offset(Cont + 1, 0).Formula <> "", rngProximo.Offset(Cont + 1, 1).Formula, rngProximo.Offset(1, 1).Formula)
            End If
            Exit Do
        End If
    Loop Until rngProximo.Offset(Cont, 0).Formula = ""
    
    IteraNucleoAdv = strAdv
    
End Function


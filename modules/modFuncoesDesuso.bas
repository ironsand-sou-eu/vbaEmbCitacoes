Attribute VB_Name = "modFuncoesDesuso"
Option Explicit

Function xxIteraNucleoAdv(strGrupo As String, strNucleoOuAdv As String) As String ' Coloquei "xx" na frente para que, caso seja chamada, eu veja a tela de erro.
''
'' A vari�vel strNucleoouAdv indica se estamos buscando um N�cleo ou um Advogado. Se for um N�cleo, retornar� uma string
''  com o nome do n�cleo e o nome do advogado respons�vel, separados por "-,-,-". Se for um Advogado, retornar� apenas o
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
    
    ' Aponta a vari�vel rngProximo para a c�lula que cont�m o n�cleo/advogado da vez. Coloca o nome do
    ' n�cleo/advogado da vez em strAdv.
    Set rngProximo = plan.Cells().Find(strGrupo).Offset(1, 0)
    strAdv = Replace(rngProximo.Formula, "Pr�ximo: ", "")
    
    ' Se for um n�cleo, adiciona o Advogado do N�cleo, separado por "-,-,-"
    If strNucleoOuAdv = "Nucleo" Then
        strAdv = strAdv & "-,-,-" & rngProximo.Offset(0, 1).Formula
    End If
    
    ' Itera as pessoas abaixo do grupo at� achar a que foi utilizada em strAdv; coloca a pr�xima
    ' na c�lula "pr�ximo".
    Cont = 0
    Do
        Cont = Cont + 1
        If InStr(1, strAdv, rngProximo.Offset(Cont, 0).Formula) = 1 Then
            strProximoAdv = IIf(rngProximo.Offset(Cont + 1, 0).Formula <> "", rngProximo.Offset(Cont + 1, 0).Formula, rngProximo.Offset(1, 0).Formula)
            rngProximo.Formula = "Pr�ximo: " & strProximoAdv
            If strNucleoOuAdv = "Nucleo" Then
                rngProximo.Offset(0, 1).Formula = IIf(rngProximo.Offset(Cont + 1, 0).Formula <> "", rngProximo.Offset(Cont + 1, 1).Formula, rngProximo.Offset(1, 1).Formula)
            End If
            Exit Do
        End If
    Loop Until rngProximo.Offset(Cont, 0).Formula = ""
    
    IteraNucleoAdv = strAdv
    
End Function


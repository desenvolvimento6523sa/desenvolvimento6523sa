Attribute VB_Name = "D_StatusCidade"
Option Private Module
Sub StatusPorCidadePortugues()
Attribute StatusPorCidadePortugues.VB_ProcData.VB_Invoke_Func = " \n14"
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With
Dim Usuario As Object
Dim LogonUsuario As Variant
Dim Tempoinicio As Double
Planilha4.Visible = True

Set Usuario = CreateObject("WScript.Network")
LogonUsuario = Usuario.UserName
Tempoinicio = Time

Planilha11.Select
For i = 4 To 44
CodMun = Range("A" & i)
    Range("I" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PILOTO")
    Range("J" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PROJETO GSED -  PRIORIDADE")
    Range("K" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PROJETO  GSED - NÃO PRIORIDADE")
    
    Range("O" & i) = WorksheetFunction.CountIfs(Planilha16.Range("D:D"), CodMun, Planilha16.Range("R:R"), "SIM") 'Realizada f2f >> BD ColarControleCampo
    
    Range("S" & i) = WorksheetFunction.CountIfs(Planilha16.Range("D:D"), CodMun, Planilha16.Range("E:E"), 1) 'Agendadas >> BD ColarControleCampo
    
    
Next i

Call AgendamentoCidade

Range("C2") = "COTAS E CONTATOS DISPONÍVEIS - CATI E GSED"
Range("F2") = "STATUS CAMPO - CATI"
Range("I2") = "GSED NO CATI"
Range("N2") = "STATUS CAMPO - F2F GSED"
Range("S2") = "VISITAS AGENDADAS POR DIA -  GSED"

Cells(3, 3) = "Cota TOTAL - PILOTO + PROJETO - CATI E GSED"
Cells(3, 4) = "Nº de contatos GSED Prioridade na listagem CATI"
Cells(3, 5) = "Nº de contatos GSED Não Prioridade na listagem CATI"
Cells(3, 6) = "UNIVERSO"
Cells(3, 7) = "TOTAL REALIZADAS PILOTO + PROJETO"
Cells(3, 8) = "FALTA "
Cells(3, 9) = "Realizadas na etapa no PILOTO"
Cells(3, 10) = "Realizadas PROJETO GSED -  PRIORIDADE"
Cells(3, 11) = "Realizadas NO PROJETO  GSED - NÃO PRIORIDADE"
Cells(3, 12) = "TOTAL GSED NO CATI"
Cells(3, 13) = "FALTA"
Cells(3, 14) = "Realizadas na etapa no PILOTO"
Cells(3, 15) = "Realizadas na etapa no PROJETO"
Cells(3, 16) = "TOTAL REALIZADAS PILOTO + PROJETO"
Cells(3, 17) = "FALTA (CATI)"
Cells(3, 18) = "FALTA (COTA)"
Cells(3, 19) = "TOTAL VISITAS AGENDADAS PROJETO"

Planilha4.Visible = xlSheetVeryHidden
Planilha11.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Status por Cidade Cati  F2F Atualizados Português em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"

End Sub
Sub StatusPorCidadeIngles()
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
End With
Dim Usuario As Object
Dim LogonUsuario As Variant
Dim Tempoinicio As Double
Planilha4.Visible = True

Set Usuario = CreateObject("WScript.Network")
LogonUsuario = Usuario.UserName
Tempoinicio = Time

Planilha11.Select

For i = 4 To 44
CodMun = Range("A" & i)
    Range("I" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PILOTO")
    Range("J" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PROJETO GSED -  PRIORIDADE")
    Range("K" & i) = WorksheetFunction.CountIfs(Planilha1.Range("X:X"), CodMun, Planilha1.Range("D:D"), 1, Planilha1.Range("W:W"), "PROJETO  GSED - NÃO PRIORIDADE")

    Range("O" & i) = WorksheetFunction.CountIfs(Planilha16.Range("D:D"), CodMun, Planilha16.Range("R:R"), "SIM") 'Realizada f2f >> BD ColarControleCampo
    
    Range("S" & i) = WorksheetFunction.CountIfs(Planilha16.Range("D:D"), CodMun, Planilha16.Range("E:E"), 1) 'Agendadas >> BD ColarControleCampo
        
Next i

Call AgendamentoCidade

Range("C2") = "QUOTAS AND AVAILABLE  CONTACTS - CATI & GSED     "
Range("F2") = "STATUS FIELDWORK - CATI      "
Range("I2") = "GSED WITHIN CATI             "
Range("N2") = "STATUS FIELDWORK - F2F GSED              "
Range("S2") = "GSED SCHEDULED VISITS"

Cells(3, 3) = "QUOTATOTAL - PILOT + PROJECT - CATI & F2F"
Cells(3, 4) = "# of GSED Priority contacts on CATI list"
Cells(3, 5) = "# of GSED Non Priority contacts on CATI list"
Cells(3, 6) = "UNIVERSE"
Cells(3, 7) = "TOTAL COMPLETES PILOT + PROJECT"
Cells(3, 8) = "# TO ACHIEVE"
Cells(3, 9) = "Completes - PILOT"
Cells(3, 10) = "Completes PROJECT GSED -  PRIORITY"
Cells(3, 11) = "Completes PROJECT GSED -  NON PRIORITY"
Cells(3, 12) = "TOTAL GSED WITHIN CATI"
Cells(3, 13) = "# TO ACHIEVE"
Cells(3, 14) = "Completes - PILOT"
Cells(3, 15) = "Completes - PROJECT"
Cells(3, 16) = "TOTAL COMPLETES PILOT + PROJECT"
Cells(3, 17) = "# TO ACHIEVE (CATI)"
Cells(3, 18) = "# TO ACHIEVE (QUOTA)"
Cells(3, 19) = "TOTAL SCHEDULED VISITS - PROJECT"

Planilha4.Visible = xlSheetVeryHidden
Planilha11.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Status por Cidade Cati  F2F Atualizados Inglês em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub


Sub AgendamentoCidade()
Planilha11.Select
'Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row
Range("T4:BH44").ClearContents
For i = 20 To 60
Datax = Cells(3, i)

        For X = 4 To 44
            Cidade = Cells(X, 1)
                Cells(X, i) = WorksheetFunction.CountIfs(Planilha16.Range("F:F"), Datax, Planilha16.Range("D:D"), Cidade)
        Next X

Next i
End Sub

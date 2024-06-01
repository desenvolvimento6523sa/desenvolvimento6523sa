Attribute VB_Name = "C_Producao"
Option Private Module

Sub ProdutividPortugues()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With

Dim Usuario As Object
Dim LogonUsuario As Variant
Dim Tempoinicio As Double

Set Usuario = CreateObject("WScript.Network")
LogonUsuario = Usuario.UserName
Tempoinicio = Time

Planilha10.Select


'Atualizar Dinamica
ActiveSheet.PivotTables("OcorrenciaTab").PivotCache.Refresh

Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row
Range("B5:AW96").ClearContents
For i = 3 To 96
Datax = Cells(3, i)

    If Datax <> "" Then
        For X = 5 To Ultima
            Ocorr = Cells(X, 1)
                If Ocorr <> "" And Ocorr <> "(vazio)" And Ocorr <> "Total Geral" Then
                    Cells(X, i) = WorksheetFunction.CountIfs(Planilha1.Range("R:R"), Datax, Planilha1.Range("U:U"), Ocorr)
                End If
        Next X
    End If
Next i
Call SomaProdutividade
Call LabelProdutividadePortuguesA
Cells(1, 1).Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With

MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Produtividade Calculada em - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI ENDLINE"
End Sub
Sub ProdutividIngles()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With

Dim Usuario As Object
Dim LogonUsuario As Variant
Dim Tempoinicio As Double

Set Usuario = CreateObject("WScript.Network")
LogonUsuario = Usuario.UserName
Tempoinicio = Time

Planilha10.Select


'Atualizar Dinamica
ActiveSheet.PivotTables("OcorrenciaTab").PivotCache.Refresh

Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row
Range("B5:AW96").ClearContents
For i = 3 To 96
Datax = Cells(3, i)

    If Datax <> "" Then
        For X = 5 To Ultima
            Ocorr = Cells(X, 1)
                If Ocorr <> "" And Ocorr <> "(vazio)" And Ocorr <> "Total Geral" Then
                    Cells(X, i) = WorksheetFunction.CountIfs(Planilha1.Range("R:R"), Datax, Planilha1.Range("U:U"), Ocorr)
                End If
        Next X
    End If
Next i
Call SomaProdutividade
Call LabelProdutividadeInglesB
Cells(1, 1).Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With

MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Produtividade Calculada em - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI ENDLINE"
End Sub
Sub SomaProdutividade()

Planilha10.Select
Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row
For i = 5 To Ultima
    Range("B" & i) = WorksheetFunction.Sum(Range("C" & i, "CR" & i))
Next i

End Sub

Sub LabelProdutividadePortuguesA()
Planilha10.Select
Cells(2, 1) = "PRODUTIVIDADE AO DIA SOMENTE ÚLTIMA OCORRÊNCIA - HISTÓRICO DETALHADO - CATI"
For i = 3 To 96
    If Cells(4, i) = "Mon" Then Cells(4, i) = "seg"
    If Cells(4, i) = "Tue" Then Cells(4, i) = "ter"
    If Cells(4, i) = "Wed" Then Cells(4, i) = "qua"
    If Cells(4, i) = "Thu" Then Cells(4, i) = "qui"
    If Cells(4, i) = "Fri" Then Cells(4, i) = "sex"
    If Cells(4, i) = "Sat" Then Cells(4, i) = "sáb"
    If Cells(4, i) = "Sun" Then Cells(4, i) = "dom"
Next i
End Sub
Sub LabelProdutividadeInglesB()
Planilha10.Select
Cells(2, 1) = "PRODUCTIVITY PER DAY - LAST OCCURRENCE ONLY - DETAILED HISTORY - CATI"
For i = 3 To 96
    If Cells(4, i) = "seg" Then Cells(4, i) = "Mon"
    If Cells(4, i) = "ter" Then Cells(4, i) = "Tue"
    If Cells(4, i) = "qua" Then Cells(4, i) = "Wed"
    If Cells(4, i) = "qui" Then Cells(4, i) = "Thu"
    If Cells(4, i) = "sex" Then Cells(4, i) = "Fri"
    If Cells(4, i) = "sáb" Then Cells(4, i) = "Sat"
    If Cells(4, i) = "dom" Then Cells(4, i) = "Sun"
Next i
End Sub





















Attribute VB_Name = "G_VisaoGeral"
Option Private Module
Sub VisaoGeral()
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

Planilha13.Select

For i = 4 To 1602
id = Cells(i, 1) & "_1"
idx = Cells(i, 1).Value


Range("G" & i) = Application.IfError(Application.VLookup(id, Planilha1.Range("AB:AC"), 2, 0), "")
Range("H" & i) = Application.IfError(Application.VLookup(idx, Planilha16.Range("A:M"), 8, 0), "")

Next i

Planilha4.Visible = xlSheetVeryHidden
Planilha13.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Visão geral atualizado em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub

Attribute VB_Name = "E_FoneErrado"
Option Private Module
Sub SeparaTelefones01()
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

'Telefones errados
Planilha12.Select

ultimalinha = Planilha1.Cells(Rows.Count, "a").End(xlUp).Row
lin = 2
For i = 2 To ultimalinha
    If (Planilha1.Cells(i, 4) = "121" Or Planilha1.Cells(i, 4) = "123") And Planilha1.Cells(i, 16) = 1 Then
        Planilha12.Cells(lin, 1) = Planilha1.Cells(i, 25) 'ID IPEC
        Planilha12.Cells(lin, 2) = Planilha1.Cells(i, 26) 'ID CHILD
        Planilha12.Cells(lin, 3) = Planilha1.Cells(i, 4) 'OCORRÊNCIA
        If Planilha1.Cells(i, 4) = 121 Then Planilha12.Cells(lin, 4) = "PHONE DOES NOT EXIST"
        If Planilha1.Cells(i, 4) = 123 Then Planilha12.Cells(lin, 4) = "INCORRECT PHONE NUMBER"
        Planilha12.Cells(lin, 5) = Planilha1.Cells(i, 27) 'DATA
        
        
        lin = lin + 1
End If
Next i
Call SeparaTelefones02

Planilha4.Visible = xlSheetVeryHidden
Planilha12.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Telefones Errados separados em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub

Sub SeparaTelefones02()
'Telefones errados
Planilha12.Select
Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row

For i = 2 To Ultima
id = Cells(i, 1).Value
Datax = Cells(i, 5).Value

    If id <> "" And Datax = "" Then Cells(i, 5) = Date
    

Next i


End Sub

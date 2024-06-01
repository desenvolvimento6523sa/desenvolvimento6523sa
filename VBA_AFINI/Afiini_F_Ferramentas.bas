Attribute VB_Name = "F_Ferramentas"
Option Private Module
Sub ExcluirDadosHD()
Planilha1.Select
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
Next ws

'Excluir (42) ENTREVISTA INICIADA POR BUSCA (13) SESSÃO EXPIRADA
linha = Cells(Cells.Rows.Count, "a").End(xlUp).Row
Range("A1:T1").Select
Selection.AutoFilter
ActiveSheet.Range("A1:T" & linha).AutoFilter Field:=4, Criteria1:="=13", Operator:=xlOr, Criteria2:="=42"
Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
Selection.AutoFilter

'Excluir Nome =Estudo 5608
Range("A1:T1").Select
Selection.AutoFilter
ActiveSheet.Range("A1:T" & linha).AutoFilter Field:=2, Criteria1:="Estudo 5608"
Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
Selection.AutoFilter

Range("A1").Select
End Sub



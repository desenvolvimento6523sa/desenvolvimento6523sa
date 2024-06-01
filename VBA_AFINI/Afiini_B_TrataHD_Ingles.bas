Attribute VB_Name = "B_TrataHD_Ingles"
Option Private Module
Sub TratamentoHD_01_Ingles()
Planilha1.Select
Range("N1") = "CODX"
Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row

    Columns("A:T").Select
    ActiveWorkbook.Worksheets("ColarHD").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ColarHD").Sort.SortFields.Add2 Key:=Range("A2:A" & Ultima), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("ColarHD").Sort.SortFields.Add2 Key:=Range("K2:K" & Ultima), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ColarHD").Sort
        .SetRange Range("A1:T" & Ultima)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


Range("C2:C" & Ultima).ClearContents
Range("C2").FormulaR1C1 = "=IF(R[-1]C1=RC1,R[-1]C3+1,1)"
Range("C2").AutoFill Range(Range("C2"), Range("A2").End(xlDown).Offset(0, 2))
Range("C:C").Copy
Range("C:C").PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False

Columns("P:AD").Clear
Columns("K:K").Select
    Selection.Copy
Range("R1").Select
    ActiveSheet.Paste
    Columns("R:R").Select
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Range("R1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("R:R").NumberFormat = "m/d/yyyy"


Cells(1, 16) = "Última Ocorrência"
Cells(1, 17) = "Total de visitas"
Cells(1, 18) = "Data da Ocorrência"
Cells(1, 19) = "Concat id_discagem"
Cells(1, 20) = "Código da Ocorrência"
Cells(1, 21) = "OcorrenciaX"
Cells(1, 22) = "Apoio 2"
Cells(1, 23) = "Apoio 3"
Cells(1, 24) = "Apoio 4"
Cells(1, 25) = "Apoio 5"
Cells(1, 26) = "Apoio 6"
Cells(1, 27) = "Apoio 7"
Cells(1, 28) = "Apoio 8"
Cells(1, 29) = "Apoio 9"

For i = 2 To Ultima
id = Range("A" & i) 'id
Discagem1 = Range("C" & i)
Dataocorrencia = Range("K" & i)
If id <> "" Then
    If id <> Range("A" & i).Offset(1, 0) Then Range("P" & i) = 1 'Ultima ocorrencia

If Discagem1 <> "" And id = "" Then Range("C" & i) = ""
End If

Next i


'Tradutor Portugues Ingles
Call Labels_Ingles_Novo1

For j = 2 To Ultima
Ultimavisita = Range("P" & j)
Discagem = Range("C" & j)
If Ultimavisita = 1 Then
    If Discagem = 1 Then Range("Q" & j) = "1 visitas"
    If Discagem = 2 Then Range("Q" & j) = "2 visitas"
    If Discagem = 3 Then Range("Q" & j) = "3 visitas"
    If Discagem = 4 Then Range("Q" & j) = "4 visitas"
    If Discagem = 5 Then Range("Q" & j) = "5 ou mais visitas"
    Range("AB" & j) = Range("A" & j) & "_" & Ultimavisita
    Range("AC" & j) = Range("U" & j)
     
End If
Next j


Range("S1") = "Concat id_discagem"
Range("T:T").ClearContents
Range("T1") = "Código da Ocorrência"
For k = 2 To Ultima
idEntrevista = Range("A" & k)
OrdemVisita = Range("C" & k)
Ocorrencias = Range("U" & k)
Data_Ocorrencia = Range("K" & k)
Data_agendamento = Range("O" & k)

    Range("S" & k) = idEntrevista & "_" & OrdemVisita
    If Ocorrencias = "SCHEDULE" Then
        Range("T" & k) = Ocorrencias & " | " & Data_Ocorrencia & "| Date Hour Schedule | " & Data_agendamento
    Else
        Range("T" & k) = Ocorrencias & " | " & Data_Ocorrencia
    
    End If
Next k


End Sub
Sub TratamentoHD_02_Ingles()

Planilha4.Select
    Range("CN5").FormulaR1C1 = "=VLOOKUP(RC[-84],ColarHD!C19:C20,2,0)"

    Range("CN5").AutoFill Destination:=Range("CN5:FO5")
    Range("CN5:FO5").AutoFill Destination:=Range("CN5:FO1603")

    Range("CN5:FO1603").Copy
    Range("CN5:FO1603").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Range("CN5:FO1603").Replace What:="#", Replacement:=""
    Range("CN5:FO1603").Replace What:="N/A", Replacement:=""
     

End Sub

Sub TratamentoHD_03_Ingles()
Range("CJ5:CM1603").ClearContents
    Range("FQ5").FormulaR1C1 = "=HLOOKUP(COUNTA(RC92:RC171),R3C92:R3416C171,RC172,0)"
    Range("FQ5").AutoFill Destination:=Range("FQ5:FQ1603")
    Range("FQ5:FQ1603").Copy
    Range("FQ5:FQ1603").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False


    Range("FQ5:FQ1603").Replace What:="#", Replacement:=""
    Range("FQ5:FQ1603").Replace What:="N/A", Replacement:=""
     
'TOTAL DE OCORRÊNCIAS
    Range("FR5").FormulaR1C1 = "=COUNTA(RC92:RC171)"
    Range("FR5").AutoFill Destination:=Range("FR5:FR1603")
    Range("FR5:FR1603").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("FR4").Select
     
'////////////////////////////////////////////////////////////////

Ultima = Cells(Cells.Rows.Count, "a").End(xlUp).Row

For m = 5 To 1603
Resultado = Range("CJ" & m)
Acao_2 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*PHONE DOESN'T EXIST*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*INCORRECT PHONE NUMBER*")

Acao_3 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*RETURN*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*SCHEDULE*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATSAPP MESSAGE SENT AND ANSWERED*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATSAPP MESSAGE SENT AND NOT ANSWERED*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATSAPP CALL - DID NOT PICK UP*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATS APP SIGN BUSY*")
    
Acao_4 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*NO ANSWER*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*PHONE BUSY*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*PHONE OUT OF AREA/ OFF*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*CONNECTION COULD NOT BE COMPLETED*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*ELECTRONIC SECRETARY / VOICEMAIL*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FAX SIGNAL*")


Ocorrencia = Cells(m, 173)
 '//COMPLETED_OK
            If InStr(1, Ocorrencia, "COMPLETED_OK", vbTextCompare) > 0 Then Cells(m, 88) = "COMPLETED ACCOMPLISHED"
            
 '//FINISHED - LOSS
            If InStr(1, Ocorrencia, "NEVER CALL THIS NUMBER", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "DOES NOT WANT TO PARTICIPATE", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "REQUESTS THE PHONE TO BE DELETED FROM THE LIST", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "FILTER - CAREGIVER'S AGE UNDER 18 YEARS", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "NAME OF THE CHILD DIVERGING FROM THE REGISTRATION", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "ABANDONMENT", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"
            If InStr(1, Ocorrencia, "WHATSAPP/ BLOCKED", vbTextCompare) > 0 Then Cells(m, 88) = "FINISHED - LOSS"

            
  '//Não passível de recontato...Após 1 ocorrência contatar via WhatsApp - total de tentativas
    
            If InStr(1, Ocorrencia, "PHONE DOESN'T EXIST", vbTextCompare) > 0 Then Cells(m, 89) = "(" & Acao_2 & " Contacts) - Not recontactable...After 1 occurrence contact via WhatsApp - total attempts"
            If InStr(1, Ocorrencia, "REFUSAL", vbTextCompare) > 0 Then Cells(m, 89) = "(" & Acao_2 & " Contacts) - Not recontactable...After 1 occurrence contact via WhatsApp - total attempts"
  
  '//Passível de recontato...Pelo menos 3 tentativas
            If InStr(1, Ocorrencia, "RETURN", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
            If InStr(1, Ocorrencia, "SCHEDULE", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
            If InStr(1, Ocorrencia, "WHATSAPP MESSAGE SENT AND ANSWERED", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
            If InStr(1, Ocorrencia, "WHATSAPP MESSAGE SENT AND NOT ANSWERED", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
            If InStr(1, Ocorrencia, "WHATSAPP CALL - DID NOT PICK UP", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
            If InStr(1, Ocorrencia, "WHATS APP SIGN BUSY", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " contacts) - Recontactable...at least 3 attempts"
  
  '//Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "NO ANSWER", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
            If InStr(1, Ocorrencia, "PHONE BUSY", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
            If InStr(1, Ocorrencia, "PHONE OUT OF AREA/ OFF", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
            If InStr(1, Ocorrencia, "CONNECTION COULD NOT BE COMPLETED", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
            If InStr(1, Ocorrencia, "ELECTRONIC SECRETARY / VOICEMAIL", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
            If InStr(1, Ocorrencia, "FAX SIGNAL", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_3 & "(" & Acao_4 & " contacts) - Recontactable...at least 3 attempts - via WhatsApp"
Next m
'Ocorrências ação
Planilha4.Select
    Range("CJ5:CM1603").Select
    Selection.Copy
Planilha8.Select
    Range("F5").Select
    ActiveSheet.Paste
'Ocorrências ação
Planilha4.Select
    Range("CN5:FO1603").Select
    Selection.Copy
Planilha8.Select
    Range("L5").Select
    ActiveSheet.Paste
    
'Ultima Ocorrência
Planilha4.Select
    Range("FQ5:FQ1603").Select
    Selection.Copy
Planilha8.Select
    Range("K5").Select
    ActiveSheet.Paste
    
'TOTAL DE OCORRÊNCIAS
Planilha4.Select
    Range("FR5:FR1603").Select
    Selection.Copy
Planilha8.Select
    Range("J5").Select
    ActiveSheet.Paste

    Application.CutCopyMode = False

Planilha8.Select
Call TratamentoCabecalhoIngles
Cells(1, 1).Select

End Sub


Sub TratamentoCabecalhoIngles()
'Controle OCORRÊNCIAS CATI
Planilha8.Select
Range("A1") = "GENERAL CONTROL BY CONTACT"
Range("F3") = "SUMMARY OF DISPOSITIONS AND ACTIONS - CATI"
Range("J3") = "OCCURRENCE SUMMARY PER CONTACT"
Range("L3") = "DISPOSITIONS PER CONTACT - CATI"

Range("B4") = "CA2 - MUNICIPALITY"
Range("C4") = "CA2 - MUNICIPALITY_2"
Range("D4") = "CA3 - FAMILY ID"
Range("E4") = "ID_CHILD"

Range("F4") = "COMPLETES"
Range("G4") = "Cannot be recontacted...After 1 disposition contact via WhatsApp - total attempts"
Range("H4") = "Can be recontacted... At least 3 attempts"
Range("I4") = "Can be recontacted... After 3 attempts, contact via WhatsApp"
Range("J4") = "TOTAL NUMBER OF CONTACTS MADE"
Range("K4") = "STATUS OF THE LAST DISPOSITION - CATI"


Range("L4") = "DISPOSITION 1"
    Range("L4").Select
    Selection.AutoFill Destination:=Range("L4:CM4"), Type:=xlFillDefault
    Range("L4:CM4").Select
Cells(1, 1).Select


End Sub
Sub Labels_Ingles_Novo1()
Planilha1.Select

'=======================================
'Ingles
    Range("U2").FormulaR1C1 = "=VLOOKUP(RC4,'LABEL_COD AÇOES _CATI'!C1:C5,4,0)"
    Range("V2").FormulaR1C1 = "=VLOOKUP(RC4,'LABEL_COD AÇOES _CATI'!C1:C5,5,0)"
    Range("W2").FormulaR1C1 = "=VLOOKUP(RC1,Listagem!C1:C14,14,0)"
    Range("X2").FormulaR1C1 = "=VLOOKUP(RC1,Listagem!C1:C14,6,0)"
    Range("Y2").FormulaR1C1 = "=VLOOKUP(RC1,Listagem!C1:C14,2,0)"
    Range("Z2").FormulaR1C1 = "=VLOOKUP(RC1,Listagem!C1:C14,3,0)"
    Range("AA2").FormulaR1C1 = "=IF(RC16=1,VLOOKUP(RC25,'TELEFONES ERRADOS'!C1:C6,5,0),"""")"


    Range("U2").AutoFill Range(Range("U2"), Range("A2").End(xlDown).Offset(0, 20))
    Range("V2").AutoFill Range(Range("V2"), Range("A2").End(xlDown).Offset(0, 21))
    Range("W2").AutoFill Range(Range("W2"), Range("A2").End(xlDown).Offset(0, 22))
    Range("X2").AutoFill Range(Range("X2"), Range("A2").End(xlDown).Offset(0, 23))
    Range("Y2").AutoFill Range(Range("Y2"), Range("A2").End(xlDown).Offset(0, 24))
    Range("Z2").AutoFill Range(Range("Z2"), Range("A2").End(xlDown).Offset(0, 25))
    Range("AA2").AutoFill Range(Range("AA2"), Range("A2").End(xlDown).Offset(0, 26))
         
Columns("U:AA").Select
    Selection.Copy
Columns("U:AA").Select

    Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
       
'APAGAR #N/D
Columns("U:AA").Replace What:="#", Replacement:=""
Columns("U:AA").Replace What:="N/A", Replacement:=""


End Sub

Sub GeraControleIngles()
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

Planilha1.Select

Call TratamentoHD_01_Ingles
Call TratamentoHD_02_Ingles
Call TratamentoHD_03_Ingles

Planilha4.Visible = xlSheetVeryHidden
Planilha8.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Ocorrências Cati Atualizados Inglês em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub


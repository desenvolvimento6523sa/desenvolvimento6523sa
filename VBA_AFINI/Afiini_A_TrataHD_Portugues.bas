Attribute VB_Name = "A_TrataHD_Portugues"
Option Private Module
Sub TratamentoHD_01_Portugues()
Attribute TratamentoHD_01_Portugues.VB_ProcData.VB_Invoke_Func = " \n14"
'Aba ColarHD
Planilha1.Select
Call ExcluirDadosHD
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

'Tradutor Ingles Portugues
Call Labels_Portugues_Novo1

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
    If Ocorrencias = "AGENDAR" Then
        Range("T" & k) = Ocorrencias & " | " & Data_Ocorrencia & "| Data hora agendado | " & Data_agendamento
    Else
        Range("T" & k) = Ocorrencias & " | " & Data_Ocorrencia
    
    End If
Next k


End Sub
Sub TratamentoHD_02_Portugues()
'Trazer ocorrências do Histórico Detalhado
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


Sub TratamentoHD_03_Portugues()
Planilha4.Select
Range("CJ5:CM1603").ClearContents
    Range("FQ5").FormulaR1C1 = "=HLOOKUP(COUNTA(RC92:RC171),R3C92:R3416C171,RC172,0)"
    Range("FQ5").AutoFill Destination:=Range("FQ5:FQ1603")
    Range("FQ5:FQ1603").Copy
    Range("FQ5:FQ1603").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False


    Range("FQ5:FQ1603").Replace What:="#", Replacement:=""
    Range("FQ5:FQ31603").Replace What:="N/A", Replacement:=""
     
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
Acao_2 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FONE NÃO EXISTE*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FONE ERRADO*")

Acao_3 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*ENTREVISTA AGENDADA*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*RETORNO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*MENSAGEM ENVIADA E RESPONDIDA - EM CONTATO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*MENSAGEM ENVIADA OU ENTREGUE E SEM RETORNO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATS APP NÃO ATENDE*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*WHATSAPP DANDO OCUPADO*")
    
Acao_4 = WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FONE NÃO ATENDE*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FONE OCUPADO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*FORA DE ÁREA / DESLIGADO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*NÃO FOI POSSÍVEL COMPLETAR A LIGAÇÃO*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*SECRETÁRIA ELETRÔNICA / CAIXA POSTAL*") _
    + WorksheetFunction.CountIf(Range("CN" & m, "FO" & m), "*SINAL DE FAX*")
                              
Ocorrencia = Cells(m, 173)
 '//REALIZADO
            If InStr(1, Ocorrencia, "REALIZADA", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - REALIZADO"
            
 '//FINALIZADO - PERDA
            If InStr(1, Ocorrencia, "NUNCA LIGAR PARA ESTE NUMERO", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "RECUSA DO RESPONDENTE", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "SOLICITA A EXCLUSÃO DO TELEFONE DE NOSSO CADASTRO", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "FILTRO - IDADE DO CUIDADOR INFERIOR A 18 ANOS", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "NOME DA CRIANÇA DIVERGENTE DO CADASTRO", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "ABANDONO", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            If InStr(1, Ocorrencia, "TELEFONE NÃO TEM WHATSAPP/ BLOQUEADO", vbTextCompare) > 0 Then Cells(m, 88) = "FINALIZADO - PERDA"
            
  '//Não passível de recontato...Após 1 ocorrência contatar via WhatsApp - total de tentativas
    
            If InStr(1, Ocorrencia, "FONE NÃO EXISTE", vbTextCompare) > 0 Then Cells(m, 89) = "(" & Acao_2 & " Contatos) - Não passível de recontato...Após 1 ocorrência contatar via WhatsApp - total de tentativas"
            If InStr(1, Ocorrencia, "FONE ERRADO", vbTextCompare) > 0 Then Cells(m, 89) = "(" & Acao_2 & " Contatos) - Não passível de recontato...Após 1 ocorrência contatar via WhatsApp - total de tentativas"
  
  '//Passível de recontato...Pelo menos 3 tentativas
            If InStr(1, Ocorrencia, "ENTREVISTA AGENDADA", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
            If InStr(1, Ocorrencia, "RETORNO", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
            If InStr(1, Ocorrencia, "MENSAGEM ENVIADA E RESPONDIDA - EM CONTATO", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
            If InStr(1, Ocorrencia, "MENSAGEM ENVIADA OU ENTREGUE E SEM RETORNO", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
            If InStr(1, Ocorrencia, "WHATS APP NÃO ATENDE", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
            If InStr(1, Ocorrencia, "WHATSAPP DANDO OCUPADO", vbTextCompare) > 0 Then Cells(m, 90) = "(" & Acao_3 & " Contatos) - Passível de recontato...Pelo menos 3 tentativas"
  
  '//Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "FONE NÃO ATENDE", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "FONE OCUPADO", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "FORA DE ÁREA / DESLIGADO", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "NÃO FOI POSSÍVEL COMPLETAR A LIGAÇÃO", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "SECRETÁRIA ELETRÔNICA / CAIXA POSTAL", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
            If InStr(1, Ocorrencia, "SINAL DE FAX", vbTextCompare) > 0 Then Cells(m, 91) = "(" & Acao_3 & "(" & Acao_4 & " Contatos) - Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
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
Call TratamentoCabecalhoPortugues
Cells(1, 1).Select

End Sub

Sub TratamentoCabecalhoPortugues()
'Controle OCORRÊNCIAS CATI
Planilha8.Select
Range("A1") = "CONTROLE GERAL POR CONTATO"
Range("F3") = "RESUMO DAS OCORRÊNCIAS E AÇÕES - CATI"
Range("J3") = "RESUMO DA OCORRENCIA POR CONTATO"
Range("L3") = "OCORRÊNCIAS POR CONTATO - CATI"

Range("B4") = "CA2 - MUNICÍPIO"
Range("C4") = "CA2 - MUNICÍPIO_2"
Range("D4") = "CA3 - Código Familiar"
Range("E4") = "ID_Criança"

Range("F4") = "FINALIZADOS"
Range("G4") = "Não passível de recontato...Após 1 ocorrência contatar via WhatsApp - total de tentativas"
Range("H4") = "Passível de recontato...Pelo menos 3 tentativas"
Range("I4") = "Passível de recontato...Após 3 tentativas, contatar via WhatsApp"
Range("J4") = "TOTAL DE CONTATOS REALIZADOS"
Range("K4") = "STATUS DA ULTIMA OCORRENCIA - CATI"


Range("B4") = "FINALIZADOS"


Range("L4") = "OCORRÊNCIA 1"
    Range("L4").Select
    Selection.AutoFill Destination:=Range("L4:CM4"), Type:=xlFillDefault
    Range("L4:CM4").Select

Cells(1, 1).Select

End Sub
Sub Labels_Portugues_Novo1()
Planilha1.Select

'=======================================
'Português
    Range("U2").FormulaR1C1 = "=VLOOKUP(RC4,'LABEL_COD AÇOES _CATI'!C1:C5,2,0)"
    Range("V2").FormulaR1C1 = "=VLOOKUP(RC4,'LABEL_COD AÇOES _CATI'!C1:C5,3,0)"
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

Sub GeraControlePortugues()
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

Call TratamentoHD_01_Portugues
Call TratamentoHD_02_Portugues
Call TratamentoHD_03_Portugues

Planilha4.Visible = xlSheetVeryHidden
Planilha8.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Controle Ocorrências Cati Atualizados Português em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub




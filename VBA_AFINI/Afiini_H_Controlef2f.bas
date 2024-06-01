Attribute VB_Name = "H_Controlef2f"
Option Private Module
Sub ControleOcorrGsedPortugues()
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

Planilha9.Select
    Range("I5:CJ1603").ClearContents

'ColarControleCampo
Planilha16.Select
    Range("I3:N1601").Select
    Selection.Copy

'CONTROLE_OCORRÊNCIAS_GSED
Planilha9.Select
    Range("I5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


'ColarControleCampo >> Total de contatos
Planilha16.Select
    Range("G3:H1601").Select
    Selection.Copy

'CONTROLE_OCORRÊNCIAS_GSED >> última ocorrência
Planilha9.Select
    Range("G5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Application.CutCopyMode = False
Call ControleOcorrGeedPortugues

Planilha4.Visible = xlSheetVeryHidden
Planilha9.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> CONTROLE_OCORRÊNCIAS_GSED atualizado em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub
Sub ControleOcorrGsedingles()
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

Planilha9.Select
    Range("I5:CJ1603").ClearContents

'ColarControleCampo
Planilha16.Select
    Range("I3:N1601").Select
    Selection.Copy

'CONTROLE_OCORRÊNCIAS_GSED
Planilha9.Select
    Range("I5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


'ColarControleCampo >> Total de contatos
Planilha16.Select
    Range("G3:H1601").Select
    Selection.Copy

'CONTROLE_OCORRÊNCIAS_GSED >> última ocorrência
Planilha9.Select
    Range("G5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Application.CutCopyMode = False
Call ControleOcorrGeedIngles

Planilha4.Visible = xlSheetVeryHidden
Planilha9.Select
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With
MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> CONTROLE_OCORRÊNCIAS_GSED atualizado em  - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- Obrigado!!!", , "BANCO MUNDIAL AFINI"
End Sub
Sub ControleOcorrGeedPortugues()

Planilha9.Select
Range("D1") = "CONTROLE GERAL POR CONTATO"
Range("G3") = "RESUMO DA OCORRENCIA POR CONTATO "
Range("I3") = "OCORRÊNCIAS POR CONTATO - GSED"


Range("A4") = "ID_IPEC"
Range("B4") = "CA2 - MUNICÍPIO"
Range("C4") = "CA2 - MUNICÍPIO_2"
Range("D4") = "CA3 - Código Familiar"
Range("E4") = "ID_Criança"
Range("F4") = "GSED  Sim, Prioritário, Sim, não prioritários, Não é GSED (Inelegível)"
Range("G4") = "TOTAL DE CONTATOS REALIZADOS"
Range("H4") = "STATUS DA ULTIMA OCORRENCIA"
Range("I4") = "OCORRÊNCIA 1"
    Range("I4").AutoFill Destination:=Range("I4:CJ4"), Type:=xlFillDefault
    Range("I4:CJ4").Select
    Range("C4").Select
    
For i = 5 To 1603
ColunaGSED = Cells(i, 6).Value
ColunaOcorr = Cells(i, 8).Value

    If InStr(1, ColunaGSED, "GSED YES PRIORITY", vbTextCompare) > 0 Then Cells(i, 6) = "GSED SIM - PRIORITÁRIO"
    If InStr(1, ColunaGSED, "YES NON PRIORITY", vbTextCompare) > 0 Then Cells(i, 6) = "GSED SIM - NÃO PRIORITÁRIOS"
    If InStr(1, ColunaGSED, "NOT GSED", vbTextCompare) > 0 Then Cells(i, 6) = "NÃO É GSED (INELEGÍVEL)"
    If InStr(1, ColunaGSED, "PILOT", vbTextCompare) > 0 Then Cells(i, 6) = "PILOTO"
        
    If InStr(1, ColunaOcorr, "IN CONFIRMATION FOR SCHEDULING", vbTextCompare) > 0 Then Cells(i, 8) = "Em confirmação para agendamento"
    If InStr(1, ColunaOcorr, "TEST CARRIED OUT", vbTextCompare) > 0 Then Cells(i, 8) = "TESTE REALIZADO"
    If InStr(1, ColunaOcorr, "CLOSED QUOTA", vbTextCompare) > 0 Then Cells(i, 8) = "COTA FECHADA"
    If InStr(1, ColunaOcorr, "TRAVELING", vbTextCompare) > 0 Then Cells(i, 8) = "Viajando"
    If InStr(1, ColunaOcorr, "CONFIRMED FOR", vbTextCompare) > 0 Then Cells(i, 8) = "Confirmada para"
    If InStr(1, ColunaOcorr, "REFUSAL", vbTextCompare) > 0 Then Cells(i, 8) = "RECUSA"
    If InStr(1, ColunaOcorr, "INELIGIBLE", vbTextCompare) > 0 Then Cells(i, 8) = "Inelegível"
    If InStr(1, ColunaOcorr, "LIVES IN ANOTHER CITY OUTSIDE THE PROJECT SAMPLE", vbTextCompare) > 0 Then Cells(i, 8) = "MORA EM OUTRA CIDADE FORA DA AMOSTRA DO PROJETO"
    If InStr(1, ColunaOcorr, "CONFIRMED FOR", vbTextCompare) > 0 Then Cells(i, 8) = "Confirmada para"
    If InStr(1, ColunaOcorr, "RETURN - RESCHEDULED (INFORM REASON IN COMMENTS)", vbTextCompare) > 0 Then Cells(i, 8) = "RETORNO - REAGENDADA (INFORME MOTIVO EM OBSERVAÇÕES)"
    If InStr(1, ColunaOcorr, "HOUSEHOLDS IN OTHER CITIES", vbTextCompare) > 0 Then Cells(i, 8) = "Domicilios em Outras cidades"
    If InStr(1, ColunaOcorr, "ACCORDING TO HIS MOTHER, HE CAME TO FORTALEZA TO SPEND THE WEEKEND WITH HIS FATHER AND HAS NOT YET RETURNED TO CASCAVEL.", vbTextCompare) > 0 Then Cells(i, 8) = " Segundo a mãe, veio para Fortaleza passar o fim de semana com o pai e ainda não voltou para Cascavel."
    If InStr(1, ColunaOcorr, "THE CHILD WAS UNWELL, CRYING A LOT, AND DID NOT LEAVE HIS MOTHER'S ARMS TO INTERACT AT ALL, DESPITE ATTEMPTS TO PLEASE HIM.", vbTextCompare) > 0 Then Cells(i, 8) = "A  criança estava indisposta, chorando muito, não saia dos braços da mãe pra interagir em nada, apesar de tentativas de aggrado"
    If InStr(1, ColunaOcorr, "INTERVIEWER TRYING TO CONTACT THE MOTHER, BUT SHE DOESN'T ANSWER OR RESPOND TO MESSAGES.", vbTextCompare) > 0 Then Cells(i, 8) = "Entrevistadora tentando contato com a mãe, mas não atende nem responde mensagem."
    If InStr(1, ColunaOcorr, "NOT PART OF GSED", vbTextCompare) > 0 Then Cells(i, 8) = "NÃO FAZ PARTE DO GSED"
        
Next i

For j = 9 To 88
    For k = 5 To 1603
    Ocorrencia = Cells(k, j).Value
        If InStr(1, Ocorrencia, "IN CONFIRMATION FOR SCHEDULING", vbTextCompare) > 0 Then Cells(k, j) = "Em confirmação para agendamento"
        If InStr(1, Ocorrencia, "TEST CARRIED OUT", vbTextCompare) > 0 Then Cells(k, j) = "TESTE REALIZADO"
        If InStr(1, Ocorrencia, "CLOSED QUOTA", vbTextCompare) > 0 Then Cells(k, j) = "COTA FECHADA"
        If InStr(1, Ocorrencia, "TRAVELING", vbTextCompare) > 0 Then Cells(k, j) = "Viajando"
        If InStr(1, Ocorrencia, "CONFIRMED FOR", vbTextCompare) > 0 Then Cells(k, j) = "Confirmada para"
        If InStr(1, Ocorrencia, "REFUSAL", vbTextCompare) > 0 Then Cells(k, j) = "RECUSA"
        If InStr(1, Ocorrencia, "INELIGIBLE", vbTextCompare) > 0 Then Cells(k, j) = "Inelegível"
        If InStr(1, Ocorrencia, "LIVES IN ANOTHER CITY OUTSIDE THE PROJECT SAMPLE", vbTextCompare) > 0 Then Cells(k, j) = "MORA EM OUTRA CIDADE FORA DA AMOSTRA DO PROJETO"
        If InStr(1, Ocorrencia, "CONFIRMED FOR", vbTextCompare) > 0 Then Cells(k, j) = "Confirmada para"
        If InStr(1, Ocorrencia, "RETURN - RESCHEDULED (INFORM REASON IN COMMENTS)", vbTextCompare) > 0 Then Cells(k, j) = "RETORNO - REAGENDADA (INFORME MOTIVO EM OBSERVAÇÕES)"
        If InStr(1, Ocorrencia, "HOUSEHOLDS IN OTHER CITIES", vbTextCompare) > 0 Then Cells(k, j) = "Domicilios em Outras cidades"
        If InStr(1, Ocorrencia, "ACCORDING TO HIS MOTHER, HE CAME TO FORTALEZA TO SPEND THE WEEKEND WITH HIS FATHER AND HAS NOT YET RETURNED TO CASCAVEL.", vbTextCompare) > 0 Then Cells(k, j) = "Segundo a mãe, veio para Fortaleza passar o fim de semana com o pai e ainda não voltou para Cascavel."
        If InStr(1, Ocorrencia, "THE CHILD WAS UNWELL, CRYING A LOT, AND DID NOT LEAVE HIS MOTHER'S ARMS TO INTERACT AT ALL, DESPITE ATTEMPTS TO PLEASE HIM.", vbTextCompare) > 0 Then Cells(k, j) = "A  criança estava indisposta, chorando muito, não saia dos braços da mãe pra interagir em nada, apesar de tentativas de aggrado"
        If InStr(1, Ocorrencia, "INTERVIEWER TRYING TO CONTACT THE MOTHER, BUT SHE DOESN'T ANSWER OR RESPOND TO MESSAGES.", vbTextCompare) > 0 Then Cells(k, j) = "Entrevistadora tentando contato com a mãe, mas não atende nem responde mensagem."
        If InStr(1, Ocorrencia, "NOT PART OF GSED", vbTextCompare) > 0 Then Cells(k, j) = "NÃO FAZ PARTE DO GSED"
        
    Next k

Next j
    
    
End Sub
Sub ControleOcorrGeedIngles()

Planilha9.Select
Range("D1") = "GENERAL CONTROL BY CONTACT"
Range("G3") = "SUMMARY OF DISPOSITIONS AND ACTIONS - GSED   "
Range("I3") = "DISPOSITIONS PER CONTACT - GSED"

Range("A4") = "ID_IPEC"
Range("B4") = "CA2 - MUNICIPALITY"
Range("C4") = "CA2 - MUNICIPALITY_2"
Range("D4") = "CA3 - FAMILY ID"
Range("E4") = "ID_CHILD"
Range("F4") = "GSED  Yes, priority, Yes, non-priority, Not GSED"
Range("G4") = "TOTAL NUMBER OF CONTACTS MADE"
Range("H4") = "STATUS OF THE LAST DISPOSITION - CATI"
Range("I4") = "DISPOSITION 1"
    Range("I4").AutoFill Destination:=Range("I4:CJ4"), Type:=xlFillDefault
    Range("I4:CJ4").Select
    Range("C4").Select

For i = 5 To 1603
ColunaGSED = Cells(i, 6).Value
ColunaOcorr = Cells(i, 8).Value
    If InStr(1, ColunaGSED, "GSED SIM - PRIORITÁRIO", vbTextCompare) > 0 Then Cells(i, 6) = "GSED YES PRIORITY"
    If InStr(1, ColunaGSED, "GSED SIM - NÃO PRIORITÁRIOS", vbTextCompare) > 0 Then Cells(i, 6) = "YES NON PRIORITY"
    If InStr(1, ColunaGSED, "NÃO É GSED (INELEGÍVEL)", vbTextCompare) > 0 Then Cells(i, 6) = "NOT GSED"
    If InStr(1, ColunaGSED, "PILOTO", vbTextCompare) > 0 Then Cells(i, 6) = "PILOT"
    
    If InStr(1, ColunaOcorr, "Em confirmação para agendamento", vbTextCompare) > 0 Then Cells(i, 8) = "IN CONFIRMATION FOR SCHEDULING"
    If InStr(1, ColunaOcorr, "TESTE REALIZADO", vbTextCompare) > 0 Then Cells(i, 8) = "TEST CARRIED OUT"
    If InStr(1, ColunaOcorr, "COTA FECHADA", vbTextCompare) > 0 Then Cells(i, 8) = "CLOSED QUOTA"
    If InStr(1, ColunaOcorr, "Viajando", vbTextCompare) > 0 Then Cells(i, 8) = "TRAVELING"
    If InStr(1, ColunaOcorr, "Confirmada para", vbTextCompare) > 0 Then Cells(i, 8) = "CONFIRMED FOR"
    If InStr(1, ColunaOcorr, "RECUSA", vbTextCompare) > 0 Then Cells(i, 8) = "REFUSAL"
    If InStr(1, ColunaOcorr, "Inelegível", vbTextCompare) > 0 Then Cells(i, 8) = "INELIGIBLE"
    If InStr(1, ColunaOcorr, "MORA EM OUTRA CIDADE FORA DA AMOSTRA DO PROJETO", vbTextCompare) > 0 Then Cells(i, 8) = "LIVES IN ANOTHER CITY OUTSIDE THE PROJECT SAMPLE"
    If InStr(1, ColunaOcorr, "Confirmada para", vbTextCompare) > 0 Then Cells(i, 8) = "CONFIRMED FOR"
    If InStr(1, ColunaOcorr, "RETORNO - REAGENDADA (INFORME MOTIVO EM OBSERVAÇÕES)", vbTextCompare) > 0 Then Cells(i, 8) = "RETURN - RESCHEDULED (INFORM REASON IN COMMENTS)"
    If InStr(1, ColunaOcorr, "Domicilios em Outras cidades", vbTextCompare) > 0 Then Cells(i, 8) = "HOUSEHOLDS IN OTHER CITIES"
    If InStr(1, ColunaOcorr, "Segundo a mãe, veio para Fortaleza passar o fim de semana com o pai e ainda não voltou para Cascavel.", vbTextCompare) > 0 Then Cells(i, 8) = "ACCORDING TO HIS MOTHER, HE CAME TO FORTALEZA TO SPEND THE WEEKEND WITH HIS FATHER AND HAS NOT YET RETURNED TO CASCAVEL. "
    If InStr(1, ColunaOcorr, "A  criança estava indisposta, chorando muito, não saia dos braços da mãe pra interagir em nada, apesar de tentativas de aggrado", vbTextCompare) > 0 Then Cells(i, 8) = "THE CHILD WAS UNWELL, CRYING A LOT, AND DID NOT LEAVE HIS MOTHER'S ARMS TO INTERACT AT ALL, DESPITE ATTEMPTS TO PLEASE HIM."
    If InStr(1, ColunaOcorr, "Entrevistadora tentando contato com a mãe, mas não atende nem responde mensagem.", vbTextCompare) > 0 Then Cells(i, 8) = "INTERVIEWER TRYING TO CONTACT THE MOTHER, BUT SHE DOESN'T ANSWER OR RESPOND TO MESSAGES."
    If InStr(1, ColunaOcorr, "NÃO FAZ PARTE DO GSED", vbTextCompare) > 0 Then Cells(i, 8) = "NOT PART OF GSED"
    
Next i

For j = 9 To 88
    For k = 5 To 1603
    Ocorrencia = Cells(k, j).Value
        If InStr(1, Ocorrencia, "Em confirmação para agendamento", vbTextCompare) > 0 Then Cells(k, j) = "IN CONFIRMATION FOR SCHEDULING"
        If InStr(1, Ocorrencia, "TESTE REALIZADO", vbTextCompare) > 0 Then Cells(k, j) = "TEST CARRIED OUT"
        If InStr(1, Ocorrencia, "COTA FECHADA", vbTextCompare) > 0 Then Cells(k, j) = "CLOSED QUOTA"
        If InStr(1, Ocorrencia, "Viajando", vbTextCompare) > 0 Then Cells(k, j) = "TRAVELING"
        If InStr(1, Ocorrencia, "Confirmada para", vbTextCompare) > 0 Then Cells(k, j) = "CONFIRMED FOR"
        If InStr(1, Ocorrencia, "RECUSA", vbTextCompare) > 0 Then Cells(k, j) = "REFUSAL"
        If InStr(1, Ocorrencia, "Inelegível", vbTextCompare) > 0 Then Cells(k, j) = "INELIGIBLE"
        If InStr(1, Ocorrencia, "MORA EM OUTRA CIDADE FORA DA AMOSTRA DO PROJETO", vbTextCompare) > 0 Then Cells(k, j) = "LIVES IN ANOTHER CITY OUTSIDE THE PROJECT SAMPLE"
        If InStr(1, Ocorrencia, "Confirmada para", vbTextCompare) > 0 Then Cells(k, j) = "CONFIRMED FOR"
        If InStr(1, Ocorrencia, "RETORNO - REAGENDADA (INFORME MOTIVO EM OBSERVAÇÕES)", vbTextCompare) > 0 Then Cells(k, j) = "RETURN - RESCHEDULED (INFORM REASON IN COMMENTS)"
        If InStr(1, Ocorrencia, "Domicilios em Outras cidades", vbTextCompare) > 0 Then Cells(k, j) = "HOUSEHOLDS IN OTHER CITIES"
        If InStr(1, Ocorrencia, "Segundo a mãe, veio para Fortaleza passar o fim de semana com o pai e ainda não voltou para Cascavel.", vbTextCompare) > 0 Then Cells(k, j) = "ACCORDING TO HIS MOTHER, HE CAME TO FORTALEZA TO SPEND THE WEEKEND WITH HIS FATHER AND HAS NOT YET RETURNED TO CASCAVEL. "
        If InStr(1, Ocorrencia, "A  criança estava indisposta, chorando muito, não saia dos braços da mãe pra interagir em nada, apesar de tentativas de aggrado", vbTextCompare) > 0 Then Cells(k, j) = "THE CHILD WAS UNWELL, CRYING A LOT, AND DID NOT LEAVE HIS MOTHER'S ARMS TO INTERACT AT ALL, DESPITE ATTEMPTS TO PLEASE HIM."
        If InStr(1, Ocorrencia, "Entrevistadora tentando contato com a mãe, mas não atende nem responde mensagem.", vbTextCompare) > 0 Then Cells(k, j) = "INTERVIEWER TRYING TO CONTACT THE MOTHER, BUT SHE DOESN'T ANSWER OR RESPOND TO MESSAGES."
        If InStr(1, Ocorrencia, "NÃO FAZ PARTE DO GSED", vbTextCompare) > 0 Then Cells(k, j) = "NOT PART OF GSED"
        
    Next k

Next j

End Sub

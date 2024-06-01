Attribute VB_Name = "X_Main"
Option Private Module

Sub SalvarControle()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
End With
Dim Caminho, Arquivo As String
Dim Usuario As Object
Dim LogonUsuario As Variant
Dim Tempoinicio As Double

Set Usuario = CreateObject("WScript.Network")
LogonUsuario = Usuario.UserName
Tempoinicio = Time

Caminho = ThisWorkbook.Path

Sheets(Array("CONTROLE_OCORRÊNCIAS_CATI", "CONTROLE_OCORRÊNCIAS_GSED", "STATUS POR CIDADE CATI E F2F", "PRODUTIVIDADE", "TELEFONES ERRADOS", "VISÃO DO CAMPO cati + GSE")).Select
Sheets("CONTROLE_OCORRÊNCIAS_CATI").Activate
Sheets(Array("CONTROLE_OCORRÊNCIAS_CATI", "CONTROLE_OCORRÊNCIAS_GSED", "STATUS POR CIDADE CATI E F2F", "PRODUTIVIDADE", "TELEFONES ERRADOS", "VISÃO DO CAMPO cati + GSE")).Copy
ActiveWorkbook.SaveAs Filename:=Caminho & "\211338_01_Controle_Banco_Mundial_AFINI_ENDLINE_" & Format(Now, "dd-mm-yyyy") & "_" & Format(Now, "hh-mm-ss") & ".xlsx"
ActiveWindow.Close
Planilha8.Select
Cells(1, 1).Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
End With

MsgBox "Prezado(a): " & LogonUsuario & vbCrLf & _
">> Controle Banco Mundial AFINI ENDLINE em - (" & Time - Tempoinicio & ")<<" & vbCrLf & _
"" & vbcrlt & _
"- O Controle está salvo no caminho: " & vbCrLf & _
"" & vbcrlt & _
"-" & Caminho & vbCrLf & _
"" & vbCrLf & _
"- Obrigado!!!", , "Banco Mundial AFINI"
Exit Sub
End Sub
Sub VisualizarPlanilha()
Planilha4.Visible = True
End Sub

Sub OcorrCatPortugues_onAction(control As IRibbonControl)
Call GeraControlePortugues

End Sub

'Callback for Button15 onAction
Sub OcorrCatIngles_onAction(control As IRibbonControl)
Call GeraControleIngles

End Sub

'Callback for Button23 onAction
Sub OcorrGSEDPortugues_onAction(control As IRibbonControl)
Call ControleOcorrGsedPortugues
End Sub

'Callback for Button24 onAction
Sub OcorrGSEDIngles_onAction(control As IRibbonControl)

Call ControleOcorrGsedingles
End Sub

'Callback for Button44 onAction
Sub VisaoCampoCatiPortugues_onAction(control As IRibbonControl)
Call VisaoGeral
End Sub

'Callback for Button45 onAction
Sub VisaoCampoCatiIngles_onAction(control As IRibbonControl)
End Sub

'Callback for Button33 onAction
Sub StatusCidadePortugues_onAction(control As IRibbonControl)

Call StatusPorCidadePortugues
End Sub

'Callback for Button34 onAction
Sub StatusCidadeIngles_onAction(control As IRibbonControl)
Call StatusPorCidadeIngles
End Sub

'Callback for Button53 onAction
Sub ProdutividadePortugues_onAction(control As IRibbonControl)

Call ProdutividPortugues
End Sub

'Callback for Button54 onAction
Sub ProdutividadeIngles_onAction(control As IRibbonControl)
Call ProdutividIngles
End Sub

'Callback for Button63 onAction
Sub TelefonesErrados_onAction(control As IRibbonControl)
Call SeparaTelefones01
End Sub

'Callback for Button73 onAction
Sub Exportar_onAction(control As IRibbonControl)

Call SalvarControle
End Sub



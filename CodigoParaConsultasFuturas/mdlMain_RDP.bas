Attribute VB_Name = "mdlMain_RDP"

Private Sub criarRdp(ByVal control As IRibbonControl)
Dim ws As Worksheet: Set ws = Worksheets("vms")
Dim strDominio As String: strDominio = Etiqueta("ServerDominio")
Dim obj As New clsRdp
Dim lRow As Long, x As Long
Dim strPath As String
Dim resposta As Variant


If (ActiveSheet.name <> ws.name) Then
    ws.Visible = xlSheetVisible
    ws.Activate
Else
    
    resposta = MsgBox("Deseja criar os arquivo rdp's selecionados ?", vbQuestion + vbYesNo, "Criar arquivo rdp.")
            
    If (resposta = vbYes) Then
        
        ''find  first empty row in database
        lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
        
        For x = 2 To lRow - 1
                
            With obj
                If (ws.Range("G" & x).value <> "") Then
                    '' Dados para arqivo
                    .strAddress = CStr(ws.Range("B" & x).value)
                    .strUsername = strDominio & "\" & CStr(ws.Range("C" & x).value)
                    .strUserpass = CStr(ws.Range("D" & x).value)
                    .strPath = CStr(ws.Range("F" & x).value)
                    .strRun = CStr(ws.Range("H" & x).value)
                    .gerarRdp
                    .gerarCredencial
                    
                    '' Copia de senha
                    ClipBoardThis CStr(ws.Range("D" & x).value)
                    
                    '' Limpar marcacao
                    ws.Range("G" & x).value = ""
                    ws.Range("H" & x).value = ""
                    
                End If
            End With
            
        Next x
        
'        MsgBox "Concluido!", vbInformation + vbOKOnly, "Criar RDP"
        
        Set obj = Nothing
    
    End If
    
End If

End Sub

Function Etiqueta(sEtiqueta As String) As String
On Error Resume Next

Etiqueta = Replace(Replace(ThisWorkbook.Names(sEtiqueta), "=", ""), Chr(34), "")

If err.Number <> 0 Then Etiqueta = "#N/A"
On Error GoTo 0
End Function

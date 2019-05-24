Attribute VB_Name = "mdlMain_RDP"

Sub novoCaminhoPadrao(ByVal control As IRibbonControl)

Dim sCaminho As String

Dim sTitle As String:       sTitle = "Caminho padrão"
Dim sMessage As String:     sMessage = "Deseja alterar o camimho padrão onde ficará salvas as VMs ?"
Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)

            
    If (resposta = vbYes) Then
    
            Select Case StrPtr(resposta)
            
                Case 0
                     MsgBox "Atualização cancelada.", 64, sTitle
                    Exit Sub
                    
                Case Else
                     sCaminho = GetFolder()
                     If (sCaminho <> "") Then
                        ThisWorkbook.Names("caminhoRdp").value = sCaminho
                        MsgBox "Caminho atualizado para : " & (sCaminho) & ".", 64, sTitle
                     Else
                        MsgBox "Operação cancelada", vbInformation, sTitle
                     End If
                     
            End Select

    End If
    
End Sub

Sub importarBaseDeDados(ByVal control As IRibbonControl)

Application.ScreenUpdating = False

Dim strSenha As String: strSenha = Etiqueta("SenhaPadrao")
Dim wsDest As Worksheet: Set wsDest = Worksheets("vms")
Dim wsCopy As Workbook

Dim Sheet As Worksheet

Dim lCopyLastRow As Long
Dim lDestLastRow As Long


Dim sTitle As String:       sTitle = "Importar base de VMs"
Dim sMessage As String:     sMessage = "Deseja importar base de VMs ?"
Dim resposta As Variant:    resposta = MsgBox(sMessage, vbQuestion + vbYesNo, sTitle)

            
If (resposta = vbYes) Then

        Select Case StrPtr(resposta)
        
            Case 0
                 MsgBox "Atualização cancelada.", 64, sTitle
                Exit Sub
                
            Case Else
                FileToOpen = Application.GetOpenFilename _
                            (Title:="Por favor selecione a planilha para importação de dados", _
                            FileFilter:="Report Files *.xls* (*.xls*),")
                
                If FileToOpen = False Then
                    MsgBox "Nenhum arquivo selecionado.", vbExclamation, "ERROR - Importação de dados"
                    Exit Sub
                Else
                
                Set wsCopy = Workbooks.Open(Filename:=FileToOpen)
                
                    For Each Sheet In wsCopy.Sheets
                
                          '1. Find last used row in the copy range based on data in column A
                          lCopyLastRow = Sheet.Cells(Sheet.Rows.Count, "A").End(xlUp).Row
                            
                          '2. Find first blank row in the destination range based on data in column A
                          'Offset property moves down 1 row
                          lDestLastRow = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).Offset(1).Row
                        
                          '3. Copy & Paste Data
                          Sheet.Range("A2:D" & lCopyLastRow).Copy wsDest.Range("B" & lDestLastRow)
                          wsDest.Range("A" & lDestLastRow).value = Sheet.name
                
                    Next Sheet
                
                End If
                    
                wsCopy.Close
                
        End Select

End If


Application.ScreenUpdating = True
    
End Sub

Sub criarRdpPorSelecao(ByVal control As IRibbonControl)

Dim ws As Worksheet: Set ws = Worksheets("vms")
Dim strDominio As String: strDominio = Etiqueta("ServerDominio")
Dim strCaminho As String: strCaminho = Etiqueta("caminhoRdp")
Dim obj As New clsRdp
Dim lRow As Long, x As Long
Dim strPath As String
Dim resposta As Variant

Dim cel As Range
Dim selectedRange As Range


If (ActiveSheet.name <> ws.name) Then
    ws.Visible = xlSheetVisible
    ws.Activate
Else
    
    resposta = MsgBox("Deseja abrir a(s) VMs selecionado(s) ?", vbQuestion + vbYesNo, "Abrir VMs (Selecionadas) ")
            
    If (resposta = vbYes) Then
        
        Set selectedRange = Application.Selection
        
        
        ''find  first empty row in database
        lRow = ws.Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Row
        
        For Each cel In selectedRange.Cells
            For x = 2 To lRow - 1
                    
                With obj
                    If (ws.Range("B" & x).value = cel.value) Then
                        '' Dados para arqivo
                        .strAddress = CStr(ws.Range("B" & x).value)
                        .strUsername = strDominio & "\" & CStr(ws.Range("C" & x).value)
                        .strUserpass = Trim(CStr(ws.Range("D" & x).value))
                        .strPath = IIf(CStr(ws.Range("F" & x).value) = "", strCaminho, CStr(ws.Range("F" & x).value))
                        .strRun = CStr(ws.Range("H" & x).value)
                        .gerarRdp
                        .gerarCredencial
                        
                        '' Copia de senha
                        ClipBoardThis Trim(CStr(ws.Range("D" & x).value))
                        
                    End If
                End With
                
            Next x
    
        Next cel
        
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

Function GetFolder() As String

    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Favor selecionar um novo camimho padrão."
        .AllowMultiSelect = False
        .InitialFileName = CreateObject("WScript.Shell").SpecialFolders("Desktop") ' Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Sub teste_LoopSelection()

    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        Debug.Print cel.address, cel.value
    Next cel

End Sub


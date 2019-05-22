VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRdp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private sAddress As String
Private sUsername As String
Private sUserpass As String
Private sPath As String
Private sRun As String

Property Get strAddress() As String:                        strAddress = sAddress:          End Property
Property Let strAddress(pValue As String):                  sAddress = pValue:              End Property

Property Get strUsername() As String:                       strUsername = sUsername:        End Property
Property Let strUsername(pValue As String):                 sUsername = pValue:             End Property

Property Get strUserpass() As String:                       strUserpass = sUserpass:        End Property
Property Let strUserpass(pValue As String):                 sUserpass = pValue:             End Property

Property Get strPath() As String:                           strPath = sPath:                End Property
Property Let strPath(pValue As String):                     sPath = pValue:                 End Property

Property Get strRun() As String:                           strRun = sRun:                   End Property
Property Let strRun(pValue As String):                     sRun = pValue:                   End Property

''--[Create, list or delete stored user names, passwords or credentials.]
'' https://ss64.com/nt/cmdkey.html
'' https://docs.microsoft.com/pt-br/windows-server/administration/windows-commands/mstsc
'' https://docs.microsoft.com/pt-br/windows-server/administration/windows-commands/remote-desktop-services-terminal-services-command-reference


Function gerarCredencial()
Dim col As New Collection

col.Add "cmdkey /delete TERMSRV/" & strAddress
col.Add "cmdkey /generic:TERMSRV/" & strAddress & " /user:" & strUsername & " /pass:""" & strUserpass & """"

'' Exclusão
If ((strPath) = "") Then strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
If (Dir(strPath & "\" & strAddress & ".bat") <> "") Then Kill strPath & "\" & strAddress & ".bat"

For Each c In col
    Saida strPath, strAddress & ".bat", CStr(c)
Next c

'' Adicionais
Saida strPath, strAddress & ".bat", "TIMEOUT 3"
Saida strPath, strAddress & ".bat", "DEL /Q /F " & strPath & "\" & strAddress & ".bat"

'' Pause
Sleep 500

'' Auto run
Shell strPath & "\" & strAddress & ".bat"

'' Pause
Sleep 500

'' Execução
If (strRun <> "") Then Shell "mstsc " & strPath & "\" & strAddress & ".rdp"

End Function

Function gerarRdp()
Dim colNews As New Collection
Dim c As Variant

colNews.Add "screen mode id:i:2"
colNews.Add "use multimon:i:0"
colNews.Add "desktopwidth:i:1366"
colNews.Add "desktopheight:i:768"
colNews.Add "session bpp:i:32"
colNews.Add "winposstr:s:0,1,1920,0,3286,728"
colNews.Add "compression:i:1"
colNews.Add "keyboardhook:i:2"
colNews.Add "audiocapturemode:i:0"
colNews.Add "videoplaybackmode:i:1"
colNews.Add "connection type:i:7"
colNews.Add "networkautodetect:i:1"
colNews.Add "bandwidthautodetect:i:1"
colNews.Add "displayconnectionbar:i:1"
colNews.Add "enableworkspacereconnect:i:0"
colNews.Add "disable wallpaper:i:0"
colNews.Add "allow font smoothing:i:0"
colNews.Add "allow desktop composition:i:0"
colNews.Add "disable full window drag:i:1"
colNews.Add "disable menu anims:i:1"
colNews.Add "disable themes:i:0"
colNews.Add "disable cursor setting:i:0"
colNews.Add "bitmapcachepersistenable:i:1"
colNews.Add "full address:s:" & strAddress
colNews.Add "audiomode:i:0"
colNews.Add "redirectprinters:i:1"
colNews.Add "redirectcomports:i:0"
colNews.Add "redirectsmartcards:i:1"
colNews.Add "redirectclipboard:i:1"
colNews.Add "redirectposdevices:i:0"
colNews.Add "autoreconnection enabled:i:1"
colNews.Add "authentication level:i:2"
colNews.Add "prompt for credentials:i:0"
colNews.Add "negotiate security layer:i:1"
colNews.Add "remoteapplicationmode:i:0"
colNews.Add "alternate shell:s:"
colNews.Add "shell working directory:s:"
colNews.Add "gatewayhostname:s:"
colNews.Add "gatewayusagemethod:i:4"
colNews.Add "gatewaycredentialssource:i:4"
colNews.Add "gatewayprofileusagemethod:i:0"
colNews.Add "promptcredentialonce:i:0"
colNews.Add "gatewaybrokeringtype:i:0"
colNews.Add "use redirection server name:i:0"
colNews.Add "rdgiskdcproxy:i:0"
colNews.Add "kdcproxyname:s:"
colNews.Add "drivestoredirect:s:"
colNews.Add "smart sizing:i:1"
colNews.Add "username:s:" & strUsername


'' Exclusão
If ((strPath) = "") Then strPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
If (Dir(strPath & "\" & strAddress & ".rdp") <> "") Then Kill strPath & "\" & strAddress & ".rdp"

For Each c In colNews
    Saida strPath, strAddress & ".rdp", CStr(c)
Next c

End Function



Private Function Saida(strCaminho As String, strArquivo As String, strConteudo As String)
    
'    If ((strCaminho) = "") Then strCaminho = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Open strCaminho & "\" & strArquivo For Append As #1
    Print #1, strConteudo
    Close #1
End Function

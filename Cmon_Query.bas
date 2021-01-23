Attribute VB_Name = "Cmon_Query"
Public Enum RequestType
    NoAuth = 0
    OAuthNeeded = 1
    REST = 3
End Enum

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long
                            


Public Sub HttpOpenLink(strUrl As String)
On Error GoTo wellsrLaunchError
    Dim R As Long
    R = ShellExecute(0, "open", strUrl, 0, 0, 1)
    If R = 5 Then 'if access denied, try this alternative
            R = ShellExecute(0, "open", "rundll32.exe", "url.dll,FileProtocolHandler " & strUrl, 0, 1)
    End If
    Exit Sub
wellsrLaunchError:
MsgBox "Error encountered while trying to launch URL." & vbNewLine & vbNewLine & "Error: " & Err.Number & ", " & Err.Description, vbCritical, "Error Encountered"
End Sub

Function HttpRequest(url As String, sType As String, RequestType As RequestType, Optional ByVal arguments As String = "")

 Dim http As MSXML2.ServerXMLHTTP60
 Set http = New MSXML2.ServerXMLHTTP60
 'filter empty namespace
 arguments = Replace(arguments, " xmlns=""""", "")
 
 On Error Resume Next
 
  http.Open sType, url, False
  If (RequestType = REST) Then
    http.setRequestHeader "ApiToken", Settings.ClientSecret
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  Else
     http.setRequestHeader "Content-Type", "text/xml"
  End If
  http.Send arguments
  
  If Err.Number <> 0 Then
    LogItem "[HttpGET] " & " unable to reach " & url
    LogItem "[HttpGET] (" & Err.Number & ") :" & Err.Description
  Err.Clear
  End If
  HttpRequest = http.responseText
 Set http = Nothing

 End Function
Function HttpGET(url As String)

    HttpGET = HttpRequest(url, "GET", REST)
    
End Function
Function HttpPOST(url As String, ByVal arguments)
    HttpPOST = HttpRequest(url, "POST", "REST", arguments)
End Function
Function HttpPut(url As String, ByVal arguments)
    HttpPut = HttpRequest(url, "PUT", "REST", arguments)
End Function
Function HttpPOSTXto(url As String, ByVal arguments)
    HttpPOSTXto = HttpRequest(url, "OAuthNeeded", REST, arguments)
End Function
Attribute VB_Name = "Cmon_Logs"


Public Sub DebugLine(ByVal message)
  If DEBUGMODE = "ON" Then
    Debug.Print "@" & Now & ":" & message
  End If
End Sub

Public Sub LogItem(ByVal strItem)

'if happens during initialisation only on debug line
If Settings Is Nothing Then
    Debug.Print "@" & Now & ":" & strItem
    Exit Sub
    
End If

'return on normal process

Dim sFilename
Dim LogFile
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")


    sFilename = "LOGS[" & Year(Now) & Month(Now) & Day(Now) & "].log"
    fsoCreateFolder "Logs", Settings.UserSystemFolder
          If Not fso.FileExists(Settings.UserSystemFolder & "Logs\" & sFilename) Then
            fso.CreateTextFile (Settings.UserSystemFolder & "Logs\" & sFilename)
          End If
    Set LogFile = fso.OpenTextFile(Settings.UserSystemFolder & "Logs\" & sFilename, 8, True)
        LogFile.WriteLine ("[" & UCase(Settings.userName) & "@ " & Now & strItem)
        LogFile.Close
        Debug.Print "@" & Now & ":" & strItem
    Set LogFile = Nothing


Set fso = Nothing
End Sub



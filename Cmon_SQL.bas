Attribute VB_Name = "Cmon_SQL"
Option Explicit
Public gConnection As Variant

Private Sub OpenConnection(ByVal moduleName)

     'instanciate settings
    If Settings Is Nothing Then Set Settings = New CmonSettings
    
    Set gConnection = CreateObject("ADODB.Connection")

    gConnection.ConnectionString = Settings.ConnectionString

        If Not (gConnection Is Nothing) Then
            DebugLine "[" & moduleName & "]" & gConnection.ConnectionString
            On Error Resume Next
            gConnection.Open
            If Err.Number <> 0 Then
                Application.Wait (Now + TimeValue("0:00:02"))
                LogItem "[" & moduleName & "]" & " Error when trying to open a DB connection :" & Err.Description
                MsgBox Err.Description & vbCrLf & "Please check the database node settings in " & vbCrLf & Settings.UserSystemFolder & ".", 64, "Database connection"
                Err.Clear
                
                        
        'open folder to correct connections sting
                fsoOpenExplorer Settings.UserSystemFolder
                Set gConnection = Nothing
            End If
        End If
        
 End Sub
  
 Public Function HasResults(QueryString As String, Purpose As String, ByRef records) As Boolean

  Dim rs As Variant
   OpenConnection (Purpose)
   HasResults = False
      If Not (gConnection Is Nothing) Then

          Set rs = CreateObject("ADODB.Recordset")
          'extract values in a array
          
            rs.Open QueryString, gConnection
            records = rs.GetRows()
            rs.Close
            
          'check if there are any results
             If UBound(records, 2) > -1 Then
                HasResults = True
                
            Else
                 LogItem "[" & Purpose & "]" & " No records found in query"
                 LogItem "[" & Purpose & "]" & " QueryString"
            End If

      End If

    Set rs = Nothing
    Set gConnection = Nothing
End Function

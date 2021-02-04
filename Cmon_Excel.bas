Attribute VB_Name = "Cmon_Excel"
Public Function ReadCustomProperties(strPropertyName As String, valueIfEmpty As Variant, _
                                    docType As Office.MsoDocProperties) As Variant

    ReadCustomProperties = valueIfEmpty
    On Error Resume Next
    ReadCustomProperties = ActiveWorkbook.CustomDocumentProperties(strPropertyName).Value
    If Err.Number > 0 Then
        UpdateCustomProperties strPropertyName, valueIfEmpty, msoPropertyTypeDate
    End If
End Function

Public Function UpdateCustomProperties(strPropertyName As String, _
    Value As Variant, docType As Office.MsoDocProperties) As Variant

    Dim oCustomProperty As DocumentProperty
    On Error Resume Next
    Set oCustomProperty = ActiveWorkbook.CustomDocumentProperties(strPropertyName)
    If oCustomProperty Is Nothing Then
        ActiveWorkbook.CustomDocumentProperties.Add _
            Name:=strPropertyName, _
            LinkToContent:=False, _
            Type:=docType, _
            Value:=Value
    Else
        oCustomProperty.Value = Value
    End If
End Function

Public Sub AutomatedUpdateCheck()
Dim lastUpdateCheck
Dim lastUpdateTime

If UpdatesHasBeenChecked = True Then
 DebugLine "[UpdateChecker] Already been Checked"
 Exit Sub
End If
  
  'Check if the BVDSettings contains the LastUpdateCheck node
   lastUpdateCheck = ReadCustomProperties("LastUpdateCheck", Now, msoPropertyTypeDate)
   lastUpdateTime = CDate(lastUpdateCheck)
   'handle conversion error and refill the LastUpdateObject
         If DateDiff("d", lastUpdateTime, Now) >= 5 Then
             Call UpdateInstaller
         Else
            LogItem "[UpdateChecker] less than 5 days between two checks"
         End If

    
     UpdatesHasBeenChecked = True

End Sub

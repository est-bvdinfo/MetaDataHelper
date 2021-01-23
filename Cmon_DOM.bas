Attribute VB_Name = "Cmon_DOM"
Option Explicit


'Public Function GetNodeValue(ByRef sourceNode As IXMLDOMNode, elementName As String) As String

''Dim childNode As IXMLDOMNode
''For Each childNode In sourceNode.ChildNodes
''    If childNode.nodeName = elementName Then
 ''       GetNodeValue = childNode.text
 ''       Exit For
 ''   End If
''Next childNode
''End Function
'________________________________________________________________________________________'

Public Function IsXMLValid(ByVal XML, ByVal Label)
Dim xmlDoc
Dim ret
Set xmlDoc = CreateObject("Msxml2.DOMDocument")
Dim errorString
Dim gotoLine
xmlDoc.async = "false"
xmlDoc.LoadXML XML


If xmlDoc.parseError.ErrorCode <> 0 Then
   errorString = "Parse Error Line " & xmlDoc.parseError.line & ", character " & xmlDoc.parseError.LinePos & vbCrLf & xmlDoc.parseError.reason
   LogItem "[IsXMLValid       ] for " & Label & " " & errorString
   
   MsgBox errorString, vbCritical, "Load of " & Label & " unsuccessfull"

   ret = False
Else
   ret = True
End If
Set xmlDoc = Nothing
 IsXMLValid = ret
End Function
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
Public Function fsoCreateFolder(ByVal fol, ByVal ParentFol) As String
Dim fsfol
Set fsfol = CreateObject("Scripting.FileSystemObject")
'check if it ends with a backslash
If Trim(Right(ParentFol, 1)) <> "\" Then ParentFol = ParentFol & "\"
If Not fsfol.FolderExists(ParentFol & fol) Then
    fsfol.CreateFolder (ParentFol & fol)
    LogItem "[fsoCreateFolder] SubFolder " & fol & " from " & ParentFol & " has been created! "
End If
    fsoCreateFolder = ParentFol & fol & "\"
Set fsfol = Nothing
End Function
Public Sub fsoDeleteFile(ByVal Path)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(Path) Then
    DebugLine "[fsoDeleteFile]" & Path
    On Error Resume Next
      fso.DeleteFile (Path)
    End If
    
If Err.Number = 70 Then
    DebugLine "[fsoDeleteFile] waiting for the file to be release by the system then retry the delete operation"
    fso.DeleteFile (Path)
    Err.Clear
End If
Set fso = Nothing
End Sub

Public Sub fsoWriteFile(ByVal Content, ByVal filename, ByVal sExtention, ByVal sfolder)
Dim oStreamRecoder
Dim fso

    DebugLine "[fsoWriteFile]" & filename & "." & sExtention & "  in folder '" & sfolder & "' about to be created"
Set fso = CreateObject("Scripting.FileSystemObject")
Set oStreamRecoder = fso.CreateTextFile(sfolder & filename & "." & sExtention, True, False)
    DebugLine "[fsoWriteFile] type of content object" & TypeName(Content)
    
    If TypeName(Content) <> "Null" Then oStreamRecoder.Write (Content)
         LogItem filename & "." & sExtention & " created in folder '" & sfolder & "'"
    Set oStreamRecoder = Nothing
    Set fso = Nothing
 
End Sub

Public Function fsoReadToLog() As String
Dim fsoDef
Dim tsDef
Dim sLine As String
Dim outputPath:
outputPath = Settings.UserSystemFolder & "CmdOuput.dat"
   Set fsoDef = CreateObject("Scripting.FileSystemObject")
   If fsoDef.FileExists(outputPath) Then
       Set tsDef = fsoDef.OpenTextFile(outputPath)
        Do While tsDef.AtEndOfStream <> True
            sLine = tsDef.ReadLine
            If sLine <> "" Then
                LogItem "[fsoReadToLog]" & sLine
                fsoReadToLog = fsoReadToLog & sLine & vbCrLf
            End If
        Loop
       tsDef.Close
   Else
        
        fsoReadToLog = "[fsoReadToLog] output file doesn't exists (" & outputPath & ")"
        LogItem fsoReadToLog
   End If
End Function

Public Sub fsoOpenExplorer(ByVal SPath)
Dim fsfol
Dim SH, txtFolderToOpen
Set SH = CreateObject("Shell.Application")
Set fsfol = CreateObject("Scripting.FileSystemObject")
If Not fsfol.FolderExists(SPath) Then
    LogItem "Folder [" & SPath & "] doesn't exist"
Else
    LogItem "[fsoOpenExplorer   ] " & SPath
    SH.Explore SPath
End If
Set SH = Nothing
Set fsfol = Nothing
End Sub

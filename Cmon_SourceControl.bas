Attribute VB_Name = "Cmon_SourceControl"
Option Explicit
Public UpdatesHasBeenChecked As Boolean
Public ToUpGradeVersion As String
Public Enum updateStatuses
    Uptodate
    ToUpdate
    RssNotReached
End Enum

Private Sub ExportModules(sourceFolderToBeDisplayed As Boolean)
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent   'Microsoft Visual Basic for Applications Extensibility 5.3 library

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or delete all files in the folder if it exist.
    If Settings.CurrentProjectFolder = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill Settings.CurrentProjectFolder & "*.cls"
        Kill Settings.CurrentProjectFolder & "*.frm"
        Kill Settings.CurrentProjectFolder & "*.bas"
        Kill Settings.CurrentProjectFolder & "*.frx"
        

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = Settings.CurrentProjectFolder
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            LogItem "[ExportModules] exporting " & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    If sourceFolderToBeDisplayed = True Then
        MsgBox "Modules exported in" & vbCrLf & szExportPath, , "Export is ready"
    Else
        LogItem "[ExportModules] Modules exported in" & vbCrLf & szExportPath
    End If
    
End Sub


Private Sub ImportModules()
    Dim wkbTarget As Excel.Workbook
    Dim objFSO As FileSystemObject
    Dim objFile As Scripting.File
    Dim szTargetWorkbook As String
    Dim szImportPath As String
    Dim i%, sName$
    Dim listOfModules As Dictionary
    Dim moduleName, cmpComponents


   ' If ActiveWorkbook.Name = ThisWorkbook.Name Then
   '     MsgBox "Select another destination workbook" & _
   '     "Not possible to import in this workbook "
   '     Exit Sub
   ' End If

    'Get the path to the folder with modules
    If Settings.CurrentProjectFolder = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If

    ''' NOTE: This workbook must be open in Excel.
    szTargetWorkbook = ActiveWorkbook.Name
    Set wkbTarget = Application.Workbooks(szTargetWorkbook)
    
    If wkbTarget.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

    ''' NOTE: Path where the code modules are located.
    szImportPath = Settings.CurrentProjectFolder
        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(szImportPath).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    '''1. load all the code modules in a collection
    Set listOfModules = New Dictionary
    For Each objFile In objFSO.GetFolder(szImportPath).Files
    
       If LCase(objFile.Name) <> "sourcecontrol.bas" Then
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            'check file does not exist yet
            If Not listOfModules.Exists(objFSO.GetBaseName(objFile.Name)) Then listOfModules.Add objFSO.GetBaseName(objFile.Name), objFile.Path
        End If
       End If
        
    Next objFile
    
    ' 2. Iterate all components and attempt to import their source from the network share
    With ThisWorkbook.VBProject
        '2.1 Process backwords as we are working through a live array while removing/adding items
        For i% = .VBComponents.Count To 1 Step -1
            ' Extract this component name
            sName$ = .VBComponents(i%).CodeModule.Name
            ' Do not change the source of this module which is currently running
            If LCase(sName$) <> "sourcecontrol" Then
                ' Import relevant source file if it exists
                If .VBComponents(i%).Type = 1 Then
                    ' Standard Module
                    .VBComponents.Remove .VBComponents(sName$)
                    .VBComponents.Import filename:=szImportPath & sName$ & ".bas"
                    If listOfModules.Exists(sName$) Then listOfModules.Remove sName$
                ElseIf .VBComponents(i%).Type = 2 Then
                    ' Class
                    .VBComponents.Remove .VBComponents(sName$)
                    .VBComponents.Import filename:=szImportPath & sName$ & ".cls"
                    If listOfModules.Exists(sName$) Then listOfModules.Remove sName$
                ElseIf .VBComponents(i%).Type = 3 Then
                    ' Form
                    .VBComponents.Remove .VBComponents(sName$)
                    .VBComponents.Import filename:=szImportPath & sName$ & ".frm"
                    If listOfModules.Exists(sName$) Then listOfModules.Remove sName$
                Else
                    ' UNHANDLED/UNKNOWN COMPONENT TYPE
                End If
            End If
        Next i
         '2.2 all the existing files have been replaced. Now add the new ones
         For Each moduleName In listOfModules.Keys()
            .VBComponents.Import filename:=listOfModules(moduleName)
         Next
        
        End With
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    

    
       MsgBox "Modules imported from" & szImportPath, , "importmade"
End Sub

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


Public Sub UpdateInstaller()
 Dim http
 Dim SH
 Dim oZip
 Dim zipFileName, downloadLink
  Dim fso As FileSystemObject
 Dim oFolderItem, subFolder, objFolder
 Dim Status As updateStatuses
 Dim currentUpdateFolder As String
 
 'instanciate settings
 If Settings Is Nothing Then Set Settings = New CmonSettings
 
 
 'check if update is required
 Status = CheckNewChangeset
 If Status = RssNotReached Then
     MsgBox "RSS Feed seems not to be accessible anymore" _
     & vbCrLf & "Please check with EST if a manual update is not required or your firewall is not blocking " & MODULE_NAME _
     , 48, "Check for update"
     Exit Sub
 ElseIf Status = Uptodate Then
    MsgBox MODULE_NAME & " is up to date with version " & MODULE_VERSION, vbInformation, "Check for update"
    Exit Sub
 ElseIf Status = ToUpdate Then
     If (MsgBox(MODULE_NAME & " has to be updated to version " & ToUpGradeVersion _
         & vbCrLf & "Do you agree to proceed?", vbQuestion + vbYesNo, "Check for update") = vbNo) Then
        Exit Sub
     End If
End If
 

 downloadLink = REPOSITORY & "/get/default.zip"
 zipFileName = "Update" & ToUpGradeVersion & ".zip"
 currentUpdateFolder = fsoCreateFolder("Updates", Settings.UserSystemFolder)


 On Error Resume Next

  Set http = CreateObject("Microsoft.XMLHTTP")
  Set fso = CreateObject("Scripting.FileSystemObject")
  http.Open "GET", downloadLink, False
  http.Send
  'in case of download error stop the download
  If Err.Number <> 0 Then
    LogItem "[UpdateDownloadAndExtract] " & " unable to reach " & downloadLink
    LogItem "[UpdateDownloadAndExtract] (" & Err.Number & ") :" & Err.Description
    Exit Sub
    Err.Clear
  Else
    'when download ran smoothly proceed with the download and the zip file save

    'Creating and filling binaries base on the received zip
   '! Not using fsoWriteFile since it's binary target
    Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")
    With bStrm
        .Type = 1 '//binary
        .Open
        .Write http.responseBody
        .savetofile currentUpdateFolder & zipFileName, 2 '//overwrite
    End With
    LogItem "[UpdateDownloadAndExtract] " & zipFileName & " downloaded"
    
    'open the Zip file and search for the pspad root folder under 'est_bvdinfo_persidhandler.....'
    Set SH = CreateObject("Shell.Application")
    Set oZip = SH.Namespace(Settings.UserSystemFolder & zipFileName)
    
    'loop for each FolderItem in the zip
    For Each oFolderItem In oZip.Items
        DebugLine "ExtractFilesFromZip\" & oFolderItem.Name
        'find the PspadRoot Folder
        If InStr(oFolderItem.Name, MODULE_OWNER) > 0 And oFolderItem.IsFolder Then
            'Once the root folder has been found convert FolderItem into a proper Folder Object
            Set subFolder = SH.Namespace(fso.GetAbsolutePathName(Settings.UserSystemFolder & zipFileName & "\" & oFolderItem.Name))
            'create a recipient folder for the zip files to be extracted
            Set objFolder = SH.Namespace(Settings.CurrentProjectFolder)
            'copy zip files to SysFol without progress bar
            LogItem "[UpdateDownloadAndExtract] upgrade to version:" & ToUpGradeVersion & " as soon as the user allows it"
            On Error Resume Next
                objFolder.CopyHere subFolder.Items, 4
              If Err.Number <> 0 Then
                LogItem "[UpdateDownloadAndExtract]  ERROR ! Update install failed !"
                LogItem "[UpdateDownloadAndExtract] (" & Err.Number & ") :" & Err.Description
                Exit For
                Err.Clear
              Else
              'transfer all the copied files into xlsm
               Call ImportModules
              End If
              'need to keep this exit for otherwise all the files in the zip
              'will be copied over and over
            Exit For
        End If
    Next
  End If
 
 Set subFolder = Nothing
 Set objFolder = Nothing
 Set fso = Nothing
 Set oFolderItem = Nothing
 Set bStrm = Nothing
 Set SH = Nothing
 Set http = Nothing

End Sub

Private Function CheckNewChangeset() As updateStatuses
Dim responseText
Dim xmlDoc
Dim oChannel
Dim oProperties
Dim oProperty
Dim currentTitle
Dim externVersions
Dim externVersion
Dim internVersion
Dim separator
Dim versionLine
Dim SH

CheckNewChangeset = Uptodate
ToUpGradeVersion = MODULE_VERSION

responseText = HttpGET(REPOSITORY_RSS)
Set xmlDoc = CreateObject("Msxml2.DOMDocument")
xmlDoc.async = "false"
xmlDoc.LoadXML responseText

If xmlDoc.parseError.ErrorCode = 0 Then
  Set oChannel = xmlDoc.SelectSingleNode("/rss/channel")
    If Not (oChannel Is Nothing) Then
    Set oProperties = oChannel.ChildNodes
        For Each oProperty In oProperties
            If (oProperty.nodeName = "item") Then
                currentTitle = oProperty.FirstChild.text
                separator = InStr(currentTitle, ":")
                DebugLine currentTitle & " : " & separator
                If (separator > 0) Then
                    versionLine = Trim(Left(currentTitle, separator - 1))
                    'if finishes with a exclamation mark then skip because not a publishabe version
                    If Right(versionLine, 1) <> "!" Then
                        externVersions = Split(versionLine, ".")
                        'check if does the correct version formating or skip the formating
                        If UBound(externVersions) >= 2 Then
                            'create a comparable version
                            On Error Resume Next
                            externVersion = CInt(externVersions(0)) * 1000 + CInt(externVersions(1)) * 100 + CInt(externVersions(2))
                            'find the internal version
                            Dim internVersions: internVersions = Split(MODULE_VERSION, ".")
                            internVersion = CInt(internVersions(0)) * 1000 + CInt(internVersions(1)) * 100 + CInt(internVersions(2))
                            DebugLine "RSS Version:" & externVersion & " intern version " & internVersion
                            'check if both are numeric then proceed to check
                            If (IsNumeric(externVersion) And IsNumeric(internVersion)) Then
                                If (externVersion > internVersion) Then
                                    CheckNewChangeset = ToUpdate
                                    ToUpGradeVersion = Trim(versionLine)
                                Else
                                    LogItem "[CheckNewChangeset] your current version " & MODULE_VERSION & " is up to date"
                                End If
                                Exit For
                            End If
                        'non incremented update (merge or forking)
                        End If
                    Else
                        DebugLine "Version: " & versionLine & " .Publish made in debug mode, not taken into account"
                    
                    End If
                End If
            End If
        'not an item node
        Next
            'Continue the loop
    Else
        LogItem "[CheckNewChangeset] /rss/channel not found"
        CheckNewChangeset = RssNotReached
    End If
Else
    LogItem "[CheckNewChangeset] RSS Feed seems not to be accessible anymore!!!"
    CheckNewChangeset = RssNotReached
    
End If

End Function
Public Sub CommitToGIT()

Dim stringToExecute, retBack, branchName As String
Dim commitFeedback As String
Dim commitComment, pushFeedback As String

If Settings Is Nothing Then Set Settings = New CmonSettings
Dim fsfol: Set fsfol = CreateObject("Scripting.FileSystemObject")

'check if the .hg repository exists'
If fsfol.FileExists(Settings.CurrentProjectFolder & ".gitignore") Then
 'Export all the modules to the root folder
   Call ExportModules(False)
 
  'ask the comment to add to the commit'
    If DEBUGMODE = "ON" Then
        commitComment = InputBox("Do you want to commit current version(Debug Mode)?" & vbCrLf & "Please add a comment!", "Commit version " & MODULE_VERSION, MODULE_VERSION & "!:")
    Else
        commitComment = InputBox("Do you want to commit current version " & vbCrLf & "Please add a comment!", "Commit version " & MODULE_VERSION, MODULE_VERSION & ":")
    End If
 
  If (Len(commitComment) > 3) Then
        
      'get current branch
       ShellRun "Git.exe branch --show-current", branchName, Settings.CurrentProjectFolder
       branchName = cleanString(branchName)
       
       'stage changesets'
       ShellRun "Git.exe add .", Settings.CurrentProjectFolder
       
      'execute the commit action'
       'stringToExecute = rgExCreateCommand(" hg.exe -v commit -R " & rgAddQuote(rgRemoveSlash(Settings.CurrentProjectFolder)) & " -m " & rgAddQuote(commitComment))
        retBack = ShellRun("Git.exe commit -m " & rgAddDblQuote(commitComment), commitFeedback, Settings.CurrentProjectFolder)
       
       'check they are no error on shell level'
        If retBack = 1 Then 'success
             MsgBox "No push required!", 1, "Push on SourceControl"
             LogItem "[ComitToGIT] No push required. Code :" & retBack
             
        ElseIf (retBack > 1) Then 'failed
            LogItem "ComitToHGMercurial No commit performed.  Error in the shell execution. CODE(" & retBack & ")"
            MsgBox "Commit to Repository failed", 64, "Commit failed. Error in the shell execution. CODE(" & retBack & ")"
            
        ElseIf retBack = 0 Then
            LogItem "[ComitToGIT] Commit " & MODULE_VERSION & " performed. Displayed in SourceTree as " & commitComment
            
            stringToExecute = "Git.exe push origin " & rgAddDblQuote(branchName)
            'ask is publish is required
            If MsgBox("Publish last commit on BitBucket?" & vbCrLf & commitFeedback, 1, "Push on SourceControl") = 1 Then
                retBack = ShellRun(stringToExecute, pushFeedback, Settings.CurrentProjectFolder)
                
              'execute the publish action'
                'check they are no error on shell level'

                  If retBack <> 0 Then
                      LogItem "[ComitToGIT] Push Failed.  Error in the shell execution. CODE(" & retBack & ")"
                      MsgBox "Push to BitBucket failed", 64, "Push failed. Error in the shell execution. CODE(" & retBack & ")"
                  Else
                     MsgBox "Push to SourceControl succeeded" & vbCrLf & pushFeedback, 1, "Push on SourceControl"
                     LogItem "[ComitToGIT] Push to SourceControl succeeded  Code :" & retBack
                
                  End If
            End If
      End If
  End If
Else

    LogItem "[ComitToGIT] No commit performed. Unable to find .hg repository in " & Settings.CurrentProjectFolder & ".hg"
    MsgBox "Commit failed. Your computer is not linked to a mercurial repository. ", 64, "Commit to BitBucket failed"
End If
End Sub

Public Function GetProjectDevFolder(userFolder As String) As String

Dim xmlDoc, bookmarks, bookMarkNode
Dim fso, ts, WshShell, sPath, sProject
Dim xmlString, defaultDevFolders As String
Dim sourceTreeBookMarkFile, sourceTreeAppDataFolder As String

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

'initialised required files
sourceTreeAppDataFolder = WshShell.ExpandEnvironmentStrings("%userprofile%") & "\AppData\Local\Atlassian\SourceTree"
sourceTreeBookMarkFile = sourceTreeAppDataFolder & "\bookmarks.xml"

GetProjectDevFolder = userFolder

'check whether Altassian sourcetree is installed
If fsoFolderExists(sourceTreeAppDataFolder) Then
   
   'if found then parse sourcetree bookmark to find the developpment path
    ' get the bookmark file
    
    If fso.FileExists(sourceTreeBookMarkFile) Then
      Set xmlDoc = CreateObject("Msxml2.DOMDocument")
      Set ts = fso.OpenTextFile(sourceTreeBookMarkFile, 1)
      xmlString = ts.ReadALL
      xmlDoc.async = "false"
      xmlDoc.LoadXML xmlString
      DebugLine "[GetProjectDevFolder] upload bookmark"
      
      'check that the settings are valid
      If IsXMLValid(xmlString, USERSETTINGS) = False Then Exit Function
       
       'parse bookmarks
       For Each bookMarkNode In xmlDoc.SelectNodes("/ArrayOfTreeViewNode/TreeViewNode")
            sProject = bookMarkNode.SelectSingleNode("Name").text
            sPath = bookMarkNode.SelectSingleNode("Path").text
            DebugLine "[" & sProject & "] = [" & sPath & "]"
            
            'escape when found
            If UCase(Trim(sProject)) = UCase(Trim(MODULE_NAME)) Then
                GetProjectDevFolder = Left(sPath, InStrRev(sPath, "\"))
                LogItem "[GetProjectDevFolder] path dev found " & GetProjectDevFolder
                Exit For
            End If
                 
       Next bookMarkNode
    
    Else
        DebugLine "[GetProjectDevFolder]  can't find bookmark file"
        
    End If
    
End If

xmlString = ""
Set ts = Nothing
Set WshShell = Nothing
Set xmlDoc = Nothing
Set bookmarks = Nothing
Set bookMarkNode = Nothing

End Function

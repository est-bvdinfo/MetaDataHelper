Attribute VB_Name = "Cmon_SourceControl"
Option Explicit
Public UpdatesHasBeenChecked As Boolean
Public ToUpGradeVersion As String
Public Enum updateStatuses
    Uptodate
    ToUpdate
    GitHubNotReached
End Enum
Private targetSourceFolder As String


Private Sub ExportModules(sourceFolderToBeDisplayed As Boolean)
    Dim bExport As Boolean
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim VBAEditor As VBIDE.VBE
    Dim objProject As VBIDE.VBProject
    Dim cmpComponent As VBIDE.VBComponent   'Microsoft Visual Basic for Applications Extensibility 5.3 library
    
    Set VBAEditor = Application.VBE
    Set objProject = VBAEditor.ActiveVBProject
    
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
        
    
    If objProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
        Exit Sub
    End If
    
    szExportPath = Settings.CurrentProjectFolder
    
    For Each cmpComponent In objProject.VBComponents
        
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
    Dim objFSO As New FileSystemObject
    Dim objFile As Scripting.File
    Dim i%, sName$
    Dim listOfModules As Dictionary
    Dim moduleName, cmpComponents
    Dim VBAEditor As VBIDE.VBE
    Dim objProject As VBIDE.VBProject

    Set VBAEditor = Application.VBE
    Set objProject = VBAEditor.ActiveVBProject
 
    'need to replicate many existing objects as the module is meant to work standalone
     Dim WshShell As Object: Set WshShell = CreateObject("WScript.Shell")
      
     targetSourceFolder = WshShell.ExpandEnvironmentStrings("%APPDATA%") & "\" & MODULE_AUDIENCE
     targetSourceFolder = GetProjectsDevFolder(targetSourceFolder) & UCase(MODULE_NAME)
    'load all the code modules in a collection
    Set listOfModules = New Dictionary
    For Each objFile In objFSO.GetFolder(targetSourceFolder).Files
       If LCase(objFile.Name) <> SOURCE_CONTROLER Then
            If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                (objFSO.GetExtensionName(objFile.Name) = "bas") Then
                'check file does not exist yet
                If Not listOfModules.Exists(objFSO.GetBaseName(objFile.Name)) Then
                    listOfModules.Add objFSO.GetBaseName(objFile.Name), objFile.Path
                End If
            End If
       End If
    Next objFile
   
     'all the existing files have been replaced. Now add the new ones
     For Each moduleName In listOfModules.Keys()
        objProject.VBComponents.Import filename:=listOfModules(moduleName)
     Next
        
       'remove oldsource installer
       MsgBox listOfModules.Count & "Modules imported from" & targetSourceFolder, , "import"
       'remove oldsource installer
       
       objProject.VBComponents.Remove objProject.VBComponents(SOURCE_CONTROLER & "0")
       objProject.VBComponents.Remove objProject.VBComponents(SOURCE_CONSTANTS & "0")
    
   
   Set objFSO = Nothing
   Set objFile = Nothing
   Set listOfModules = Nothing
   
End Sub


Public Sub UpdateInstaller()
 Dim http
 Dim oZip, zipFullPath
 Dim zipFileName, downloadLink As String
 Dim oFolderItem, subFolder, objFolder
 Dim Status As updateStatuses
 Dim rootSourceFolder As String
 Dim VBAEditor As VBIDE.VBE
 Dim objProject As VBIDE.VBProject

 Set VBAEditor = Application.VBE
 Set objProject = VBAEditor.ActiveVBProject
 
 'instanciate settings
 If Settings Is Nothing Then Set Settings = New CmonSettings
 rootSourceFolder = fsoCreateFolder("Updates", Settings.UserSystemFolder)
 targetSourceFolder = Settings.CurrentProjectFolder
 
 'check if the workbook is protected otherwise not possible to update
 If objProject.Protection = 1 Then MsgBox "The VBA in this workbook is protected," & vbCrLf & "not possible to Import the code": Exit Sub
 
 'check if update is required
 Status = CheckNewChangeset
 If Status = GitHubNotReached Then
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
 

 ' proceed withj update
 On Error GoTo Skip_Import:

 downloadLink = REPOSITORY & "archive/master.zip"
 zipFileName = "Update" & ToUpGradeVersion & ".zip"
  
 Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
 Dim fso As FileSystemObject: Set fso = CreateObject("Scripting.FileSystemObject")
  
  http.Open "GET", downloadLink, False
  http.send
  'in case of download error stop the download
  If Err.Number <> 0 Then
    LogItem "[UpdateDownloadAndExtract] " & " unable to reach " & downloadLink
    LogItem "[UpdateDownloadAndExtract] (" & Err.Number & ") :" & Err.Description
    Exit Sub
    Err.Clear
  Else
    'when download ran smoothly proceed with the download and the zip file save

    'Creating and filling binaries base on the received zip
    zipFullPath = rootSourceFolder & zipFileName
    
   '! Not using fsoWriteFile since it's binary target
    Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")
    With bStrm
        .Type = 1 '//binary
        .Open
        .Write http.responseBody
        .savetofile zipFullPath, 2 '//overwrite
    End With
    LogItem "[UpdateDownloadAndExtract] " & zipFileName & " downloaded"
    
    'open the Zip file and search for the pspad root folder
    Dim SH: Set SH = CreateObject("Shell.Application")
    Set oZip = SH.Namespace((zipFullPath)) 'need to keep double parenthesis
    
    'loop for each FolderItem in the zip
    For Each oFolderItem In oZip.Items
        DebugLine "[UpdateDownloadAndExtract] ExtractFilesFromZip ==>" & oFolderItem.Name
        'find the PspadRoot Folder
        If InStr(LCase(oFolderItem.Name), LCase(MODULE_NAME)) > 0 And oFolderItem.IsFolder Then
            'Once the root folder has been found convert FolderItem into a proper Folder Object
            Set subFolder = SH.Namespace(fso.GetAbsolutePathName(zipFullPath & "\" & oFolderItem.Name))
            'create a recipient folder for the zip files to be extracted
            Set objFolder = SH.Namespace((targetSourceFolder))
            
            'check that there are files to be imported
            Dim objFSO As New FileSystemObject
            If objFSO.GetFolder(targetSourceFolder).Files.Count = 0 Then MsgBox "There are no files to import": Exit Sub
    
            'copy zip files to SysFol without progress bar
            
            MsgBox "You are about to be prompted to allow a bulk files copy" _
                    & vbCrLf & "Please select the Copy & Replace option" _
                    & vbCrLf & "and apply it to all files" & vbCrLf & MODULE_OWNER
            
            'remove existing modules
             Dim i%, sName$
              With objProject
                '2.1 remove all the code except sourcecontroler
                For i% = .VBComponents.Count To 1 Step -1
                    ' Extract this component name
                    sName$ = .VBComponents(i%).CodeModule.Name
                    ' Do not change the source of this module which is currently running
                    If LCase(sName$) <> LCase(SOURCE_CONTROLER) And LCase(sName$) <> LCase(SOURCE_CONSTANTS) And (.VBComponents(i%).Type = vbext_ct_ClassModule Or _
                        .VBComponents(i%).Type = vbext_ct_MSForm Or .VBComponents(i%).Type = vbext_ct_StdModule) Then
                        ' Standard Module
                         .VBComponents.Remove .VBComponents(i%)
                    End If
                Next i
                              
              'versioning of source installer
                    .VBComponents(SOURCE_CONTROLER).Name = SOURCE_CONTROLER & "0"
                    .VBComponents(SOURCE_CONSTANTS).Name = SOURCE_CONSTANTS & "0"
              
               End With
               
               'copy files in the right folder
              LogItem "[UpdateDownloadAndExtract] upgrade to version:" & ToUpGradeVersion & " as soon as the user allows it"
              objFolder.CopyHere subFolder.Items, 4

              'transfer all the copied files into xlsm
               Application.OnTime Now + TimeValue("00:00:01"), "ImportModules"

              'need to keep this exit for otherwise all the files in the zip

              Exit For
        End If
    Next
  End If
 
 'error handeling
Skip_Import:
 If Err.Number <> 0 Then
  LogItem "[UpdateDownloadAndExtract]  ERROR ! Update install failed !"
  LogItem "[UpdateDownloadAndExtract] (" & Err.Number & ") :" & Err.Description
  Err.Clear
End If
 
 Set subFolder = Nothing
 Set objFolder = Nothing
 Set fso = Nothing
 Set oFolderItem = Nothing
 Set bStrm = Nothing
 Set SH = Nothing
 Set http = Nothing
 Set oZip = Nothing

End Sub

Private Function CheckNewChangeset() As updateStatuses

Dim html As New HTMLDocument
Dim posts As MSHTML.IHTMLElementCollection
Dim post As MSHTML.IHTMLElement
Dim responseText
Dim currentTitle
Dim currentLine
Dim externVersions
Dim externVersion
Dim internVersion
Dim separatorPos
Dim versionLine
Dim SH

CheckNewChangeset = Uptodate
ToUpGradeVersion = MODULE_VERSION

responseText = HttpGET(REPOSITORY & "commits/master")
'populate html doc from response Text
html.body.innerHTML = responseText
    
'try to find an element that is the less prone to be change overtime
Set posts = html.getElementsByTagName("a")

If InStr(html.body.innerHTML, MODULE_NAME) = 0 Then
    LogItem "[CheckNewChangeset] enable to reach github or parse the page"
    CheckNewChangeset = GitHubNotReached
    Exit Function
End If

For Each post In posts
    'check the element contains path to the commit if not go next
    If InStr(LCase(post.href), LCase(MODULE_NAME) & "/commit") = 0 Then GoTo nextIteration:
    
    currentTitle = post.innerText
    separatorPos = InStr(currentTitle, ":")
        
    'skip the line if other html markup and there is not semicolon
    If (separatorPos > 0 And InStr(currentTitle, "<") = 0) Then
        versionLine = Trim(Left(currentTitle, separatorPos - 1))
        
        DebugLine "[CheckNewChangeset] commit: " & versionLine
        'if finishes with a exclamation mark then skip for non devs
        If Right(versionLine, 1) = "!" And Not fsoFolderExists(Settings.CurrentProjectFolder & ".git") Then GoTo nextIteration:
        
        'remove exclamation mark
        versionLine = Replace(versionLine, "!", "")
        'get numbers
        externVersions = Split(versionLine, ".")
        
        'check if does the correct version formating or exit
        If UBound(externVersions) < 2 Then GoTo nextIteration:
                
        'create a comparable version
        On Error Resume Next
        externVersion = CInt(externVersions(0)) * 1000 + CInt(externVersions(1)) * 100 + CInt(externVersions(2))
        'find the internal version
        Dim internVersions: internVersions = Split(MODULE_VERSION, ".")
        internVersion = CInt(internVersions(0)) * 1000 + CInt(internVersions(1)) * 100 + CInt(internVersions(2))
        DebugLine "[CheckNewChangeset]:  extern " & externVersion & " vs intern version " & internVersion
        
        'check if both are numeric then proceed to check
        If (IsNumeric(externVersion) And IsNumeric(internVersion)) Then
            If (externVersion > internVersion) Then
                ToUpGradeVersion = Trim(versionLine)
                CheckNewChangeset = ToUpdate
                LogItem "[CheckNewChangeset] update required to " & versionLine
                Exit Function
            End If
        End If
            
    End If
          
nextIteration:
    Next post

LogItem "[CheckNewChangeset] your current version " & MODULE_VERSION & " is up to date"

End Function
Public Sub CommitToGIT()

Dim stringToExecute, retBack, branchName As String
Dim commitFeedback As String
Dim commitComment, pushFeedback As String

If Settings Is Nothing Then Set Settings = New CmonSettings

'check if the .git repository exists'

If fsoFolderExists(Settings.CurrentProjectFolder & ".git") Then

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
       
       'check they are no error on shell level²'
             
         If (retBack <> 0) Then 'failed
            LogItem "ComitToGIT No commit performed.  Error in the shell execution. CODE(" & retBack & ")"
            MsgBox "Commit to Repository failed", 64, "Commit failed. Error in the shell execution. CODE(" & retBack & ")"
            
         Else
            ' confirm whether commit is actually required
            If InStr(commitFeedback, "Your branch is up to date") > 0 Then
                MsgBox commitFeedback, 1, "No commit required"
                LogItem "[ComitToGIT] No commit required." & vbCrLf & commitFeedback
            
            Else
                'ask whether publish is required
                LogItem "[ComitToGIT] Commit " & MODULE_VERSION & " performed. Displayed in SourceTree as " & commitComment
            
                stringToExecute = "Git.exe push origin " & rgAddDblQuote(branchName)
         
                If MsgBox("Publish last commit on BitBucket?" & vbCrLf & commitFeedback, 1, "Push to GitHub") = 1 Then
                    retBack = ShellRun(stringToExecute, pushFeedback, Settings.CurrentProjectFolder)
                
                    'execute the publish action'
                    'check they are no error on shell level'

                  If retBack <> 0 Then
                      LogItem "[ComitToGIT] Push Failed. Error in the shell execution. CODE(" & retBack & ")" & vbCrLf & pushFeedback
                      MsgBox "Push to Bitbucked failed. Error in the shell execution. CODE(" & retBack & ")" & vbCrLf & pushFeedback, 64, "Push to GitHub"
                  Else
                     MsgBox "Push to GitHub" & vbCrLf & pushFeedback, 1, "Push on GitHub"
                     LogItem "[ComitToGIT] Push to GitHub " & pushFeedback
                
                  End If
                End If
            End If
      End If
  End If
Else

    LogItem "[ComitToGIT] No commit performed. Unable to find .hg repository in " & Settings.CurrentProjectFolder & ".hg"
    MsgBox "Commit failed. Your computer is not linked to a mercurial repository. ", 64, "Commit to BitBucket failed"
End If
End Sub

Public Function GetProjectsDevFolder(userFolder As String) As String

Dim WshShell
Dim line, lines
Dim gitConfigfileContent As String
'set default folder in case the below get interupted
GetProjectsDevFolder = userFolder


'fetch repository from
Set WshShell = CreateObject("WScript.Shell")
ShellRun "Git.exe config --list", gitConfigfileContent, WshShell.ExpandEnvironmentStrings("%userprofile%")
Set WshShell = Nothing

    'split output into lines
    lines = Split(gitConfigfileContent, vbCrLf)

'check that content is available
    
    If UBound(lines) > 1 Then
    For Each line In lines
        line = Trim(LCase(line))
        'DebugLine "[GetProjectsDevFolder] repo file line: " & line
        If InStr(line, "recentrepo") > 0 And Right(line, Len(MODULE_NAME)) = LCase(MODULE_NAME) Then
                GetProjectsDevFolder = Mid(line, 16, Len(line) - (Len(MODULE_NAME) + 15))
                GetProjectsDevFolder = Replace(GetProjectsDevFolder, "/", "\")
                'DebugLine "[GetProjectsDevFolder] path dev found " & GetProjectsDevFolder
                Exit For
            End If
    Next line

    Else
        'DebugLine "[GetProjectsDevFolder] can't find any repository in git config "
    End If


End Function
Private Function ShellRun(ByVal sCmd As String, ByRef sOutput As String, Optional ByVal sExecutionFolder As String)

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput, oErrors As Object
    
    If Len(sExecutionFolder) > 0 Then oShell.CurrentDirectory = sExecutionFolder
    
    Set oExec = oShell.Exec(sCmd)
    ShellRun = oExec.ExitCode
    
    Set oOutput = oExec.StdOut
    Set oErrors = oExec.StdErr
    'DebugLine "[ShellRun] exit code: (" & ShellRun & ") for command: " & sCmd
    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        'DebugLine "[ShellRun] " & sLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    While Not oErrors.AtEndOfStream
        sLine = oErrors.ReadLine
        'DebugLine "[ShellRun] " & sLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend
    sOutput = s

End Function






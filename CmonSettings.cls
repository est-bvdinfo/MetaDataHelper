VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmonSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_userName As String
Private m_computerName As String
Private m_usersysFolder As String
Private m_vbaprojectsFolder As String
Private m_currentProjectFolder As String
Private m_isInitialized As Boolean
Private m_connectionString As String
Private m_serviceURL As String
Private m_clientId As String
Private m_clientSecret As String

Private Sub Class_Initialize()
Call Init

End Sub


Public Property Get CurrentProjectFolder()
   Call Init
   CurrentProjectFolder = m_currentProjectFolder
End Property

Public Property Get userName()
    Call Init
    userName = m_userName
End Property

Public Property Get ComputerName()
    Call Init
    ComputerName = m_computerName
End Property
Public Property Get UserSystemFolder()
    Call Init
    UserSystemFolder = m_usersysFolder

End Property
Public Property Get ConnectionString()
    Call ReadUserSettings
    ConnectionString = m_connectionString

End Property


Public Property Get ServiceURL()
    Call ReadUserSettings
    ServiceURL = m_serviceURL

End Property
Public Property Get ClientId()
    Call ReadUserSettings
    ClientId = m_clientId

End Property
Public Property Get ClientSecret()
    Call ReadUserSettings
    ClientSecret = m_clientSecret

End Property


Private Sub Init()

If m_isInitialized Then Exit Sub

Dim WshShell As Object
Dim WshNetwork As Object
Dim fso As Object
Dim vbaProjectsFolder As String
    
      
Set WshShell = CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")
Set fso = CreateObject("scripting.filesystemobject")
    
    'identify user and computer
    m_userName = UCase(WshNetwork.userName)
    m_computerName = UCase(WshNetwork.ComputerName)
    
    DebugLine "Multiple folder creations"
        
    'create user settings in  appdata if not created yet
    m_usersysFolder = fsoCreateFolder(MODULE_AUDIENCE, WshShell.ExpandEnvironmentStrings("%APPDATA%"))
    
    'generic folder

    vbaProjectsFolder = GetProjectsDevFolder(m_usersysFolder)

    'specific project
    m_currentProjectFolder = fsoCreateFolder(UCase(MODULE_NAME), vbaProjectsFolder)
  
Set WshNetwork = Nothing
Set WshShell = Nothing

m_isInitialized = True
End Sub

Private Sub ReadUserSettings()

Dim defaultConfigFile
Dim oDataBase
Dim oAuthentication
Dim fso As Object
Dim xmlDoc
Dim ts
Dim xmlString

DebugLine "[GetUserVariables] reached"
Set fso = New FileSystemObject

    'get localy stored user variables from app data
If fso.FileExists(m_usersysFolder & USERSETTINGS) Then
      Set xmlDoc = CreateObject("Msxml2.DOMDocument")
      Set ts = fso.OpenTextFile(m_usersysFolder & USERSETTINGS, 1)
      xmlString = ts.ReadALL
      xmlDoc.async = "false"
      xmlDoc.LoadXML xmlString
      
      'check that the settings are valid
      If IsXMLValid(xmlString, USERSETTINGS) = False Then
         Exit Sub
      End If
       'fill connection string
        Set oDataBase = xmlDoc.SelectSingleNode("/" & MODULE_NAME & "/database")
        If (oDataBase Is Nothing) Then
            MsgBox "database node is missing or not properly formated" & vbCrLf & "from " & m_usersysFolder & vbCrLf & "Please Correct!", 64, USERSETTINGS
        Else
               m_connectionString = "Driver={SQL Server}; Server=" & UCase(oDataBase.getAttribute("server")) & "; Database=" & LCase(oDataBase.getAttribute("database")) & "; UID=" & oDataBase.getAttribute("user") & "; PWD=" & oDataBase.getAttribute("password")
               m_serviceURL = "Http://" & LCase(oDataBase.getAttribute("server")) & "/" & LCase(oDataBase.getAttribute("database")) & "/"

               
        End If
        
        Set oAuthentication = xmlDoc.SelectSingleNode("/" & MODULE_NAME & "/authentication")
        If (oAuthentication Is Nothing) Then
            MsgBox "authentication node is missing or not properly formated" & vbCrLf & "from " & m_usersysFolder & vbCrLf & "Please Correct!", 64, USERSETTINGS
        Else
               m_clientId = CStr(oAuthentication.getAttribute("client_id"))
               m_clientSecret = CStr(oAuthentication.getAttribute("client_secret"))
               
        End If
        
Else
        'if setting file not found

        'create a user default config.xml file
        defaultConfigFile = "<" & MODULE_NAME & ">" _
                        & vbCrLf & "<database server=""setup-main1.be.bvdep.net/api/orbis/"" database=""Companies"" user=""DemoLogin"" password=""demo"" />" _
                        & vbCrLf & "<authentication client_id=""abcdef"" client_secret=""token"" />" _
                        & vbCrLf & "</" & MODULE_NAME & ">"
        'create and write multiple user config.xml
        fsoWriteFile defaultConfigFile, MODULE_NAME, "Xml", m_usersysFolder
        
        'display and send to clipboard
        MsgBox USERSETTINGS & " was created for " & m_userName & "within folder " & vbCrLf & m_usersysFolder _
        & vbCrLf & "Please update your own configuration settings! (path in clipboard)"
        ClipBoard_SetData m_usersysFolder
        
        LogItem "[LoadConst] " & USERSETTINGS & " was created for " & m_userName

      
             
End If
defaultConfigFile = ""
Set oDataBase = Nothing
Set ts = Nothing
Set xmlDoc = Nothing


End Sub




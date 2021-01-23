VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CmonRESTAuth2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================================== '
'
' OAuth 2.0 Google Authenticator
' Developed by Kyle Beachill
' licence: MIT (http://www.opensource.org/licenses/mit-license.php)
'
' Inspired loosely by Tim Halls authentication classes in his Excel-Rest library:
' https://github.com/timhall/Excel-REST
'
'
' Features:
'     Simple class to handle Google OAuth 2.0 Authentication
'     Follows the Installed Application Flow
'     Returns Simply the value for the Authorization header in API requests
'
' Gotchas:
'     Tokens are held in plain text in the registry
'
' Required References:
'   - Microsoft Internet Controls
'   - Microsoft XML
'
' ================================================================================== '


Option Explicit

'// Simple enum for current authentication status
Private Enum AuthenticationStatus
    NotAuthenticated = 1
    TokenExpired = 2
    Authenticated = 3
End Enum


'// Application Client ID and Application Secret
Private strClientId As String
Private strClientSecret As String

'// Authentication codes, tokens and expiry date
Private strTokenKey As String
Private strToken As String
Private strRefreshToken As String
Private dtExpiresWhen As Date
Private strAuthCode As String

'// Url End points for the authentication
Private strAuthUrl As String
Private strTokenUrl As String
Private strRedirectUri As String


'// Internet Explorer variables for initial authentication request
Private WithEvents oIExplorer As InternetExplorer
Attribute oIExplorer.VB_VarHelpID = -1
Private blnIeComplete As Boolean

Private strResponseText As String
Private oResponse As Object

'// Save the request object to prevent being created for each token expiry
Private objXMLRequest As MSXML2.ServerXMLHTTP60

'// Since we are persisting the credentials to the registry, we need to read these in each time the class
'// is initialized, if they aren't found - these will be default values, "" for strings and 1900/01/01 for te date
Private Sub Class_Initialize()
    
    Dim sDate As String

    strToken = GetSetting("WrikeAuth", "Tokens", "Token")
    strRefreshToken = GetSetting("WrikeAuth", "Tokens", "RefreshKey")
    sDate = GetSetting("WrikeAuth", "Tokens", "TokenExpiry")
    
    If Len(sDate) > 0 Then
         dtExpiresWhen = CDate(sDate)
    Else
         dtExpiresWhen = #1/1/1900#
    End If
    
End Sub

'// Allows the overriding of the default google EndPoints - these are unlikely to change
Public Sub InitEndPoints( _
    Optional ByVal authURL As String = "https://www.wrike.com/oauth2/authorize", _
    Optional ByVal TokenUrl As String = "https://www.wrike.com/oauth2/token", _
    Optional ByVal RedirectUri As String = "https://www.myapp.com/oauth2_uri" _
)
    
    strAuthUrl = authURL
    strTokenUrl = TokenUrl
    strRedirectUri = RedirectUri
    
End Sub

'// Application ID and Secret will always need passing, since they are required for refresh calls
'// Though these *could* be persisted in the registry also
Public Sub InitClientCredentials(ByVal ClientId As String, ByVal ClientSecret As String)

    strClientId = ClientId
    strClientSecret = ClientSecret

End Sub

'// Simple function to return the authentication status of the currently held credentials

Private Function getAuthenticationStatus() As AuthenticationStatus
        
    '// If the Refresh Token Length is 0 then the initial authentication hasn't occurred
    If Len(strRefreshToken) = 0 Then
        getAuthenticationStatus = NotAuthenticated
        Exit Function
    End If
    
    '// If the refresh date is less than now (with a 10 second buffer) then the token has expired
    If dtExpiresWhen < DateAdd("s", 10, Now()) Then
        getAuthenticationStatus = TokenExpired
        Exit Function
    End If
    
    '// Otherwise the token is valid
    getAuthenticationStatus = Authenticated
    
    
End Function
Private Sub GetNewToken()
    
    Set oIExplorer = New InternetExplorer
    
    With oIExplorer
        .Navigate CreateAuthRequest()
        .AddressBar = False
        .MenuBar = False
        .Resizable = False
        .visible = True
    End With
    ' frmWebAuth.Show vbModeless
    
    '// Wait for userInteraction
    Do: DoEvents: Loop Until blnIeComplete
    
    '// Do we have an Authentication Code?
    If Len(strAuthCode) = 0 Then
        Err.Raise vbObjectError + 2, _
            Description:="User cancelled Authentication"
    End If
    
    '// Now Get a new Token
    If objXMLRequest Is Nothing Then Set objXMLRequest = New MSXML2.ServerXMLHTTP60
    
    With objXMLRequest
        .Open "POST", strTokenUrl, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send CreateTokenRequest()

        If .Status <> 200 Then
            '// Error getting OAuth2 token
            Err.Raise vbObjectError + .Status, _
                Description:="Failed to retrieve OAuth2 Token - " & .Status & ": " & .responseText
        End If
        
        '// Get the credentials from the response
        strToken = GetProp("access_token", .responseText)
        strRefreshToken = GetProp("refresh_token")
        dtExpiresWhen = DateAdd("s", CLng(GetProp("expires_in")), Now())
    
    End With
    
    '// Persist the Refresh key and expiry - the above should only ever need running once per application
    SaveSetting "WrikeAuth", "Tokens", "RefreshKey", strRefreshToken
    SaveSetting "WrikeAuth", "Tokens", "Token", strToken
    SaveSetting "WrikeAuth", "Tokens", "TokenExpiry", CStr(dtExpiresWhen)

End Sub

Private Sub RefreshToken()

    If objXMLRequest Is Nothing Then Set objXMLRequest = New MSXML2.ServerXMLHTTP60
    
    With objXMLRequest
        .Open "POST", strTokenUrl, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send CreateRefreshRequest()
        If .Status <> 200 Then
            '// Error getting OAuth2 token
            Err.Raise vbObjectError + .Status, _
                Description:="Failed to retrieve OAuth2 Token - " & .Status & ": " & .responseText
        End If
        
        '// Get the credentials from the response
        strToken = GetProp("access_token", .responseText)
        dtExpiresWhen = DateAdd("s", CLng(GetProp("expires_in")), Now())
    End With
    
    '// Persist new token in registry
    SaveSetting "WrikeAuth", "Tokens", "Token", strToken
    SaveSetting "WrikeAuth", "Tokens", "TokenExpiry", CStr(dtExpiresWhen)
        
        
End Sub
'// Simple function that gets a propery from a single depth JSON formatted string
'// Requires the property name
'// Requires te JSON string on the first pass
Private Function GetProp(strPropName As String, Optional strJSObject As String = "") As String
    
    Static oScriptControl As Object
    
    If oScriptControl Is Nothing Then Set oScriptControl = CreateObject("ScriptControl")
    
    With oScriptControl
        .Language = "JScript"
        .AddCode "function getProp(json, prop) { return json[prop]; }"
        
        If Len(strJSObject) > 0 Then
            strResponseText = strJSObject
            Set oResponse = .Eval("(" & strJSObject & ")")
        End If
        GetProp = .Run("getProp", oResponse, strPropName)
    End With

End Function
'// Public property to return the Authorisation value header for a request
Public Property Get AuthHeader() As String
    
    Dim eAuthStatus As AuthenticationStatus
    
    eAuthStatus = getAuthenticationStatus
    
    If eAuthStatus = NotAuthenticated Then
        GetNewToken
    ElseIf eAuthStatus = TokenExpired Then
        RefreshToken
    End If
    
    AuthHeader = "Bearer " & strToken
    
    
End Property

'//===========================================================================================================
'// String building functions for the requests

'// Step 1: The initial url for authentication - Note the scope attribute, this sets what the application can access
Private Function CreateAuthRequest() As String
    ' Generate initial Authentication Request
    ' Using installed application flow: https://developers.google.com/accounts/docs/OAuth2InstalledApp
    CreateAuthRequest = strAuthUrl
    If InStr(1, CreateAuthRequest, "?") < 1 Then: CreateAuthRequest = CreateAuthRequest & "?"
    CreateAuthRequest = CreateAuthRequest & "response_type=code"
    CreateAuthRequest = CreateAuthRequest & "&client_id=" & strClientId
    'CreateAuthRequest = CreateAuthRequest & "&redirect_uri=" & strRedirectUri
    LogItem vbCrLf & "Auth Request" & CreateAuthRequest
    
End Function

'// Step 2: The initial POST body to get the initial Token and refresh token
Private Function CreateTokenRequest() As String

    
    CreateTokenRequest = "&client_id=" & strClientId
    CreateTokenRequest = CreateTokenRequest & "&client_secret=" & strClientSecret
    'only for google CreateTokenRequest = CreateTokenRequest & "&redirect_uri=" & strRedirectUri
    CreateTokenRequest = CreateTokenRequest & "&grant_type=authorization_code"
    CreateTokenRequest = CreateTokenRequest & "&code=" & strAuthCode
    
    LogItem vbCrLf & "Token Request" & CreateTokenRequest
    
End Function

'// Step 3: The POST body to refresh a token after it has expired
Private Function CreateRefreshRequest() As String

    CreateRefreshRequest = "client_id=" & strClientId
    CreateRefreshRequest = CreateRefreshRequest & "&client_secret=" & strClientSecret
    CreateRefreshRequest = CreateRefreshRequest & "&grant_type=refresh_token"
    CreateRefreshRequest = CreateRefreshRequest & "&refresh_token=" & strRefreshToken

    
End Function

'//===========================================================================================================
'// Event handling for Internet Explorer Object
'// OAuth 2.0 Process flow requires a user to provide access through the browser for initial Authentication

'//Break Loop on user Quit of IE
Private Sub oIExplorer_OnQuit()
    blnIeComplete = True
End Sub

'//Check the title Window, if Success or Denied Found End the IE interaction
Private Sub oIExplorer_TitleChange(ByVal text As String)

    If InStr(1, text, "code=") > 0 Then
        strAuthCode = Right(text, Len(text) - (InStr(1, text, "code=") + 4))
        LogItem "authCode:" + strAuthCode
        oIExplorer.Quit
    ElseIf InStr(1, text, "Denied") > 0 Then
        oIExplorer.Quit
    End If
    
End Sub


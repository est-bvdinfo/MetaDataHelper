Attribute VB_Name = "Cmon_TextProcessing"
'Handle 64-bit and 32-bit Office
#If VBA7 Then
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, _
    ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As LongPtr
  Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As LongPtr
  Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, _
    ByVal hMem As LongPtr) As LongPtr
#Else
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
  Private Declare Function CloseClipboard Lib "user32" () As Long
  Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Private Declare Function EmptyClipboard Lib "user32" () As Long
  Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
    ByVal lpString2 As Any) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
    As Long, ByVal hMem As Long) As Long
#End If

Const GHND = &H42
Const CF_TEXT = 1
Const MAXSIZE = 4096

Public Function fsoFindValueIntoString(ByVal paragraph, ByVal startText, ByVal stopText)
Dim startPos: startPos = InStr(paragraph, startText)
Dim endPos
If startPos > 0 Then
startPos = startPos + Len(startText)
    endPos = InStr(startPos, paragraph, stopText)
    DebugLine "startPos: " & startPos & " endPos :" & endPos
    If endPos > 0 Then
        fsoFindValueIntoString = Mid(paragraph, startPos, endPos - startPos)
    End If
End If

End Function
Public Function CleanText(ByVal OrigString) As String

CleanText = ReplaceMultiple(OrigString, "", Chr(34), ";", ",", "*", "(", ")")

End Function


Private Function ReplaceMultiple(ByVal OrigString As String, _
     ByVal ReplaceString As String, ParamArray FindChars()) _
     As String

'*********************************************************
'PURPOSE: Replaces multiple substrings in a string with the
'character or string specified by ReplaceString

'PARAMETERS: OrigString -- The string to replace characters in
'            ReplaceString -- The replacement string
'            FindChars -- comma-delimited list of
'                 strings to replace with ReplaceString
'
'RETURNS:    The String with all instances of all the strings
'            in FindChars replaced with Replace String
'EXAMPLE:    s= ReplaceMultiple("H;*()ello", "", ";", ",", "*", "(", ")") -
             'Returns Hello
'CAUTIONS:   'Overlap Between Characters in ReplaceString and
'             FindChars Will cause this function to behave
'             incorrectly unless you are careful about the
'             order of strings in FindChars
'***************************************************************

Dim lLBound As Long
Dim lUBound As Long
Dim lCtr As Long
Dim sAns As String

lLBound = LBound(FindChars)
lUBound = UBound(FindChars)

sAns = OrigString

For lCtr = lLBound To lUBound
    sAns = Replace(sAns, CStr(FindChars(lCtr)), ReplaceString)
Next

ReplaceMultiple = sAns
        

End Function
Public Function rgExCreateCommand(ByVal command)
Dim outputPath:
outputPath = Settings.UserSystemFolder & "CmdOuput.dat"
'remove the previous version
fsoDeleteFile outputPath
'build the command
rgExCreateCommand = VBA.Environ$("COMSPEC") & " /C " & command & " > " & outputPath

End Function
'_________________________________________________'
Public Function rgAddQuote(ByVal text)
  rgAddQuote = Chr(34) & text & Chr(34)
End Function

Public Function rgAddDblQuote(ByVal text)
  rgAddDblQuote = Chr(34) & text & Chr(34)
End Function
 '_________________________________________________'
Public Sub rgWriteLine(ByRef text, ByVal line)
   text = text & line & vbCrLf
End Sub

Public Function rgRemoveSlash(ByVal text) As String
'remove exces of slash
If Trim(Right(text, 1)) = "\" Then
    rgRemoveSlash = Trim(Left(text, Len(text) - 1))
Else
    rgRemoveSlash = text
End If
End Function

Public Function Round_Up(ByVal d As Double) As Integer
    Dim result As Integer
    result = Math.Round(d)
    If result >= d Then
        Round_Up = result
    Else
        Round_Up = result + 1
    End If
End Function
Function HasKey(col As Collection, key As String) As Boolean
    Dim v As Variant
  On Error Resume Next
    v = IsObject(col.Item(key))
    HasKey = Not IsEmpty(v)
End Function



Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

Dim hGlobalMemory As Long, lpGlobalMemory As Long
Dim hClipMemory As Long, X As Long

'Allocate moveable global memory
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
  X = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Function
Function cleanString(str As String) As String
    Dim ch, bytes() As Byte: bytes = str
    For Each ch In bytes
        If Chr(ch) Like "[A-Z.a-z 0-9]" Then cleanString = cleanString & Chr(ch)
    Next ch
End Function


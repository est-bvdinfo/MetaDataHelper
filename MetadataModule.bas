Attribute VB_Name = "MetadataModule"
Option Explicit
Public Sections As SectionCollection
Public SelectPayload As String
Public WherePayload As String
Public Enum ColumnSelect
        fieldName = 1
        fieldCode = 2
        
End Enum


Public Sub GetMetadataSelect()


Dim parsedMetadata As Dictionary
Dim URLget As String
Dim currentSheet As Worksheet

'init parameters
Set Settings = New CmonSettings

If Len(Settings.ServiceURL) < 7 Then Exit Sub

'Fetch metadata select from REST API if not in cache
If Len(SelectPayload) < 10 Then
    URLget = Settings.ServiceURL & "Metadata/select"
    SelectPayload = HttpGET(URLget)
End If

'Convert JSON into object
Set parsedMetadata = ParseJson(SelectPayload)

Set Sections = New SectionCollection

Sections.InitiateCollections parsedMetadata, 0

LogItem "Select Medata Loaded, preparing to render in Excel"

' start rendering the dictionary
' delete all sheets
    DeleteSheets
    
Dim scSection As SectionCollection

For Each scSection In Sections.SubSections
    
   Set currentSheet = Sheets.Add(After:=ActiveSheet)

   currentSheet.Name = IIf(Len(scSection.Name) >= 30, Left(scSection.Name, 30), scSection.Name)

    'add header
    
    
    'add fields
    

Next



End Sub



Private Sub DeleteSheets()

Application.DisplayAlerts = False
Dim ws
For Each ws In ThisWorkbook.Worksheets
    ws.Delete
Next
Application.DisplayAlerts = True

End Sub


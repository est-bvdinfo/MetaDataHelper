Attribute VB_Name = "MetadataModule"
Option Explicit
Public Sections As SectionCollection
Public SelectPayload As String
Public WherePayload As String
Public Enum ColumnSelect
        fieldName = 1
        FieldCode = 2
        
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

Sections.InitiateCollections parsedMetadata, -2, ""

LogItem "Select Medata Loaded, preparing to render in Excel"

' start rendering the dictionary
' delete all sheets
    DeleteSheets
    
Dim topSection As SectionCollection
Dim counter As Integer
For Each topSection In Sections.SubSections
    
   Set currentSheet = Sheets.Add(After:=ActiveSheet)

   currentSheet.Name = IIf(Len(topSection.Name) >= 30, Left(topSection.Name, 30), topSection.Name)
    'start fields and section creations
    parseSection topSection, currentSheet, 2
    
    counter = counter + 1
    If counter > 2 Then Exit For
Next


End Sub

Private Sub parseSection(topSection As SectionCollection, currentSheet As Worksheet, currentRow As Integer)

Dim subSection As SectionCollection
Dim cFields As FieldAttributes

'check if there are fields to display title
    If topSection.Fields.Count > 0 Then topSection.PrintSection currentSheet, currentRow
        
    'get the fields from top section
    For Each cFields In topSection.Fields
        currentRow = currentRow + 1
        cFields.PrintLine currentSheet, currentRow
    Next cFields

    For Each subSection In topSection.SubSections
    
        currentRow = currentRow + 1
        parseSection subSection, currentSheet, currentRow

    Next subSection

End Sub

Private Sub DeleteSheets()

Application.DisplayAlerts = False

Dim ws As Object
For Each ws In ThisWorkbook.Worksheets
    If (ThisWorkbook.Worksheets.Count > 1) Then ws.Delete
Next
Application.DisplayAlerts = True

End Sub


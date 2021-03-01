Attribute VB_Name = "MetadataModule"
Option Explicit
Public Sections As SectionCollection
Public DimensionDic As Dictionary
Public SelectPayload As String
Public WherePayload As String
Public Enum ColumnSelect
        Level = 1
        Path = 2
        FieldCode = 3
        LabelEN = 4
        DataType = 5
        Model = 6
        Description = 7
        FieldLenght = 8 ' beware use for dynamic dimensions
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
Set DimensionDic = New Dictionary

Sections.InitiateCollections parsedMetadata, 0, ""

LogItem "Select Medata Loaded, preparing to render in Excel"

' start rendering the dictionary
' delete all sheets
    DeleteSheets
    
Dim topSection As SectionCollection
Dim counter As Integer
Dim key As Variant

For Each topSection In Sections.SubSections
    
   Set currentSheet = Sheets.Add(After:=ActiveSheet)

   With currentSheet
        .Name = IIf(Len(topSection.Name) >= 30, Left(topSection.Name, 30), topSection.Name)
        .Cells(1, ColumnSelect.DataType).Value = "DataType"
        .Cells(1, ColumnSelect.Description).Value = "Description"
        .Cells(1, ColumnSelect.FieldCode).Value = "FieldCode"
        .Cells(1, ColumnSelect.FieldLenght).Value = "Lenght"
        .Cells(1, ColumnSelect.LabelEN).Value = "Label_EN"
        .Cells(1, ColumnSelect.Level).Value = "Level"
        .Cells(1, ColumnSelect.Model).Value = "ModelID"
        .Cells(1, ColumnSelect.Path).Value = "Path"
        For Each key In DimensionDic.Keys
            .Cells(1, ColumnSelect.FieldLenght + DimensionDic(key)).Value = key
        Next key
       .Rows("1:1").Font.Bold = True

   End With
   
    'start fields and section creations
    parseSection topSection, currentSheet, 2
    
    'autofit
    currentSheet.Columns("A:Q").AutoFit
    
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


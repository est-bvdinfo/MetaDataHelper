VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SectionCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private m_sName As String
Private m_iDepth As Integer
Private m_cSubSections As Collection
Private m_sSectionPath As String
Private m_cFields As Collection
Private m_bIsInitiated As Boolean
Public Function InitiateCollections(parsedMetadata As Dictionary, Depth As Integer, parentSectionPath As String)
Dim vKey As Variant

m_iDepth = Depth + 1
m_sSectionPath = parentSectionPath & "\" & parsedMetadata("Name")
'init new collectionz
Set m_cFields = New Collection
Set m_cSubSections = New Collection

For Each vKey In parsedMetadata.Keys

 Select Case vKey
    
     Case "Name":
        m_sName = parsedMetadata(vKey)
        
     Case "Fields":

        ' parse each field
        Dim oField As Object
        Dim faField As FieldAttributes
        For Each oField In parsedMetadata(vKey)
                Set faField = New FieldAttributes
                faField.InitiateObject oField, m_iDepth, m_sSectionPath
                    
                'add to current object
                If faField.IsInitiated Then m_cFields.Add faField
                    
        Next
     
     Case "Sections":
        
        ' parse each subsection
        Dim oSection As Object
        Dim scSections As SectionCollection
        For Each oSection In parsedMetadata(vKey)
                Set scSections = New SectionCollection
                    scSections.InitiateCollections oSection, m_iDepth, m_sSectionPath
                    
                    'add to current object
                    If scSections.IsInitiated Then m_cSubSections.Add scSections
        Next
        
     Case Else
 
 End Select
 
Next

DebugLine "[InitiateSections] - " & Space(m_iDepth) & "Name:" & m_sName & _
                                ", Nbr. Fields: " & IIf(m_cFields Is Nothing, 0, m_cFields.Count) & _
                                ", Nbr. Sections: " & IIf(m_cSubSections Is Nothing, 0, m_cSubSections.Count)

m_bIsInitiated = True


End Function
Public Sub PrintSection(currentSheet As Worksheet, currentRow As Integer)

currentSheet.Cells(currentRow, 1).Value = m_iDepth - 2

    With currentSheet.Cells(currentRow, 2)
        .Value = m_sSectionPath
        .Font.FontStyle = "Italic"
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Size = 13
    End With
End Sub
Public Property Get IsInitiated() As Boolean: IsInitiated = m_bIsInitiated: End Property
Public Property Get SubSections() As Collection
Set SubSections = m_cSubSections

End Property

Public Property Get Fields() As Collection
Set Fields = m_cFields

End Property

Public Property Get Depth() As Integer
 Depth = m_iDepth

End Property
Public Property Get SectionPath() As String
    SectionPath = m_sSectionPath
End Property

Public Property Get Name() As String
 Name = m_sName

End Property
   


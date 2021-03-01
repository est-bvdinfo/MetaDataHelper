VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldAttributes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sLabel As String
Private m_sFieldId As String
Private m_sModelId As String
Private m_iDepth As Integer
Private m_sRequestType As String
Private m_sDescription As String
Private m_sSectionPath As String
Private m_cDimensions As Collection
Private m_bIsInitiated As Boolean
Private m_bIsLargeText As Boolean
Public Function InitiateObject(parsedMetadata As Dictionary, depth As Integer, parentSectionPath As String)
Dim vKey As Variant

m_iDepth = depth + 1
m_sSectionPath = parentSectionPath
'init new collection
Set m_cDimensions = New Collection

For Each vKey In parsedMetadata.Keys

 Select Case vKey
    
     Case "Label":
        m_sLabel = parsedMetadata(vKey)
        
     Case "FieldId":
        m_sFieldId = parsedMetadata(vKey)
     
     Case "ModelId":
        m_sModelId = parsedMetadata(vKey)
        
    Case "LargeText":
        m_bIsLargeText = CBool(parsedMetadata(vKey))
     
     Case "RequestType":
        m_sRequestType = parsedMetadata(vKey)

     Case "Dimensions":
        
        ' parse each subsection
        Dim oDimension As Object
        Dim cDimension As FieldDimension
            For Each oDimension In parsedMetadata(vKey)
                    Set cDimension = New FieldDimension
                    cDimension.InitiateObject oDimension, m_iDepth
                    
                    'add to current object
                    If cDimension.IsInitiated Then m_cDimensions.Add cDimension
             
            Next
     
     Case Else
 
 End Select
 
Next

DebugLine "[InitiateField] - " & Space(m_iDepth) & "Field:" & m_sFieldId & " ,Label: " & m_sLabel & _
                        ", Model:" & m_sModelId & ", Type: " & m_sRequestType & _
                        ", Nbr dim.: " & IIf(m_cDimensions Is Nothing, 0, m_cDimensions.Count)


m_bIsInitiated = True


End Function
Public Sub PrintLine(thisSheet As Worksheet, row As Integer)

thisSheet.Cells(row, 1).Value = m_iDepth
thisSheet.Cells(row, 2).Value = m_sSectionPath
With thisSheet.Cells(row, 3)
    .Value = m_sFieldId
    .Font.FontStyle = "Bold"
End With

thisSheet.Cells(row, 4).Value = m_sLabel
thisSheet.Cells(row, 5).Value = m_sModelId
thisSheet.Cells(row, 6).Value = m_sRequestType
thisSheet.Cells(row, 8).Value = LargeText
    
End Sub


Public Property Get IsInitiated() As Boolean: IsInitiated = m_bIsInitiated: End Property

Public Property Get Label() As String: Label = m_sLabel: End Property
Public Property Get fieldId() As String
    fieldId = m_sFieldId
End Property
Public Property Get modelId() As String
    modelId = m_sModelId
End Property
Public Property Get RequestType() As String
    RequestType = m_sRequestType
End Property
Public Property Get Description() As String
    Description = m_sDescription
End Property
Public Property Get Dimensions() As Collection
  Set Dimensions = m_cDimensions
End Property
Public Property Get depth() As Integer
 depth = m_iDepth
End Property
Public Property Get IsLargeText() As Boolean: IsLargeText = m_bIsLargeText: End Property
        

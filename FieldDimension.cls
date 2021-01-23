VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldDimension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_sName As String
Private m_sType As String
Private m_iDepth As Integer
Private m_bIsInitiated As Boolean


Public Function InitiateObject(parsedMetadata As Dictionary, depth As Integer)
Dim vKey As Variant

m_iDepth = depth + 1
'init new collection

For Each vKey In parsedMetadata.Keys

 Select Case vKey
    
     Case "Name":
        m_sName = parsedMetadata(vKey)
        
     Case "Type":
        m_sType = parsedMetadata(vKey)
    
     Case Else
 
 End Select
 
Next

DebugLine "[Initiate] - " & Space(m_iDepth) & "Name:" & m_sName & " Type: " & m_sType

m_bIsInitiated = True


End Function


Public Property Get IsInitiated() As Boolean: IsInitiated = m_bIsInitiated: End Property
Public Property Get DimName() As Collection
 Name = m_sName

End Property

Public Property Get DimType() As String
 DimType = m_sType

End Property
Public Property Get depth() As Integer
 depth = m_iDepth

End Property
    



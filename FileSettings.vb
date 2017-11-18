'--------------Members--------------
Private m_fileFormat As String
Private m_extention As String
Private m_seperator As String
Private m_structureType As String
Private m_index_start As String
Private m_Index_end As String
Private m_arrayOperator_start As String
Private m_arrayOperator_end As String

'--------------Custom Classes--------------

Private m_genInfo As GeneratorInfo

'--------------Properties--------------
'FileFormat
Property Let FileFormat(val As String)
    m_fileFormat = val
End Property

Property Get FileFormat() As String
    FileFormat = m_extention
End Property

'Extention
Property Let Extention(val As String)
    m_extention = val
End Property

Property Get Extention() As String
    Extention = m_extention
End Property

'Seperator
Property Let Seperator(val As String)
    m_seperator = val
End Property

Property Get Seperator() As String
    Seperator = m_seperator
End Property

'StructureType
Property Let StructureType(val As String)
    m_structureType = val
End Property

Property Get StructureType() As String
    StructureType = m_structureType
End Property

'Index_start
Property Let Index_start(val As String)
    m_index_start = val
End Property

Property Get Index_start() As String
    Index_start = m_index_start
End Property

'Index_End
Property Let Index_End(val As String)
    m_Index_end = val
End Property

Property Get Index_End() As String
    Index_End = m_Index_end
End Property

'ArrayOperator_Start
Property Let ArrayOperator_Start(val As String)
    m_arrayOperator_start = val
End Property

Property Get ArrayOperator_Start() As String
    ArrayOperator_Start = m_arrayOperator_start
End Property

'ArrayOperator_End
Property Let ArrayOperator_End(val As String)
    m_arrayOperator_end = val
End Property

Property Get ArrayOperator_End() As String
    ArrayOperator_End = m_arrayOperator_end
End Property

'--------------Methods--------------
Public Sub Class_Initialize()
    Set m_genInfo = New GeneratorInfo
End Sub

Public Sub Class_Terminate()
    Set m_genInfo = Nothing
End Sub

'Data sampling from "FieSettings" sheet
Public Sub Init(ByVal rowIndex As Long)
    Dim tsWidth As Long: tsWidth = m_genInfo.Range_FileSettings.Columns.Count
    Dim columnIndex As Long: columnIndex = 1

    For columnIndex = 1 To tsWidth
        Select Case columnIndex
            Case 1
                Me.FileFormat = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 2
                Me.Extention = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 3
                Me.Seperator = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 4
                Me.StructureType = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 5
                Me.Index_start = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 6
                Me.Index_End = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 7
                Me.ArrayOperator_Start = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case 8
                Me.ArrayOperator_End = m_genInfo.Range_FileSettings.Cells(rowIndex, columnIndex).Value
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(columnIndex)
        End Select
    Next
End Sub

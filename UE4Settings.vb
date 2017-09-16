'--------------Members--------------
Private m_needsGenerate As Boolean
Private m_dataTable As Range
Private m_propertyList As Range
Private m_projectRoot As String
Private m_targetDirectory As String
Private m_fileName As String
Private m_fileFormat As FileSettings
Private m_header As String
Private m_propertyTable As String

Private m_genInfo As GeneratorInfo
'--------------Properties--------------
'NeedsGenerate
Property Let NeedsGenerate(val As Boolean)
    m_needsGenerate = val
End Property

Property Get NeedsGenerate() As Boolean
    NeedsGenerate = m_needsGenerate
End Property

'Datatable
Property Let DataTable(val As Range)
    Set m_dataTable = val
End Property

Property Get DataTable() As Range
    Set DataTable = m_dataTable
End Property

'PropertyList
Property Let PropertyList(val As Range)
    Set m_propertyList = val
End Property

Property Get PropertyList() As Range
    Set PropertyList = m_propertyList
End Property

'ProjectRoot
Property Let ProjectRoot(val As String)
    If val <> "" Then
       m_projectRoot = val
    End If
End Property

Property Get ProjectRoot() As String
    ProjectRoot = m_projectRoot
End Property

'TargetDirectory
Property Let TargetDirectory(val As String)
    If val <> "" Then
       m_targetDirectory = val
    End If
End Property

Property Get TargetDirectory() As String
    TargetDirectory = m_targetDirectory
End Property

'FileName
Property Let FileName(val As String)
    If val <> "" Then
       m_fileName = val
    End If
End Property

Property Get FileName() As String
    FileName = m_fileName
End Property

'PropertyTable
Property Let PropertyTable(val As String)
    If val <> "" Then
       m_propertyTable = val
    End If
End Property

'FileFormat
Property Let FileFormat(val As FileSettings)
    Set m_fileFormat = val
End Property

Property Get FileFormat() As FileSettings
    Set FileFormat = m_fileFormat
End Property

Property Get PropertyTable() As String
    PropertyTable = m_propertyTable
End Property


'--------------Methods--------------
Public Sub Class_Initialize()
    Set m_genInfo = New GeneratorInfo
    Set m_fileFormat = New FileSettings
End Sub

Public Sub Class_Terminate()
    Set m_genInfo = Nothing
    Set m_fileFormat = Nothing
End Sub

'Data sampling from "UE4Settings" sheet
Public Sub Init(ByVal rowIndex As Long)
    Dim tsWidth As Long: tsWidth = m_genInfo.Range_UE4Settings.Columns.Count
    Dim columnIndex As Long: columnIndex = 1

    For columnIndex = 1 To tsWidth
        If columnIndex > 1 And Me.NeedsGenerate = False Then
            Exit For
        End If

        Select Case columnIndex
            Case 1
                If m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value = "Generate" Then
                    Me.NeedsGenerate = True
                Else
                    Me.NeedsGenerate = False
                End If
            Case 2
                If m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value <> "" Then
                    Me.DataTable = Worksheets("UE4Datatable").Range(m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value)
                End If
            Case 3
                If m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value <> "" Then
                    Me.PropertyList = Worksheets("UE4Datatable").Range(m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value)
                End If
            Case 4
                Me.ProjectRoot = m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value
            Case 5
                Me.TargetDirectory = m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value
            Case 6
                Me.FileName = m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value
            Case 7
                Dim fileFormatIndex As Long
                fileFormatIndex = m_genInfo.Range_FileSettings.Find(What:=m_genInfo.Range_UE4Settings.Cells(rowIndex, columnIndex).Value).Row - _
                m_genInfo.Range_FileSettings.Cells(1, 1).Row + 1
                Call Me.FileFormat.Init(fileFormatIndex)
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(columnIndex)
        End Select
    Next
End Sub

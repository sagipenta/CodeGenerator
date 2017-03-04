'--------------Members--------------
Private m_needsGenerate As Boolean
Private m_setName As String
Private m_scriptTable As Range
Private m_propertyList As Range
Private m_projectRoot As String
Private m_targetDirectory As String
Private m_fileFormat As FileSettings
Private m_headerComment As String
Private m_header As String
Private m_propertyTable As String
Private m_footer As String

Private m_folderPaths() As String

Private m_genInfo As GeneratorInfo
'--------------Properties--------------
'NeedsGenerate
Property Let NeedsGenerate(val As Boolean)
    m_needsGenerate = val
End Property

Property Get NeedsGenerate() As Boolean
    NeedsGenerate = m_needsGenerate
End Property

'SetName
Property Let SetName(val As String)
    If val <> "" Then
        m_setName = val
    End If
End Property

Property Get SetName() As String
    SetName = m_setName
End Property

'Scripttable
Property Let ScriptTable(val As Range)
    Set m_scriptTable = val
End Property

Property Get ScriptTable() As Range
    Set ScriptTable = m_scriptTable
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
Property Let targetDirectory(val As String)
    If val <> "" Then
       m_targetDirectory = val
    End If
End Property

Property Get targetDirectory() As String
    targetDirectory = m_targetDirectory
End Property

'FileFormat
Property Let FileFormat(val As FileSettings)
    Set m_fileFormat = val
End Property

Property Get FileFormat() As FileSettings
    Set FileFormat = m_fileFormat
End Property

'HeaderComment
Property Let HeaderComment(val As String)
    If val <> "" Then
       m_headerComment = val
    End If
End Property

Property Get HeaderComment() As String
    HeaderComment = m_headerComment
End Property

'Header
Property Let Header(val As String)
    If val <> "" Then
       m_header = val
    End If
End Property

Property Get Header() As String
    Header = m_header
End Property

'PropertyTable
Property Let PropertyTable(val As String)
    If val <> "" Then
       m_propertyTable = val
    End If
End Property

Property Get PropertyTable() As String
    PropertyTable = m_propertyTable
End Property

'Footer
Property Let Footer(val As String)
    If val <> "" Then
       m_footer = val
    End If
End Property

Property Get Footer() As String
    Footer = m_footer
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

Public Function GetFolderPaths() As String()
    GetFolderPaths = m_folderPaths
End Function

Public Sub SetFolderPaths(val() As String)
    m_folderPaths = val
End Sub

Public Function GetFolderPath(index As Long) As String
    GetFolderPath = m_folderPaths(index)
End Function

Public Sub SetFolderPath(index As Long, val As String)
    m_folderPaths(index) = val
End Sub

'Data sampling from "TableSettings" sheet
Public Sub Init(ByVal rowIndex As Long)
    Dim tsWidth As Long: tsWidth = m_genInfo.Range_ScriptSettings.Columns.Count
    Dim columnIndex As Long: columnIndex = 1

    For columnIndex = 1 To tsWidth
        If columnIndex > 1 And Me.NeedsGenerate = False Then
            Exit For
        End If

        Select Case columnIndex
            Case 1
                If m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value = "Generate" Then
                    Me.NeedsGenerate = True
                Else
                    Me.NeedsGenerate = False
                End If
            Case 2
                Me.SetName = m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value
            Case 3
                If m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value <> "" Then
                    Me.ScriptTable = Range(m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value)
                End If
                ReDim m_folderPaths(Me.ScriptTable.Columns.Count)
            Case 4
                If m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value <> "" Then
                    Me.PropertyList = Range(m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value)
                End If
            Case 5
                Me.ProjectRoot = m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value
            Case 6
                Me.targetDirectory = m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value
            Case 7
                Dim fileFormatIndex As Long
                fileFormatIndex = m_genInfo.Range_FileSettings.Find(What:=m_genInfo.Range_ScriptSettings.Cells(rowIndex, columnIndex).Value).Row - _
                m_genInfo.Range_FileSettings.Cells(1, 1).Row + 1
                Call Me.FileFormat.Init(fileFormatIndex)
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(columnIndex)
        End Select
    Next

    'Data sampling from "Templates" sheet
    Dim slWidth As Long: slWidth = m_genInfo.Range_ScriptTemplates.Columns.Count
    Dim slIndex As Long: slIndex = m_genInfo.Range_ScriptTemplates.Find(What:=SetName).Row - m_genInfo.Range_ScriptTemplates.Cells(1, 1).Row + 1

    For slColumnIndex = 1 To slWidth
        Select Case slColumnIndex
            Case 1
            Case 2
                Me.HeaderComment = m_genInfo.Range_ScriptTemplates.Cells(slIndex, slColumnIndex).Value
            Case 3
                Me.Header = m_genInfo.Range_ScriptTemplates.Cells(slIndex, slColumnIndex).Value
            Case 4
                Me.PropertyTable = m_genInfo.Range_ScriptTemplates.Cells(slIndex, slColumnIndex).Value
            Case 5
                Me.Footer = m_genInfo.Range_ScriptTemplates.Cells(slIndex, slColumnIndex).Value
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(columnIndex)
        End Select
    Next
End Sub

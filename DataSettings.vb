'===============================================================================
'   @ScriptName: DataSettings
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================

'--------------Members--------------
Private m_needsGenerate As Boolean
Private m_generationFormat As String
Private m_exportingType As String
Private m_dataTable As Range
Private m_propertyList As Range
Private m_template As TemplateSettings
Private m_projectRoot As String
Private m_targetDirectory As String
Private m_fileName As String

Private m_folderPaths() As String

'--------------Custom Classes--------------

Private m_genInfo As GeneratorInfo
Private m_fileFormat As FileSettings


'--------------Properties--------------
'NeedsGenerate
Property Let NeedsGenerate(val As Boolean)
    m_needsGenerate = val
End Property

Property Get NeedsGenerate() As Boolean
    NeedsGenerate = m_needsGenerate
End Property

'GenerationFormat
Property Let GenerationFormat(val As String)
    If val <> "" Then
        m_generationFormat = val
    End If
End Property

Property Get GenerationFormat() As String
    GenerationFormat = m_generationFormat
End Property

'ExportingType
Property Let ExportingType(val As String)
    If val <> "" Then
        m_exportingType = val
    End If
End Property

Property Get ExportingType() As String
    ExportingType = m_exportingType
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

'Template
Property Let Template(val As TemplateSettings)
    If val <> "" Then
        Set m_template = val
    End If
End Property

Property Get Template() As TemplateSettings
    Set Template = m_template
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

'FileFormat
Property Let FileFormat(val As FileSettings)
    Set m_fileFormat = val
End Property

Property Get FileFormat() As FileSettings
    Set FileFormat = m_fileFormat
End Property

'--------------Methods--------------
Public Sub Class_Initialize()
    Set m_genInfo = New GeneratorInfo
    Set m_template = New TemplateSettings
    Set m_fileFormat = New FileSettings
End Sub

Public Sub Class_Terminate()
    Set m_genInfo = Nothing
    Set m_fileFormat = Nothing
    Set m_template = Nothing
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
Public Sub Init(ByVal targetRange As Range)
    'System'
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path

    Dim dsHeight As Long: dsHeight = targetRange.Rows.Count
    Dim dsColumnIndex As Long: dsColumnIndex = 1

    For dsRowIndex = 1 To dsHeight
        If dsRowIndex > 1 And Me.NeedsGenerate = False Then
            Exit For
        End If

        Select Case dsRowIndex
            Case 1 'NeedsGenerate'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value = "Generate" Then
                    Me.NeedsGenerate = True
                Else
                    Me.NeedsGenerate = False
                End If
            Case 2 'GenerationFormat'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value <> "" Then
                    Me.GenerationFormat = targetRange.Cells(dsRowIndex, dsColumnIndex).Value
                End If
            Case 3 'ExportingType'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value <> "" Then
                    Me.ExportingType = targetRange.Cells(dsRowIndex, dsColumnIndex).Value
                End If
            Case 4 'DataTable'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value <> "" Then
                    Me.DataTable = Range(targetRange.Cells(dsRowIndex, dsColumnIndex).Value)
                End If
                ReDim m_folderPaths(Me.DataTable.Columns.Count)
            Case 5 'PropertyList'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value <> "" Then
                    Me.PropertyList = Range(targetRange.Cells(dsRowIndex, dsColumnIndex).Value)
                End If
            Case 6 'Template'
                If targetRange.Cells(dsRowIndex, dsColumnIndex).Value <> "" Then
                    Call Me.Template.Init(targetRange.Cells(dsRowIndex, dsColumnIndex).Value)
                End If
            Case 7 'FileFormat'
                Dim fileFormatIndex As Long
                fileFormatIndex = m_genInfo.Range_FileSettings.Find(What:=targetRange.Cells(dsRowIndex, dsColumnIndex).Value).Row - _
                m_genInfo.Range_FileSettings.Cells(1, 1).Row + 1
                Call Me.FileFormat.Init(fileFormatIndex)
            Case 8 'ProjectRoot'
                Me.ProjectRoot = fso.GetAbsolutePathName(targetRange.Cells(dsRowIndex, dsColumnIndex).Value)
            Case 9 'TargetDirectory'
                Me.TargetDirectory = targetRange.Cells(dsRowIndex, dsColumnIndex).Value
            Case 10 'FileName'
                Me.FileName = targetRange.Cells(dsRowIndex, dsColumnIndex).Value
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(dsRowIndex)
        End Select
    Next

    Set fso = Nothing

End Sub

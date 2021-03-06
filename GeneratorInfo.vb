'Define Global info here
Public Range_Templates As Range
Public Range_FileSettings As Range

Public CommentIndex As Long
Public AutoGenerateIndex As Long
Public DatatableInitialIndex As Long

Public DataSettingsRange As string

'Initializes Generator Global Information
Public Sub Class_Initialize()

    Set Range_FileSettings = Range("FS_FileSettings")
    Set Range_Templates = Range("T_Templates")

    CommentIndex = 1
    DatatableInitialIndex = 2

    DataSettingsRange = "D3:D12"

End Sub

Public Function GetIsCategoryStart(cell As Range) As Boolean
    If cell.Value <> "" _
      And cell.Offset(1, 0).Value = "" _
      And cell.Offset(-1, 0).Value = "" _
      And cell.Offset(0, 1).Value = "" Then
        If cell.Column <> 1 Then
            If cell.Offset(0, -1).Value <> "" Then 'If it is right end
                GetIsCategoryStart = False
                Exit Function
            End If
        End If

        GetIsCategoryStart = True
        Exit Function
    End If
    GetIsCategoryStart = False
End Function

Public Function GetIsCategoryEnd(cell As Range) As Boolean
    If cell.Value <> "" _
        And cell.Offset(1, 0).Value = "" _
        And cell.Offset(0, 1).Value = "" Then
        GetIsCategoryEnd = True
        Exit Function
    End If
    GetIsCategoryEnd = False
End Function

Public Function GetIsFileName(cell As Range) As Boolean
    If cell.Value <> "" _
        And cell.Offset(0, -1).Value = "" _
        And cell.Offset(0, 1).Value <> "" Then
            GetIsFileName = True
            Exit Function
    End If
    GetIsFileName = False
End Function

Public Function GetIsFolderHead(cell As Range) As Boolean
    If cell.Value <> "" _
        And cell.Offset(0, -1).Value = "" _
        And cell.Offset(0, 1).Value <> "" Then
        If cell.Row <> 1 Then
            If cell.Offset(-1, 0).Value = "" Then
                GetIsFolderHead = True
                Exit Function
            End If
        End If
    End If
    GetIsFolderHead = False
End Function

Public Function GetDepth(targetRange As Range, rowIndex As Long) As Long
    Dim xOffset As Long: xOffset = Me.DatatableInitialIndex

    Do While targetRange.Cells(rowIndex, xOffset) = ""
        If xOffset > targetRange.Columns.Count - Me.DatatableInitialIndex + 1 Then
            GetDepth = 1
            Exit Function
        End If
        xOffset = xOffset + 1
    Loop
    GetDepth = xOffset - Me.DatatableInitialIndex + 1
End Function

Public Function ReplaceKeys(replaceTarget As String, PropertyList As Range, p_repKeys() As String, Optional extentnion = "") As String
    Dim typePrefix As String: typePrefix = ""
    Dim typeSuffix As String: typeSuffix = ""
    For i = Me.DatatableInitialIndex To UBound(p_repKeys)
        Select Case extentnion
        Case "UE4CSV"
        Case Else
          typePrefix = Me.GetTypePrefix(PropertyList.Cells(2, i))
          typeSuffix = Me.GetTypeSuffix(PropertyList.Cells(2, i))
        End Select

        replaceTarget = Replace(replaceTarget, "<" & PropertyList.Cells(1, i) & ">", typePrefix & p_repKeys(i) & typeSuffix)
    Next i

    ' [　の後にスペースが空くのを矯正
    replaceTarget = Replace(replaceTarget, "[ ", "[")
    ReplaceKeys = replaceTarget
 End Function

Public Function GetTabByDepth(depth As Long) As String
    Dim valByDepth As String
    valByDepth = ""
    For i = 1 To depth
        valByDepth = valByDepth & vbTab
    Next
    GetTabByDepth = valByDepth
End Function

Public Function GetFolderName(targetRange As Range, rowIndex As Long) As String
    Dim folderPath As String: folderPath = ""
    Dim yOffset As Long: yOffset = 0
    Do While GetParentFolder(targetRange, rowIndex - yOffset) <> ""
        folderPath = "\" & GetParentFolder(targetRange, rowIndex - yOffset) & folderPath
        yOffset = yOffset + 1
    Loop
    GetFolderName = folderPath
End Function

Public Function GetDirectories(folderPaths() As String, targetRange As Range, rowIndex As Long) As String()
    Dim folderName As String: folderName = ""
    Dim yOffset As Long: yOffset = 0
    Do While GetParentFolder(targetRange, rowIndex - yOffset) <> ""
        folderName = GetParentFolder(targetRange, rowIndex - yOffset)
        yOffset = yOffset + 1
        folderPaths(GetDepth(targetRange, rowIndex - yOffset)) = folderName
    Loop
    GetDirectories = folderPaths
End Function

Public Function CreateDirectories(folderPaths() As String, TargetDirectory As String, targetRange As Range, rowIndex As Long) As String
    Dim folderName As String: folderName = ""
    Dim depth As Long: depth = GetDepth(targetRange, rowIndex)

    For dirIndex = 1 To depth
        folderName = folderName & folderPaths(dirIndex)
        If Dir(TargetDirectory & folderName, vbDirectory) = "" Then
            MkDir TargetDirectory & folderName
        End If
    Next
    CreateDirectories = folderName
End Function

Public Function GetParentFolder(targetRange As Range, currentRow As Long) As String
    If currentRow = 1 Then
        GoTo BreakFunction
    End If

    Dim currentRowDepth As Long: currentRowDepth = GetDepth(targetRange, currentRow)
    Dim parentRowDepth As Long: parentRowDepth = GetDepth(targetRange, currentRow - 1)

    If parentRowDepth > currentRowDepth Then
        GoTo BreakFunction
    End If

    GetParentFolder = "\" & targetRange.Cells(currentRow - 1, parentRowDepth + Me.DatatableInitialIndex - 1)
    Exit Function

BreakFunction:
    GetParentFolder = ""
    Exit Function

End Function

Public Function GetTypePrefix(propertyType As String) As String
    Select Case propertyType
    Case "array"
        GetTypePrefix = "{"
    Case "string"
        GetTypePrefix = """"
    Case "String"
        GetTypePrefix = """"
    Case "Name"
        GetTypePrefix = """"
    Case "Text"
        GetTypePrefix = """"
    Case Else
        GetTypePrefix = ""
    End Select
End Function

Public Function GetTypeSuffix(propertyType As String) As String
    Select Case propertyType
    Case "array"
        GetTypeSuffix = "}"
    Case "string"
        GetTypeSuffix = """"
    Case "String"
        GetTypeSuffix = """"
    Case "Name"
        GetTypeSuffix = """"
    Case "Text"
        GetTypeSuffix = """"
    Case Else
        GetTypeSuffix = ""
    End Select
End Function

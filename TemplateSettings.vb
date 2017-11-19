'--------------Members--------------
Private m_templateRow As Range

Private m_headerComment As String
Private m_header As String
Private m_propertyTable As String
Private m_footer As String

'--------------Custom Classes--------------

Private m_genInfo As GeneratorInfo

'--------------Properties--------------

Property Let TemplateRow(val As Range)
    Set m_templateRow = val
End Property

Property Get TemplateRow() As Range
    Set TemplateRow = m_templateRow
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
End Sub

Public Sub Class_Terminate()
    Set m_genInfo = Nothing
End Sub

'Data sampling from "TableSettings" sheet
Public Sub Init(ByVal templateName As String)
    Me.TemplateRow = m_genInfo.Range_Templates.Rows(m_genInfo.Range_Templates.Find(What:=templateName).Row - _
                m_genInfo.Range_Templates.Cells(1, 1).Row + 1)
    'Data sampling from "Template"
    Dim tWidth As Long: tWidth = Me.TemplateRow.Columns.Count
    Dim tIndex As Long: tIndex = 1

    For tColumnIndex = 1 To tWidth
        Select Case tColumnIndex
            Case 1
            Case 2
                Me.HeaderComment = Me.TemplateRow.Cells(tIndex, tColumnIndex).Value
            Case 3
                Me.Header = Me.TemplateRow.Cells(tIndex, tColumnIndex).Value
            Case 4
                Me.PropertyTable = Me.TemplateRow.Cells(tIndex, tColumnIndex).Value
            Case 5
                Me.Footer = Me.TemplateRow.Cells(tIndex, tColumnIndex).Value
            Case 6
            Case Else
                MsgBox "Property doesn't exist. Index:" & Str(tColumnIndex)
        End Select
    Next

End Sub

'===============================================================================
'   @ScriptName: ParseAllTables
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================

Sub ParseAllTables()
    Dim InitialSheetName As String: InitialSheetName = ActiveSheet.Name
    For Each sheet In WorkSheets
        sheet.Activate
        sheet.Select
        If ActiveSheet.Range("D3").Value = "Generate" Then
            Call ParseThisTable.ParseThisTable
        End If
    Next sheet
    WorkSheets(InitialSheetName).Select
End Sub

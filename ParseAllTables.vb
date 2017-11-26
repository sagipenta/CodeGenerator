'===============================================================================
'   @ScriptName: ParseAllTables
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================

Sub ParseAllTables()

    For Each sheet In WorkSheets
        sheet.Activate
        sheet.Select
        If ActiveSheet.Range("D3").Value = "Generate" Then
            Call ParseThisTable.ParseThisTable
        End If
    Next sheet
End Sub

'===============================================================================
'   @ScriptName: ImportUE4CSV
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================

Sub ImportUE4CSV()
    Dim fileToOpen As String

    fileToOpen = Application.GetOpenFilename("CSV to Import,*.csv?")
    If fileToOpen = "False" Then
        GoTo Cancelled
    End If

    Dim settings As DataSettings
    Set settings = New DataSettings
    Call settings.Init(ActiveSheet.Range("D3:D11"))

    Dim dtColumn As Long: dtColumn = 2

    Dim csvBuf As String
    Dim csvRows As Variant

    Open fileToOpen For Input As #1
    Do Until EOF(1)
		Line Input #1, csvBuf
        csvRows = Split(csvBuf,",")

        If csvRows(0) <> "---" And settings.DataTable.Rows(dtColumn).Cells(1, 1).Value = "" Then
            With settings.DataTable.Rows(dtColumn)
                For i = 1 To .Columns.Count
                    If i <> 1 And i <> .Columns.Count Then
                        settings.DataTable.Rows(dtColumn).Cells(1, i).Value = csvRows(i-2)
                    End If
                Next i
            End With
        End If
        dtColumn = dtColumn + 1

	Loop

    MsgBox "Finished Importing: " & fileToOpen

    Cancelled:
	Close #1

End Sub

'===============================================================================
'   @ScriptName: ParseThisTable
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================



Sub ParseThisTable()
    Dim genInfo As GeneratorInfo
    Set genInfo = New GeneratorInfo
    Dim settings As DataSettings
    Set Settings = New DataSettings
    Call settings.Init(ActiveSheet.Range(genInfo.DataSettingsRange))
    Select Case settings.GenerationFormat

    Case "UE4Datatable"
        Call UE4DatatableGenerator.OutputUE4Datatable(settings)
    Case "ScriptGenerator"
        Call ScriptGenerator.OutputScript(settings)
    Case "TableGenerator"
        Call OutputTable(settings)
    Case Else
        MsgBox "Irregal GenerationFormat: " & settings.GenerationFormat
    End Select

    Set settings = Nothing
End Sub

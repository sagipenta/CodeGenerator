'===============================================================================
'   @ScriptName: ParseThisTable
'   @Developed by: Shigeki Saito
'   @Date:2017 Nov.
'===============================================================================

Sub ParseThisTable()
    Dim settings As DataSettings
    Set Settings = New DataSettings
    Call settings.Init(ActiveSheet.Range("D3:D11"))
    Select Case settings.GenerationType

    Case "UE4Datatable"
        Call UE4DatatableGenerator.OutputUE4Datatable(settings)
    Case "LuaScript"
        Call ScriptGenerator.OutputLuaScript(settings)
    Case "LuaTable"
        Call OutputLuaTable(settings)
    Case Else
        MsgBox "Irregal GenerationType: " & settings.GenerationType
    End Select

    Set settings = Nothing
End Sub

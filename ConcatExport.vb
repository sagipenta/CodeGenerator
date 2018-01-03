'===============================================================================
'   @ScriptName: ConcatExport
'   @Developed by: Shigeki Saito
'   @Date:2017 Dec.
'===============================================================================

Public Sub ExportConcatTable()
    Dim genInfo As GeneratorInfo: Set genInfo = New GeneratorInfo
    Set genInfo = New GeneratorInfo

    Dim settings As DataSettings: Set settings = New DataSettings

    Dim initialSheetName As String

    Dim fSkipHiddenCell As Boolean: fSkipHiddenCell = False
    Dim fSkipNullCell As Boolean: fSkipNullCell = False
    Dim fSkipLineEndPartition As Boolean: fSkipLineEndPartition = True

    Dim dtWidth As Long      'レンジの横サイズ
    Dim dtHeight As Long     'レンジの縦サイズ
    Dim dtColumnIndex As Long
    Dim dtRowIndex As Long
    Dim plIndex As Long

    Dim lWriteSize As Long  '読み込み要素数
    Dim nfile As Integer    'ファイル番号
    Dim szLineWord As String
    Dim szCellWord As String
    Dim szPartitionTemp As String
    Dim targetFullPath As String

    Dim writeTarget As Object '文字列を書き込むオブジェクト
    Set writeTarget = CreateObject("ADODB.Stream") '文字コードをUTF-8にするのでADODBを使用

    lWriteSize = -1


    'エラー処理の設定
    On Error GoTo ErrorHandler

    initialSheetName = ActiveSheet.Name
    Call settings.Init(ActiveSheet.Range(genInfo.DataSettingsRange))

    '相対ターゲットパスとファイル名を結合
    targetFullPath = settings.ProjectRoot & settings.TargetDirectory & "\" & settings.FileName & settings.FileFormat.Extention
    If Dir(settings.ProjectRoot & settings.TargetDirectory, vbDirectory) = "" Then
        MkDir settings.ProjectRoot & settings.TargetDirectory
    End If

    'すでにファイルがある場合は削除する
    If Dir(targetFullPath) <> "" Then
        Kill targetFullPath 'ファイルの削除
    End If

    writeTarget.Type = adTypeText '保存型を文字列に
    writeTarget.Charset = "UTF-8" '文字コードを指定
    writeTarget.Open 'オブジェクトをインスタンス化

    dtWidth = settings.DataTable.Columns.Count
    dtHeight = settings.DataTable.Rows.Count
    lWriteSize = 0

    '検索置換用の文字列を入れる配列
    Dim rowValues() As String
    ReDim rowValues(dtWidth)

    '開始処理。Datatable構造体メンバの定義部分作成
    For plIndex = 3 To settings.PropertyList.Columns.Count
        szLineWord = szLineWord & settings.PropertyList.Cells(1, plIndex).Value & ","
    Next

    szLineWord = "---," & szLineWord
    writeTarget.WriteText szLineWord, adWriteLine
    For Each Sheet In Worksheets
        Sheet.Activate
        Sheet.Select
        Call settings.Init(ActiveSheet.Range(genInfo.DataSettingsRange))

        If ActiveSheet.Range("D5").Value = "Concat" Then

            '-------------------------------
            '出力
            '-------------------------------

            '1行ごとに文字列を抽出してデータ化していく
            For dtRowIndex = 1 To dtHeight
                '非表示セルのチェックありで非表示セルだったら出力しない
                If fSkipHiddenCell And settings.DataTable.Rows(dtRowIndex).Hidden Then
                    GoTo ToNextRow
                End If

                szPartitionTemp = ""    '最初は区切り文字無し
                szLineWord = ""

                '行のコンバート
                For dtColumnIndex = 1 To dtWidth

                    '非表示セルのチェックありで非表示セルだったら出力しない
                    If fSkipHiddenCell And settings.DataTable.Columns(dtColumnIndex).Hidden Then
                        GoTo ToNextColumn
                    End If

                    '行内の1列の文字列抽出
                    szCellWord = settings.DataTable.Cells(dtRowIndex, dtColumnIndex)

                    '空白スキップかつ空セルだったら出力しない
                    If fSkipNullCell And szCellWord = "" Then
                        GoTo ToNextColumn
                    End If
                    rowValues(dtColumnIndex) = szCellWord

                    '-------------------------------
                    '列ごとに処理を掛けていく
                    '-------------------------------
                    Select Case dtColumnIndex
                    Case genInfo.CommentIndex
                        If szCellWord <> "" Then
                            GoTo ToNextRow
                        End If
                    End Select

                    '-------------------------------
                    '結合処理
                    '-------------------------------
                    szLineWord = szLineWord & szCellWord & szPartitionTemp
                    szCellWord = ""
                    lWriteSize = lWriteSize + 1
                    szPartitionTemp = settings.FileFormat.Seperator   '次から区切り文字あり

                '行の出力
ToNextColumn:
                Next

                '空白スキップかつ全て空行だったら行を出力しない
                If fSkipNullCell And szLineWord = "" Then
                    GoTo ToNextRow
                End If

                '行末の区切り文字のスキップなしだったら区切り文字をつける
                If fSkipLineEndPartition = False Then szLineWord = szLineWord + settings.FileFormat.Seperator

                '１行出力
                szLineWord = Replace(szLineWord, vbLf, vbCrLf)
                szLineWord = Replace(szLineWord, "  ", vbTab)
                writeTarget.WriteText szLineWord, adWriteLine

ToNextRow:
            Next

        End If
    Next Sheet

    '-------------------------------
        'BOM取って書き込み
    '-------------------------------

    writeTarget.Position = 0
    writeTarget.Type = adTypeBinary
    writeTarget.Position = 3

    Dim bin: bin = writeTarget.Read

    writeTarget.Close

    Dim newStream: Set newStream = CreateObject("ADODB.Stream")
    newStream.Type = adTypeBinary
    newStream.Open
    newStream.Write (bin)


    '書き込んだファイルの保存
    newStream.SaveToFile targetFullPath, adSaveCreateOverwrite
    newStream.Close

    Set genInfo = Nothing
    Worksheets(initialSheetName).Select

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox "Abort creating" & targetFullPath & "（Error Number:" & Err.Number & "）"
    End If

    MsgBox targetFullPath & " has created"


End Sub

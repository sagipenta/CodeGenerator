'===============================================================================
'   TableGenerator
'   Developed by Shigeki Saito
'   2017 Feb.
'===============================================================================

'Sampling Targets(Range) have to be defined here
'Each range's prefix samples initial letters of its sheet name
Dim genInfo As GeneratorInfo

Sub GenerateTables()
    Set genInfo = New GeneratorInfo
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ChDrive ThisWorkbook.Path
    ChDir ThisWorkbook.Path

    Dim tsWidth As Long: tsWidth = genInfo.Range_TableSettings.Columns.Count
    Dim tsHeight As Long: tsHeight = genInfo.Range_TableSettings.Rows.Count

    Dim sets() As TableSettings
    ReDim sets(tsHeight)

    For setIndex = 1 To tsHeight
        Dim settings As TableSettings
        Set settings = New TableSettings
        Call settings.Init(setIndex)
        If settings.NeedsGenerate Then
            Set sets(setIndex) = settings
            settings.ProjectRoot = fso.GetAbsolutePathName(settings.ProjectRoot)
            Call OutputTable(sets(setIndex))
        End If
    Next
    Set fso = Nothing
    Set genInfo = Nothing

End Sub

Public Function OutputTable( _
    settings As TableSettings, _
    Optional fSkipHiddenCell As Boolean = False, _
    Optional fSkipNullCell As Boolean = False, _
    Optional fSkipLineEndPartition As Boolean = True _
) As Long

    Dim dtWidth As Long      'レンジの横サイズ
    Dim dtHeight As Long     'レンジの縦サイズ
    Dim dtColumnIndex As Long
    Dim dtRowIndex As Long

    Dim lWriteSize As Long  '読み込み要素数
    Dim nfile As Integer    'ファイル番号
    Dim szLineWord As String
    Dim szCellWord As String
    Dim szPartitionTemp As String
    Dim targetFullPath As String
    Dim headerTab As String
    Dim categoryHeaderStartDepth As Long: categoryHeaderStartDepth = 0

    Dim writeTarget As Object '文字列を書き込むオブジェクト
    Set writeTarget = CreateObject("ADODB.Stream") '文字コードをUTF-8にするのでADODBを使用

    lWriteSize = -1

    'エラー処理の設定
    On Error GoTo ErrorHandler

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

    dtWidth = settings.dataTable.Columns.Count
    dtHeight = settings.dataTable.Rows.Count
    lWriteSize = 0

    '検索置換用の文字列を入れる配列
    Dim rowValues() As String
    ReDim rowValues(dtWidth)

    '-------------------------------
    '出力
    '-------------------------------

    '開始処理。クラスヘッダの作成を行う
    szLineWord = settings.HeaderComment & vbLf & settings.Header

    '改行コードをCRLF、スペースをタブに矯正
    szLineWord = Replace(szLineWord, vbLf, vbCrLf)
    szLineWord = Replace(szLineWord, "  ", vbTab)
    writeTarget.WriteText szLineWord, adWriteLine

    '1行ごとに文字列を抽出してLuaコード化していく
    For dtRowIndex = 1 To dtHeight
        '非表示セルのチェックありで非表示セルだったら出力しない
        If fSkipHiddenCell And settings.dataTable.Rows(dtRowIndex).Hidden Then
            GoTo ToNextRow
        End If

        szPartitionTemp = ""    '最初は区切り文字無し
        szLineWord = ""

        '行のコンバート
        For dtColumnIndex = 1 To dtWidth

            '非表示セルのチェックありで非表示セルだったら出力しない
            If fSkipHiddenCell And settings.dataTable.Columns(dtColumnIndex).Hidden Then
                GoTo ToNextColumn
            End If

            '行内の1列の文字列抽出
            szCellWord = settings.dataTable.Cells(dtRowIndex, dtColumnIndex)

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
            Case dtWidth
                '行末までループが来たところで置き換え
                szCellWord = genInfo.ReplaceKeys(settings.PropertyTable, settings.propertyList, rowValues)

                'カテゴリ終端であればフッタを生成
                If genInfo.GetIsCategoryEnd(settings.dataTable.Cells(dtRowIndex, dtColumnIndex)) Then
                    Dim categoryDepth As Long
                    categoryDepth = genInfo.GetDepth(settings.dataTable, dtRowIndex) - _
                                                genInfo.GetDepth(settings.dataTable, dtRowIndex + 1)
                    For categoryIndex = 1 To categoryDepth
                        szCellWord = szCellWord & genInfo.GetTabByDepth(genInfo.GetDepth(settings.dataTable, dtRowIndex) - categoryIndex) & "}," & vbLf
                    Next
                End If
            Case Else
                'コードを自動生成
                If genInfo.GetIsCategoryStart(settings.dataTable.Cells(dtRowIndex, dtColumnIndex)) Then
                    categoryHeaderStartDepth = categoryHeaderStartDepth + 1
                    headerTab = ""
                    For chIndex = 1 To (dtColumnIndex - genInfo.DatatableInitialIndex + 1)
                        headerTab = vbTab & headerTab
                    Next
                    szLineWord = headerTab & szCellWord & vbLf & headerTab & "{"
                    szLineWord = Replace(szLineWord, vbLf, vbCrLf)
                    szLineWord = Replace(szLineWord, "  ", vbTab)
                    writeTarget.WriteText szLineWord, adWriteLine
                    GoTo ToNextRow
                End If
                szCellWord = ""

            End Select

            '-------------------------------
            '結合処理
            '-------------------------------
            szLineWord = szLineWord & szPartitionTemp & szCellWord

            lWriteSize = lWriteSize + 1
            szPartitionTemp = szPartition   '次から区切り文字あり

        '行の出力
ToNextColumn:
        Next

        '空白スキップかつ全て空行だったら行を出力しない
        If fSkipNullCell And szLineWord = "" Then
            GoTo ToNextRow
        End If

        '行末の区切り文字のスキップなしだったら区切り文字をつける
        If fSkipLineEndPartition = False Then szLineWord = szLineWord + szPartition

        '１行出力
        szLineWord = Replace(szLineWord, vbLf, vbCrLf)
        szLineWord = Replace(szLineWord, "  ", vbTab)
        writeTarget.WriteText szLineWord, adWriteLine

ToNextRow:
    Next

    'EOL
    szLineWord = settings.Footer
    szLineWord = Replace(szLineWord, vbLf, vbCrLf)
    szLineWord = Replace(szLineWord, "  ", vbTab)
    writeTarget.WriteText szLineWord, adWriteLine

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

ErrorHandler:
    If Err.Number <> 0 Then
        MsgBox targetFullPath & "の書き込みを中断します。（エラー番号" & Err.Number & "）"
    End If

    MsgBox targetFullPath & "の作成が完了しました"

    '戻り値は読み込んだ要素数
    OutputTable = lWriteSize

End Function

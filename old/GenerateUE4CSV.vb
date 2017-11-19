Sub GenerateUE4CSV()
'
' CSV出力
'
'コンパイルを掛ける--------------------------
    Application.SendKeys ("%{F11}")

    Application.SendKeys ("%DL")

    Application.SendKeys ("%FC") '
'宣言-----------------------------------
    Dim objNetWork As Object    '.NET機能アクセス用
    Dim userName As String           'ユーザー名
    Dim fileLocation As String               '保存先
    Dim absPath As String               'フルパス
    Dim fileName As String              'ファイル名
    Dim lastRow As Long, row As Long, lastCol As Integer, col As Integer
    Dim fileNumber As Integer
    Dim cell As Range

'初期化---------------------------------------
    'ネットワークオブジェクトの作成
    Set objNetWork = CreateObject("WScript.Network")
    'ユーザー名を取得
    userName = objNetWork.userName
    'ファイル名指定
    fileName = ActiveSheet.Name & ".csv"
    '保存先指定
    fileLocation = "C:\Users\" & userName & "\Desktop\convert\"
    'フルパス指定
    absPath = fileLocation & fileName

  '"Convert"フォルダを作成
    If Dir(fileLocation, vbDirectory) = "" Then
        MkDir fileLocation
    End If
        ChDir fileLocation

    lastRow = Range("A5").End(xlDown).row
    lastCol = Range("A5").End(xlToRight).Column

    'ファイル番号を適当に割当
    fileNumber = FreeFile
'実行---------------------------------------
  'スクリーンupdate停止
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Open absPath For Output As #fileNumber

    For row = 6 To lastRow
        For col = 2 To lastCol - 1
            Write #fileNumber, Cells(row, col);
        Next
        Write #fileNumber, Cells(row, col)
    Next

Close fileNumber

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

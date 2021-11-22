Sub webdata貼付_main()
    FM_ary = excel読み込み
    csv_ary = フレッセイcsv読み込み
    
    'JANコード検索　フレッセイ数量合算
    For i = 1 To UBound(FM_ary)
        For h = 0 To UBound(csv_ary)
            If FM_ary(i, 2) = csv_ary(h, 82) Then
                FM_ary(i, 20) = csv_ary(h, 141)
            End If
        Next h
    Next i
    
    ThisWorkbook.Worksheets("Webdata").Activate
    Range("A1").Select
    Worksheets("Webdata").Range(Cells(1, 1), Cells(UBound(FM_ary), 22)) = FM_ary
    
    ThisWorkbook.Worksheets("ピッキング表").Activate
    MsgBox ("Webdata貼付完了しました。")
End Sub

Function excel読み込み() As Variant
    Dim file As String, max_n As Long
    Dim buf As String, tmp As Variant, ary() As Variant
    Dim i As Long, n As Long, val As Long
    max_n = 0
    
    '準備
    'file = "C:\test.csv" 'ファイル指定
    file = csvファイル名探索()
    pos = InStrRev(file, "\")
    FMname = Mid(file, pos + 1)
    
    '  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
    
    Application.DisplayAlerts = False
    
    Workbooks.Open file
      'アクティブ
    Worksheets("Sheet1").Activate
    Worksheets("Sheet1").Select
    Range("A1").Select
    
    Dim ALastRow As Long
    ALastRow = Cells(Rows.Count, 1).End(xlUp).Row
    read_Col = 22
    
    Dim sheetData As Variant
    
    sheetData = Worksheets("Sheet1").Range(Cells(1, 1), Cells(ALastRow, read_Col)) 'シートデータ
    
    Workbooks(FMname).Close
    
    Application.DisplayAlerts = True
  
    excel読み込み = sheetData

End Function

Function csvファイル名探索() As Variant


    csvfilepath = "\\192.168.100.151\ｃｇｃ生産管理データ"
    
    Dim f As Object, cnt As Long
    Dim filename() As Variant    '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月(ファイル名)

    cnt = 0
    s = 0
    '２次元配列の格納する数を求める
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvfilepath).Files
            s = s + 1
        Next f
    End With
    
    'csv存在チェック
    If s = 0 Then
        MsgBox "excelファイルが空です。"
        End
    End If
    
    ReDim filename(s - 1, 1) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        now_date = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2) '現在の日付
        For Each f In .GetFolder(csvfilepath).Files
            If f.Name Like "*" + now_date + "*" Then    'ファイル名に現在の日付が含まれているか
                filename(cnt, 0) = f.Name
                filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
                cnt = cnt + 1
            End If
        Next f
    End With
    
    '含まれていないので終了
    If cnt = 0 Then
        MsgBox ("本日のFMエクスポートしたエクセルがありません。終了します。")
        End
    End If
    
    Dim Max As Integer
    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
    For i = 0 To UBound(filename)
        If Max = 9999 Then
            Max = i
        End If
        If filename(i, 1) > filename(Max, 1) Then
            Max = i
        End If
    Next i
   
    csvファイル名探索 = csvfilepath & "\" & filename(Max, 0)
    
End Function

Sub test()
    now_date = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
End Sub

Function フレッセイcsv読み込み() As Variant
    Dim csvPath As String
    Dim csvName As String
    
    csvName = csv探索()
    
    '本日のフレッセイDLデータがない場合
    now_date = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
    If Not (csvName Like "*" + now_date + "*") Then
        MsgBox ("本日のフレッセイダウンロードデータがありません。FMエクスポートのみ貼付します。")
        ReDim ary(1, 160) As Variant '取得した行数で2次元配列の再定義
        フレッセイcsv読み込み = ary
        Exit Function
    End If
    
    Open csvName For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
    Close #1 'CSVファイルを閉じる
  
    ReDim ary(max_n - 1, 160) As Variant '取得した行数で2次元配列の再定義
    
    Open csvName For Input As #1 'CSVファイルを開く
        Do Until EOF(1) '最終行までループ
        Line Input #1, buf '読み込んだデータを1行ずつみていく
        tmp = Split(buf, ",") 'カンマで分割
        For i = 0 To UBound(tmp) '項目数ぶんループ
            ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
        Next i
        n = n + 1 '配列の次の行へ
        Loop
    Close #1 'CSVファイルを閉じる

    フレッセイcsv読み込み = ary
    
End Function

Function csv探索() As Variant

'    ship_date = Sheets("受注入力").Range("A1")  '出荷日
'
'    'パスの検索
'    sFileFullPath = ThisWorkbook.Path
'    For i = Len(sFileFullPath) To 0 Step -1
'        If InStr(i, sFileFullPath, "\") > 0 Then
'            '現在のフォルダ名を取得
'            sFolderName = Mid(sFileFullPath, InStr(i, sFileFullPath, "\") + 1)
'            '1つ上の階層のフォルダのまでのフルパスを取得
'            sParentFolderPath = Mid(sFileFullPath, 1, InStr(1, sFileFullPath, sFolderName) - 2)
'            Exit For
'        End If
'    Next
    csvPath = "\\192.168.100.105\新rev_files"
    
    'ディレクトリ存在チェック
'    If Dir(csvFilePath, vbDirectory) = "" Then
'        MsgBox "excelディレクトリが存在しません。"
'        End
'    Else
'        Debug.Print "ディレクトリが存在します。"
'    End If
    
    Dim f As Object, cnt As Long
    Dim filename() As Variant    '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月(ファイル名)

    cnt = 0
    s = 0
    '２次元配列の格納する数を求める
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvPath).Files
            s = s + 1
        Next f
    End With
    
    'csv存在チェック
    If s = 0 Then
        MsgBox "excelファイルが空です。"
        End
    End If
    
    ReDim filename(s - 1, 1) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvPath).Files
            If f.Name Like "*PICKING*" Then
                filename(cnt, 0) = f.Name
                filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
                cnt = cnt + 1
            End If
        Next f
    End With
    
    Dim Max As Integer
    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
    For i = 0 To UBound(filename)
        If Max = 9999 Then
            Max = i
        End If
        If filename(i, 1) > filename(Max, 1) Then
            Max = i
        End If
    Next i
   
    csv探索 = csvPath & "\" & filename(Max, 0)
    
End Function



Sub 計画_読み込みmain()
    'アクティブ
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A6").Select  'A1は作業者によるキーボード押下で出荷日が変更されるトラブルがあるため禁止 2021.03.30
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    
    csv_data = 計画_csv読み込み      '該当のcsvデータ
End Sub

Function 計画_csv読み込み() As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = 計画_csvファイル名探索
  
'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
  
  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる
  
  ReDim ary(max_n - 1, 3) As Variant '取得した行数で2次元配列の再定義
    
'   Open file For Input As #1 'CSVファイルを開く
'       Do Until EOF(1) '最終行までループ
'       Line Input #1, buf '読み込んだデータを1行ずつみていく
'       tmp = Split(buf, ",") 'カンマで分割
'       For i = 0 To UBound(tmp) '項目数ぶんループ
'         ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
'       Next i
'       n = n + 1 '配列の次の行へ
'     Loop
'   Close #1 'CSVファイルを閉じる

    'utf8でcsv読み込み
    Dim buf As String, i As Long
    Dim tmp As Variant, j As Long

    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file
        Do Until .EOS
            buf = .ReadText(-2)
            i = i + 1
            tmp = Split(buf, ",")
            For j = 0 To UBound(tmp)
                'Cells(i, j + 1) = tmp(j)
                ary(i, j + 1) = tmp(j)
            Next j
        Loop
        .Close
    End With

    計画_csv読み込み = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function

Function 計画_csvファイル名探索() As Variant

    csvFilePath = "\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\csv\コープデリピッキング表用"
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath
        End
    Else
        Debug.Print "ディレクトリが存在します。"
    End If
    
    Dim f As Object, cnt As Long
    Dim filename() As Variant    '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月(ファイル名)

    cnt = 0
    s = 0
    '２次元配列の格納する数を求める
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            s = s + 1
        Next f
    End With
    
    'csv存在チェック
    If s = 0 Then
        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath
        End
    End If
    
    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
            ' filename(cnt, 2) = Mid(filename(cnt, 0), 1, 4)  '年 ファイル名
            ' filename(cnt, 3) = Mid(filename(cnt, 0), 6, 2) '月 ファイル名
            ' filename(cnt, 4) = Mid(filename(cnt, 0), 9, 2) '日 ファイル名
            cnt = cnt + 1
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
    
    'csv存在チェック
    If Max = 9999 Then
        MsgBox "該当の出荷日のcsvファイルがありません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name
        End
    End If
    
    計画_csvファイル名探索 = csvFilePath & "\" & filename(Max, 0)
    
End Function



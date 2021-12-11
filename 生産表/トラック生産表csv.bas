Sub トラック更新main()
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    ActiveSheet.UnProtect      '保護解除

    '処理すべきかどうか
    seisan_ship_date = Sheets("受注入力").Range("A1")  '生産表の出荷日
    truck_ship_date = Sheets("トラック").Range("F1")  'トラック出荷日

    If seisan_ship_date = truck_ship_date Then
        Debug.Print "トラック開始"
    Else
        Debug.Print "トラック終了"
        MsgBox "トラックの出荷日と生産表の出荷日が一致していません。" & vbCrLf & vbCrLf & _
        "生産表　" & seisan_ship_date & vbCrLf & _
        "トラック　" & truck_ship_date & vbCrLf & vbCrLf & _
        "\\Afnewt320-kyoyu\社内共有\【購買部】\【トラック関連】\【トラック配送計算用】 .xlsm" & vbCrLf & vbCrLf & _
        "を「開く→保存→閉じる」を行って下さい。" & vbCrLf & _
        "その後「トラック更新」ボタンを押して下さい。"

        '終了
        Exit Sub
    End If

'    'アクティブ
'    Worksheets("受注入力").Activate
'    Worksheets("受注入力").Select
'    Range("A6").Select  'A1は作業者によるキーボード押下で出荷日が変更されるトラブルがあるため禁止 2021.03.30


    'アクティブ
    Worksheets("トラック").Activate
    Worksheets("トラック").Select
    'csv_data = csv読み込み_トラック       '該当のcsvデータ
    'A列最終行取得
    Dim TLastRow As Long
    TLastRow = Cells(Rows.Count, 1).End(xlUp).Row

    csv_data = Range(Cells(1, 1), Cells(TLastRow, 3)) 'csvのデータも考えていたが今回は「トラック」シートから取得。

    ship_date = Sheets("受注入力").Range("A1")  '出荷日

    'アクティブ
    Worksheets("手入力").Activate
    Worksheets("手入力").Select
    Range("A2").Select  '複数セル選択の解除。A1はセル入力されているので作業者の誤入力防止のためA2。

    'A列最終行取得
    Dim ALastRow As Long
    ALastRow = Cells(Rows.Count, 1).End(xlUp).Row

    A_start_row = 4
    A_END_row = ALastRow

    '削除
    Range(Cells(A_start_row, 15), Cells(A_END_row, 15)).Clear

    O_column = Range(Cells(A_start_row, 15), Cells(A_END_row, 15))
    A_col_code = Range(Cells(A_start_row, 1), Cells(A_END_row, 1))

    'csvデータの出荷数をD列へ格納
    For i = 1 To UBound(csv_data)
        For a = 1 To UBound(A_col_code)
            If csv_data(i, 1) = 2619 Then
                Debug.Print 32
            End If

            If csv_data(i, 1) = "" Then
                'pass
            ElseIf Str(A_col_code(a, 1)) = Str(csv_data(i, 1)) Then
                Debug.Print 34
                '既に入力されていた場合、足す
                'O_column(a, 1) = val(O_column(a, 1)) + val(csv_data(i, 1))
                '上書き
                O_column(a, 1) = val(csv_data(i, 2))
            End If
        Next a
    Next i
    
    'O列へ貼付
    Range(Cells(A_start_row, 15), Cells(A_END_row, 15)) = O_column

    MsgBox "トラック更新完了しました。" & vbCrLf & vbCrLf & "【出荷日】" & ship_date

    Application.ScreenUpdating = True                  '画面起動
    Application.Calculation = xlCalculationAutomatic  '自動計算
    ActiveSheet.Protect       '保護

End Sub

Sub トラック削除()
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    ActiveSheet.UnProtect      '保護解除

    'アクティブ
    Worksheets("手入力").Activate
    Worksheets("手入力").Select
    Range("A2").Select  '複数セル選択の解除。A1はセル入力されているので作業者の誤入力防止のためA2。

    'A列最終行取得
    Dim ALastRow As Long
    ALastRow = Cells(Rows.Count, 1).End(xlUp).Row

    A_start_row = 4
    A_END_row = ALastRow
    
    '削除
    Range(Cells(A_start_row, 15), Cells(A_END_row, 15)).Clear
    
    Application.ScreenUpdating = True                  '画面起動
    Application.Calculation = xlCalculationAutomatic  '自動計算
    ActiveSheet.Protect       '保護
End Sub

'
'Function csv読み込み_トラック()
'  Dim file As String, max_n As Long
'  Dim buf As String, tmp As Variant, ary() As Variant
'  Dim i As Long, n As Long, val As Long
'  max_n = 0
'
'  '準備
'  'file = "C:\test.csv" 'ファイル指定
'  file = csvファイル名探索_トラック
'
''  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
'
'  Open file For Input As #1 'CSVファイルを開く
'        Do Until EOF(1)
'            Line Input #1, buf
'            max_n = max_n + 1
'        Loop
'  Close #1 'CSVファイルを閉じる
'
'  ReDim ary(max_n - 1, 2) As Variant '取得した行数で2次元配列の再定義
'
'  Open file For Input As #1 'CSVファイルを開く
'      Do Until EOF(1) '最終行までループ
'      Line Input #1, buf '読み込んだデータを1行ずつみていく
'      tmp = Split(buf, ",") 'カンマで分割
'      For i = 0 To UBound(tmp) '項目数ぶんループ
'        ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
'      Next i
'      n = n + 1 '配列の次の行へ
'    Loop
'  Close #1 'CSVファイルを閉じる
'
'    csv読み込み_トラック = ary
''
''  For i = 1 To UBound(ary)
''    Debug.Print ary(i, 0)
''  Next
'End Function
'
'Function csvファイル名探索_トラック()
'
'    ship_date = Sheets("受注入力").Range("A1")  '出荷日
'
'    csvFilePath = "\\Afnewt320-kyoyu\社内共有\【購買部】\【トラック関連】\csv\" & Customer_name & "\" & Year(ship_date) & "年\" & Right("0" & Month(ship_date), 2) & "月"
'
'    'ディレクトリ存在チェック
'    If Dir(csvFilePath, vbDirectory) = "" Then
'        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
'        End
'    Else
'        Debug.Print "ディレクトリが存在します。"
'    End If
'
'    Dim f As Object, cnt As Long
'    Dim filename() As Variant    '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月(ファイル名)
'
'    cnt = 0
'    s = 0
'    '２次元配列の格納する数を求める
'    With CreateObject("Scripting.FileSystemObject")
'        For Each f In .GetFolder(csvFilePath).Files
'            s = s + 1
'        Next f
'    End With
'
'    'csv存在チェック
'    If s = 0 Then
'        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
'        End
'    End If
'
'    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
'
'    With CreateObject("Scripting.FileSystemObject")
'        For Each f In .GetFolder(csvFilePath).Files
'            filename(cnt, 0) = f.Name
'            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
'            filename(cnt, 2) = Mid(filename(cnt, 0), 5, 4)  '年 ファイル名
'            filename(cnt, 3) = Mid(filename(cnt, 0), 10, 2) '月 ファイル名
'            filename(cnt, 4) = Mid(filename(cnt, 0), 13, 2) '日 ファイル名
'            cnt = cnt + 1
'        Next f
'    End With
'
'    Dim Max As Integer
'    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
'    For i = 0 To UBound(filename)
'        If Str(Year(ship_date)) & Str(Right("0" & Month(ship_date), 2)) & Str(Right("0" & Day(ship_date), 2)) = Str(filename(i, 2)) & Str(filename(i, 3)) & Str(filename(i, 4)) Then
'            If Max = 9999 Then
'                Max = i
'            End If
'            If filename(i, 1) > filename(Max, 1) Then
'                Max = i
'            End If
'        End If
'    Next i
'
'    'csv存在チェック
'    If Max = 9999 Then
'        MsgBox "該当の出荷日のcsvファイルがありません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
'        End
'    End If
'
'    csvファイル名探索_トラック = csvFilePath & "\" & filename(Max, 0)
'
'End Function




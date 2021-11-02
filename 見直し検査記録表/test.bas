Option Base 0

Sub リセット()
    Dim rc As VbMsgBoxResult
    rc = MsgBox("リセットしますか？", vbYesNo + vbQuestion)
    
    If rc = vbYes Then
        Call 日付更新
    Else
        Exit Sub
    End If
    
     MsgBox ("リセット完了しました。")
    
End Sub

Sub 日付更新()
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    
    Call 記録コピー
    
    Range("I5:J50,L5:U50").ClearContents
    Range("I1").Value = Date
    Range("I2").Formula = Now
    
    Customer_name = "コープデリ"
    csv_data = 事前入力csv読み込み(Customer_name)       '該当のcsvデータ
    item_list = 重複なしリスト_事前入力(csv_data)
    paste_data = 合計数量計算(csv_data, item_list)
    
    Worksheets("貼付").Activate
    Worksheets("貼付").Select
    Range("A1").Select
    ActiveSheet.Unprotect
    Sheets("貼付").Cells.Clear
    Worksheets("貼付").Range(Cells(1, 1), Cells(UBound(item_list) + 1, 4)) = paste_data

    Worksheets("見直し検査記録").Activate
    Worksheets("見直し検査記録").Select
    Range("A1").Select

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ActiveSheet.Protect
    Application.Calculation = xlAutomatic
End Sub

Sub test()
    Worksheets("Sheet1").Activate
    B_LastRow = Cells(Rows.Count, 1).End(xlUp).Row  ''B列の最終行取得
    read_Col = 1                                   'CB列まで
    
    sheetData = Worksheets("Sheet1").Range(Cells(1, 1), Cells(B_LastRow, read_Col))   'シートデータ取得
    f = 重複なしリスト_事前入力(sheetData)
    l = UBound(f)
End Sub

Function 重複なしリスト_事前入力(csv_data As Variant) As Variant
    
    Dim 辞書 As Object
    Set 辞書 = CreateObject("Scripting.Dictionary")
   
    Dim dataList As Variant
    ReDim dataList(UBound(csv_data), 0) As Variant

    For b = 0 To UBound(csv_data)
        dataList(b, 0) = csv_data(b, 0)
    Next b
   
    For i = 0 To UBound(dataList)
        '辞書に登録されていない時は
          If Not 辞書.Exists(dataList(i, 0)) Then
              '辞書に登録する。
              辞書(dataList(i, 0)) = Empty
          End If
    Next i
    
    重複なしリスト_事前入力 = 辞書.keys
    
End Function
'
'Sub Weekday_sample_02() '数値(シリアル値)を求めて代入
'
'    MsgBox "2016/1/1の曜日の整数値は" & vbCrLf & Weekday(Date)
'l = Date
'
'End Sub

Function 事前入力csv読み込み(Customer_name As Variant) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = 事前入力csvファイル名探索(Customer_name)
  
'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
  
  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる
  
  ReDim ary(max_n - 1, 3) As Variant '取得した行数で2次元配列の再定義
    
  Open file For Input As #1 'CSVファイルを開く
      Do Until EOF(1) '最終行までループ
      Line Input #1, buf '読み込んだデータを1行ずつみていく
      tmp = Split(buf, ",") 'カンマで分割
      For i = 0 To UBound(tmp) '項目数ぶんループ
        ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
      Next i
      n = n + 1 '配列の次の行へ
    Loop
  Close #1 'CSVファイルを閉じる
  
    事前入力csv読み込み = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function

Function 事前入力csvファイル名探索(Customer_name As Variant) As Variant

    ship_date = Date  '出荷日
    
    'パスの検索
    sFileFullPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表"
    For i = Len(sFileFullPath) To 0 Step -1
        If InStr(i, sFileFullPath, "\") > 0 Then
            '現在のフォルダ名を取得
            sFolderName = Mid(sFileFullPath, InStr(i, sFileFullPath, "\") + 1)
            '1つ上の階層のフォルダのまでのフルパスを取得
            sParentFolderPath = Mid(sFileFullPath, 1, InStr(1, sFileFullPath, sFolderName) - 2)
            Exit For
        End If
    Next
    csvFilePath = sFileFullPath & "\コープ事前入力csv\" & Customer_name & "\" & Year(ship_date) & "年\" & Right("0" & Month(ship_date), 2) & "月"
    
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name
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
        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name
        End
    End If
    
    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
            filename(cnt, 2) = Mid(filename(cnt, 0), 5, 4)  '年 ファイル名
            filename(cnt, 3) = Mid(filename(cnt, 0), 10, 2) '月 ファイル名
            filename(cnt, 4) = Mid(filename(cnt, 0), 13, 2) '日 ファイル名
            cnt = cnt + 1
        Next f
    End With
    
    Dim Max As Integer
    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
    For i = 0 To UBound(filename)
        If Str(Year(ship_date)) & Str(Right("0" & Month(ship_date), 2)) & Str(Right("0" & Day(ship_date), 2)) = Str(filename(i, 2)) & Str(filename(i, 3)) & Str(filename(i, 4)) Then
            If Max = 9999 Then
                Max = i
            End If
            If filename(i, 1) > filename(Max, 1) Then
                Max = i
            End If
        End If
    Next i
    
    'csv存在チェック
    If Max = 9999 Then
        MsgBox "該当の出荷日のcsvファイルがありません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name
        End
    End If
    
    事前入力csvファイル名探索 = csvFilePath & "\" & filename(Max, 0)
    
End Function

Function 合計数量計算(csv_data, item_list) As Variant
    Dim paste_data As Variant
    ReDim paste_data(UBound(item_list), 3) As Variant
    
    kuruko_sum = 0
    deli_sum = 0
    
    
    For i = 0 To UBound(item_list)
        For c = 0 To UBound(csv_data)
            If csv_data(c, 0) = item_list(i) Then
                If csv_data(c, 2) = "6" Then
                    kuruko_sum = kuruko_sum + csv_data(c, 1)
                ElseIf csv_data(c, 2) = "7" Then
                    deli_sum = deli_sum + csv_data(c, 1) + 2
                Else
                     deli_sum = deli_sum + csv_data(c, 1)
                End If
            End If
        Next c
        
        If Weekday(Date) = 2 Or Weekday(Date) = 3 Then
            deli_sum = deli_sum + 4
        Else
            deli_sum = deli_sum + 3
        End If
        
        paste_data(i, 0) = i + 1
        paste_data(i, 1) = item_list(i)
        paste_data(i, 2) = deli_sum
        paste_data(i, 3) = kuruko_sum
        kuruko_sum = 0
        deli_sum = 0
    Next i
    
    合計数量計算 = paste_data
End Function

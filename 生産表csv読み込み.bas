
Function 生産表更新main(Customer_name As Variant)
    
    'アクティブ
    Worksheets("受注入力").Activate
    Worksheets("受注入力").Select
    Range("A6").Select  'A1は作業者によるキーボード押下で出荷日が変更されるトラブルがあるため禁止 2021.03.30
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    ActiveSheet.UnProtect      '保護解除
    
    
    csv_data = csv読み込み(Customer_name)       '該当のcsvデータ
    ship_date = Sheets("受注入力").Range("A1")  '出荷日
    
    'A列最終行取得
    Dim ALastRow As Long
    ALastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'A列格納
    A_column = Range(Cells(1, 1), Cells(ALastRow, 1))

    A_start_row = 0
    A_END_row = 0
    
    'A列から取引名を検索
    For i = 1 To UBound(A_column)
            If A_column(i, 1) = Customer_name Then
                A_start_row = i + 1
                e = i
                Exit For
            End If
    Next i
    'A列からENDを検索 ※↑の検索後の行から開始
    For i = e To UBound(A_column)
            If A_column(i, 1) = "END" Then
                A_END_row = i - 1
                Exit For
            End If
    Next

    'D列変数宣言
    Dim D_column() As Variant
    ReDim D_column(1 To A_END_row - A_start_row + 1, 1 To 1)
    '0で埋める
    For i = 1 To UBound(D_column)
        D_column(i, 1) = 0
    Next i
    
    'D_column = Range(Cells(A_start_row, 4), Cells(A_END_row, 4))    '宣言が面倒なのでコピーして持ってくる
    A_col_code = Range(Cells(A_start_row, 1), Cells(A_END_row, 1))  'A列の商品コード　検索用

    ship_sum = 0    '合計数
    'csvデータの出荷数をD列へ格納
    For i = 0 To UBound(csv_data)
        For a = 1 To UBound(A_col_code)
            If A_col_code(a, 1) = csv_data(i, 0) Then
                D_column(a, 1) = csv_data(i, 1)
                ship_sum = ship_sum + csv_data(i, 1)
            End If
        Next a
    Next i
    
    Debug.Print 2234
    'D列へ貼付
    Range(Cells(A_start_row, 4), Cells(A_END_row, 4)) = D_column
    
    MsgBox "更新完了しました。" & vbCrLf & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【合計数】" & ship_sum & vbCrLf & vbCrLf & "※OKボタンクリック後シートの再計算が行われます。" & vbCrLf & "右下が１００％になるまでお待ち下さい。"
    
    If Customer_name = "CGC" Then
        'pass
    ElseIf 1 = 1 Then
        Application.ScreenUpdating = True                  '画面起動
        Application.Calculation = xlCalculationAutomatic  '自動計算
        ActiveSheet.Protect       '保護
    End If
        
'    Debug.Print LastRow
'    Debug.Print 12
End Function

Function csv読み込み(Customer_name As Variant) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = csvファイル名探索(Customer_name)
  
'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得
  
  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる
  
  ReDim ary(max_n - 1, 2) As Variant '取得した行数で2次元配列の再定義
    
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
  
    csv読み込み = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function

Function csvファイル名探索(Customer_name As Variant) As Variant

    ship_date = Sheets("受注入力").Range("A1")  '出荷日
    
    'パスの検索
    sFileFullPath = ThisWorkbook.Path
    For i = Len(sFileFullPath) To 0 Step -1
        If InStr(i, sFileFullPath, "\") > 0 Then
            '現在のフォルダ名を取得
            sFolderName = Mid(sFileFullPath, InStr(i, sFileFullPath, "\") + 1)
            '1つ上の階層のフォルダのまでのフルパスを取得
            sParentFolderPath = Mid(sFileFullPath, 1, InStr(1, sFileFullPath, sFolderName) - 2)
            Exit For
        End If
    Next
    csvFilePath = sParentFolderPath & "\ピッキング表\csv\" & Customer_name & "\" & Year(ship_date) & "年\" & Right("0" & Month(ship_date), 2) & "月"
    
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
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
        MsgBox "csvファイルが空です。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
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
        MsgBox "該当の出荷日のcsvファイルがありません。" & vbCrLf & csvFilePath & vbCrLf & vbCrLf & "【出荷日】" & ship_date & vbCrLf & "【出荷先】" & Customer_name & vbCrLf & "①ピッキング表の印刷を行っていない可能性があります。" & vbCrLf & "②生産表A1セルの出荷日が間違えている可能性があります。"
        End
    End If
    
    csvファイル名探索 = csvFilePath & "\" & filename(Max, 0)
    
End Function



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''ファイルバックアップ''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function データバックアップ(Customer_name As Variant)
    file_pass_to = ディレクトリ作成コピー先()
    
    'コピー先のパス
    file_pass_to = file_pass_to & "\" & Customer_name & "_ピッキング表.xlsm"
    
    'コピー元のパス
    file_pass_from = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\" & Customer_name & "_ピッキング表.xlsm"
    
    'ファイルのコピー
    FileCopy file_pass_from, file_pass_to
    
End Function


Function ディレクトリ作成コピー先() As Variant
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'A1出荷日
    Dim ship_date As Date
    
    ship_date = Workbooks("生産表.xlsm").Sheets("受注入力").Range("A1")

    ' root = ActiveWorkbook.Path & "\csv"
    root = "\\Afnewt320-kyoyu\社内共有\AFSKS\データ保管"
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    dd = Format(Day(ship_date), "00日")

    'F2 出荷先
    Dim Customer_name As String
    Customer_name = Range("F2")
    
    Dim rtn As Long
    rtn = ディレクトリ作成2(root, yyyy, mm, mm & dd)
    file_pass = root & "\" & yyyy & "\" & mm & "\" & mm & dd
    
    ディレクトリ作成コピー先 = file_pass
'    Select Case rtn
'        Case 0
'            MsgBox "フォルダを作成しました。"
'        Case 1
'            MsgBox "フォルダは存在します。"
'        Case Else
'            MsgBox "フォルダの作成に失敗しました。"
'    End Select
End Function

Function ディレクトリ作成2(ParamArray arg()) As Long
    On Error GoTo ErrExit
    If Dir(Join(arg, "\"), vbDirectory) <> "" Then
        CreateDirectory = 1
        Exit Function
    End If
  
    Dim ary As Variant
    Dim i As Long
    For i = LBound(arg) To UBound(arg)
        ary = arg
        ReDim Preserve ary(i)
        If Dir(Join(ary, "\"), vbDirectory) = "" Then
            MkDir Join(ary, "\")
        End If
    Next
  
    CreateDirectory = 0
    Exit Function
  
ErrExit:
    CreateDirectory = 9
End Function


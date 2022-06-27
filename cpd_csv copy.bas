Sub 事前入力_コープデリ読み込み()
    Customer_name = "コープデリ"                       '出荷先
    Call 事前入力_読み込みmain(Customer_name)
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    'Call 事前入力データバックアップ(Customer_name)
End Sub

Function 事前入力_読み込みmain(Customer_name As Variant)
    
    'アクティブ
    Worksheets("Dピッキング表").Activate
    Worksheets("Dピッキング表").Select
    Range("A6").Select
    

    Dim rc As VbMsgBoxResult
    rc = MsgBox("事前入力の読み込みを開始します。" & vbCrLf & "カレンダーより該当の製造日をクリックしてください", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "事前入力の読み込みを中止します", vbCritical
        Exit Function
    End If
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    Worksheets("Dピッキング表").Unprotect   '保護解除
    
    modosu_ship = Worksheets("Dピッキング表").Range("D6")   '戻す用
    CalenderForm2.Show   'カレンダー
'    Application.Calculate   '再計算
    kakunin_ship = Worksheets("Dピッキング表").Range("D6")  '出荷日取得
    
    rc = MsgBox(str(kakunin_ship) + "の事前入力したデータを読み込みますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        Worksheets("Dピッキング表").Range("D6") = modosu_ship    '戻す
            Application.Calculation = xlCalculationAutomatic    '自動計算
            Application.ScreenUpdating = True                 '画面
            Worksheets("Dピッキング表").Protect   '保護
                
        MsgBox "事前入力の読み込みを中止します", vbCritical
        Exit Function
    End If
    
    
    csv_data = 事前入力csv読み込み(Customer_name)       '該当のcsvデータ
    
    センター数量反映 (csv_data)
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Application.ScreenUpdating = True                 '画面
    Worksheets("Dピッキング表").Protect '保護
    
    MsgBox "読み込み完了しました。"
    
End Function

Function 事前入力_読み込みリセット用main(Customer_name As Variant)
    
    'アクティブ
    Worksheets("Dピッキング表").Activate
    Worksheets("Dピッキング表").Select
    Range("A6").Select
    

'    Dim rc As VbMsgBoxResult
'    rc = MsgBox("事前入力の読み込みを開始します。" & vbCrLf & "カレンダーより該当の製造日をクリックしてください", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        MsgBox "事前入力の読み込みを中止します", vbCritical
'        Exit Function
'    End If
    
'    modosu_ship = Worksheets("Dピッキング表").Range("D6")   '戻す用
'    CalenderForm.Show   'カレンダー
'    Application.Calculate   '再計算
'    kakunin_ship = Worksheets("Dピッキング表").Range("D6")  '出荷日取得
'
'    rc = MsgBox(Str(kakunin_ship) + "の事前入力したデータを読み込みますか？", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        Worksheets("Dピッキング表").Range("D6") = modosu_ship    '戻す
'        MsgBox "事前入力の読み込みを中止します", vbCritical
'        Exit Function
'    End If
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    Worksheets("Dピッキング表").Unprotect   '保護解除
    
    
    csv_data = 事前入力csv読み込み(Customer_name)       '該当のcsvデータ
    
    センター数量反映 (csv_data)
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Worksheets("Dピッキング表").Protect '保護
    
'    MsgBox "読み込み完了しました。"
    
End Function
Function センター数量反映(csv_data As Variant)
    'アクティブ
    Worksheets("Dピッキング表").Activate
    Worksheets("Dピッキング表").Select
    ship_date = Range("D6")
    
    C_column = 重複なしリスト_事前入力(csv_data)
    'C列変数宣言
'    Dim C_column() As Variant
'    ReDim C_column(UBound(csv_data), 0) As Variant
'
    
'    For i = 0 To UBound(csv_data)
'        C_column(i, 0) = csv_data(i, 0)
'    Next i
    
    'C列貼付け
    Range(Cells(8, 3), Cells(38, 3)).ClearContents
'    Range(Cells(8, 3), Cells(UBound(csv_data) + 8, 3)) = C_column
'
    '数量削除
    Range("F8:I38,K8:L38").ClearContents
    
    'C列変数宣言
'    Dim C_column() As Variant
'    ReDim C_column(31, 3) As Variant
'    C_column = Range(Cells(8, 3), Cells(38, 3))

    Dim center_num As Variant
    ReDim center_num(220, 3) As Variant '取得した行数で2次元配列の再定義
    For s = 0 To UBound(C_column)
        For i = 0 To UBound(csv_data)
            Cells(8 + s, 3) = C_column(s)
            For c = 1 To 7
                If c = 5 Then '塩尻の数式
                       '数式が入っているので飛ばす
                Else
                    If Left(csv_data(i, 0), 1) = "Ｉ" And Cells(8 + s, 3) = csv_data(i, 0) And Int(csv_data(i, 2)) = c Then
                        Cells(8 + s, 5 + c) = csv_data(i, 1)
                    End If
                End If
            Next c
        Next i
    Next s
    
End Function

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
    
    重複なしリスト_事前入力 = 辞書.Keys
    
End Function



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

    ship_date = Sheets("Dピッキング表").Range("D6")  '出荷日
    
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
    csvFilePath = ActiveWorkbook.Path & "\コープ事前入力csv\" & Customer_name & "\" & Year(ship_date) & "年\" & Right("0" & Month(ship_date), 2) & "月"
    
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
        If str(Year(ship_date)) & str(Right("0" & Month(ship_date), 2)) & str(Right("0" & Day(ship_date), 2)) = str(filename(i, 2)) & str(filename(i, 3)) & str(filename(i, 4)) Then
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



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''ファイルバックアップ''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function 事前入力データバックアップ(Customer_name As Variant)
    file_pass_to = 事前入力ディレクトリ作成コピー先()
    
    'コピー先のパス
    file_pass_to = file_pass_to & "\" & Customer_name & "_ピッキング表.xlsm"
    
    'コピー元のパス
    file_pass_from = "\\Afnewt320-kyoyu\社内共有\AFSKS\Dピッキング表\" & Customer_name & "_ピッキング表.xlsm"
    
    'ファイルのコピー
    FileCopy file_pass_from, file_pass_to
    
End Function


Function 事前入力ディレクトリ作成コピー先() As Variant
    Dim root As String
    Dim yyyy As String
    Dim mm As String
    Dim dd As String
    'A1出荷日
    Dim ship_date As Date
    
    ship_date = Workbooks("生産表.xlsm").Sheets("Dピッキング表").Range("D6")

    ' root = ActiveWorkbook.Path & "\csv"
    root = "\\Afnewt320-kyoyu\社内共有\AFSKS\データ保管"
    yyyy = Format(Year(ship_date), "0000年")
    mm = Format(Month(ship_date), "00月")
    dd = Format(Day(ship_date), "00日")

    'F2 出荷先
    Dim Customer_name As String
    Customer_name = Range("F2")
    
    Dim rtn As Long
    rtn = 事前入力ディレクトリ作成2(root, yyyy, mm, mm & dd)
    file_pass = root & "\" & yyyy & "\" & mm & "\" & mm & dd
    
    事前入力ディレクトリ作成コピー先 = file_pass
'    Select Case rtn
'        Case 0
'            MsgBox "フォルダを作成しました。"
'        Case 1
'            MsgBox "フォルダは存在します。"
'        Case Else
'            MsgBox "フォルダの作成に失敗しました。"
'    End Select
End Function

Function 事前入力ディレクトリ作成2(ParamArray arg()) As Long
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







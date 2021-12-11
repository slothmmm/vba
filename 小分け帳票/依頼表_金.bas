Sub 依頼表更新_金()

'    Dim rc As VbMsgBoxResult
'    rc = MsgBox("データ更新を行いますか？「**」シートが更新されます。", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        MsgBox "データ更新を中止します", vbCritical
'        Exit Sub
'    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算
    
    CalenderForm.Show   'カレンダー
    
    Worksheets("小分け品").Activate
    
    A_LastRow = Cells(Rows.Count, 1).End(xlUp).Row  ''A列の最終行取得
    read_Col = 40
    
    sheetData = Worksheets("小分け品").Range(Cells(1, 1), Cells(A_LastRow, read_Col))   'シートデータ取得
    
    Worksheets("小分け依頼表_金").Activate
    work_date = Day(Worksheets("小分け依頼表_金").Range("O1"))
    work_date_2 = Day(DateAdd("d", 1, Worksheets("小分け依頼表_金").Range("O1")))
    
    Dim paste_data() As Variant         '貼付け
    ReDim paste_data(A_LastRow, 2)
    
    add_No = 0
    
    '初日
    For i = 1 To UBound(sheetData) - 10
        If IsError(sheetData(i, work_date + 7)) Or sheetData(i, 4) = "アイソニーフーズ福島" Or sheetData(i + 2, 7) = "煮物" Then
            GoTo Continue
        End If
        
        If sheetData(i, 5) = "指示数" And sheetData(i, work_date + 7) <> Empty Then
            For s = 0 To 4
                paste_data(add_No, 0) = Worksheets("小分け依頼表_金").Range("O1")
                paste_data(add_No, 1) = sheetData(i + s, 2)
                If s = 0 Then
                    paste_data(add_No, 2) = sheetData(i, work_date + 7)
                End If
                add_No = add_No + 1
            Next s
        End If
Continue:
    Next i
    
'    '２日目
'    For i = 1 To UBound(sheetData) - 10
'        If work_date_2 = 1 Then
'            MsgBox ("月末により１日のデータがこのエクセルにないため反映しません。翌月の小分け在庫表を開き、N1セルへ日付を１日で入力し、C列から１日でフィルターをかけて印刷して下さい。")
'            Exit For
'        End If
'
'        If IsError(sheetData(i, work_date_2 + 7)) Or sheetData(i, 4) = "アイソニーフーズ福島" Or sheetData(i + 2, 7) = "煮物" Then
'            GoTo Continue2
'        End If
'
'        If sheetData(i, 5) = "指示数" And sheetData(i, work_date_2 + 7) <> Empty Then
'            For s = 0 To 4
'                paste_data(add_No, 0) = DateAdd("d", 1, Worksheets("小分け依頼表_金").Range("O1"))
'                paste_data(add_No, 1) = sheetData(i + s, 2)
'                If s = 0 Then
'                    paste_data(add_No, 2) = sheetData(i, work_date_2 + 7)
'                End If
'                add_No = add_No + 1
'            Next s
'        End If
'Continue2:
'    Next i

    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData  'フィルター解除
    Worksheets("小分け依頼表_金").Range("C6:D760").ClearContents   'C列D列　日付け、部品
    Worksheets("小分け依頼表_金").Range("J6:J760").ClearContents   'J列　指示数
    Worksheets("小分け依頼表_金").Range("L6:L760").ClearContents   'L列　順番
    
    C_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 1) 'paste_dateの1列目を１次元配列へ変換 日付け
    C_paste = WorksheetFunction.Transpose(C_paste)                                 '２次元配列へ変換
    
    D_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 2) 'paste_dateの2列目を１次元配列へ変換 部品
    D_paste = WorksheetFunction.Transpose(D_paste)                                 '２次元配列へ変換
    
    J_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 3) 'paste_dateの10列目を１次元配列へ変換 指示数
    J_paste = WorksheetFunction.Transpose(J_paste)                                 '２次元配列へ変換
    
    Worksheets("小分け依頼表_金").Range(Cells(6, 3), Cells(UBound(C_paste) + 5, 3)) = C_paste  '日付け
    Worksheets("小分け依頼表_金").Range(Cells(6, 4), Cells(UBound(C_paste) + 5, 4)) = D_paste  '部品
    Worksheets("小分け依頼表_金").Range(Cells(6, 10), Cells(UBound(C_paste) + 5, 10)) = J_paste    '指示数
    
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        Application.Calculate
        With ActiveSheet
            .Range("A3").Select
            .Range("A3").AutoFilter Field:=6, Criteria1:="<>"
    End With

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算
    
End Sub

Sub 依頼表並び替え_金()
    Application.Calculation = xlCalculationAutomatic  '自動計算
    Application.Calculate                             '再計算
    
    If Worksheets("小分け依頼表_金").Range("T1") = "X" Or Worksheets("小分け依頼表_金").Range("T1") = "エラー" Then
        MsgBox ("L列の順番に重複したデータが入力されています。同じ数字を入力しないよう、再度見直して下さい。U列が「１」以外の行を探して下さい。")
        Exit Sub
    End If
    

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual     '手動計算
    Worksheets("小分け依頼表_金").Activate
    
    '*********************L列の空白を埋める************************************************
    empty_No = 1001 'L列は空白の場合、自動で1001から連番する
    L_order_col = Worksheets("小分け依頼表_金").Range(Cells(6, 12), Cells(700, 12))   'L列
    D_parts_col = Worksheets("小分け依頼表_金").Range(Cells(6, 4), Cells(700, 4))    'D列
    
    For i = 1 To UBound(L_order_col) Step 5
        If D_parts_col(i, 1) <> Empty And L_order_col(i, 1) = Empty Then
            L_order_col(i, 1) = empty_No
            empty_No = empty_No + 1
        End If
    Next i
    
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData  'フィルター解除
    Worksheets("小分け依頼表_金").Range("L6:L760").ClearContents    'L列削除
    Worksheets("小分け依頼表_金").Range(Cells(6, 12), Cells(700, 12)) = L_order_col     'L列貼付け
    
    '*********************L列ソート************************************************
    date_col = Worksheets("小分け依頼表_金").Range(Cells(6, 3), Cells(700, 3))   'C列シートデータ取得
    parts_col = Worksheets("小分け依頼表_金").Range(Cells(6, 4), Cells(700, 4))   'D列シートデータ取得
    shiji_col = Worksheets("小分け依頼表_金").Range(Cells(6, 10), Cells(700, 10))   'J列シートデータ取得
    order_col = Worksheets("小分け依頼表_金").Range(Cells(6, 12), Cells(700, 12))   'L列シートデータ取得
    black_col = Worksheets("小分け依頼表_金").Range(Cells(6, 13), Cells(700, 13))   'M列シートデータ取得
    
    ordar_data = WorksheetFunction.Transpose(order_col) 'L列１次元配列へ変換
    
    add_MAX = 0 '貼付け計算用
    add_No = 0  '貼付け計算用
    
    Dim order_sort() As Variant         '貼付け
    ReDim order_sort(700)
    
    '要素数を計算する
    For i = 1 To UBound(ordar_data)
        If ordar_data(i) <> Empty Then
            order_sort(add_MAX) = ordar_data(i)
            add_MAX = add_MAX + 1
        End If
    Next i
    
    ReDim order_sort(add_MAX - 1)
    
    'L列に含まれているデータを抽出
    For i = 1 To UBound(ordar_data)
        If ordar_data(i) <> Empty Then
            order_sort(add_No) = ordar_data(i)
            add_No = add_No + 1
        End If
    Next i
    
    Call クイックソート_金(order_sort, LBound(order_sort), UBound(order_sort))  '戻り値ないけどorder_sortは変わる
   
   '*********************貼り付けデータの計算************************************************
    Dim paste_data() As Variant         '貼付け
    ReDim paste_data(700, 4)

    paste_No = 0
    For e = 0 To UBound(order_sort)
        For i = 1 To UBound(order_col) Step 5
            If order_col(i, 1) = order_sort(e) Then
                paste_data(paste_No, 0) = date_col(i, 1)            '日付け
                paste_data(paste_No + 1, 0) = date_col(i + 1, 1)    '日付け
                paste_data(paste_No + 2, 0) = date_col(i + 2, 1)    '日付け
                paste_data(paste_No + 3, 0) = date_col(i + 3, 1)    '日付け
                paste_data(paste_No + 4, 0) = date_col(i + 4, 1)    '日付け
                
                paste_data(paste_No, 1) = parts_col(i, 1)           '指示数
                paste_data(paste_No + 1, 1) = parts_col(i + 1, 1)   '指示数
                paste_data(paste_No + 2, 1) = parts_col(i + 2, 1)   '指示数
                paste_data(paste_No + 3, 1) = parts_col(i + 3, 1)   '指示数
                paste_data(paste_No + 4, 1) = parts_col(i + 4, 1)   '指示数

                paste_data(paste_No, 2) = shiji_col(i, 1)           '指示数
                paste_data(paste_No, 3) = order_col(i, 1)            '順番
                paste_data(paste_No, 4) = black_col(i, 1)            '黒

                paste_No = paste_No + 5
            End If
        Next i
    Next e
    
    '*********************貼り付けC列D列J列L列************************************************
    C_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 1) 'paste_dateの1列目を１次元配列へ変換
    C_paste = WorksheetFunction.Transpose(C_paste)                                 '２次元配列へ変換
    
    D_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 2) 'paste_dateの2列目を１次元配列へ変換
    D_paste = WorksheetFunction.Transpose(D_paste)                                 '２次元配列へ変換
    
    J_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 3) 'paste_dateの10列目を１次元配列へ変換
    J_paste = WorksheetFunction.Transpose(J_paste)                                 '２次元配列へ変換
    
    L_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 4) 'paste_dateの12列目を１次元配列へ変換
    L_paste = WorksheetFunction.Transpose(L_paste)                                 '２次元配列へ変換
    
    M_paste = WorksheetFunction.Index(WorksheetFunction.Transpose(paste_data), 5) 'paste_dateの13列目を１次元配列へ変換
    M_paste = WorksheetFunction.Transpose(M_paste)                                 '２次元配列へ変換
    
    
    Worksheets("小分け依頼表_金").Range(Cells(6, 3), Cells(700, 3)) = C_paste 'C列シートデータ貼付
    Worksheets("小分け依頼表_金").Range(Cells(6, 4), Cells(700, 4)) = D_paste  'D列シートデータ貼付
    Worksheets("小分け依頼表_金").Range(Cells(6, 10), Cells(700, 10)) = J_paste  'J列シートデータ貼付
    Worksheets("小分け依頼表_金").Range(Cells(6, 12), Cells(700, 12)) = L_paste  'L列シートデータ貼付
    Worksheets("小分け依頼表_金").Range(Cells(6, 13), Cells(700, 13)) = M_paste  'M列シートデータ貼付
    
    
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        Application.Calculate
        With ActiveSheet
            .Range("A3").Select
            .Range("A3").AutoFilter Field:=6, Criteria1:="<>"
    End With
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic  '自動計算
   
End Sub

Sub クイックソート_金(ByRef argAry() As Variant, _
                   ByVal lngMin As Long, _
                   ByVal lngMax As Long)
    Dim i As Long
    Dim j As Long
    Dim vBase As Variant
    Dim vSwap As Variant
    vBase = argAry(Int((lngMin + lngMax) / 2))
    i = lngMin
    j = lngMax
    Do
        Do While argAry(i) < vBase
            i = i + 1
        Loop
        Do While argAry(j) > vBase
            j = j - 1
        Loop
        If i >= j Then Exit Do
        vSwap = argAry(i)
        argAry(i) = argAry(j)
        argAry(j) = vSwap
        i = i + 1
        j = j - 1
    Loop
    If (lngMin < i - 1) Then
        Call クイックソート_金(argAry, lngMin, i - 1)
    End If
    If (lngMax > j + 1) Then
        Call クイックソート_金(argAry, j + 1, lngMax)
    End If
End Sub


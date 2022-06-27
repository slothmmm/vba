Sub csv_main()

    'アクティブ
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    ' ************************    コープ  ************************************
    'ファイル取得 計画数
    csvFilePath_cope = "\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】\csv\Ｉ\"
    file_list_cope = csvファイル名探索(csvFilePath_cope)
    
    'コープ
    sh_name = "コープ計画数"
    Dim cope_data As Variant
    cope_data = getCSV_utf8(sh_name, file_list_cope, csvFilePath_cope)
    
    ' ************************    ユーコープ  ************************************
    'ファイル取得 計画数
    csvFilePath_Ucope = "\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】\csv\Ｑ\"
    file_list_Ucope = csvファイル名探索(csvFilePath_Ucope)
    
    'ユーコープ
    sh_name = "ユーコープ計画数"
    Dim Ucope_data As Variant
    Ucope_data = getCSV_utf8(sh_name, file_list_Ucope, csvFilePath_Ucope)
    
    ' ************************    戻し  ************************************
    'ファイル取得 計画数
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\戻し\"
    file_list_modoshi = csvファイル名探索(csvFilePath_modoshi, "包材")
    
    '戻し
    sh_name = "戻し"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)

    ' ************************    入荷数  ************************************
    'ファイル取得 計画数
    csvFilePath_nyuuka = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\入荷数\"
    file_list_nyuuka = csvファイル名探索(csvFilePath_nyuuka, "包材")
    
    '入荷数
    sh_name = "入荷数"
    Dim nyuuka_data As Variant
    nyuuka_data = getCSV_utf8(sh_name, file_list_nyuuka, csvFilePath_nyuuka)

    ' ************************    在庫数  ************************************
    'ファイル取得 在庫数
    csvFilePath_zaiko = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\在庫数\"
    file_list_zaiko = csvファイル名探索(csvFilePath_zaiko, "包材")
    
    '在庫数
    sh_name = "在庫数"
    Dim zaiko_data As Variant
    zaiko_data = getCSV_utf8(sh_name, file_list_zaiko, csvFilePath_zaiko)

    Worksheets("レシピ予定表").Activate
    
    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
End Sub

Sub tomorrow_add()
    Range("G2").Select
    Range("G2").Value = DateAdd("d", 1, Range("G2"))
    Call one_search
    
End Sub

Sub yesterday_add()
    Range("G2").Select
    Range("G2").Value = DateAdd("d", -1, Range("G2"))
    Call one_search
End Sub

Sub one_search()
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    ActiveSheet.Unprotect      '保護解除

    ThisWorkbook.Activate

    date_G2 = Worksheets("レシピ予定表").Range("M18") '日付
    Dim paste_one_aduke As Variant  '貼り付けデータ
    Dim paste_one_modoshi As Variant  '貼り付けデータ
    Dim LastRow As Long '最終行取得
    Dim LastCol As Long '最終列取得
    
    '*****************コープ取得************************************
    Worksheets("コープ計画数").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    cope_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「コープ計画数」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_cope(1 To UBound(cope_data, 1), 1 To UBound(cope_data, 2) + 2) '(行,列)
    r = 1

    CP_hiduke = DateAdd("d", 5, date_G2) '月曜日から５日後は土曜日（初日）
    column_b = True
    No_cope = 1

    For d = 1 To 5
        For i = 1 To UBound(cope_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(cope_data, 2)
                    paste_one_cope(r, 1) = "No"
                    paste_one_cope(r, c + 1) = cope_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf cope_data(i, 2) = CP_hiduke Then
                For c = 1 To UBound(cope_data, 2)
                    'paste_one_cope(r, 1) = Trim(str(d) & Right("00" & str(r - 1), 2)) 'A列No
                    paste_one_cope(r, 1) = No_cope 'A列No
                    paste_one_cope(r, c + 1) = cope_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d

    'クリアして貼り付け
    Worksheets("コープ形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("コープ形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_cope, 1), UBound(paste_one_cope, 2))) = paste_one_cope
    
    '*****************ユーコープ取得************************************
    Worksheets("ユーコープ計画数").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    Ucope_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「ユーコープ計画数」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_udata(1 To UBound(Ucope_data, 1), 1 To UBound(Ucope_data, 2) + 2) '(行,列)
    r = 1

    CP_hiduke = DateAdd("d", 5, date_G2) '月曜日から５日後は土曜日（初日）
    column_b = True
    No_cope = 1

    For d = 1 To 5
        For i = 1 To UBound(Ucope_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(Ucope_data, 2)
                    paste_one_udata(r, 1) = "No"
                    paste_one_udata(r, c + 1) = Ucope_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf Ucope_data(i, 2) = CP_hiduke Then
                For c = 1 To UBound(Ucope_data, 2)
                    paste_one_udata(r, 1) = No_cope     'A列No
                    paste_one_udata(r, c + 1) = Ucope_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    'クリアして貼り付け
    Worksheets("ユーコープ形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("ユーコープ形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_udata, 1), UBound(paste_one_udata, 2))) = paste_one_udata

    '*****************戻し取得************************************
    Worksheets("戻し").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    modoshi_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「戻し」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_modoshi(1 To UBound(modoshi_data, 1), 1 To UBound(modoshi_data, 2) + 2) '(行,列)
    r = 1

    CP_hiduke = date_G2 '月曜日から５日後は土曜日（初日）
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(modoshi_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(modoshi_data, 2)
                    paste_one_modoshi(r, 1) = "No"
                    paste_one_modoshi(r, c + 1) = modoshi_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf modoshi_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(modoshi_data, 2)
                    paste_one_modoshi(r, 1) = No_cope    'A列No
                    paste_one_modoshi(r, c + 1) = modoshi_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    'クリアして貼り付け
    Worksheets("戻し形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("戻し形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_modoshi, 1), UBound(paste_one_modoshi, 2))) = paste_one_modoshi

    '*****************入荷数取得************************************
    Worksheets("入荷数").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    nyuuka_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「入荷数」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_nyuuka(1 To UBound(nyuuka_data, 1), 1 To UBound(nyuuka_data, 2) + 2) '(行,列)
    r = 1

    CP_hiduke = date_G2 '月曜日から５日後は土曜日（初日）
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(nyuuka_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(nyuuka_data, 2)
                    paste_one_nyuuka(r, 1) = "No"
                    paste_one_nyuuka(r, c + 1) = nyuuka_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf nyuuka_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(nyuuka_data, 2)
                    paste_one_nyuuka(r, 1) = No_cope    'A列No
                    paste_one_nyuuka(r, c + 1) = nyuuka_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    'クリアして貼り付け
    Worksheets("入荷数形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("入荷数形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_nyuuka, 1), UBound(paste_one_nyuuka, 2))) = paste_one_nyuuka


    '*****************在庫数取得************************************
    Worksheets("在庫数").Activate
    
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「在庫数」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_zaiko(1 To UBound(zaiko_data, 1), 1 To UBound(zaiko_data, 2) + 2) '(行,列)
    r = 1

    CP_hiduke = date_G2 '月曜日から５日後は土曜日（初日）
    column_b = True
    No_cope = 1

    For d = 1 To 10
        For i = 1 To UBound(zaiko_data)
            If i = 1 And column_b Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = "No"
                    paste_one_zaiko(r, c + 1) = zaiko_data(i, c)
                Next c
                r = r + 1
                column_b = False
            ElseIf zaiko_data(i, 12) = CP_hiduke Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = No_cope    'A列No
                    paste_one_zaiko(r, c + 1) = zaiko_data(i, c)
                Next c
                r = r + 1
                No_cope = No_cope + 1
            End If
        Next i
        CP_hiduke = DateAdd("d", 1, CP_hiduke)
        No_cope = 1
    Next d
    
    'クリアして貼り付け
    Worksheets("在庫数形成").Activate
    ActiveSheet.Unprotect      '保護解除
    Worksheets("在庫数形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_zaiko, 1), UBound(paste_one_zaiko, 2))) = paste_one_zaiko

    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
    'ActiveSheet.Protect      '保護
    Worksheets("レシピ予定表").Activate
    
End Sub
Sub paste_one_data()
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算

    ThisWorkbook.Activate
    Worksheets("レシピ予定表").Activate
    Worksheets("レシピ予定表").Unprotect
    
    Application.Calculate '再計算
    
    ' ************************    データ変数へ格納  ************************************
    'レシピNo
    u_recipe_no = Range(Cells(1, 26), Cells(15, 26))
    c_recipe_no = Range(Cells(18, 26), Cells(41, 26))
    
    '計画数
    u_syo = Range(Cells(1, 7), Cells(15, 7))
    c_syo = Range(Cells(18, 7), Cells(41, 7))

    '戻しペーストデータ
    Range(Cells(2, 30), Cells(15, 39)).ClearContents
    Range(Cells(19, 30), Cells(41, 39)).ClearContents
    modoshi_U_paste = Range(Cells(1, 30), Cells(15, 39)).Formula
    modoshi_U_date = Range(Cells(1, 30), Cells(1, 39))              '上の日付
    modoshi_C_paste = Range(Cells(18, 30), Cells(41, 39)).Formula
    modoshi_C_date = Range(Cells(18, 30), Cells(18, 39))            '上の日付

    '入荷ペーストデータ
    Range(Cells(2, 41), Cells(15, 50)).ClearContents
    Range(Cells(19, 41), Cells(41, 50)).ClearContents
    nyuuka_U_paste = Range(Cells(1, 41), Cells(15, 50)).Formula
    nyuuka_U_date = Range(Cells(1, 41), Cells(1, 50))              '上の日付
    nyuuka_C_paste = Range(Cells(18, 41), Cells(41, 50)).Formula
    nyuuka_C_date = Range(Cells(18, 41), Cells(18, 50))            '上の日付
    '計画数ペーストデータ
    Range(Cells(2, 52), Cells(15, 56)).ClearContents
    Range(Cells(19, 52), Cells(41, 56)).ClearContents
    keikaku_U_paste = Range(Cells(1, 52), Cells(15, 56)).Formula
    keikaku_U_date = Range(Cells(1, 52), Cells(1, 56))              '上の日付
    keikaku_C_paste = Range(Cells(18, 52), Cells(41, 56)).Formula
    keikaku_C_date = Range(Cells(18, 52), Cells(18, 56))            '上の日付
    
    ' ************************    ペーストデータ収集  ************************************
    '戻しペーストデータ
    modoshi_U_paste = get_paste_data(u_recipe_no, modoshi_U_paste, "戻し形成", modoshi_U_date)
    modoshi_C_paste = get_paste_data(c_recipe_no, modoshi_C_paste, "戻し形成", modoshi_C_date)
    '入荷ペーストデータ
    nyuuka_U_paste = get_paste_data(u_recipe_no, nyuuka_U_paste, "入荷数形成", nyuuka_U_date)
    nyuuka_C_paste = get_paste_data(c_recipe_no, nyuuka_C_paste, "入荷数形成", nyuuka_C_date)
    '計画数ペーストデータ
    keikaku_U_paste = get_keikaku(u_syo, keikaku_U_paste, "ユーコープ形成", keikaku_U_date)
    keikaku_C_paste = get_keikaku(c_syo, keikaku_C_paste, "コープ形成", keikaku_C_date)

    ' ************************    ペースト  ************************************
    Worksheets("レシピ予定表").Activate
    '戻し
    Range(Cells(1, 30), Cells(15, 39)) = modoshi_U_paste
    Range(Cells(18, 30), Cells(41, 39)) = modoshi_C_paste
    '入荷
    Range(Cells(1, 41), Cells(15, 50)) = nyuuka_U_paste
    Range(Cells(18, 41), Cells(41, 50)) = nyuuka_C_paste
    
    '計画
     Range(Cells(1, 52), Cells(15, 56)) = keikaku_U_paste
     Range(Cells(18, 52), Cells(41, 56)) = keikaku_C_paste

    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算

End Sub

Function get_paste_data(recipe_no As Variant, paste_data As Variant, sheet_name As Variant, ueno_date) As Variant
    Worksheets(sheet_name).Activate
    Dim this_sh_data As Variant  '貼り付けデータ
    Dim LastRow As Long '最終行取得
    Dim LastCol As Long '最終列取得
    LastRow = Cells(Rows.Count, 3).End(xlUp).Row
    LastCol = Cells(1, 1).End(xlToRight).Column + 5
    
    this_sh_data = Range(Cells(1, 1), Cells(LastRow, LastCol))
    
    For r = 1 To UBound(recipe_no)  'レシピNo回す
        For t = 1 To UBound(this_sh_data)   '形成データを回す
            If this_sh_data(t, 7) = recipe_no(r, 1) Then    'レシピNoと形成データの7列目
                For p = 1 To UBound(paste_data, 2)  'ペーストの列数回す
                    If ueno_date(1, p) = this_sh_data(t, 13) Then
                        If sheet_name Like "*戻し形成*" Then
                            paste_data(r, p) = "貸倉庫 " & this_sh_data(t, 12)
                        ElseIf sheet_name Like "*入荷数形成*" Then
                            paste_data(r, p) = "入荷数 " & this_sh_data(t, 12)
                        End If
                    End If
                Next p
            End If
        Next t
    Next r
    
    get_paste_data = paste_data

End Function

Function get_keikaku(keikaku_no As Variant, paste_data As Variant, sheet_name As Variant, ueno_date) As Variant
    Worksheets(sheet_name).Activate
    Dim this_sh_data As Variant  '貼り付けデータ
    Dim LastRow As Long '最終行取得
    Dim LastCol As Long '最終列取得
    LastRow = Cells(Rows.Count, 3).End(xlUp).Row
    LastCol = Cells(1, 1).End(xlToRight).Column + 5
    
    this_sh_data = Range(Cells(1, 1), Cells(LastRow, LastCol))
    
    For r = 1 To UBound(keikaku_no)  '商品コード回す
        For t = 1 To UBound(this_sh_data)   '形成データを回す
            If this_sh_data(t, 2) = keikaku_no(r, 1) Then    '商品コードと形成データの7列目
                For p = 1 To UBound(paste_data, 2)  'ペーストの列数回す
                    If ueno_date(1, p) = this_sh_data(t, 3) Then
                        paste_data(r, p) = this_sh_data(t, 4)
                    End If
                Next p
            End If
        Next t
    Next r
    
    get_keikaku = paste_data

End Function


Sub test()
    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
End Sub

Sub パレット削除()
    Worksheets("移動明細").Activate
    Range("Q5").Select
    Union(Selection, Range("Q5:Q64")).Select
    Union(Selection, Range("W65:W124")).Select
    Selection.ClearContents
    Range("G2").Select
End Sub

Sub filter_paste(sh_name As Variant, paste_one_aduke As Variant, paste_one_modoshi As Variant)
    Dim paste_d As Variant
    date_G2 = Worksheets("レシピ予定表").Range("M18") '日付
    
    '*****************【預け】sh_nameでフィルター************************************
    '格納する２次元配列サイズ設定
    ReDim paste_d(1 To UBound(paste_one_aduke, 1), 1 To UBound(paste_one_aduke, 2)) '(行,列)

    r = 1
    For i = 1 To UBound(paste_one_aduke)
        If i = 1 Then
            For c = 1 To UBound(paste_one_aduke, 2)
                paste_d(r, c) = paste_one_aduke(i, c)
            Next c
            r = r + 1
        ElseIf paste_one_aduke(i, 13) = sh_name And paste_one_aduke(i, 12) = date_G2 Then
            For c = 1 To UBound(paste_one_aduke, 2)
                paste_d(r, 1) = r - 1      'A列No
                paste_d(r, c) = paste_one_aduke(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    'クリアして貼り付け
    Worksheets(sh_name & "預け" & "形成").Activate
    Worksheets(sh_name & "預け" & "形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_d, 1), UBound(paste_d, 2))) = paste_d

    '*****************【戻し】sh_nameでフィルター************************************
    '格納する２次元配列サイズ設定
    ReDim paste_d(1 To UBound(paste_one_modoshi, 1), 1 To UBound(paste_one_modoshi, 2)) '(行,列)

    r = 1
    For i = 1 To UBound(paste_one_modoshi)
        If i = 1 Then
            For c = 1 To UBound(paste_one_modoshi, 2)
                paste_d(r, c) = paste_one_modoshi(i, c)
            Next c
            r = r + 1
        ElseIf paste_one_modoshi(i, 13) = sh_name And paste_one_modoshi(i, 12) = date_G2 Then
            For c = 1 To UBound(paste_one_modoshi, 2)
                paste_d(r, 1) = r - 1      'A列No
                paste_d(r, c) = paste_one_modoshi(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    'クリアして貼り付け
    Worksheets(sh_name & "戻し" & "形成").Activate
    Worksheets(sh_name & "戻し" & "形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_d, 1), UBound(paste_d, 2))) = paste_d

End Sub

Function getCSV_utf8(sh_name As Variant, file_list As Variant, csvFilePath As Variant) As Variant
    
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Worksheets(1)
    
    'Dim strPath As String
    'strPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\預け\【預け】在庫_ダンボール_2022.3.xlsm.csv"
    
    Dim i As Long, j As Long
    Dim strLine As String
    Dim arrLine As Variant 'カンマでsplitして格納
    
    'D列変数宣言
    Dim paste_data() As Variant
    
    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")
    
    'シートクリア
    ThisWorkbook.Activate
    Worksheets(sh_name).Activate
    Worksheets(sh_name).Cells.ClearContents
    max_n = 0
    i = 1
    
    For n = 0 To UBound(file_list)
        With adoSt
            .Charset = "UTF-8"        'Streamで扱う文字コートをutf-8に設定
            .Open                             'Streamをオープン
            .LoadFromFile (csvFilePath & file_list(n, 0)) 'ファイルからStreamにデータを読み込む
            
            Do Until .EOS           'Streamの末尾まで繰り返す
                strLine = .ReadText(adReadLine) 'Streamから1行取り込み
                arrLine = Split(Replace(replaceColon(strLine), """", ""), ":") 'strLineをカンマで区切りarrLineに格納

                max_n = max_n + 1
            Loop

            .Close
        End With
    Next n

    '格納する２次元配列サイズ設定
    ReDim paste_data(1 To max_n, 1 To 30) '(行,列)
    
    csv_column_name = 1 'カラム名を１行目に追加
    
    For n = 0 To UBound(file_list)
        csv_row_num = 1
        
        With adoSt
            .Charset = "UTF-8"        'Streamで扱う文字コートをutf-8に設定
            .Open                             'Streamをオープン
            .LoadFromFile (csvFilePath & file_list(n, 0)) 'ファイルからStreamにデータを読み込む
            
            Do Until .EOS           'Streamの末尾まで繰り返す
                
                    strLine = .ReadText(adReadLine) 'Streamから1行取り込み
                    arrLine = Split(Replace(replaceColon(strLine), """", ""), ":") 'strLineをカンマで区切りarrLineに格納
                    
                    If csv_column_name = 1 Then 'カラム名を１行目に追加
                        For j = 0 To UBound(arrLine)
                            paste_data(i, j + 1) = arrLine(j)
                        Next j
                        i = i + 1
                        csv_column_name = 2
                    ElseIf csv_row_num <> 1 Then 'データの部分を追加
                        For j = 0 To UBound(arrLine)
                            paste_data(i, j + 1) = arrLine(j)
                        Next j
                        i = i + 1
                    End If
                    
                csv_row_num = csv_row_num + 1
            Loop
        
            .Close
        End With
        
    Next n

    Range(Cells(1, 1), Cells(max_n, 30)) = paste_data

    getCSV_utf8 = paste_data

End Function

'受け取った文字列のカンマをコロンに置き換える
'ダブルクォーテーションで囲まれているカンマは置き換えない
Function replaceColon(ByVal str As String) As String
    
    Dim strTemp As String
    Dim quotCount As Long
    
    Dim l As Long
    For l = 1 To Len(str)  'strの長さだけ繰り返す
    
        strTemp = Mid(str, l, 1) 'strから現在の1文字を切り出す
    
        If strTemp = """" Then   'strTempがダブルクォーテーションなら
    
            quotCount = quotCount + 1   'ダブルクォーテーションのカウントを1増やす
    
        ElseIf strTemp = "," Then   'strTempがカンマなら
    
            If quotCount Mod 2 = 0 Then   'quotCountが2の倍数なら
    
                str = Left(str, l - 1) & ":" & Right(str, Len(str) - l)   '現在の1文字をコロンに置き換える
    
            End If
    
        End If
    
    Next l
    
    replaceColon = str

End Function

Function csvファイル名探索(csvFilePath As Variant, Optional csvname_ichibu As Variant = "無し") As Variant
    'csvFilePath = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\預け"
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox "csvディレクトリが存在しません。"
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
        MsgBox "csvファイルが空です。"
        End
    End If
    
    ReDim filename(s - 1, 4) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvFilePath).Files
            If csvname_ichibu = "無し" Then
                filename(cnt, 0) = f.Name
                filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
                cnt = cnt + 1
            Else
                If f.Name Like "*" & csvname_ichibu & "*" Then
                    filename(cnt, 0) = f.Name
                    filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
                    cnt = cnt + 1
                End If
            End If
        Next f
    End With
    
    ReDim tmp(cnt - 1, 4)
    For i = 0 To cnt - 1
        For x = 0 To LBound(filename)
            tmp(i, x) = filename(i, x)
        Next x
    Next i
    
    csvファイル名探索 = tmp
    
End Function





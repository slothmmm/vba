Sub csv_main()
    'アクティブ
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
'    ActiveSheet.Unprotect      '保護解除
    
    '*********************  ファイル取得 ****************************************
    'ファイル取得預け
    csvFilePath_aduke = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\預け\"
    file_list_aduke = csvファイル名探索(csvFilePath_aduke)
    
    'ファイル取得戻し
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\戻し\"
    file_list_modoshi = csvファイル名探索(csvFilePath_modoshi)
    
    'ファイル取得戻し
    csvFilePath_gaibuzaiko = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\外部在庫数\"
    file_list_gaibuzaiko = csvファイル名探索(csvFilePath_gaibuzaiko)

    '*********************  貼付 ****************************************
    'csv貼り付け_預け
    sh_name = "預けcsv"
    Dim aduke_data As Variant
    aduke_data = getCSV_utf8(sh_name, file_list_aduke, csvFilePath_aduke)
    
    'csv貼り付け_戻し
    sh_name = "戻しcsv"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)
    
    'csv貼り付け_戻し
    sh_name = "外部在庫数csv"
    Dim gaibu_data As Variant
    gaibu_data = getCSV_utf8(sh_name, file_list_gaibuzaiko, csvFilePath_gaibuzaiko)

    Worksheets("移動明細").Activate
    
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

    date_G2 = Worksheets("移動明細").Range("G2") '日付
    Dim paste_one_aduke As Variant  '貼り付けデータ
    Dim paste_one_modoshi As Variant  '貼り付けデータ

    ThisWorkbook.Activate
    '*****************預けデータ取得************************************
    Worksheets("預けcsv").Activate
    
    Dim LastRow As Long '最終行取得
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    aduke_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「預けcsv」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_aduke(1 To UBound(aduke_data, 1), 1 To UBound(aduke_data, 2)) '(行,列)
    r = 1
    For i = 1 To UBound(aduke_data)
        If i = 1 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_aduke(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        ElseIf aduke_data(i, 10) = "預け" And aduke_data(i, 12) = date_G2 And gaibu_data(i, 13) = "新木商事" Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_aduke(r, 1) = r - 1      'A列No
                paste_one_aduke(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        End If
    Next i
    
    'クリアして貼り付け
    Worksheets("預け形成1日").Activate
    Worksheets("預け形成1日").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_aduke, 1), UBound(paste_one_aduke, 2))) = paste_one_aduke
    
    '*****************戻しデータ取得************************************
    Worksheets("戻しcsv").Activate
    
    'Dim LastRow As Long '最終行取得
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    aduke_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「預けcsv」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_modoshi(1 To UBound(aduke_data, 1), 1 To UBound(aduke_data, 2)) '(行,列)
    r = 1
    For i = 1 To UBound(aduke_data)
        If i = 1 Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_modoshi(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        ElseIf aduke_data(i, 10) = "戻し" And aduke_data(i, 12) = date_G2 And gaibu_data(i, 13) = "新木商事" Then
            For c = 1 To UBound(aduke_data, 2)
                paste_one_modoshi(r, 1) = r - 1      'A列No
                paste_one_modoshi(r, c) = aduke_data(i, c)
            Next c
            r = r + 1
        End If
    Next i
   
    'クリアして貼り付け
    Worksheets("戻し形成1日").Activate
    Worksheets("戻し形成1日").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_modoshi, 1), UBound(paste_one_modoshi, 2))) = paste_one_modoshi

    '*****************外部在庫数データ取得************************************
    Worksheets("外部在庫数csv").Activate
    
    'Dim LastRow As Long '最終行取得
    LastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    gaibu_data = Range(Cells(1, 1), Cells(LastRow, LastCol))    '「外部在庫数形成1日」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_gaibu(1 To UBound(gaibu_data, 1), 1 To UBound(gaibu_data, 2)) '(行,列)
    r = 1
    For i = 1 To UBound(gaibu_data)
        If i = 1 Then
            For c = 1 To UBound(gaibu_data, 2)
                paste_one_gaibu(r, c) = gaibu_data(i, c)
            Next c
            r = r + 1
        ElseIf gaibu_data(i, 10) = "外部在庫数" And gaibu_data(i, 12) = date_G2 And gaibu_data(i, 13) = "新木商事" Then
            For c = 1 To UBound(gaibu_data, 2)
                paste_one_gaibu(r, 1) = r - 1      'A列No
                paste_one_gaibu(r, c) = gaibu_data(i, c)
            Next c
            r = r + 1
        End If
    Next i

    'クリアして貼り付け
    Worksheets("外部在庫数形成1日").Activate
    Worksheets("外部在庫数形成1日").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_gaibu, 1), UBound(paste_one_gaibu, 2))) = paste_one_gaibu

    '******************************* 各場所 *****************************************
    '貸倉庫貼り付け
    sh_name = "貸倉庫"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    'スーパーレックス貼り付け
    sh_name = "スーパーレックス"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '新木商事
    sh_name = "新木商事"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    'タドコロ物流
    sh_name = "タドコロ物流"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)

    '自社トラック
    sh_name = "自社トラック"
    Call filter_paste(sh_name, paste_one_aduke, paste_one_modoshi)
    
    '全シート再計算
    Application.Calculate
    
    '指定日付でフィルター
    Worksheets("移動明細").Activate
    Call フィルタークリア
    Range("E4").AutoFilter Field:=3, Criteria1:="<>"
    
    Worksheets("貸倉庫").Activate
    Call フィルタークリア
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("スーパーレックス").Activate
    Call フィルタークリア
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("新木商事").Activate
    Call フィルタークリア
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("タドコロ物流").Activate
    Call フィルタークリア
    Range("D6").AutoFilter Field:=11, Criteria1:="<>"

    Worksheets("自社トラック").Activate
    Call フィルタークリア
    Range("D7").AutoFilter Field:=11, Criteria1:="<>"
    
    '戻る
    Worksheets("移動明細").Activate
    
    Call パレット削除
    
    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
    ActiveSheet.Protect      '保護
    
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
    date_G2 = Worksheets("移動明細").Range("G2") '日付
    
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

Function csvファイル名探索(csvFilePath As Variant) As Variant
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
            filename(cnt, 0) = f.Name
            filename(cnt, 1) = f.DateLastModified   '作成日はf.DateCreated 更新日はf.DateLastModified  アクセス日はf.DateLastAccessed
            cnt = cnt + 1
        Next f
    End With
    
    csvファイル名探索 = filename
    
End Function
Sub csv_main()

    'アクティブ
    Workbooks("在庫金額.xlsm").Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
'    ActiveSheet.Unprotect      '保護解除
    
    'ファイル取得在庫数
    csvFilePath_aduke = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\在庫数\"
    file_list_aduke = csvファイル名探索(csvFilePath_aduke)
    
    'ファイル取得外部在庫数
    csvFilePath_modoshi = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\外部在庫数\"
    file_list_modoshi = csvファイル名探索(csvFilePath_modoshi)
    
    'csv貼り付け_在庫数
    sh_name = "在庫数csv"
    Dim zaiko_data As Variant
    zaiko_data = getCSV_utf8(sh_name, file_list_aduke, csvFilePath_aduke)
    
    'csv貼り付け_外部在庫数
    sh_name = "外部在庫数csv"
    Dim modoshi_data As Variant
    modoshi_data = getCSV_utf8(sh_name, file_list_modoshi, csvFilePath_modoshi)
    
    'one更新
    Call one_search
    
    Worksheets("棚卸明細表").Activate
    
    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
End Sub

Sub tomorrow_add()
    Range("J3").Select
    Range("J3").Value = DateAdd("d", 1, Range("J3"))
    Call one_search
    
End Sub

Sub yesterday_add()
    Range("J3").Select
    Range("J3").Value = DateAdd("d", -1, Range("J3"))
    Call one_search
End Sub

Sub one_search()
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    ActiveSheet.Unprotect      '保護解除

    date_G2 = Worksheets("棚卸明細表").Range("J3") '日付
    Dim paste_one_zaiko As Variant  '貼り付けデータ
    Dim paste_one_kowake As Variant  '貼り付けデータ

    Workbooks("在庫金額.xlsm").Activate
    '*****************在庫数データ取得************************************
    Worksheets("在庫数csv").Activate
    
    Dim lastRow As Long '最終行取得
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(lastRow, LastCol))    '「在庫数csv」シートのデータ
    
    
    '*****************外部在庫数データ取得************************************
    Worksheets("外部在庫数csv").Activate
    
    
    'Dim LastRow As Long '最終行取得
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    gaibu_row = Range(Cells(1, 1), Cells(lastRow, LastCol))    '「外部在庫数csv」シートのデータ

    '格納する２次元配列サイズ設定
    ReDim paste_one_zaiko(1 To (UBound(zaiko_data, 1) + UBound(gaibu_row, 1)), 1 To UBound(zaiko_data, 2))    '(行,列)
    ReDim paste_one_kowake(1 To (UBound(zaiko_data, 1) + UBound(gaibu_row, 1)), 1 To UBound(zaiko_data, 2))  '(行,列)

    r = 1
    k = 1
    For i = 1 To UBound(zaiko_data)
        If i = 1 Then
            For c = 1 To UBound(zaiko_data, 2)
                paste_one_zaiko(r, c) = zaiko_data(i, c)
                paste_one_kowake(k, c) = zaiko_data(k, c)
            Next c
            r = r + 1   '通常在庫数
            k = k + 1   '小分け
        ElseIf zaiko_data(i, 10) = "在庫数" And zaiko_data(i, 12) = date_G2 Then
            If zaiko_data(i, 2) Like "*小分け品*" Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_kowake(k, 1) = k - 1      'A列No
                    paste_one_kowake(k, c) = zaiko_data(i, c)
                Next c
                k = k + 1
            Else
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = r - 1      'A列No
                    paste_one_zaiko(r, c) = zaiko_data(i, c)
                Next c
                r = r + 1
            End If
        End If
    Next i
    
    '*****************外部在庫数データ取得************************************
    Worksheets("外部在庫数csv").Activate
    
    'Dim LastRow As Long '最終行取得
    lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    'Dim LastCol As Long '最終列取得
    LastCol = Cells(1, 1).End(xlToRight).Column + 5

    zaiko_data = Range(Cells(1, 1), Cells(lastRow, LastCol))    '「外部在庫数csv」シートのデータ

'    r = 1
'    k = 1
    
    For i = 1 To UBound(zaiko_data, 1)
        If i = 1 Then
            ' For c = 1 To UBound(zaiko_data, 2)
            '     paste_one_zaiko(r, c) = zaiko_data(i, c)
            '     paste_one_kowake(k, c) = zaiko_data(k, c)
            ' Next c
'            r = r + 1   '通常在庫数
'            k = k + 1   '小分け
        ElseIf zaiko_data(i, 10) = "外部在庫数" And zaiko_data(i, 12) = date_G2 Then
            If zaiko_data(i, 2) Like "*小分け品*" Then
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_kowake(k, 1) = k - 1      'A列No
                    paste_one_kowake(k, c) = zaiko_data(i, c)
                Next c
                k = k + 1
            Else
                For c = 1 To UBound(zaiko_data, 2)
                    paste_one_zaiko(r, 1) = r - 1      'A列No
                    paste_one_zaiko(r, c) = zaiko_data(i, c)
                Next c
                r = r + 1
            End If
        End If
    Next i
    
    'クリアして貼り付け
    Worksheets("在庫数形成").Activate
    Worksheets("在庫数形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_zaiko, 1), UBound(paste_one_zaiko, 2))) = paste_one_zaiko
     'クリアして貼り付け
    Worksheets("小分け在庫数形成").Activate
    Worksheets("小分け在庫数形成").Cells.ClearContents
    Range(Cells(1, 1), Cells(UBound(paste_one_kowake, 1), UBound(paste_one_kowake, 2))) = paste_one_kowake
 
    '全シート再計算
    Application.Calculate
     
    '戻る
    Worksheets("棚卸明細表").Activate
    
    Application.ScreenUpdating = True                 '画面起動
    Application.Calculation = xlCalculationAutomatic       '自動計算
    ActiveSheet.Protect AllowFiltering:=True       '保護
    
End Sub

Function getCSV_utf8(sh_name As Variant, file_list As Variant, csvFilePath As Variant) As Variant
    
    'Dim ws As Worksheet
    'Set ws = ThisWorkbook.Worksheets(1)
    
    'Dim strPath As String
    'strPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\在庫数\【在庫数】在庫_ダンボール_2022.3.xlsm.csv"
    
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
    'csvFilePath = "\\Afnewt320-kyoyu\社内共有\AFSKS\在庫表\csv\在庫数"
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



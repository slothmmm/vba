Const delimiter = "~"   'ダブルクォーテーション内のカンマを一時的に変える文字指定
Const lineFeedCode = vbCrLf   '読み込むファイルの改行コード指定 CRLF or LF

Sub csv読み込みmain()
    'アクティブ
    ThisWorkbook.Activate
    first_sheet = ActiveSheet.Name
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    '*******************    先に該当データがあるか見る。    **************************************************************************************
    '************************   丸大 csvファイル名探索   ***********************************
    marudai_code = "25726549"        '丸大宛先コード
    csvname_marudai = csv探索(marudai_code)
    
    '************************   IY csvファイル名探索   ***********************************
    IYmain1_code = "25726573"        'IY宛先コード
    csvname_IY = csv探索(IYmain1_code)
    
    '****************************************************************************************************************************************************
    '************************   丸大 csvファイル取得、貼付   ***********************************
    ' 'ファイル取得+貼り付け
    ' paste_sheetname = "受注データcsv"
    ' Call getCSV_utf8(csvname_marudai, paste_sheetname)
    
    '************************   IY csvファイル取得、貼付   ***********************************
    'ファイル取得+貼り付け
    paste_sheetname = "受注データcsv"
    Call getCSV_utf8(csvname_IY,csvname_marudai, paste_sheetname)
    
    '************************   END   ***********************************
    Worksheets(first_sheet).Activate

    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '自動計算
    
    MsgBox ("受注データcsv取り込み完了しました。")
    
End Sub

Function csv探索(csvname) As Variant
    'csvPath = "\\192.168.100.105\新rev_files"
    csvPath = "\\Afnewt320-kyoyu\社内共有\個人フォルダ\笠間\IYテスト\csv"
    
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
        MsgBox ("csvファイルがありません。" & vbCrLf & vbCrLf & "csvのパス : " & csvPath)
        End
    End If
    
    ReDim filename(s - 1, 1) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvPath).Files
            If f.Name Like "*" + csvname + "*" + "juchu.csv" Then
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
    
    If IsEmpty(filename(Max, 0)) Then
        MsgBox ("csvファイルがありません。" & vbCrLf & vbCrLf & "csvのパス : " & csvPath & vbCrLf & vbCrLf & "csvファイル名に 「" & csvname & "」が含まれているcsvファイルがありません。")
        End
    End If
    
    csv探索 = csvPath & "\" & filename(Max, 0)
    
End Function

Sub getCSV_utf8(csvname_IY,csvname_marudai, paste_sheetname)
    Dim i As Long, j As Long
    Dim strLines, arrLine, strLine As Variant
    Dim strAll As String
    Dim strBuf As String

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")


    '***************************   IYの csvの行数をカウントする ***********************************
    row_no_IY = -2
    With adoSt
        .Charset = "shift_jis"  'Streamで扱う文字コートをutf-8に設定
        .Open   'Streamをオープン
        .LoadFromFile (csvname_IY) 'ファイルからStreamにデータを読み込む

        strAll = .ReadText(adReadAll)
        strLines = Split(strAll, lineFeedCode)
        For Each strLine In strLines    'Streamの末尾まで繰り返す
            row_no_IY = row_no_IY + 1
        Next strLine

        .Close
    End With

    strAll = ""
    strLines = ""
    
    If row_no_IY <= 0 Then
        MsgBox ("csvのデータ行数が0です！終了します。" & vbCrLf & vbCrLf & csvname_IY)
        End     'csvのデータ行数が0なら終了
    ' Else
    '     '格納する２次元配列サイズ設定
    '     ReDim paste_data(1 To row_no_IY, 1 To 60) '(行,列)
    End If

    '***************************   丸大の csvの行数をカウントする ***********************************
    row_no_maru = -2
    With adoSt
        .Charset = "shift_jis"  'Streamで扱う文字コートをutf-8に設定
        .Open   'Streamをオープン
        .LoadFromFile (csvname_marudai) 'ファイルからStreamにデータを読み込む

        strAll = .ReadText(adReadAll)
        strLines = Split(strAll, lineFeedCode)
        For Each strLine In strLines    'Streamの末尾まで繰り返す
            row_no_maru = row_no_maru + 1
        Next strLine

        .Close
    End With

    strAll = ""
    strLines = ""
    
    If row_no_maru <= 0 Then
        MsgBox ("csvのデータ行数が0です！終了します。" & vbCrLf & vbCrLf & csvname_marudai)
        End     'csvのデータ行数が0なら終了
    Else
        '格納する２次元配列サイズ設定
        ReDim paste_data(1 To row_no_maru + row_no_IY, 1 To 60) '(行,列)
    End If

    ' ***************************   IYの csvのデータを配列に格納する ***********************************
    i = 1
    With adoSt
        '.Charset = "UTF-8"  'Streamで扱う文字コートをutf-8に設定
        .Charset = "shift_jis"  'Streamで扱う文字コートをutf-8に設定
        .Open   'Streamをオープン
        .LoadFromFile (csvname_IY) 'ファイルからStreamにデータを読み込む

        strAll = .ReadText(adReadAll)
        strLines = Split(strAll, lineFeedCode)
        For Each strLine In strLines    'Streamの末尾まで繰り返す

            If strBuf <> "" Then
                strBuf = strBuf & lineFeedCode & strLine
            Else
                strBuf = strLine
            End If
            If double_quotation_count(strBuf) Mod 2 = 0 Then
              arrLine = Split(Replace(replaceDelimiter(strBuf), """", ""), delimiter) 'strLineをカンマで区切りarrLineに格納
                If i <> 1 Then  '1行目はスキップ
                    For j = 0 To UBound(arrLine)
                        paste_data(i - 1, j + 1) = arrLine(j)
                    Next j
                End If
              i = i + 1
              strBuf = ""
            End If

        Next strLine

        .Close
    End With

    IY_csv_date = paste_data(i-1,19)

    ' ***************************   丸大の csvのデータを配列に格納する ***********************************
    i = 1
    With adoSt
        '.Charset = "UTF-8"  'Streamで扱う文字コートをutf-8に設定
        .Charset = "shift_jis"  'Streamで扱う文字コートをutf-8に設定
        .Open   'Streamをオープン
        .LoadFromFile (csvname_marudai) 'ファイルからStreamにデータを読み込む

        strAll = .ReadText(adReadAll)
        strLines = Split(strAll, lineFeedCode)
        For Each strLine In strLines    'Streamの末尾まで繰り返す

            If strBuf <> "" Then
                strBuf = strBuf & lineFeedCode & strLine
            Else
                strBuf = strLine
            End If
            If double_quotation_count(strBuf) Mod 2 = 0 Then
              arrLine = Split(Replace(replaceDelimiter(strBuf), """", ""), delimiter) 'strLineをカンマで区切りarrLineに格納
                If i <> 1 Then  '1行目はスキップ
                    For j = 0 To UBound(arrLine)
                        paste_data(i - 1, j + 1) = arrLine(j)
                    Next j
                End If
              i = i + 1
              strBuf = ""
            End If

        Next strLine

        .Close
    End With

    marudai_csv_date = paste_data(i-1,19)

    '********************************* 取り込みデータのチェック *****************************************
    if IY_csv_date <> marudai_csv_date then
        MsgBox ("IYと丸大のcsvのデータが異なります！終了します。" & vbCrLf & vbCrLf & csvname_IY & vbCrLf& vbCrLf & csvname_marudai)
        End     'csvのデータ行数が0なら終了
    end if

    '********************************* 取り込み済みか判定 *****************************************
    ThisWorkbook.Activate
    Worksheets(paste_sheetname).Activate
    
    B_LastRow = Cells(Rows.Count, 2).End(xlUp).Row          'B列の最終行取得
    column_s = Range(Cells(2, 19), Cells(B_LastRow, 19))
    
    csv_date = paste_data(1, 19)
    
    For i = 1 To UBound(column_s)
        If CStr(paste_data(1, 19)) = CStr(column_s(i, 1)) Then
            MsgBox ("既に取り込み済みです！" & vbCrLf & vbCrLf & "csvの発注日 : " & CStr(paste_data(1, 19)) & vbCrLf & paste_sheetname & "シートの発注日 : " & CStr(column_s(i, 1)) & vbCrLf & vbCrLf & "csvのパス : " & csvname_IY)
            End
        End If
    Next i
    
    '貼り付け
    
    ' '丸大は背景色
    ' If csvname_IY Like "*25726549*" Then
    '     'Range(Cells(B_LastRow + 1, 1), Cells(B_LastRow + row_no_IY, 60)).Interior.Color = RGB(200, 100, 0)
    '     Range(Cells(B_LastRow + 1, 1), Cells(B_LastRow + row_no_IY, 60)).Interior.ColorIndex = 20
    ' Else
    
    End If
    
    Range(Cells(B_LastRow + 1, 1), Cells(B_LastRow + row_no_IY, 60)) = paste_data

End Sub

'受け取った文字列のカンマをdelimiterに置き換える
'ダブルクォーテーションで囲まれているカンマは置き換えない
Function replaceDelimiter(ByVal str As String) As String

    Dim strTemp As String
    Dim quotCount As Long

    Dim l As Long
    For l = 1 To Len(str)  'strの長さだけ繰り返す

        strTemp = Mid(str, l, 1) 'strから現在の1文字を切り出す

        If strTemp = """" Then   'strTempがダブルクォーテーションなら
            quotCount = quotCount + 1   'ダブルクォーテーションのカウントを1増やす
        ElseIf strTemp = "," Then   'strTempがカンマなら
            If quotCount Mod 2 = 0 Then   'quotCountが2の倍数なら
                str = Left(str, l - 1) & delimiter & Right(str, Len(str) - l)   '現在の1文字をdelimiterに置き換える
            End If
        End If

    Next l

    replaceDelimiter = str

End Function

Function double_quotation_count(target)
    Dim buf As String, i As Long, cnt As Long
    For i = 1 To Len(target)
        If Mid(target, i, 1) = """" Then cnt = cnt + 1
    Next i
    double_quotation_count = cnt
End Function






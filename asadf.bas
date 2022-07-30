Const delimiter = "~"   'ダブルクォーテーション内のカンマを一時的に変える文字指定
Const lineFeedCode = vbLf   '読み込むファイルの改行コード指定 CRLF or LF

Sub csv読み込みmain()

    'アクティブ
    ThisWorkbook.Activate
    first_sheet = ActiveSheet.Name
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    'range("N2") 検索値
    Worksheets("ラベル").Activate
    csvname = Range("N2").Value
    
    csvname = csv探索(csvname)
    
    'ファイル取得 計画数
    'csvFilePath_lavel = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM\"
    paste_sheetname = "csv"
    Call getCSV_utf8(csvname, paste_sheetname)
    
    'ラベル100x80
    Worksheets("ラベル").Activate
    Call 印刷.フィルターIP67

    'コンテナ明細票
    Worksheets("コンテナ明細票").Activate
    Call 印刷.フィルターIP67

    Worksheets(first_sheet).Activate

    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
    
End Sub

Function csv探索(csvname) As Variant
    csvPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM"
    
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
        MsgBox "csvファイルが空です。"
        End
    End If
    
    ReDim filename(s - 1, 1) '【0】ファイル名.csv【1】作成日【2】年(ファイル名)【3】月0埋め(ファイル名)【4】月0埋め(ファイル名)
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(csvPath).Files
            If f.Name Like "*" + csvname + "*" Then
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
   
    csv探索 = csvPath & "\" & filename(Max, 0)
    
End Function


Sub getCSV_utf8(strPath, paste_sheetname)

    '格納する２次元配列サイズ設定
    ReDim paste_data(1 To 1000, 1 To 60) '(行,列)

    'strPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM\混載ラベル_フレッセイ_220702_161717.csv"
    Dim i As Long, j As Long
    Dim strLines, arrLine, strLine As Variant
    Dim strAll As String
    Dim strBuf As String

    'ADODB.Streamオブジェクトを生成
    Dim adoSt As Object
    Set adoSt = CreateObject("ADODB.Stream")

    i = 1
    With adoSt
        .Charset = "UTF-8"  'Streamで扱う文字コートをutf-8に設定
        .Open   'Streamをオープン
        .LoadFromFile (strPath) 'ファイルからStreamにデータを読み込む

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

              For j = 0 To UBound(arrLine)
                  paste_data(i, j + 1) = arrLine(j)
              Next j

              i = i + 1
              strBuf = ""
            End If

        Next strLine

        .Close
    End With

        'シートクリア
    ThisWorkbook.Activate
    Worksheets(paste_sheetname).Activate
    Worksheets(paste_sheetname).Cells.ClearContents
    
    Range(Cells(1, 1), Cells(1000, 60)) = paste_data

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



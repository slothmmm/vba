Const delimiter = "~"   'ダブルクォーテーション内のカンマを一時的に変える文字指定
Const lineFeedCode = vbLf   '読み込むファイルの改行コード指定 CRLF or LF

Sub getCSV_utf8()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("csv")

    '格納する２次元配列サイズ設定
    ReDim paste_data(1 To 1000, 1 To 60) '(行,列)

    Dim strPath As String

    strPath = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM\混載ラベル_フレッセイ_220702_161717.csv"
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
                  paste_data(i, j + 1).Value = arrLine(j)
              Next j

              i = i + 1
              strBuf = ""
            End If

        Next strLine

        .Close
    End With

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


Sub csv_main()
    'アクティブ
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    'range("N2") 検索値
    Worksheets("ラベル60x80").Activate
    csv_name = range("N2").Value
    'ファイル取得 計画数
    csvFilePath_lavel = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM\"
    file_list_cope = csvファイル名探索(csvFilePath_lavel,csv_name)
    
    '
    sh_name = "csv"
    Dim cope_data As Variant
    cope_data = getCSV_utf8(sh_name, file_list_cope, csvFilePath_lavel)
    
    Application.ScreenUpdating = True                 '画面停止
    Application.Calculation = xlCalculationAutomatic       '手動計算
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
    ReDim paste_data(1 To max_n, 1 To 60) '(行,列)
    
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



Sub 前月シートコピーmain()    
    Dim rc As VbMsgBoxResult
    rc = MsgBox("このシートを削除して、前月シートからコピーしますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "中止します", vbCritical
        End
    End If

    copy_moto_sh = 翌月シートチェック
    this_sheet_name = ActiveSheet.Name  '現在のシート名を格納
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ActiveSheet.Delete                  '現在のシート削除
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Worksheets(copy_moto_sh).Copy before:=Worksheets(copy_moto_sh)  '前月シート複製
    ActiveSheet.Name = this_sheet_name                              'シート名を戻す
    Range("C4") = DateAdd("m", 1, Range("C4"))                      '日付変更
    
    Debug.Print "コピー元 : " & copy_moto_sh
                
    Call 実績数式の置換
    MsgBox ("前月からシートコピー完了しました。")
End Sub

Function 翌月シートチェック() As Variant
    
    this_sheet_num = Replace(ActiveSheet.Name, "月", "")
    
    '現在のシート名チェック
    If Not IsNumeric(this_sheet_num) Then   'このシート名が数値化出来るか
        MsgBox ("このシート名に余計な文字列が含まれています" & vbCrLf & "「1月」～「12月」と指定して下さい。" & vbCrLf & vbCrLf & "このシート名 : " & ActiveSheet.Name)
        End
    ElseIf Not (1 <= this_sheet_num And this_sheet_num <= 12) Then
        MsgBox ("シート名を1～12月にしてください。") & vbCrLf & vbCrLf & "このシート名 : " & ActiveSheet.Name
        End
    ElseIf this_sheet_num Like "*.*" Then
        MsgBox (".を含めないでください。") & vbCrLf & vbCrLf & "このシート名 : " & ActiveSheet.Name
        End
    End If
    
    'コピー元シート名作成
    copy_moto_sh = Trim(Month(DateAdd("m", -1, "2000/" & this_sheet_num & "/1"))) & "月"
    
    'コピー元のシート名が存在するかどうか
    Dim ws As Worksheet, flag As Boolean
    For Each ws In Worksheets
        If ws.Name = copy_moto_sh Then flag = True
    Next ws
    If flag = False Then
        MsgBox "コピー元となる前月のシート" & vbCrLf & vbCrLf & "「" & copy_moto_sh & "」" & "シートがありません"
        End
    End If
    
    '前月のC4セルに日付けが入っているか
     If Not IsDate(Worksheets(copy_moto_sh).Range("C4")) Then
        MsgBox "「" & copy_moto_sh & "」シートC4セルに 「yyyy/mm/mm」形式で日付を入力して下さい。"
        End
     End If

    翌月シートチェック = copy_moto_sh
End Function


Sub 実績数式の置換()
'    Dim rc As VbMsgBoxResult
'    rc = MsgBox("実績の数式の置換を行います。対象列 : E列～AI列", vbYesNo + vbQuestion)
'    If rc = vbNo Then
'        MsgBox "中止します", vbCritical
'        End
'    End If
    
    Debug.Print "実績の数式の置換開始"

    '***********************************シートチェック**********************************************************
    '最終行取得
    Dim E_LastRow As Long
    E_LastRow = Cells(Rows.Count, 5).End(xlUp).Row

    '値を配列へ格納     A列からAN列
    value_cell = Range(Cells(1, 1), Cells(E_LastRow, 35))
    
    'E列で日付がある検索し、ヒットしたら３１個の日付データがある行か検索。その行数を格納。
    E_date_row = 0
    
    For i = 1 To E_LastRow
        If IsError(value_cell(i, 5)) Then
            GoTo skipError
        End If
        'E列に日付データがあった場合
        If IsDate(value_cell(i, 5)) Then
            row_date_cnt = 1            '行で検索し、31個の日付データがあるか
            For d = 1 To 30
                If IsDate(value_cell(i, d + 5)) Then
                    row_date_cnt = row_date_cnt + 1
                End If
            Next d
            
            If row_date_cnt = 31 Then
                E_date_row = i                  'E列の日付けの行数
                Debug.Print "３１個の日付データが見つかりました。"
                Exit For
            End If
        End If

        Debug.Print str(i) & "行目 可能かどうか" & str(IsDate(value_cell(i, 5))) & "長さ : " & str(Len(value_cell(i, 5))) & "  値 : " & value_cell(i, 5)
        
skipError:  'エラーの場合はここへ
    Next i
    
    If E_date_row = 0 Then
        MsgBox ("日付セルが３１個ありません。通常は8行目に１日～３１日（←は例）のデータがあります。前月シートを見直して下さい。")
        End
    End If
    
    
    '***********************************置換開始**********************************************************
    '数式を配列へ格納    A列からAN列
    tikan_cell = Range(Cells(1, 1), Cells(E_LastRow, 35)).Formula
    
    Dim E_nen As regMsg   ' オブジェクト型の変数を宣言
    Dim E_tsuki As regMsg   ' オブジェクト型の変数を宣言
    
    Dim F_AN_nen As regMsg   ' オブジェクト型の変数を宣言
    Dim F_AN_tsuki As regMsg   ' オブジェクト型の変数を宣言
    
    '***********************E列***********************
    '数式の「****年」「**月」は何か
    For i = 1 To UBound(tikan_cell)
        If tikan_cell(i, 5) Like "*年*" Then
            Set E_nen = 正規表現_数式_年(tikan_cell(i, 5))    ' インスタンス生成
            Set E_tsuki = 正規表現_数式_月(tikan_cell(i, 5))    ' インスタンス生成
            Exit For
        End If
    Next i
    
    
    E8_nen = Trim(str(Year(value_cell(E_date_row, 5)))) & "年"    '置き換える年
    E8_tsuki = Trim(str(Month(value_cell(E_date_row, 5)))) & "月"  '置き換える月
    
    'E列置換
    
    For i = 1 To UBound(tikan_cell)
        tikan_cell(i, 5) = Replace(tikan_cell(i, 5), E_nen.Value, E8_nen)
        tikan_cell(i, 5) = Replace(tikan_cell(i, 5), E_tsuki.Value, E8_tsuki)
    Next i
    
    '***********************F列からAN列***********************
     '数式の「****年」「**月」は何か
    For i = 1 To UBound(tikan_cell)
        If tikan_cell(i, 6) Like "*年*" Then
            Set F_AN_nen = 正規表現_数式_年(tikan_cell(i, 6))    ' インスタンス生成
            Set F_AN_tsuki = 正規表現_数式_月(tikan_cell(i, 6))    ' インスタンス生成
            Exit For
        End If
    Next i
    
    F8_AN8_nen = Trim(str(Year(value_cell(E_date_row, 6)))) & "年"    '置き換える年
    F8_AN8_tsuki = Trim(str(Month(value_cell(E_date_row, 6)))) & "月"  '置き換える月
    
    'F列からAN列置換
    For i = 1 To UBound(tikan_cell)
        For s = 1 To 30
            tikan_cell(i, 5 + s) = Replace(tikan_cell(i, 5 + s), F_AN_nen.Value, F8_AN8_nen)
            tikan_cell(i, 5 + s) = Replace(tikan_cell(i, 5 + s), F_AN_tsuki.Value, F8_AN8_tsuki)
        Next s
    Next i
    
    '数式を貼付け
    Range(Cells(1, 1), Cells(E_LastRow, 35)) = tikan_cell
    'MsgBox ("実績の数式置換が終了しました。")
    
End Sub


Function 正規表現_数式_年(formula_str As Variant) As regMsg
    'formula_str = "aaaaaa2021年abc" '対象文字列
    
    Dim nen As regMsg   ' オブジェクト型の変数を宣言
    Set nen = New regMsg    ' インスタンス生成
    
    'RegExpオブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    '正規表現の指定
    With reg
        .Pattern = "\d{4}年"      'パターンを指定
        .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
        .Global = True          '文字列全体を検索するか(True)、しないか(False)
    End With
    
    Dim Matches
    Set Matches = reg.Execute(formula_str) '正規表現でのマッチングを実行
    
    For Each Match In Matches
        nen.Value = Match.Value '値
        nen.FirstIndex = Match.FirstIndex '開始位置
        nen.Length = Match.Length   '長さ
    Next Match
    
    Set 正規表現_数式_年 = nen
    
End Function

Function 正規表現_数式_月(formula_str As Variant) As regMsg
    'formula_str = "aaaaaa2021月abc" '対象文字列
    
    Dim tsuki As regMsg   ' オブジェクト型の変数を宣言
    Set tsuki = New regMsg    ' インスタンス生成
    
    'RegExpオブジェクトの作成
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    '正規表現の指定
    With reg
        .Pattern = "\d{2}月|\d{1}月"      'パターンを指定 1桁or2桁
        .IgnoreCase = False     '大文字と小文字を区別するか(False)、しないか(True)
        .Global = True          '文字列全体を検索するか(True)、しないか(False)
    End With
    
    Dim Matches
    Set Matches = reg.Execute(formula_str) '正規表現でのマッチングを実行
    
    For Each Match In Matches
        tsuki.Value = Match.Value '値
        tsuki.FirstIndex = Match.FirstIndex '開始位置
        tsuki.Length = Match.Length   '長さ
    Next Match
    
    Set 正規表現_数式_月 = tsuki
    
End Function



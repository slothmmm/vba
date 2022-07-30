Sub RPA_CSV読み込みmain()
    'アクティブ
    ThisWorkbook.Activate
    
    Application.ScreenUpdating = False                 '画面停止
    Application.Calculation = xlCalculationManual     '手動計算
    'ActiveSheet.Unprotect      '保護解除
    
    Worksheets("ラベル60x80").Activate
    csv_name = range("N2").Value
    
    csv_data = RPA_CSV読み込み(csv_name)       '該当のcsvデータ
    
    
    
    '非表示
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Application.Calculation = xlCalculationAutomatic    '自動計算
    Application.ScreenUpdating = True                 '画面
    
End Sub


Function RPA_CSV読み込み(csv_name As Variant) As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0
  
  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = RPA_CSVファイル名探索(csv_name)
  

  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる
  
  ReDim ary(max_n - 1, 30) As Variant '取得した行数で2次元配列の再定義
    
  Open file For Input As #1 'CSVファイルを開く
      Do Until EOF(1) '最終行までループ
      Line Input #1, buf '読み込んだデータを1行ずつみていく
      
      '↓ダブルコーテーション無しのcsv
      'tmp = Split(buf, ",") 'カンマで分割
      '↓ダブルコーテーション有りのcsv
      tmp = Split(Replace(buf, """", ""), ",") 'strLineをカンマで区切りarrLineに格納

      For i = 0 To UBound(tmp) '項目数ぶんループ
        ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
      Next i
      n = n + 1 '配列の次の行へ
    Loop
  Close #1 'CSVファイルを閉じる
  
    RPA_CSV読み込み = ary

End Function


Function RPA_重複なし貴社商品CDリスト(csv_data As Variant) As Variant
    
    Dim 辞書 As Object
    Set 辞書 = CreateObject("Scripting.Dictionary")
   
    Dim dataList As Variant
    ReDim dataList(UBound(csv_data) - 1, 0) As Variant

    For b = 0 To UBound(csv_data) - 1
        dataList(b, 0) = csv_data(b + 1, 16)
    Next b
   
    For i = 0 To UBound(dataList)
        '辞書に登録されていない時は
          If Not 辞書.Exists(dataList(i, 0)) Then
              '辞書に登録する。
              辞書(dataList(i, 0)) = Empty
          End If
    Next i
    
    RPA_重複なし貴社商品CDリスト = 辞書.keys
    
End Function

Function RPA_CSVファイル名探索(csv_name As Variant) As Variant

    csvFilePath = "\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\ラベル\SATOFM"
    'ディレクトリ存在チェック
    If Dir(csvFilePath, vbDirectory) = "" Then
        MsgBox csvFilePath & "ディレクトリが存在しません。"
        End
    Else
        Debug.Print csvFilePath & "ディレクトリが存在します。"
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
        MsgBox csvFilePath & "のcsvファイルが空です。"
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
    
    Dim Max As Integer
    Max = 9999 '初期値を設定 下記のfor文で9999のままなら、csvデータに該当の出荷日がなかったということになる。
    For i = 0 To UBound(filename)
        If filename(i, 0) Like "*"+ csv_name +"*" And Right(filename(i, 0), 3) = "csv" Then
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
        MsgBox csv_name + "csvファイルがありません。"
        End
    End If
    
    RPA_CSVファイル名探索 = csvFilePath & "\" & filename(Max, 0)
    
End Function


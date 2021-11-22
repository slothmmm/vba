Function csv読み込み() As Variant
  Dim file As String, max_n As Long
  Dim buf As String, tmp As Variant, ary() As Variant
  Dim i As Long, n As Long, val As Long
  max_n = 0

  '準備
  'file = "C:\test.csv" 'ファイル指定
  file = "\\192.168.100.105\新rev_files\order_YOK.csv"

'  max_n = CreateObject("Scripting.FileSystemObject").OpenTextFile(file, 8).Line 'ファイルの行数取得

  Open file For Input As #1 'CSVファイルを開く
        Do Until EOF(1)
            Line Input #1, buf
            max_n = max_n + 1
        Loop
  Close #1 'CSVファイルを閉じる

  ReDim ary(max_n - 1, 2) As Variant '取得した行数で2次元配列の再定義

  Open file For Input As #1 'CSVファイルを開く
      Do Until EOF(1) '最終行までループ
      Line Input #1, buf '読み込んだデータを1行ずつみていく
      tmp = Split(buf, ",") 'カンマで分割
      For i = 0 To UBound(tmp) '項目数ぶんループ
        ary(n, i) = tmp(i) '分割した内容を配列の項目へ入れる（0→ID, 1→名称, 2→値）
      Next i
      n = n + 1 '配列の次の行へ
    Loop
  Close #1 'CSVファイルを閉じる

    csv読み込み = ary
'
'  For i = 1 To UBound(ary)
'    Debug.Print ary(i, 0)
'  Next
End Function
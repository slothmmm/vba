Sub test()
    csvname = "セイミヤ"
    aaa = csv探索(csvname)
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

Attribute VB_Name = "共通マクロ読み込み"
' ワークブックを開く時のイベント
Sub load_macro_main()
    ' txtに書いてある外部ライブラリを読み込み
    load_from_conf ".\共通マクロ\lib.txt"
    MsgBox "読み込み終了しました"
End Sub



' -------------------- モジュール読み込みに関する関数 --------------------



' 設定ファイルに書いてある外部ライブラリを読み込みます。
Sub load_from_conf(conf_path)
    
    ' 全モジュールを削除
    clear_modules
    
    ' 絶対パスに変換
    conf_path = abs_path(conf_path)
    If Dir(conf_path) = "" Then
        MsgBox "外部ライブラリ定義" & conf_path & "が存在しません。"
        Exit Sub
    End If
    
    ' 読み取り
    fp = FreeFile
    Open conf_path For Input As #fp
    Do Until EOF(fp)
        ' １行ずつ
        Line Input #fp, temp_str
        If Len(temp_str) > 0 Then
            module_path = abs_path(temp_str)
            If Dir(module_path) = "" Then
                ' エラー
                MsgBox "モジュール" & module_path & "は存在しません。"
                Exit Do
            Else
                ' モジュールとして取り込み
                include module_path
            End If
        End If
    Loop
    Close #fp

    ThisWorkbook.Save
    
End Sub


' あるモジュールを外部から読み込みます。
' パスが.で始まる場合は，相対パスと解釈されます。
Sub include(file_path)
    ' 絶対パスに変換
    file_path = abs_path(file_path)
    
    ' 標準モジュールとして登録
    ThisWorkbook.VBProject.VBComponents.Import file_path
End Sub


' 全モジュールを初期化します。
Private Sub clear_modules()
    On Error Resume Next
    With ThisWorkbook.VBProject.VBComponents
        .Remove .Item("在庫表作成")
        .Remove .Item("翌月置換")
        .Remove .Item("保護")
        .Remove .Item("dateClass")
        .Remove .Item("fileClass")
        .Remove .Item("colorClass")
    End With
End Sub

' ファイルパスを絶対パスに変換します。
Function abs_path(file_path)
    ' 絶対パスに変換
    If Left(file_path, 1) = "." Then
        file_path = ThisWorkbook.Path & Mid(file_path, 2, Len(file_path) - 1)
    End If
    
    abs_path = file_path

End Function



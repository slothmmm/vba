Sub sheet_copy()
    Dim FilePath As String
    Dim FileName As String
    Dim wb As Workbook
    
    Application.ScreenUpdating = False
    
    'ファイルの入っているフォルダをパスを設定
    FilePath = ThisWorkbook.Path
    'ファイル名を設定
    FileName = "aaa.xlsm"
    
    'コピー元のブックが存在するか確認
   If Dir(FilePath & "\" & FileName) = "" Then
      '既に開いていたらメッセージを表示してSubを抜ける
      MsgBox FileName & "というファイルが存在しません" & vbCrLf & _
         "指定のフォルダに該当のファイルを入れて実行し直してください"
      Exit Sub
   End If
   
   '既に開いているかをチェック
   For Each wb In Workbooks
      If wb.Name = FileName Then
         '既に開いていたらメッセージを表示してSubを抜ける
         MsgBox FileName & "は既に開いています"
         Exit Sub
      End If
   Next wb
    
    
    
    'ブックを開く
    Workbooks.Open FilePath & "\aaa.xlsm" 'コピー元のブックのパス
    'シートをコピー
    Workbooks("aaa.xlsm").Worksheets("a").Copy after:=ThisWorkbook.Worksheets("c")
    
    'ブックを閉じる
    Application.DisplayAlerts = False
    'コピー元のブックを閉じる(セーブしない)
    Workbooks("aaa.xlsm").Close savechanges:=False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

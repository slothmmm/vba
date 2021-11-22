Attribute VB_Name = "自動"
Sub 自動()
        ThisWorkbook.Activate
        Call 保護.全保護解除
        Worksheets("設定").Range("X1").Value = Now
End Sub

Sub ばっくあっぷ()
    On Error GoTo Error1 '一応エラー対策
    Dim res As Integer
    Dim CopyBook As String
    Dim hiduke As String
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
        
    'Backup後の保存先＆ブック名
    CopyBook = ThisWorkbook.Path & "\バックアップ\" & "BackupFile" & hiduke & ".xlsm"
    ActiveWorkbook.SaveCopyAs CopyBook  'Backup保存するコード

    Exit Sub
Error1:         'エラーが発生した場合はここへ飛ぶ
    MsgBox "エラー番号:" & Err.Number & vbLf & _
    "エラー内容：" & Err.Description & vbLf
    Exit Sub
End Sub

Sub 日付の入力()
    Worksheets("設定").Range("A1").Value = Format(Year(Now), "00")
    Worksheets("設定").Range("C1").Value = Format(Month(Now), "00")
    Worksheets("設定").Range("E1").Value = Format(Day(Now), "00")
End Sub

Sub 使用中()
    Dim use_file_name(2) As String
    use_file_name(1) = "使用中.txt"
    use_file_name(2) = "空いてます.txt"
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo myError
    FSO.GetFile(ThisWorkbook.Path & "\" & use_file_name(2)).Name = use_file_name(1)
    Exit Sub
    Set FSO = Nothing
    
myError:
    MsgBox "「使用中」「 空いてます」 が機能していない可能性あり(無視して良い)"
End Sub

Sub 空いてます()
    Dim use_file_name(2) As String
    use_file_name(1) = "使用中.txt"
    use_file_name(2) = "空いてます.txt"
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error GoTo myError
    FSO.GetFile(ThisWorkbook.Path & "\" & use_file_name(1)).Name = use_file_name(2)
    Exit Sub
    Set FSO = Nothing
    
myError:
    MsgBox "「使用中」「 空いてます」 が機能していない可能性あり(無視して良い)"
End Sub


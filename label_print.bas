Sub CGCカード印刷()
   
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
    Dim i As Integer
    Dim LAST As Integer
    'Dim myPrinter As String

    'myPrinter = Application.ActivePrinter
    'Application.ActivePrinter = "SATO SG408R-ex_190 on Ne01:"
    
    ActiveSheet.PrintPreview
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        'Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Range("A4") = "=MAX(M!$A:$A)"
    
    LAST = Worksheets("出力").Range("A4")
    
    For i = 1 To LAST
      
    Range("A2") = i
    
    ActiveSheet.Range("$A$5:$X$2270").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False, ActivePrinter:="SATO SG408R-ex_190"
    
    Next i
    
    'Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
End Sub
Sub CGCカード再印刷()
        
    If MsgBox("発行範囲指定はしましたか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
    
        Exit Sub
               
    End If
  
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
    Dim i As Integer
    Dim START As Integer
    Dim LAST As Integer
    Dim myPrinter As String

    'myPrinter = Application.ActivePrinter

    'Application.ActivePrinter = "SATO SG408R-ex_190 on Ne01:"
    
    ActiveSheet.PrintPreview
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        'Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    START = Worksheets("出力").Range("A2")
    LAST = Worksheets("出力").Range("A4")
    
    For i = START To LAST
      
    Range("A2") = i
    
    ActiveSheet.Range("$A$5:$X$2270").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False, ActivePrinter:="SATO SG408R-ex_190"
    
    Next i
    
    'Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
End Sub
Sub 予備カード印刷()
   
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
    Dim myPrinter As String
 
    'myPrinter = Application.ActivePrinter
    'Application.ActivePrinter = "SATO SG408R-ex_190 on Ne01:"
    
    ActiveSheet.PrintPreview
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        'Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False, ActivePrinter:="SATO SG408R-ex_190"
    
    'Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
End Sub
Sub 日付調整()
   
    Range("k3") = "='\\Afnewt320-kyoyu\社内共有\AFSKS\ピッキング表\[CGC_ピッキング表.xlsm]ピッキング表'!$D$6"
    
End Sub
Sub 発行範囲指定()

Dim ans1 As String
Dim ans2 As String

    ans1 = InputBox("印刷開始№を入力してください", "印刷範囲確認", "")

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"

    Range("A2").Value = ans1
    
    ans2 = InputBox("印刷終了№を入力してください", "印刷範囲確認", "")

    Range("A4").Value = ans2
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
    False, Password:="0001"
    

End Sub
Sub もしもの場合()

    If MsgBox("読み取り専用で開いていますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
             
            Exit Sub
               
    End If

    Sheets("M").Select
    ActiveSheet.Unprotect

    MsgBox ("データのリンク編集を行ってください")

End Sub
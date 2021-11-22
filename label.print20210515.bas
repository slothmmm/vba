Sub コープカード印刷()
   
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
    Dim i As Integer
    Dim LAST As Integer
       
    '20210515
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    
    If myPrinter Like "*195*" Then
    Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定をIP195へ変更して下さい。"
        Exit Sub
    End If
    

    'Application.ActivePrinter = "SATO SG408R-ex_195 on Ne00:"
    'Application.ActivePrinter = "SATO SG408R-ex_IP195 on Ne02:"
    
    ActiveSheet.PrintPreview
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
        
    Range("A4") = "=MAX(M!$A:$A)"
    
    LAST = Worksheets("出力").Range("A4")
    
    For i = 1 To LAST
      
    Range("A2") = i
          
    ActiveSheet.Range("$A$5:$X$4025").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="SATO SG408R-ex_IP195", Collate:=True, _
    IgnorePrintAreas:=False
    
    Next i
    
    Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
End Sub
Sub CGCカード再印刷()
    '20210515
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    
    If myPrinter Like "*195*" Then
    Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定をIP195へ変更して下さい。"
        Exit Sub
    End If
        
        
    If MsgBox("発行範囲指定はしましたか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
    
        Exit Sub
               
    End If
  
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
    Dim i As Integer
    Dim START As Integer
    Dim LAST As Integer
    
    'Dim myPrinter As String

    'myPrinter = Application.ActivePrinter

    'Application.ActivePrinter = "SATO SG408R-ex_IP195 on Ne02:"
    
    ActiveSheet.PrintPreview
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    START = Worksheets("出力").Range("A2")
    LAST = Worksheets("出力").Range("A4")
    
    For i = START To LAST
      
    Range("A2") = i
         
    ActiveSheet.Range("$A$5:$X$4025").AutoFilter Field:=1, Criteria1:="<>"
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="SATO SG408R-ex_IP195", Collate:=True, _
    IgnorePrintAreas:=False
    
    Next i
    
    Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
End Sub
Sub 予備カード印刷()
       '20210515
    Dim myPrinter As String
    myPrinter = Application.ActivePrinter
    
    If myPrinter Like "*195*" Then
    Application.ActivePrinter = myPrinter
    Else
        MsgBox myPrinter & "が選択されています。" & vbCrLf & "プリンターの設定をIP195へ変更して下さい。"
        Exit Sub
    End If
    
    
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect Password:="0001"
   
'    Dim myPrinter As String
'
'    myPrinter = Application.ActivePrinter
'
'    Application.ActivePrinter = "SATO SG408R-ex_IP195 on Ne01:"
'
    ActiveSheet.PrintPreview
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    If MsgBox("印刷処理を行いますか？", vbApplicationModel + vbInformation + vbOKCancel, "確認") = vbCancel Then
        
        Application.ActivePrinter = myPrinter
    
        ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
        False, Password:="0001"
       
            Exit Sub
               
    End If
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="SATO SG408R-ex_IP195", Collate:=True, _
    IgnorePrintAreas:=False
    
    Application.ActivePrinter = myPrinter
    
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:= _
    False, Password:="0001"
    
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



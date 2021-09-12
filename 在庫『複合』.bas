Sub Login()

    Dim ie As InternetExplorer
    
    Dim txtInput As HTMLInputElement
    Dim txtInput1 As HTMLInputElement
    Dim txtInput2 As HTMLInputElement
    Dim txtInput3 As HTMLInputElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    
    ie.Navigate "https://www.flsupply.jp/vdrwebsystem/Login/login.cfm"
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "USERID" Then
            txtInput.Value = "ms05830v0110"
            Exit For
        End If
    Next
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "pass" Then
     
     '      txtInput.Value = "aysny0057"
           txtInput.Value = Range("C" & Range("B1:B500").Find("pass").Row).Value
    
         
  
            Exit For
        End If
    Next
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput3 In ie.document.all.tags("input")
        If txtInput3.Value = "OK" Then
            txtInput3.Click
            Exit For
        End If
    Next
    
End Sub
Sub login2()

    Dim ie As InternetExplorer
    
    Dim txtInput As HTMLInputElement
    Dim txtInput1 As HTMLInputElement
    Dim txtInput2 As HTMLInputElement
    Dim txtInput3 As HTMLInputElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    
    ie.Navigate "https://www.kanto-syokuryo.jp/shop/customer/menu.aspx"
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "uid" Then
            txtInput.Value = "yokms350"
            Exit For
        End If
    Next
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "pwd" Then
           txtInput.Value = "219aubej"
    
         
  
            Exit For
        End If
    Next
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput3 In ie.document.all.tags("input")
        If txtInput3.Name = "order" Then
            txtInput3.Click
            Exit For
        End If
    Next
    
End Sub

Sub コメント()

    With ActiveCell
        If TypeName(.Comment) = "Nothing" Then
            .AddComment.Text ""
        End If
        
        Application.SendKeys "+{F2}"
    End With
    
End Sub


Sub hLogin()

    Dim ie As InternetExplorer
    
    Dim txtInput As HTMLInputElement
    Dim txtInput1 As HTMLInputElement
    Dim txtInput2 As HTMLInputElement
    Dim txtInput3 As HTMLInputElement
     
    Dim ps As String
    ps = Worksheets("三菱食品㈱_NB").Cells(2, 42).Value
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    
    ie.Navigate "https://www.flsupply.jp/vdrwebsystem/Login/login.cfm"
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "USERID" Then
            txtInput.Value = "ms05830v0210"
            Exit For
        End If
    Next
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "pass" Then
            txtInput.Value = ps
            Exit For
        End If
    Next
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop

    
    For Each txtInput3 In ie.document.all.tags("input")
        If txtInput3.Value = "OK" Then
            txtInput3.Click
            Exit For
        End If
    Next

    
End Sub

Sub webEDI_login()

    Dim ie As InternetExplorer
    
    Dim txtInput As HTMLInputElement
    Dim txtInput1 As HTMLInputElement
    Dim txtInput2 As HTMLInputElement
    Dim txtInput3 As HTMLInputElement
    
    Set ie = CreateObject("InternetExplorer.Application")
    
    ie.Visible = True
    
    ie.Navigate "https://webpf.finet.co.jp/webedi/user/"
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "uidcompanyId" Then
            txtInput.Value = "ABF7768000"
            Exit For
        End If
    Next

  For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "userId" Then
            txtInput.Value = "ABFZ21S"
            Exit For
        End If
    Next

    For Each txtInput In ie.document.getElementsByTagName("input")
        If txtInput.Name = "password" Then
           txtInput.Value = "af2000"
            Exit For
        End If
    Next
    
    Do While ie.Busy Or ie.readyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    For Each txtInput3 In ie.document.all.tags("input")
        If txtInput3.Name = "ログイン" Then
            txtInput3.Click
            Exit For
        End If
    Next
    
End Sub
Sub Login()

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
     
     '      txtInput.Value = "aysny0057"
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
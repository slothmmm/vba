Sub シート非表示()

   Dim WS As Worksheet, flag As Boolean
    For Each WS In Worksheets
        If WS.Name = "" Or _
           WS.Name = "" Or _
           WS.Name = "" _
        Then
        Else: WS.Visible = False
        End If
    Next
   
End Sub
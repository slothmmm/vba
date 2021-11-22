'「マスター更新日時」シート「B２」セル
'Function 販売商品マスターの更新日時()
'
'    販売商品マスターの更新日時 = FileDateTime("\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\販売商品マスター.xlsm")
'
'End Function

Sub 印刷判定ボタン()
        masterTime = FileDateTime("\\Afnewt320-kyoyu\社内共有\AFSKS\生産管理\販売商品マスター.xlsm")
        openTime = Worksheets("マスター更新日時").Range("B3").Value
        hantei = ""
        Debug.Print (masterTime)
        Debug.Print (openTime)
        If openTime <= masterTime Then
            hantei = "印刷時は生産表を保存し、開き直して下さい。"
        Else
            hantei = "印刷可能です。"
        End If
        MsgBox "販売商品マスターの更新日時：" & masterTime & vbCrLf & "生産表を開いた時間　　　 ：" & openTime & vbCrLf & vbCrLf & hantei
End Sub

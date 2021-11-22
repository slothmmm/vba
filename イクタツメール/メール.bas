Public Function メール送信手動(sijisuu As Integer, seizoubi As Date, shikabi As Date)
    Set oApp = CreateObject("Outlook.Application")
    Set myNameSpace = oApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(6)
    'myFolder.Display 'OUTLOOK起動
    Set objMail = oApp.CreateItem(olMailItem)

    'メール
    meado = "hashimoto@aysny.co.jp; oba@aysny.co.jp;"
    kenmei = "【更新通知】コープデリフローズン"
    naiyou = "お疲れ様です。" + vbCrLf + "更新完了をお知らせ致します。"

    'メールへ反映
    With objMail
        .To = meado
        '.CC = mailList.meado(1)
        .Subject = kenmei
        .Body = naiyou
        .BodyFormat = 2
        .Display            'OUTLOOK送信画面の起動
    End With

End Function
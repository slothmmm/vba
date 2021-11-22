Sub メールボタン()
    Set oApp = CreateObject("Outlook.Application")
    Set myNameSpace = oApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(6)
    'myFolder.Display 'OUTLOOK起動
    Set objMail = oApp.CreateItem(olMailItem)

    '入力シートからデータ抽出
    Dim henList  As henClass
    Set henList = 入力データ抽出()
    Worksheets("入力").Activate

    'メールクラス
    Dim mailList  As mailClass
    Set mailList = New mailClass

    'メール
    mailList.meado = メール送信リスト()
    mailList.kenmei = henList.bunrui + "変更" + " " + henList.syouCD
    mailList.naiyou = Worksheets("メール文").Range("B2").Value + vbCrLf + henList.naiyouCD + henList.naiyouBuhin + vbCrLf + henList.naiyou + vbCrLf + vbCrLf + Worksheets("メール文").Range("B3").Value + vbCrLf + henList.iraiName

    'メールへ反映
    With objMail
        .To = mailList.meado(0)
        .CC = mailList.meado(1)
        .Subject = mailList.kenmei
        .Body = mailList.naiyou
        .BodyFormat = 2
        .Display            'OUTLOOK送信画面の起動
    End With

    Worksheets("入力").Activate
End Sub

Public Function メール送信リスト() As Variant
    Worksheets("リスト").Activate
    Dim meadoSeizoTemp As Range, meadoSystemTemp As Range, meadoEigyoTemp As Range, meadoKanriTemp As Range

    '「リスト」シートの一番下
    Dim Last_Row_Seizo As Long, Last_Row_System As Long, Last_Row_kanri As Long
    Last_Row_Seizo = Worksheets("リスト").Cells(Rows.Count, 5).End(xlUp).Row
    Last_Row_System = Worksheets("リスト").Cells(Rows.Count, 7).End(xlUp).Row
    Last_Row_kanri = Worksheets("リスト").Cells(Rows.Count, 14).End(xlUp).Row

    'メアドそれぞれ取得
    Set meadoSeizoTemp = Worksheets("リスト").Range(Cells(4, 5), Cells(Last_Row_Seizo, 5))
    Set meadoSystemTemp = Worksheets("リスト").Range(Cells(4, 7), Cells(Last_Row_System, 7))
    Set meadoKanriTemp = Worksheets("リスト").Range(Cells(4, 14), Cells(Last_Row_kanri, 14))

    'メアドそれぞれ１文に成形
    Dim meadoSeizo As String, meadoSystem As String, meadoEigyo As String, meadoKanri As String
    Dim i As Long
            '製造メアド
            For i = 0 To UBound(meadoSeizoTemp(), 1)
                If meadoSeizoTemp(i) = "" Then

                Else
                    meadoSeizo = meadoSeizo + meadoSeizoTemp(i)
                    meadoSeizo = meadoSeizo + ";"
                End If
            Next i

            'システムメアド
            For i = 0 To UBound(meadoSystemTemp(), 1)
                If meadoSystemTemp(i) = "" Then

                Else
                    meadoSystem = meadoSystem + meadoSystemTemp(i)
                    meadoSystem = meadoSystem + ";"
                End If
            Next i

            '管理メアド
            For i = 0 To UBound(meadoKanriTemp(), 1)
                If meadoKanriTemp(i) = "" Then

                Else
                    meadoKanri = meadoKanri + meadoKanriTemp(i)
                    meadoKanri = meadoKanri + ";"
                End If
            Next i

    '送信先メアドまとめ
    Dim meadoMatome(1) As String
    '宛先
        meadoMatome(0) = meadoMatome(0) + meadoKanri
    'CC
        meadoMatome(1) = meadoMatome(1) + meadoSeizo
        meadoMatome(1) = meadoMatome(1) + meadoSystem

    メール送信リスト = meadoMatome

End Function
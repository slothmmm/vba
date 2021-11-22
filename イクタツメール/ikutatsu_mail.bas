Sub イクタツ判定()
        ActiveWorkbook.Sheets("一覧②").Activate            'アクティブ
        Application.Calculation = xlCalculationAutomatic  '自動計算
        Application.Calculate                             '再計算
        '最終行取得
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 11).End(xlUp).Row
        '数式を配列へ格納     F列からAJ列
        tikan_cell = Range(Cells(1, 11), Cells(LastRow, 27))

        Dim judgeIkutatsu, sijisuu As Integer
        judgeIkutatsu = 0
        sijisuu = 0
        '置換
        For s = 1 To (LastRow - 1)
                    If tikan_cell(s, 1) Like "*イクタツ*" Then
                        judgeIkutatsu = 1
                        sijisuu = sijisuu + tikan_cell(s, 16)
                    End If
        Next s

        Debug.Print judgeIkutatsu
        Debug.Print sijisuu

        'メール
        Call メール送信(sijisuu)
End Sub

Public Function メール送信(sijisuu As Integer)
    Set oApp = CreateObject("Outlook.Application")
    Set myNameSpace = oApp.GetNamespace("MAPI")
    Set myFolder = myNameSpace.GetDefaultFolder(6)
    'myFolder.Display 'OUTLOOK起動
    Set objMail = oApp.CreateItem(olMailItem)

    'メールクラス
    Dim mailList  As mailClass
    Set mailList = New mailClass

    'メール
    mailList.meado = メール送信リスト()
    mailList.kenmei = "けんめい"
    mailList.naiyou = Worksheets("メール文").Range("B2").Value + vbCrLf + vbCrLf + Worksheets("メール文").Range("B3").Value + vbCrLf + vbCrLf + "発行枚数 : " + Str(sijisuu / 10) + "枚 + 予備数" + vbCrLf + " (製造パック数 " + Str(sijisuu) + "pc)"

    'メールへ反映
    With objMail
        .To = mailList.meado(0)
        .CC = mailList.meado(1)
        .Subject = mailList.kenmei
        .Body = mailList.naiyou
        .BodyFormat = 2
        .Display            'OUTLOOK送信画面の起動
    End With

End Function



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

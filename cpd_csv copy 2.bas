'」」」」」」」 API declarations: 」」」」」」」
Private Declare Sub keybd_event Lib "user32" _
 (ByVal bVk As Byte, _
 ByVal bScan As Byte, _
 ByVal dwFlags As Long, _
 ByVal dwExtraInfo As Long)

Private Declare Function GetKeyboardState Lib "user32" _
 (pbKeyState As Byte) As Long

Const VK_NUMLOCK = &H90           '「NumLock」キー
Const KEYEVENTF_EXTENDEDKEY = &H1 'キーを押す
Const KEYEVENTF_KEYUP = &H2       'キーを放す
Sub numLockOn()
  Dim NumLockState As Boolean
  Dim keys(0 To 255) As Byte

  GetKeyboardState keys(0)
  NumLockState = keys(VK_NUMLOCK)

'「NumLock」キーがオフの場合はオンにする。
  If NumLockState <> True Then
    'キーを押す
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'キーを放す
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
  End If
End Sub

'Sub データ保存()
''」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
''」」」」」 ブックを閉じるときに自動実行のためのマクロ
'    Sheets("ピッキング表").Select
''宣言
'    Dim SaveDir1, SaveDir2, SaveDir3, fn, mypath As String, d As Date
'
'    d = Worksheets("ピッキング表").Range("D6") + 2
'    fn = Worksheets("データ中継").Range("C2") & "_ピッキング表.xlsm"
'    mypath = ActiveWorkbook.Path
'    If d >= Date - 14 Then
'    Else
'    Exit Sub
'    End If
''読み取り判定
'    If ActiveWorkbook.ReadOnly = True Then
'    Exit Sub
'    End If
'
''    If Range("AN2") = "×" Then
''    GoTo 強制保存
''    End If
'
''メッセージ
'
''          "※通常は必ずOKをクリックしてください。" & vbNewLine & _
''          "" & vbNewLine & _
''          "(保存しない場合はキャンセルしてください。)"
''    kesu = MsgBox(msg, 1, "自動保存")
'    kesu = 1
'    If kesu = 2 Then
'    Exit Sub
'    End If
'
'強制保存:
''上書き保存
'    Application.ScreenUpdating = False
'    Application.DisplayAlerts = False
'    ActiveWorkbook.Save
'
''フォルダ存在確認･作成 階層①
'    SaveDir1 = "\\AFNEWT320-KYOYU\社内共有\AFSKS\データ保管\" & Format(d, "yyyy年")
'    If Dir(SaveDir1, vbDirectory) = "" Then
'        MkDir SaveDir1
'    End If
''フォルダ存在確認･作成 階層②
'    SaveDir2 = SaveDir1 & "\" & Format(d, "m月")
'    If Dir(SaveDir2, vbDirectory) = "" Then
'        MkDir SaveDir2
'    End If
''フォルダ存在確認･作成 階層③
'    SaveDir3 = SaveDir2 & "\" & Format(d, "m.d")
'    If Dir(SaveDir3, vbDirectory) = "" Then
'       MkDir SaveDir3
'    End If
''ファイルが元かコピーか確認
'    If mypath Like "*AFSKS\ピッキング表*" Then
'    Else
'    Exit Sub
'    End If
'
''階層③フォルダに元Bookを名前を付けて保存
'    If Dir(SaveDir3 & "\" & fn) = "" Then
'    Else
'    Kill SaveDir3 & "\" & fn
'    End If
'    With ActiveWorkbook
'        .SaveAs filename:=SaveDir3 & "\" & fn, _
'                          FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    End With
'    On Error Resume Next
'    With ActiveWorkbook
'        .SaveAs filename:=mypath & "\" & fn, _
'                          FileFormat:=xlOpenXMLWorkbookMacroEnabled
'    End With
'    Application.DisplayAlerts = True
'    Application.ScreenUpdating = True
'    On Error GoTo 0
'
''」」」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'End Sub

Sub 注文数入力後振分表印刷()

    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    'ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Range("BI6").Value = Now
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
'    Sheets("振分").Select
'    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
'        IgnorePrintAreas:=False
'    Sheets("ピッキング表").Select
'
    
    Dim strCellArea As String
    strCellArea = "C2:BE41"
      
    '選択したセル範囲で印刷プレビュー実施
    With ActiveSheet
        .PageSetup.PrintArea = strCellArea '印刷範囲を設定
        
        .PageSetup.PrintArea = ""          '印刷範囲をクリア
        '印刷の向きを「横向き」
       .PageSetup.Orientation = xlLandscape
        '用紙サイズを「A4」
       .PageSetup.PaperSize = xlPaperA4
       '.PrintPreview
       .PrintOut From:=1, To:=1, Copies:=1
    End With
    
    
    Application.ScreenUpdating = True
    
    MsgBox "振分表の印刷が完了しました"
    
End Sub
Sub 印刷前処理()

    Application.ScreenUpdating = False

    ActiveSheet.Unprotect
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Range("BI7").Value = Now
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    Application.ScreenUpdating = True
    
    MsgBox "帳票印刷前処理が完了しました→②帳票印刷処理へ"
    
End Sub
Sub シート全表示()

   Dim WS As Worksheet
    For Each WS In Worksheets
        WS.Visible = True
    Next
    
End Sub
Sub シート非表示()

   Dim WS As Worksheet, flag As Boolean
    For Each WS In Worksheets
        If WS.Name = "ピッキング表" Or _
           WS.Name = "箱数中継" Or _
           WS.Name = "データ中継" Or _
           WS.Name = "センターマスター" _
        Then
        Else: WS.Visible = False
        End If
    Next
   
End Sub
Sub コープデリP印刷()
    '2021.07.19
    Call 抜けチェック

    '2021.01.31
    Call csv出力.csv_main
    
    '2022.01.28
    Call フォント調整
    Call フォント調整_クルコ
    
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    
'    2021.08.28
'    Application.Run "データ保存"
'
'」」」」」」」」」」」」」」」」」」」」」」」」」」」」
'データエラーチェック
    
    If Range("BG10") = "OK" Then
    Else
    MsgBox "※データ異常あり(印刷をキャンセルします)"
    Exit Sub
    End If
    
    If Worksheets("ピッキング表").Range("BG11") = "NG" Then
    MsgBox "▲▲▲データ更新に異常があります▲▲▲" & vbNewLine & _
           "ファイルを閉じて更新しなおしてください。"
    Exit Sub
    End If
    
    Dim p1, p2, p3, p4 As Long, n1, n2, n3, n4 As String
    p1 = Range("BG14")
    p2 = Range("BG15")
    p3 = Range("BG16")
    p4 = Range("BG17")

    n1 = Range("BF14")
    n2 = Range("BF15")
    n3 = Range("BF16")
    n4 = Range("BF17")
   
    msg = "印刷内容を確認してください。" & _
          vbNewLine & n1 & "   " & p1 & _
          vbNewLine & n2 & "   " & p2 & _
          vbNewLine & n3 & "   " & p3 & _
          vbNewLine & n4 & "   " & p4
    kesu = MsgBox(msg, 1, "帳票発行")
    If kesu = 2 Then
    Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Worksheets("ピッキング表").Unprotect
    ActiveWorkbook.UpdateLink Name:=ActiveWorkbook.LinkSources
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Application.DisplayAlerts = False
    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    Worksheets("ピッキング表").Unprotect
    
    If p1 > 0 Then
    Else
    GoTo myerror
    End If

'印刷処理1
    With Worksheets("振分")
     .Visible = True
     .PrintOut Copies:=p1, Collate:=True, IgnorePrintAreas:=False
'     .Visible = False
    End With
    
    With Worksheets("レシピ用")
     .Visible = True
     .PrintOut Copies:=p1, Collate:=True, IgnorePrintAreas:=False
'     .Visible = False
    End With
    
    With Worksheets("レシピ看板(クルコ)")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("レシピ看板")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    If Worksheets("ラベル用").Range("A2") > 0 Then
    With Worksheets("ラベル用")
     .Visible = True
     .PrintOut Copies:=p1, Collate:=True, IgnorePrintAreas:=False
'     .Visible = False
    End With
    End If
        
'印刷回数履歴
    Worksheets("ピッキング表").Range("BJ14").Formula = Range("BJ14") + 1
        
myerror:
    If p2 > 0 Then
    Else
    GoTo myerror2
    End If
    
'印刷処理2
    
    With Worksheets("チェックシート")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p2, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("チェックシート(クルコ)")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p2, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With

'印刷回数履歴
    Worksheets("ピッキング表").Range("BJ15").Formula = Range("BJ15") + 1
         
myerror2:
    If p3 > 0 Then
    Else
    GoTo myerror3
    End If
        
    With Worksheets("ローラー掛け")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("ロットメモクルコ")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("ロットメモ")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("作業順番表")
     .Visible = True
     .Range("I6").AutoFilter Field:=9, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("I6").AutoFilter Field:=9
'     .Visible = False
    End With
    
    With Worksheets("看板クルコ")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    '2022.06.21
'    With Worksheets("看板 (全)")
'     .Visible = True
'     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
'     .PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
'     .Range("A1").AutoFilter Field:=1
''     .Visible = False
'    End With
    
    With Worksheets("看板2デリ")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板2クルコ")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板3")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板4")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
       
    With Worksheets("看板5")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板4a")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("看板5a")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p3, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("ラベルチェック(クルコ)")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    If Worksheets("ラベルチェック(クルコ)②").Range("AK1") > 0 Then
    With Worksheets("ラベルチェック(クルコ)②")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    End If
    
    With Worksheets("ラベルチェック")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    If Worksheets("ラベルチェック②").Range("AK1") > 0 Then
    With Worksheets("ラベルチェック②")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    End If

    With Worksheets("デリ日別")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With

    With Worksheets("クルコ日別")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=1, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With

'印刷処理3
    With Worksheets("振分(出荷)")
     .Visible = True
     .PageSetup.PaperSize = xlPaperA4
     .PrintOut Copies:=8, Collate:=True, IgnorePrintAreas:=False
     .PageSetup.PaperSize = xlPaperA3
     .PrintOut Copies:=2, Collate:=True, IgnorePrintAreas:=False
'     .Visible = False
    End With

    If Worksheets("ラベル確認").Range("AB1") > 0 Then
        With Worksheets("ラベル確認")
        .Visible = True
        .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
        .PrintOut Copies:=p1, Collate:=True, IgnorePrintAreas:=False
        .Range("A1").AutoFilter Field:=1
    '     .Visible = False
        End With
    End If

'印刷回数履歴
    Worksheets("ピッキング表").Range("BJ16").Formula = Range("BJ16") + 1
    
myerror3:
    If p4 > 0 Then
    Else
    GoTo myerror4
    End If

'指示書的印刷処理
    With Worksheets("払い出し一覧")
     .Visible = True
      With .AutoFilter.Sort
       .SortFields.Clear
       .SortFields.Add Key:=Range("E4"), SortOn:=xlSortOnValues, _
         Order:=xlAscending, DataOption:=xlSortNormal
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
      End With
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
      With .AutoFilter.Sort
       .SortFields.Clear
       .SortFields.Add Key:=Range("B4"), SortOn:=xlSortOnValues, _
         Order:=xlAscending, DataOption:=xlSortNormal
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
      End With
'     .Visible = False
    End With
       
    With Worksheets("払い出し")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
        
    With Worksheets("加工表")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("包装表")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    With Worksheets("包装表(クルコ)")
     .Visible = True
     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
     .Range("A1").AutoFilter Field:=1
'     .Visible = False
    End With
    
    '2022.06.21
'    With Worksheets("包装表 (全)")
'     .Visible = True
'     .Range("A1").AutoFilter Field:=1, Criteria1:="<>"
'     .PrintOut Copies:=p4, Collate:=True, IgnorePrintAreas:=False
'     .Range("A1").AutoFilter Field:=1
''     .Visible = False
'    End With
    
    
    
'印刷回数履歴
    Worksheets("ピッキング表").Range("BJ17").Formula = Range("BJ17") + 1
    

myerror4:
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Sheets("ピッキング表").Select
    Range("F8").Select
    Call データ保管.データ保管
    Worksheets("ピッキング表").Activate
    Worksheets("ピッキング表").Select
    Range("A1").Select
    Application.ScreenUpdating = True
    
End Sub

Sub コープデリ全表示()

    msg = "Sheet内全表示" & Chr$(10) & _
          "※無断使用禁止"
    kesu = MsgBox(msg, 1, "メンテナンス用")
    If kesu = 2 Then
    Exit Sub
    End If
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    Cells.EntireRow.Hidden = False
    Cells.EntireColumn.Hidden = False
    Range("F8").Select
    Application.ScreenUpdating = True
End Sub
Sub コープデリ非表示()

    msg = "Sheet内使用箇所以外非表示" & Chr$(10) & _
          "※無断使用禁止"
    kesu = MsgBox(msg, 1, "メンテナンス用")
    If kesu = 2 Then
    Exit Sub
    End If
    Application.ScreenUpdating = False
    Rows("4:5").EntireRow.Hidden = True
    Rows("30:38").EntireRow.Hidden = True
    Range("B:B,E:E,J:J,N:BC,BK:BL").EntireColumn.Hidden = True
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Range("F8").Select
    
    Application.ScreenUpdating = True
End Sub
Sub コープデリリセット()

    msg = "数値クリア" & Chr$(10) & _
          "出荷日更新"
    kesu = MsgBox(msg, 1, "受注数入力")
    If kesu = 2 Then
    Exit Sub
    End If
'
''日別データ保存済みか確認
'    If Range("AM2").Formula = "☆" Then
'    Range("AN2").Formula = "×"
''    Call データ保存
'    Range("AN2").Formula = ""
'    End If
    
    Application.ScreenUpdating = False

    Worksheets("ピッキング表").Unprotect
    Range("F8:I38,K8:L38,BJ14:BJ17,BI6:BI7").ClearContents
    Range("D6").Formula = Date + 1
    Range("AM2").Formula = "☆"
    
'商品コードC列に数式をコピー
    Range("C7").Copy
    Range("C8:C38").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
'@@@@@@@@@@
    
'販売計画コピー
    Dim d As Date, 対象 As String, WS As Worksheet, flag As Boolean
    d = Range("D6")
    対象 = Format(d, "m月")
    Workbooks.Open filename:="\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】\【販売計画集計表】.xlsm", _
                             UpdateLinks:=0, ReadOnly:=True
    
    For Each WS In Worksheets
    If WS.Name = 対象 Then flag = True
    Next WS
    If flag = True Then
    Else
    MsgBox "対象シートが見つかりません。以下を確認してください。" & vbNewLine & _
           "" & vbNewLine & _
           "Pass ：\\Afnewt320-kyoyu\社内共有\【生産管理】\【システム】" & vbNewLine & _
           "Book ：【販売計画集計表】.xlsm" & vbNewLine & _
           "Sheet：" & 対象 & vbNewLine & _
           "", vbInformation
    GoTo skip
    End If
    Worksheets(対象).Range("A:AQ").Copy
    Workbooks("コープデリ_ピッキング表.xlsm").Activate
    Worksheets("計画").Visible = True
    Worksheets("計画").Range("A:AQ").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
'    Worksheets("計画").Visible = False
skip:
    Application.DisplayAlerts = False
    Workbooks("【販売計画集計表】.xlsm").Close SaveChanges:=False
    
    Workbooks("コープデリ_ピッキング表.xlsm").Activate
    Sheets("ピッキング表").Select
    Call コード固定
    Range("F8").Select
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    
    kakunin_ship = Worksheets("ピッキング表").Range("D6")  '出荷日取得
    Dim rc As VbMsgBoxResult
    rc = MsgBox(Str(kakunin_ship) + "で事前入力したデータを出力しますか？", vbYesNo + vbQuestion)
    If rc = vbNo Then
        MsgBox "出力を中止します", vbCritical
        ActiveWorkbook.Save
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Exit Sub
    End If
    Customer_name = "コープデリ"                       '出荷先
    Call 事前入力_読み込みリセット用main(Customer_name)

    ActiveWorkbook.Save
    Application.DisplayAlerts = True
    
    Application.ScreenUpdating = True
End Sub
Sub コード固定()
  
    Worksheets("ピッキング表").Unprotect
    Range("C8:C38").Copy
    Range("C8:C38").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Worksheets("ピッキング表").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Worksheets("ピッキング表").Range("A1").Select

End Sub


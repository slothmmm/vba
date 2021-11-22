Attribute VB_Name = "PDF"
'''''''''''''''''''''''''''''''          PDF                     '''''''''''''''''''''''''''''''''''
Sub PDF副原材料()
    Call 保護.全保護解除
    
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=26, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="副原材料"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "副原材料" & hiduke & ".pdf"

    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9016").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call PDF_賞味期限
    Call 保護.複数保護
    
    MsgBox "副原材料の印刷完了"
End Sub
Sub PDF日曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=39, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(日)" & hiduke & ".pdf"

    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "日曜日のPDF完了"
End Sub

Sub PDF月曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=40, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(月)" & hiduke & ".pdf"
    
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "月曜日のPDF完了"
End Sub

Sub PDF火曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=41, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(火)" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "火曜日のPDF完了"
End Sub

Sub PDF水曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=42, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(水)" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "水曜日のPDF完了"
End Sub

Sub PDF木曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=43, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(木)" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "木曜日のPDF完了"
End Sub

Sub PDF金曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=44, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(金)" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "金曜日のPDF完了"
End Sub

Sub PDF土曜日()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=45, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="諸口"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "諸口(土)" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9019").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "土曜日のPDF完了"
End Sub

Sub PDF_IY()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=31, Criteria1:="<>"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="IY"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "IY_" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9023").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "IYのPDF完了"
End Sub

Sub PDF_CN()
    Call 保護.全保護解除
    ''''''''''''''''''''''''''''''''''''''''''''''   形成１  '''''''''''''''''''''''''''''''''''''''''''''
    Worksheets("形成1").Activate
    
    Call ソート形成1.ソート形成1_ALL
    
    With ActiveSheet
        .ListObjects("新館").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("商品管理").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷蔵庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("冷凍庫").Range.AutoFilter Field:=25, Criteria1:="<>"
        .ListObjects("その他").Range.AutoFilter Field:=25, Criteria1:="<>"
    End With
        
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "コープNo1_" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("形成1").Range("D4:V183").Address
        .Orientation = xlLandscape '印刷向きを横方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''   形成２  '''''''''''''''''''''''''''''''''''''''''''''
    Worksheets("形成2").Activate
    
    Call ソート形成2.ソート形成2_レシピ

    With ActiveSheet
        .Range("C5").Select
        .Range("C5").AutoFilter Field:=25, Criteria1:="<>"
    End With
    
    'PDF設定
    Dim fileName2 As String '保存先フォルダパス＆ファイル名
    
    fileName2 = ThisWorkbook.Path & "\PDF\" & "コープNo2(レシピ)_" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("形成2").Range("D4:V58").Address
        .Orientation = xlLandscape '印刷向きを横方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName2
    End With
    Call 保護.複数保護
    Worksheets("印刷他").Activate
    MsgBox "CNのPDF完了"
End Sub

Sub PDF_賞味期限()
    Call 保護.全保護解除
    Worksheets("賞味期限").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .Range("A10").AutoFilter Field:=1, Criteria1:="<>"
    End With
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "賞味期限" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("賞味期限").Range("A8:J55").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "賞味期限のPDF完了"
End Sub

Sub PDFシール()
    Call 保護.全保護解除
    Worksheets("印刷他").Activate
    
    With ActiveSheet
        .Range("A10").Select
        If .FilterMode Then .ShowAllData
        .ListObjects("テーブル2").Range.AutoFilter Field:=38, Criteria1:="シール"
        .ListObjects("テーブル2").Range.AutoFilter Field:=6, Criteria1:="<>-"
        .ListObjects("テーブル5").Range.AutoFilter Field:=1, Criteria1:="シール"
    End With
    
    Call ソート印刷他.ソート仕入先名
    
    'PDF設定
    Dim fileName As String '保存先フォルダパス＆ファイル名
    hiduke = Format(Month(Now), "00") & Format(Day(Now), "00") & "_" & Format(Hour(Now), "00") & Format(Minute(Now), "00") & Format(Second(Now), "00")
    
    fileName = ThisWorkbook.Path & "\PDF\" & "シール" & hiduke & ".pdf"
    
    With ActiveSheet.PageSetup
        .PrintArea = Sheets("印刷他").Range("A8:J9026").Address
        .Orientation = xlPortrait '印刷向きを縦方向に設定
        .Zoom = False '拡大縮小を設定（しない）
        .FitToPagesWide = 1 'すべての列を1ページに印刷
        .FitToPagesTall = False 'シートを1ページに印刷
        
        .LeftMargin = Application.CentimetersToPoints(0.5)
        .RightMargin = Application.CentimetersToPoints(0.5)
        .TopMargin = Application.CentimetersToPoints(0.5)
        .BottomMargin = Application.CentimetersToPoints(0.5)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
    End With
    
    With ActiveSheet
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName
    End With
    Call 保護.複数保護
    MsgBox "シールのPDF完了"
End Sub

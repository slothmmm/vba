Sub 印刷_小分払出()
        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        Application.CalculateFull
        With ActiveSheet
            .Range("A5").Select
            'If .FilterMode Then .ShowAllData
            .Range("A5").AutoFilter Field:=12, Criteria1:="<>"
            'Application.CalculateFull
        End With
        
        With ActiveSheet
             '.PageSetup.Zoom = 55 '倍率55%
             '印刷の向きを「横向き」
             .PageSetup.Orientation = xlLandscape
              '用紙サイズを「A4」
             .PageSetup.PaperSize = xlPaperA3
              '印刷を1部実行
             .PrintOut Copies:=1
        End With
        
        
        'If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub

Sub 印刷_小分保管表()
        Call 更新_形成用小分け(2)

        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        Application.CalculateFull
        With ActiveSheet
            .Range("A2").Select
            .Range("A2").AutoFilter Field:=1, Criteria1:="<>"
        End With
        
        With ActiveSheet
             .PageSetup.Orientation = xlPortrait
             .PageSetup.PaperSize = xlPaperA4
             .PrintOut Copies:=1
        End With
        
        'If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub

Sub 印刷_払出記入表()
        If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
        Application.CalculateFull
        With ActiveSheet
            .Range("A2").Select
            .Range("A2").AutoFilter Field:=1, Criteria1:="<>"
        End With
        
        With ActiveSheet
             .PageSetup.Orientation = xlPortrait
             .PageSetup.PaperSize = xlPaperA4
             .PrintOut Copies:=1
        End With
        
        'If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
End Sub


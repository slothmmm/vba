Attribute VB_Name = "ソート形成1"
Sub ソート形成1_新館()
    With ActiveSheet
        Range("新館[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("形成1").ListObjects("新館").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("新館").Sort.SortFields.Add Key _
        :=Range("新館[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成1").ListObjects("新館").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成1_商品管理()
    With ActiveSheet
        Range("商品管理[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("形成1").ListObjects("商品管理").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("商品管理").Sort.SortFields.Add Key _
        :=Range("商品管理[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成1").ListObjects("商品管理").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成1_冷蔵庫()
    With ActiveSheet
        Range("冷蔵庫[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷蔵庫").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷蔵庫").Sort.SortFields.Add Key _
        :=Range("冷蔵庫[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成1").ListObjects("冷蔵庫").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成1_冷凍庫()
    With ActiveSheet
        Range("冷凍庫[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷凍庫").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷凍庫").Sort.SortFields.Add Key _
        :=Range("冷凍庫[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成1").ListObjects("冷凍庫").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成1_その他()
    With ActiveSheet
        Range("その他[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With
    
    ActiveWorkbook.Worksheets("形成1").ListObjects("その他").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("その他").Sort.SortFields.Add Key _
        :=Range("その他[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成1").ListObjects("その他").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成1_ALL()
    Call ソート形成1_新館
    Call ソート形成1_商品管理
    Call ソート形成1_冷蔵庫
    Call ソート形成1_冷凍庫
    Call ソート形成1_その他
    
    ActiveSheet.Range("M1").Select
End Sub

Sub ソート形成1クリア_ALL()
    ActiveWorkbook.Worksheets("形成1").ListObjects("新館").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("商品管理").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷蔵庫").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("冷凍庫").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成1").ListObjects("その他").Sort.SortFields.Clear
    
    ActiveSheet.Range("P1").Select
End Sub



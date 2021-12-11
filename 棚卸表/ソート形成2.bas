Attribute VB_Name = "ソート形成2"
Sub ソート形成2_レシピ()
    With ActiveSheet
        Range("レシピ[[#Headers],[仕入先名]]").Select
        If .FilterMode Then .ShowAllData
    End With

    ActiveWorkbook.Worksheets("形成2").ListObjects("レシピ").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("形成2").ListObjects("レシピ").Sort.SortFields.Add Key _
        :=Range("レシピ[[#All],[仕入先名]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("形成2").ListObjects("レシピ").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ソート形成2クリア()
    ActiveWorkbook.Worksheets("形成2").ListObjects("レシピ").Sort.SortFields.Clear
    ActiveSheet.Range("B2").Select
End Sub



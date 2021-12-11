Sub csvシート出力main()
    'ファイル名
    csvFileName = csvファイル名

    'ワークシート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("出力用")

    Call csvシート出力(csvFileName, ws)

End Sub

Function csvファイル名() As Variant
    'csvファイル名_現在時刻
    hiduke = Format(Year(Now), "0000") & "年" & Format(Month(Now), "00") & "月" & Format(Day(Now), "00") & "日" & Format(Hour(Now), "00") & "時" & Format(Minute(Now), "00") & "分" & Format(Second(Now), "00") & "秒"

    'csvファイル名
    Dim csvFileName As String
    csvFileName = ActiveWorkbook.Path & "\csv\" & ThisWorkbook.Name & hiduke & ".csv"

    csvファイル名 = csvFileName
End Function

Function csvシート出力(csvFileName As Variant, ws As Variant)
    'CSV Open >> Close
    Open csvFileName For Output As #1

    Dim i As Long, j As Long
    i = 1

    Do While ws.Cells(i, 1).Value <> ""

        j = 1
        Do While ws.Cells(i, j + 1).Value <> ""

            Print #1, ws.Cells(i, j).Value & ",";
            j = j + 1

        Loop

        Print #1, ws.Cells(i, j).Value & vbCr;
        i = i + 1

    Loop

    Close #1
End Function


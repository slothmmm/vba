Sub janHelp()
    
    Application.MacroOptions macro:="Jan", _
        Description:="Janコードをエクセル上でバーコード表示します。()の中に7桁、8桁、12桁、13桁のJANを入力してください。もしくはJANコードが入力されているセルを指定してください。" & vbCrLf & "表示された文字のフォントをJANCODE-nicotanに変更するとバーコードが表示されます。", _
        Category:="JANCODE"
    Application.MacroOptions macro:="JanCD", _
        Description:="7桁、12桁のJANコードのチェックデジットを計算して0～9の数字を返します。", _
        Category:="JANCODE"
    Application.MacroOptions macro:="reJan", _
        Description:="JANコードのフォント用文字列をJANコードの数値に戻します。", _
        Category:="JANCODE"
    Application.MacroOptions macro:="JanW", _
        Description:="JANCODE-nicWabun フォントの数字なしのバーコードを表示するため数式です。" & vbCrLf & "=JAN()の答えを全角化した文字列を返します。", _
        Category:="JANCODE"
    Application.MacroOptions macro:="ITF", _
        Description:="ITFコードの数式です。" & vbCrLf & "第１引数にはインジケータ(外箱)" & vbCrLf & "第２引数にはJANコードを指定してください。" & vbCrLf & "チェックデジットを付与したITFコード(14桁)を返します。", _
        Category:="JANCODE"
    Application.MacroOptions macro:="ITFCD", _
        Description:="ITFコードのチェックデジットを返す数式です。" & vbCrLf & "第１引数にはインジケータ(外箱)" & vbCrLf & "第２引数にはJANコード(12,13)を指定してください。" & vbCrLf & "チェックデジットを返します。", _
        Category:="JANCODE"
                
    Application.MacroOptions macro:="CD", Description:="", Category:=0
    Application.MacroOptions macro:="Eight", Description:="", Category:=0
    Application.MacroOptions macro:="Thirteen", Description:="", Category:=0
    Application.MacroOptions macro:="undo", Description:="", Category:=0

End Sub
Public Function Jan(JANCODE)

If JANCODE = "" Then
    Jan = CVErr(xlErrNA)
    Exit Function
End If

If Not IsNumeric(JANCODE) Then
    Jan = CVErr(xlErrValue)
    Exit Function
End If

On Error GoTo Err
    Select Case Len(JANCODE)
        Case 7:
            Jan = Eight(JANCODE)
        Case 8:
            Jan = Eight(JANCODE)
        Case 12:
            Jan = Thirteen(JANCODE)
        Case 13:
            Jan = Thirteen(JANCODE)
        Case Else
            GoTo Err
    End Select
Exit Function

Err:
    Jan = CVErr(xlErrValue)
    Exit Function

End Function
Public Function JanW(JANCODE) As String

Dim ans As Variant
ans = Jan(JANCODE)
If IsError(ans) Then
    JanW = ans
Else
    JanW = StrConv(ans, vbWide)
End If

End Function
Public Function JanCD(JANCODE)

If JANCODE = "" Then
    JanCD = CVErr(xlErrNA)
    Exit Function
End If

If Not IsNumeric(JANCODE) Then
    JanCD = CVErr(xlErrValue)
    Exit Function
End If
On Error GoTo Err
    Select Case Len(JANCODE)
        Case 7:
            JanCD = CD("00000" + CStr(JANCODE))
        Case 8:
            JanCD = CD("00000" + CStr(JANCODE))
        Case 12:
            JanCD = CD(JANCODE)
        Case 13:
            JanCD = CD(JANCODE)
        Case Else
            GoTo Err
    End Select
Exit Function

Err:
    JanCD = CVErr(xlErrValue)
    Exit Function

End Function
Public Function reJan(strJan) As String

Dim JANCODE As String, ans As Variant
On Error GoTo Err
strJan = StrConv(strJan, vbNarrow)
Select Case Len(strJan)
    Case 11:
        ans = undo(strJan, 11)
        If ans <> False Then
            JANCODE = ans
        End If
    Case 15:
        ans = undo(strJan, 15)
        If ans <> False Then
            JANCODE = ans
        End If
    Case Else
        GoTo Err
End Select

reJan = JANCODE

Exit Function
Err:
    reJan = CVErr(xlErrValue)
    Exit Function

End Function
Private Function CD(strJancode) As Byte

Dim g As Byte, k As Byte, h As Byte
g = 0
k = 0
h = 0
For i = 12 To 1 Step -2
    g = g + Val(Mid(strJancode, i, 1))
    k = k + Val(Mid(strJancode, i - 1, 1))
Next
h = (g * 3 + k) Mod 10
If h = 0 Then
    CD = 0
Else
    CD = 10 - h
End If
End Function
Private Function Eight(n)

Dim strJanfont As String, CheckDigit As Byte, BAR As Variant
BAR = getBar
strJanfont = "Y"
For i = 1 To 4
    strJanfont = strJanfont + Mid(n, i, 1)
Next
strJanfont = strJanfont + "K"
For i = 5 To 7
    strJanfont = strJanfont + BAR(2)(CByte(Mid(n, i, 1)))
Next
CheckDigit = CD("00000" + CStr(n))
strJanfont = strJanfont + BAR(2)(CheckDigit)
strJanfont = strJanfont + "Z"
Eight = strJanfont

End Function
Private Function Thirteen(n)

Dim strJanfont As String
Dim Initial(9) As Variant
Dim nIni As Byte
Dim CheckDigit As Byte
Dim BAR As Variant

Initial(0) = "000000"
Initial(1) = "001011"
Initial(2) = "001101"
Initial(3) = "001110"
Initial(4) = "010011"
Initial(5) = "011001"
Initial(6) = "011100"
Initial(7) = "010101"
Initial(8) = "010110"
Initial(9) = "011010"
strJanfont = ""
nIni = 0
BAR = getBar
    
nIni = Left(n, 1)
strJanfont = getStartCode(nIni)

For i = 1 To 6
    strJanfont = strJanfont + BAR(CByte(Mid(Initial(nIni), i, 1)))(CByte(Mid(n, i + 1, 1)))
Next
strJanfont = strJanfont + "K"
For i = 8 To 12
    strJanfont = strJanfont + BAR(2)(CByte(Mid(n, i, 1)))
Next
CheckDigit = CD(n)
strJanfont = strJanfont + BAR(2)(CheckDigit)
strJanfont = strJanfont + "Z"
Thirteen = strJanfont

End Function
Private Function undo(strJan, n)

Dim temp As String
temp = ""
For i = 1 To n
    Select Case Mid(strJan, i, 1)
        Case "a", "0", "A", "L"
            temp = temp + "0"
        Case "b", "1", "B", "M"
            temp = temp + "1"
        Case "c", "W", "2", "C", "N"
            temp = temp + "2"
        Case "d", "3", "D", "O"
            temp = temp + "3"
        Case "e", "X", "4", "E", "P"
            temp = temp + "4"
        Case "f", "5", "F", "Q"
            temp = temp + "5"
        Case "g", "6", "G", "R"
            temp = temp + "6"
        Case "h", "7", "H", "S"
            temp = temp + "7"
        Case "i", "8", "I", "T"
            temp = temp + "8"
        Case "j", "9", "J", "U"
            temp = temp + "9"
        Case "K", "Y", "Z"
        
        Case Else
            undo = False
            Exit Function
    End Select
Next
undo = temp

End Function
Private Function getStartCode(n) As String

Dim Startbar(9) As String
Startbar(0) = "a"
Startbar(1) = "b"
Startbar(2) = "W"
Startbar(3) = "d"
Startbar(4) = "X"
Startbar(5) = "f"
Startbar(6) = "g"
Startbar(7) = "h"
Startbar(8) = "i"
Startbar(9) = "j"
getStartCode = Startbar(n)

End Function
Private Function getBar() As Variant

Dim BAR(2) As Variant
Dim k(9) As String
Dim g(9) As String
Dim r(9) As String
k(0) = "0"
k(1) = "1"
k(2) = "2"
k(3) = "3"
k(4) = "4"
k(5) = "5"
k(6) = "6"
k(7) = "7"
k(8) = "8"
k(9) = "9"
g(0) = "A"
g(1) = "B"
g(2) = "C"
g(3) = "D"
g(4) = "E"
g(5) = "F"
g(6) = "G"
g(7) = "H"
g(8) = "I"
g(9) = "J"
r(0) = "L"
r(1) = "M"
r(2) = "N"
r(3) = "O"
r(4) = "P"
r(5) = "Q"
r(6) = "R"
r(7) = "S"
r(8) = "T"
r(9) = "U"
BAR(0) = k
BAR(1) = g
BAR(2) = r
getBar = BAR

End Function
Public Function ITF(indicator, JANCODE)

Dim ind As String

If indicator = "" Or JANCODE = "" Then
    ITF = CVErr(xlErrNA)
    Exit Function
End If

If Not IsNumeric(indicator) And IsNumeric(JANCODE) Then
    ITF = CVErr(xlErrValue)
    Exit Function
End If
ind = CStr(indicator Mod 10)

On Error GoTo Err
    Select Case Len(JANCODE)
        Case 7:
            ITF = ind & "00000" & JANCODE & CStr(ITFCD(ind, CStr("00000" & JANCODE)))
        Case 8:
            ITF = ind & "00000" & Mid(CStr(JANCODE), 1, 7) & CStr(ITFCD(ind, CStr("00000" & Mid(CStr(JANCODE), 1, 7))))
        Case 12:
            ITF = ind & CStr(JANCODE) & CStr(ITFCD(ind, CStr(JANCODE)))
        Case 13:
            ITF = ind & Mid(CStr(JANCODE), 1, 12) & CStr(ITFCD(ind, CStr(JANCODE)))
        Case Else
            GoTo Err
    End Select
Exit Function

Err:
    ITF = CVErr(xlErrValue)
    Exit Function

End Function
Public Function ITFCD(indicator As String, JANCODE As String) As Byte

Dim g As Byte, k As Byte, h As Byte
Dim strCODE As String
g = 0
k = 0
h = 0
strCODE = Mid(indicator & JANCODE, 1, 13)
For i = 1 To 13 Step 2
    g = g + Val(Mid(strCODE, i, 1))
    k = k + Val(Mid(strCODE, i + 1, 1))
Next
h = (g * 3 + k) Mod 10
If h = 0 Then
    ITFCD = 0
Else
    ITFCD = 10 - h
End If

End Function



Sub 使用者の参照設定のGUIDを調べる()
    'https://kouten0430.hatenablog.com/entry/2017/10/22/134852
    Dim myRef As Variant
    Dim i As Integer
        i = 1
    For Each myRef In ActiveWorkbook.VBProject.References
        i = i + 1
        Debug.Print (myRef.Name)
        Debug.Print (myRef.GUID)
        Debug.Print (myRef.Major)
        Debug.Print (myRef.Minor)
    Next
End Sub
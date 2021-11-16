Option Explicit

Sub Microsoft_Scripting_Runtime()

    On Error GoTo Err
    
    'Microsoft Scripting RuntimeのGUID
    Const MSR_GUID = "{420B2830-E718-11CF-893D-00A0C9054228}"
    '参照設定を追加
    Application.VBE.ActiveVBProject.References.AddFromGuid MSR_GUID, 1, 0
    
'    MsgBox "参照設定を追加しました！"
    
    Exit Sub
    
Err:
'    MsgBox "エラーが発生しました！" & vbCrLf & Err.Description
 
End Sub
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
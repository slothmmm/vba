Sub 止められるものなら止めてみろ()
　　Application.Interactive = False
　　Application.EnableCancelKey = xlDisabled
　　
　　Dim i As Long
　　For i = 1 To 10
　　　　'以下は適当に入れたもので特段の意味はありません。
　　　　Application.Wait Now() + TimeSerial(0, 0, 1)
　　　　DoEvents
　　　　MsgBox WorksheetFunction.Rept("ムダ！", i)
　　Next
　　
　　Application.Interactive = True
　　Application.EnableCancelKey = xlInterrupt
　　MsgBox "ムダな悪あがきでしたね！"
End Sub


'Application.EnableCancelKey = XlEnableCancelKey
'XlEnableCancelKey
'xldisabled	割り込みを無視します。
'xlErrorHandler	このエラーは?On Error GoTo?ステートメントでトラップできます。
'エラー コードは 18 です。
'xlInterrupt	デバッグ、終了などを行えるように、実行中のプロシージャを停止します。


'Application.Interactive
'Microsoft Excel が対話モードの場合はTrue。
'通常、このプロパティはTrueです。
'このプロパティをFalseに設定すると、キーボードおよびマウスからのすべての入力がブロックされます (コードによって表示されるダイアログボックスへの入力を除く)。

'Application.Interactive = True/False

'ユーザー入力をブロックすると、ユーザーがExcel オブジェクトを移動したり、アクティブにしたりしても、マクロを妨害することはできません。
'マクロの終了前に必ず設定をTrueに戻してください。
'マクロ実行が終了しても、このプロパティは自動的にはTrueに戻りません。

'「Esc」や「Ctrl」+「Break」を止めるだけであれば、
'Application.EnableCancelKey = xlDisabled
'これだけでも良いと思われますが、更に念には念を入れて、
'Application.Interactive = False
'こちらも入れています。
'Interactiveについては、マクロ終了後も戻らないので、必ずTrueに戻してください。

'上記VBAは、10回の「OK」で終了します。
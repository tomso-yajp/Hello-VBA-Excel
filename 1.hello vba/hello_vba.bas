Attribute VB_Name = "hello_vba"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' 1.イミディエイトウィンドウを表示します
'   ショートカットキーは、Ctrl + G です
' 2.call helloVBA をコンソールに入力して Enter キーを押します
'  helloVBA:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub helloVBA()
  'Debug.Print "hello vba ♪"
  Debug.Print "call helloVBA"
  MsgBox "hello vba ♪", vbInformation, "hello massege"
End Sub

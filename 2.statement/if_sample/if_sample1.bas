Attribute VB_Name = "if_sample1"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If 条件 Then
'   '条件と一致する場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_true()
Dim n1 As Integer: n1 = 0
Dim n2 As Integer: n2 = 10
'sample1 code
Debug.Print "--if_true: sample code----------------------"
If n2 > n1 Then
  Debug.Print "イミディエイトウィンドウに表示されます"
  Debug.Print "n1: " & n1 & " より n2: " & n2 & " が大きいです"
  Debug.Print "n2 > n1 の条件式に一致します"
End If
Debug.Print vbLf
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If 条件 Then
'    条件と一致する場合の処理
' else
'　　条件と一致しない場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_else()
Dim n As Integer
Dim str As String:

n = 0
'sample1 code is true(if)
Debug.Print "--if_else: sample1 n=10---------------------"
If n > 5 Then
  Debug.Print "true: n > 0 の条件式に一致します"
Else
  Debug.Print "false: n > 0 の条件式に一致しません"
End If
Debug.Print "n = " & n

n = 5
'sample2 code is false(else)
Debug.Print "--if_else: sample2 n=10---------------------"
If n > 11 Then
  Debug.Print "true: n > 11 の条件式に一致します"
Else
  Debug.Print "false: n > 11 の条件式に一致しません"
End If
Debug.Print "n = " & n

'sample3 code is true(if)
Debug.Print "--if_else: sample3 str=""apple""-------------"
str = "apple"
If str = "apple" Then
  Debug.Print "true: str = ""apple"" の条件に一致します"
Else
  Debug.Print "false: str = ""apple"" の条件に一致しません"
End If

'sample4 code is false(else)
Debug.Print "--if_else: sample4 str=""apple""-------"
If str = "orange" Then
  Debug.Print "true: str = ""orange"" の条件に一致します"
Else
  Debug.Print "false: str = ""orange"" の条件に一致しません"
End If
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If 条件1 Then
'    条件1と一致する場合の処理
' elseif 条件2
'    条件2と一致する場合の処理
' else
'　　条件と一致しない場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_elseif()
Dim n As Integer
Dim str As String
'sample1 code is the if
Debug.Print "--if_elseif: sample1 n=0-------------------"
n = 0
If n = 0 Then
  Debug.Print "true: n 変数の値は、0 です"
ElseIf n = 1 Then
  Debug.Print "true: n 変数の値は、1 です"
Else
  Debug.Print "false: n 変数の値は、条件と一致しませんでした"
End If

'sample2 code is the elseif
Debug.Print "--if_elseif: sample2 n=1-------------------"
n = 1
If n = 0 Then
  Debug.Print "true: n 変数の値は、0 です"
ElseIf n = 1 Then
  Debug.Print "true: n 変数の値は、1 です"
Else
  Debug.Print "false: n 変数の値は、条件と一致しませんでした"
End If

'sample3 code is the else
Debug.Print "--if_elseif: sample3 n=2------------------"
n = 2
If n = 0 Then
  Debug.Print "true: n 変数の値は、0 です"
ElseIf n = 1 Then
  Debug.Print "true: n 変数の値は、1 です"
Else
  Debug.Print "false: n 変数の値は、条件と一致しませんでした"
End If
End Sub









Attribute VB_Name = "if_sample2"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If 条件1 and 条件1 Then
'   条件1と条件2に一致する場合の処理
' Else
'   条件に一致しない場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Sub if_and()
Dim n As Integer:
'sample1 code is if,And
Debug.Print "--if_and: sample1 if,And n=5----------------"
n = 5
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 の条件式に一致します"
End If

'sample2 code is if,if
Debug.Print "--if_and: sample2 if,true n=5---------------"
If 10 > n Then
  If n > 0 Then
    Debug.Print "true: 10 > n, n > 0 の条件式に一致します"
  End If
End If

'sample3 code is if,And
Debug.Print "--if_and: sample3 if,And n=11--------------"
n = 11
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 の条件式に一致します"
Else
  Debug.Print "false: 条件に一致しません"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If 条件1 or 条件1 Then
'   条件1と条件2のいずれかに一致する場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_or()
Dim str As String
'sample1 code is if,Or
Debug.Print "--if_or: sample1 if,or str=""apple""------"
str = "apple"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: 条件に一致しました"
End If

'sample2 code is if,Or
Debug.Print "--if_or: sample2 if,or str=""banana""------"
str = "banana"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: 条件に一致しました"
Else
  Debug.Print "false: 条件に一致しませんでした"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If Not 条件 Then
'   条件を否定する場合の処理
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_not()
Dim n As Integer
'sample1 code is if,not
Debug.Print "--if_not: sample1 if,not n=1---------------"
n = 1
If Not n = 0 Then
  Debug.Print "true: Not n=0: 条件式と一致します"
End If

'sample2 code is if,not
Debug.Print "--if_not: sample2 if,not n=1---------------"
If Not n = 0 And Not n > 10 Then
  Debug.Print "true: n=0,n>10 条件と一致します"
End If
End Sub




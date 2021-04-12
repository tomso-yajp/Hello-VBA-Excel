Attribute VB_Name = "if_sample2"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ğŒ1 and ğŒ1 Then
'   ğŒ1‚ÆğŒ2‚Éˆê’v‚·‚éê‡‚Ìˆ—
' Else
'   ğŒ‚Éˆê’v‚µ‚È‚¢ê‡‚Ìˆ—
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Sub if_and()
Dim n As Integer:
'sample1 code is if,And
Debug.Print "--if_and: sample1 if,And n=5----------------"
n = 5
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 ‚ÌğŒ®‚Éˆê’v‚µ‚Ü‚·"
End If

'sample2 code is if,if
Debug.Print "--if_and: sample2 if,true n=5---------------"
If 10 > n Then
  If n > 0 Then
    Debug.Print "true: 10 > n, n > 0 ‚ÌğŒ®‚Éˆê’v‚µ‚Ü‚·"
  End If
End If

'sample3 code is if,And
Debug.Print "--if_and: sample3 if,And n=11--------------"
n = 11
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 ‚ÌğŒ®‚Éˆê’v‚µ‚Ü‚·"
Else
  Debug.Print "false: ğŒ‚Éˆê’v‚µ‚Ü‚¹‚ñ"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ğŒ1 or ğŒ1 Then
'   ğŒ1‚ÆğŒ2‚Ì‚¢‚¸‚ê‚©‚Éˆê’v‚·‚éê‡‚Ìˆ—
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_or()
Dim str As String
'sample1 code is if,Or
Debug.Print "--if_or: sample1 if,or str=""apple""------"
str = "apple"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: ğŒ‚Éˆê’v‚µ‚Ü‚µ‚½"
End If

'sample2 code is if,Or
Debug.Print "--if_or: sample2 if,or str=""banana""------"
str = "banana"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: ğŒ‚Éˆê’v‚µ‚Ü‚µ‚½"
Else
  Debug.Print "false: ğŒ‚Éˆê’v‚µ‚Ü‚¹‚ñ‚Å‚µ‚½"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If Not ğŒ Then
'   ğŒ‚ğ”Û’è‚·‚éê‡‚Ìˆ—
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_not()
Dim n As Integer
'sample1 code is if,not
Debug.Print "--if_not: sample1 if,not n=1---------------"
n = 1
If Not n = 0 Then
  Debug.Print "true: Not n=0: ğŒ®‚Æˆê’v‚µ‚Ü‚·"
End If

'sample2 code is if,not
Debug.Print "--if_not: sample2 if,not n=1---------------"
If Not n = 0 And Not n > 10 Then
  Debug.Print "true: n=0,n>10 ğŒ‚Æˆê’v‚µ‚Ü‚·"
End If
End Sub




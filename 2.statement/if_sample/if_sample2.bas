Attribute VB_Name = "if_sample2"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ����1 and ����1 Then
'   ����1�Ə���2�Ɉ�v����ꍇ�̏���
' Else
'   �����Ɉ�v���Ȃ��ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Sub if_and()
Dim n As Integer:
'sample1 code is if,And
Debug.Print "--if_and: sample1 if,And n=5----------------"
n = 5
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 �̏������Ɉ�v���܂�"
End If

'sample2 code is if,if
Debug.Print "--if_and: sample2 if,true n=5---------------"
If 10 > n Then
  If n > 0 Then
    Debug.Print "true: 10 > n, n > 0 �̏������Ɉ�v���܂�"
  End If
End If

'sample3 code is if,And
Debug.Print "--if_and: sample3 if,And n=11--------------"
n = 11
If 10 > n And n > 0 Then
  Debug.Print "true: 10 > n, n > 0 �̏������Ɉ�v���܂�"
Else
  Debug.Print "false: �����Ɉ�v���܂���"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ����1 or ����1 Then
'   ����1�Ə���2�̂����ꂩ�Ɉ�v����ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_or()
Dim str As String
'sample1 code is if,Or
Debug.Print "--if_or: sample1 if,or str=""apple""------"
str = "apple"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: �����Ɉ�v���܂���"
End If

'sample2 code is if,Or
Debug.Print "--if_or: sample2 if,or str=""banana""------"
str = "banana"
If str = "orange" Or str = "apple" Then
  Debug.Print "true: �����Ɉ�v���܂���"
Else
  Debug.Print "false: �����Ɉ�v���܂���ł���"
End If
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If Not ���� Then
'   ������ے肷��ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_not()
Dim n As Integer
'sample1 code is if,not
Debug.Print "--if_not: sample1 if,not n=1---------------"
n = 1
If Not n = 0 Then
  Debug.Print "true: Not n=0: �������ƈ�v���܂�"
End If

'sample2 code is if,not
Debug.Print "--if_not: sample2 if,not n=1---------------"
If Not n = 0 And Not n > 10 Then
  Debug.Print "true: n=0,n>10 �����ƈ�v���܂�"
End If
End Sub




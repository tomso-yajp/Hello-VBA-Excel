Attribute VB_Name = "if_sample1"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ���� Then
'   '�����ƈ�v����ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_true()
Dim n1 As Integer: n1 = 0
Dim n2 As Integer: n2 = 10
'sample1 code
Debug.Print "--if_true: sample code----------------------"
If n2 > n1 Then
  Debug.Print "�C�~�f�B�G�C�g�E�B���h�E�ɕ\������܂�"
  Debug.Print "n1: " & n1 & " ��� n2: " & n2 & " ���傫���ł�"
  Debug.Print "n2 > n1 �̏������Ɉ�v���܂�"
End If
Debug.Print vbLf
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ���� Then
'    �����ƈ�v����ꍇ�̏���
' else
'�@�@�����ƈ�v���Ȃ��ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_else()
Dim n As Integer
Dim str As String:

n = 0
'sample1 code is true(if)
Debug.Print "--if_else: sample1 n=10---------------------"
If n > 5 Then
  Debug.Print "true: n > 0 �̏������Ɉ�v���܂�"
Else
  Debug.Print "false: n > 0 �̏������Ɉ�v���܂���"
End If
Debug.Print "n = " & n

n = 5
'sample2 code is false(else)
Debug.Print "--if_else: sample2 n=10---------------------"
If n > 11 Then
  Debug.Print "true: n > 11 �̏������Ɉ�v���܂�"
Else
  Debug.Print "false: n > 11 �̏������Ɉ�v���܂���"
End If
Debug.Print "n = " & n

'sample3 code is true(if)
Debug.Print "--if_else: sample3 str=""apple""-------------"
str = "apple"
If str = "apple" Then
  Debug.Print "true: str = ""apple"" �̏����Ɉ�v���܂�"
Else
  Debug.Print "false: str = ""apple"" �̏����Ɉ�v���܂���"
End If

'sample4 code is false(else)
Debug.Print "--if_else: sample4 str=""apple""-------"
If str = "orange" Then
  Debug.Print "true: str = ""orange"" �̏����Ɉ�v���܂�"
Else
  Debug.Print "false: str = ""orange"" �̏����Ɉ�v���܂���"
End If
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' If ����1 Then
'    ����1�ƈ�v����ꍇ�̏���
' elseif ����2
'    ����2�ƈ�v����ꍇ�̏���
' else
'�@�@�����ƈ�v���Ȃ��ꍇ�̏���
' End If
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_elseif()
Dim n As Integer
Dim str As String
'sample1 code is the if
Debug.Print "--if_elseif: sample1 n=0-------------------"
n = 0
If n = 0 Then
  Debug.Print "true: n �ϐ��̒l�́A0 �ł�"
ElseIf n = 1 Then
  Debug.Print "true: n �ϐ��̒l�́A1 �ł�"
Else
  Debug.Print "false: n �ϐ��̒l�́A�����ƈ�v���܂���ł���"
End If

'sample2 code is the elseif
Debug.Print "--if_elseif: sample2 n=1-------------------"
n = 1
If n = 0 Then
  Debug.Print "true: n �ϐ��̒l�́A0 �ł�"
ElseIf n = 1 Then
  Debug.Print "true: n �ϐ��̒l�́A1 �ł�"
Else
  Debug.Print "false: n �ϐ��̒l�́A�����ƈ�v���܂���ł���"
End If

'sample3 code is the else
Debug.Print "--if_elseif: sample3 n=2------------------"
n = 2
If n = 0 Then
  Debug.Print "true: n �ϐ��̒l�́A0 �ł�"
ElseIf n = 1 Then
  Debug.Print "true: n �ϐ��̒l�́A1 �ł�"
Else
  Debug.Print "false: n �ϐ��̒l�́A�����ƈ�v���܂���ł���"
End If
End Sub









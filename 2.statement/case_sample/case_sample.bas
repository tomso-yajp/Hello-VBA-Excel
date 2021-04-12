Attribute VB_Name = "case_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  case_num1:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub case_num1()
Dim n As Integer

n = WorksheetFunction.RandBetween(1, 10)
Select Case n
Case 1: Debug.Print "1�`10�F"
Case 2: Debug.Print "2�̔{���F"
Case 3: Debug.Print "3�̔{���F"
Case 4: Debug.Print "4�̔{���F"
Case 5: Debug.Print "5�̔{���F"
Case 6: Debug.Print "6�̔{���F"
Case 7: Debug.Print "7�̔{���F"
Case 8: Debug.Print "8�̔{���F"
Case 9: Debug.Print "9�̔{���F"
Case Else
Debug.Print "no thanks"
End Select
Call case_sub(n)

End Sub
Sub num_sub(Optional n As Integer = 1)
Dim i As Integer
Debug.Print ;
For i = 1 To 9
  If i < 9 Then Debug.Print n * i; Else Debug.Print n * i
Next
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  case_num2:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub case_num2()
Dim vs As Variant, vn As Variant
Dim n As Variant

n = WorksheetFunction.RandBetween(1, 40)
Select Case n
Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
  Debug.Print "1-10 ���l�ł�"
Case 11 To 20
  Debug.Print "11-20 ���l�ł�"
Case 21 To 30
  Debug.Print "21-30 ���l�ł�"
Case Is > 30
  Debug.Print "30 ���傫�����l�ł�"
End Select

End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  case_str:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub case_str()
Dim v As Variant
Dim str As String
Dim n As Integer: n = WorksheetFunction.RandBetween(0, 8)
v = "����,���Ђ�,�˂�,����,������,�ւ�,���ʂ�,�炢����,����"
v = Split(v, ",")
str = v(n)
Select Case str
Case "horse", v(0): Debug.Print "�q�q�[��"
Case "duck", v(1): Debug.Print "�K�[�A�K�[�A"
Case "cat", v(2): Debug.Print "�j���[���A�j���I"
Case "dog", v(3): Debug.Print "�����A���I��"
Case "rabbit", v(4): Debug.Print "�s����"
Case "snake", v(5): Debug.Print "�j�����A�j����"
Case "raccoon", v(6): Debug.Print "�|���|�R�|��"
Case "lion", v(7): Debug.Print "�K�I�["
Case "fox", v(8): Debug.Print "�R���A�R�[��"
Case Else
End Select
End Sub








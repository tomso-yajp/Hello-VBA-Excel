Attribute VB_Name = "case_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  case_num1:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub case_num1()
Dim n As Integer

n = WorksheetFunction.RandBetween(1, 10)
Select Case n
Case 1: Debug.Print "1～10："
Case 2: Debug.Print "2の倍数："
Case 3: Debug.Print "3の倍数："
Case 4: Debug.Print "4の倍数："
Case 5: Debug.Print "5の倍数："
Case 6: Debug.Print "6の倍数："
Case 7: Debug.Print "7の倍数："
Case 8: Debug.Print "8の倍数："
Case 9: Debug.Print "9の倍数："
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
  Debug.Print "1-10 数値です"
Case 11 To 20
  Debug.Print "11-20 数値です"
Case 21 To 30
  Debug.Print "21-30 数値です"
Case Is > 30
  Debug.Print "30 より大きい数値です"
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
v = "うま,あひる,ねこ,いぬ,うさぎ,へび,たぬき,らいおん,きつね"
v = Split(v, ",")
str = v(n)
Select Case str
Case "horse", v(0): Debug.Print "ヒヒーン"
Case "duck", v(1): Debug.Print "ガー、ガーア"
Case "cat", v(2): Debug.Print "ニャーン、ニャオ"
Case "dog", v(3): Debug.Print "ワン、ワオン"
Case "rabbit", v(4): Debug.Print "ピョン"
Case "snake", v(5): Debug.Print "ニョロ、ニョロ"
Case "raccoon", v(6): Debug.Print "ポンポコポン"
Case "lion", v(7): Debug.Print "ガオー"
Case "fox", v(8): Debug.Print "コン、コーン"
Case Else
End Select
End Sub








Attribute VB_Name = "if_sample3"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' IsArray: returns true if the variable is an array
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_isarray()
Dim v As Variant
v = "orange,apple,banana"
'sample2 code is if,isarray
Debug.Print "--if_isarray: sample2 if,isarray------------"
If IsArray(v) Then
  Debug.Print "true: is an array: �z��ł�"
Else
  Debug.Print "false: not an array: �z��ł́A����܂���"
  Debug.Print "�ϐ�:v �́A" & gettype(v) & "�ł�"
End If

v = Split(v, ",")
'sample2 code is if,isarray
Debug.Print "--if_isarray: sample2 if,isarray------------"
If IsArray(v) Then
  Debug.Print "true: is an array: �z��ł�"
Else
  Debug.Print "false: not an array: �z��ł́A����܂���"
  Debug.Print "�ϐ�:v �́A" & gettype(v) & "�ł�"
End If

Erase v
'sample3 code is if,isarray
Debug.Print "--if_isarray: sample3 if,isarray-------------"
If IsArray(v) Then
  Debug.Print "true: is an array: �z��ł�"
Else
  Debug.Print "false: not an array: �z��ł́A����܂���"
  Debug.Print "�ϐ�:v �́A" & gettype(v) & "�ł�"
End If
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' IsNumeric: return true if the variable is a number
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_isnumeric()
Dim n As Variant
n = 0
'sample1 code is if,isnumeric
Debug.Print "--if_isnumeric: sample1 if,isnumeric----------"
If IsNumeric(n) Then
  Debug.Print "true: is a number: ���l�ł�"
Else
  Debug.Print "false: not a number: ���l�ł͂���܂���"
  Debug.Print "�ϐ�:n �́A" & gettype(n) & "�ł�"
End If

n = "orange"
'sample2 code is if,isnumeric
Debug.Print "--if_isnumeric: sample2 if,isnumeric----------"
If IsNumeric(n) Then
  Debug.Print "true: is a number: ���l�ł�"
Else
  Debug.Print "false: not a number: ���l�ł͂���܂���"
  Debug.Print "�ϐ�:n �́A" & gettype(n) & "�ł�"
End If

n = ""
'sample3 code is if,isnumeric
Debug.Print "--if_isnumeric: sample3 if,isnumeric----------"
If IsNumeric(n) Then
  Debug.Print "true: is a number: ���l�ł�"
Else
  Debug.Print "false: not a number: ���l�ł͂���܂���"
  Debug.Print "�ϐ�:n �́A" & gettype(n) & "�ł�"
End If
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' IsDate: return true if the variable is a date
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub if_isdate()
Dim d As Variant
d = Date
'sample1 code is if,isdate
Debug.Print "--if_isdate: sample1 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: ���t�ł�"
Else
  Debug.Print "false: not a date: ���t�ł͂���܂���"
  Debug.Print "�ϐ�:d �́A" & gettype(d) & "�ł�"
End If

d = Time()
'sample2 code is if,isdate
Debug.Print "--if_isdate: sample2 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: �����ł�"
Else
  Debug.Print "false: not a date: �����ł͂���܂���"
  Debug.Print "�ϐ�:d �́A" & gettype(d) & "�ł�"
End If

d = Now()
'sample3 code is if,isdate
Debug.Print "--if_isdate: sample3 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: ���t�Ǝ����ł�"
Else
  Debug.Print "false: not a date: ���t�Ǝ����ł͂���܂���"
  Debug.Print "�ϐ�:d �́A" & gettype(d) & "�ł�"
End If

d = ""
'sample4 code is if,isdate
Debug.Print "--if_isdate: sample4 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: ���t�^�ł�"
Else
  Debug.Print "false: not a date: ���t�^�ł͂���܂���"
  Debug.Print "�ϐ�:d �́A" & gettype(d) & "�ł�"
End If
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' IsObject: returns true if the variable is an object
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub is_isobject()
Dim o As Variant
Set o = ThisWorkbook
'sample1 code is if,isobject
Debug.Print "--if_isobject: sample1 if,isobject----------"
If IsObject(o) Then
  Debug.Print "true: is an object: �I�u�W�F�N�g�ł�"
Else
  Debug.Print "false: not a object: �I�u�W�F�N�g�ł͂���܂���"
  Debug.Print "�ϐ�:o �́A" & gettype(o) & "�ł�"
End If

o = Split("orange,apple,banana", ",")
'sample2 code is if,isobject
Debug.Print "--if_isobject: sample2 if,isobject----------"
If IsObject(o) Then
  Debug.Print "true: is an object: �I�u�W�F�N�g�ł�"
Else
  Debug.Print "false: not a object: �I�u�W�F�N�g�ł͂���܂���"
  Debug.Print "�ϐ�:o �́A" & gettype(o) & "�ł�"
End If

o = ThisWorkbook.Name
'sample3 code is if,isobject
Debug.Print "--if_isobject: sample3 if,isobject----------"
If IsObject(o) Then
  Debug.Print "true: is an object: �I�u�W�F�N�g�ł�"
Else
  Debug.Print "false: not a object: �I�u�W�F�N�g�ł͂���܂���"
  Debug.Print "�ϐ�:o �́A" & gettype(o) & "�ł�"
End If

Set o = Nothing
End Sub
Function gettype(Optional v As Variant)
Dim ret  As String
If IsNumeric(v) Then
  ret = "�����^"
ElseIf IsDate(v) Then
  ret = "���t�^"
ElseIf IsNull(v) Then
  ret = "��"
ElseIf IsArray(v) Then
  ret = "�z��^"
ElseIf TypeName(v) = "String" Then
  ret = "String �^"
ElseIf TypeName(v) = "Varaiant" Then
  ret = "String �^"
ElseIf TypeName(v) = "Object" Then
  ret = "String �^"
End If
gettype = ret
End Function

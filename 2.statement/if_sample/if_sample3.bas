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
  Debug.Print "true: is an array: 配列です"
Else
  Debug.Print "false: not an array: 配列では、ありません"
  Debug.Print "変数:v は、" & gettype(v) & "です"
End If

v = Split(v, ",")
'sample2 code is if,isarray
Debug.Print "--if_isarray: sample2 if,isarray------------"
If IsArray(v) Then
  Debug.Print "true: is an array: 配列です"
Else
  Debug.Print "false: not an array: 配列では、ありません"
  Debug.Print "変数:v は、" & gettype(v) & "です"
End If

Erase v
'sample3 code is if,isarray
Debug.Print "--if_isarray: sample3 if,isarray-------------"
If IsArray(v) Then
  Debug.Print "true: is an array: 配列です"
Else
  Debug.Print "false: not an array: 配列では、ありません"
  Debug.Print "変数:v は、" & gettype(v) & "です"
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
  Debug.Print "true: is a number: 数値です"
Else
  Debug.Print "false: not a number: 数値ではありません"
  Debug.Print "変数:n は、" & gettype(n) & "です"
End If

n = "orange"
'sample2 code is if,isnumeric
Debug.Print "--if_isnumeric: sample2 if,isnumeric----------"
If IsNumeric(n) Then
  Debug.Print "true: is a number: 数値です"
Else
  Debug.Print "false: not a number: 数値ではありません"
  Debug.Print "変数:n は、" & gettype(n) & "です"
End If

n = ""
'sample3 code is if,isnumeric
Debug.Print "--if_isnumeric: sample3 if,isnumeric----------"
If IsNumeric(n) Then
  Debug.Print "true: is a number: 数値です"
Else
  Debug.Print "false: not a number: 数値ではありません"
  Debug.Print "変数:n は、" & gettype(n) & "です"
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
  Debug.Print "true: is a date: 日付です"
Else
  Debug.Print "false: not a date: 日付ではありません"
  Debug.Print "変数:d は、" & gettype(d) & "です"
End If

d = Time()
'sample2 code is if,isdate
Debug.Print "--if_isdate: sample2 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: 時刻です"
Else
  Debug.Print "false: not a date: 時刻ではありません"
  Debug.Print "変数:d は、" & gettype(d) & "です"
End If

d = Now()
'sample3 code is if,isdate
Debug.Print "--if_isdate: sample3 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: 日付と時刻です"
Else
  Debug.Print "false: not a date: 日付と時刻ではありません"
  Debug.Print "変数:d は、" & gettype(d) & "です"
End If

d = ""
'sample4 code is if,isdate
Debug.Print "--if_isdate: sample4 if,isdate--------------"
If IsDate(d) Then
  Debug.Print "true: is a date: 日付型です"
Else
  Debug.Print "false: not a date: 日付型ではありません"
  Debug.Print "変数:d は、" & gettype(d) & "です"
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
  Debug.Print "true: is an object: オブジェクトです"
Else
  Debug.Print "false: not a object: オブジェクトではありません"
  Debug.Print "変数:o は、" & gettype(o) & "です"
End If

o = Split("orange,apple,banana", ",")
'sample2 code is if,isobject
Debug.Print "--if_isobject: sample2 if,isobject----------"
If IsObject(o) Then
  Debug.Print "true: is an object: オブジェクトです"
Else
  Debug.Print "false: not a object: オブジェクトではありません"
  Debug.Print "変数:o は、" & gettype(o) & "です"
End If

o = ThisWorkbook.Name
'sample3 code is if,isobject
Debug.Print "--if_isobject: sample3 if,isobject----------"
If IsObject(o) Then
  Debug.Print "true: is an object: オブジェクトです"
Else
  Debug.Print "false: not a object: オブジェクトではありません"
  Debug.Print "変数:o は、" & gettype(o) & "です"
End If

Set o = Nothing
End Sub
Function gettype(Optional v As Variant)
Dim ret  As String
If IsNumeric(v) Then
  ret = "整数型"
ElseIf IsDate(v) Then
  ret = "日付型"
ElseIf IsNull(v) Then
  ret = "空白"
ElseIf IsArray(v) Then
  ret = "配列型"
ElseIf TypeName(v) = "String" Then
  ret = "String 型"
ElseIf TypeName(v) = "Varaiant" Then
  ret = "String 型"
ElseIf TypeName(v) = "Object" Then
  ret = "String 型"
End If
gettype = ret
End Function

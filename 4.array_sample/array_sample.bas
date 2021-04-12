Attribute VB_Name = "array_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' array_num:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub array_num()
Dim num As Variant
Dim num1() As Integer
Dim num2(5) As Integer
Dim num3(1 To 5) As Integer
Dim i As Integer

ReDim num1(4)
num1(0) = 0
num1(1) = 1
num1(2) = 2
num1(3) = 3
num1(4) = 4
Debug.Print num1(1) & " : " & num1(4)
Erase num1

num = Array(0, 1, 2, 3, 4)
Debug.Print num(1) & " : " & num(4)
Erase num

'sample1 code is array,int
Debug.Print "--array_int: sample1 array,int num1(0-5)----"
ReDim num1(5)
Debug.Print "num1: " & UBound(num1) & "  num1(0-4)"
For i = 0 To 5
  num1(i) = i
Next
Debug.Print "print out ::"
For i = LBound(num1) To UBound(num1)
  Debug.Print "num1(" & i & ") = " & num1(i)
Next

'sample2 code is array,int
Debug.Print "--array_int: sample2 array,int num1(0-5)----"
For i = 0 To 5
  num2(i) = i
Next
Debug.Print "print out ::"
For i = LBound(num2) To UBound(num2)
  Debug.Print "num2(" & i & ") = " & num2(i)
Next

'sample3 code is array,int
Debug.Print "--array_int: sample3 array,int num1(1-5)----"
For i = 1 To 5
  num3(i) = i
Next
Debug.Print "print out ::"
For i = LBound(num3) To UBound(num3)
  Debug.Print "num3(" & i & ") = " & num3(i)
Next
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' array_str:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub array_str()
Dim str As Variant
Dim str1() As String
Dim str2(5) As String
Dim str3(1 To 5) As String
Dim i As Integer

'arrays1
ReDim str1(4)
str1(0) = "orange"
str1(1) = "apple"
str1(2) = "banana"
str1(3) = "kiwi"
str1(4) = "mango"
Debug.Print str1(1) & " : " & str1(4)
Erase str1

'arrays2
str = Array("orange", "apple", "banana", "kiwi", "mango", "peach")
Debug.Print str(1) & " : " & str(4)
Erase str1

'arrays3
str = "orange,apple,banana,kiwi,mango,peach"
str = Split(str, ",")
Debug.Print str(1) & " : " & str(4)

'sample1 code is array,string
Debug.Print "--array_string: sample1 array,string str1(0-5)----"
ReDim str1(5)
For i = 0 To 5
  str1(i) = str(i)
Next
Debug.Print "print out ::"
For i = LBound(str1) To UBound(str1)
  Debug.Print "str1(" & i & ") = " & str1(i)
Next

'sample2 code is array,string
Debug.Print "--array_string: sample2 array,string str2(0-5)----"
For i = 0 To 5
  str2(i) = str(i)
Next
Debug.Print "print out ::"
For i = LBound(str2) To UBound(str2)
  Debug.Print "str2(" & i & ") = " & str2(i)
Next

'sample3 code is array,string
Debug.Print "--array_string: sample3 array,string str3(1-5)----"
For i = 1 To 5
  str3(i) = str(i)
Next
Debug.Print "print out ::"
For i = LBound(str3) To UBound(str3)
  Debug.Print "str3(" & i & ") = " & str3(i)
Next
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' array_var:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub array_var()
Dim v As Variant, var1 As Variant
Dim var2 As Variant, var3 As Variant
v = "orange,apple,banana,kiwi,mango,peach"

'sample1 code is array,variant
Debug.Print "--array_variant: sample1 array,variant var(0-5)----"
v = Split(v, ",")
For i = LBound(v) To UBound(v)
  If IsArray(var1) Then
    ReDim Preserve var1(UBound(var1) + 1)
  ElseIf i = LBound(v) Or Not IsArray(var1) Then
    var1 = Array("") ' Redim var1(0)
  End If
  var1(UBound(var1)) = v(i)
Next
Debug.Print "print out ::"
For i = LBound(var1) To UBound(var1)
  Debug.Print "var1(" & i & ") = " & var1(i)
Next

'sample2 code is array,variant
Debug.Print "--array_variant: sample2 array,variant var2(0-5)----"
var2 = Array("orange", "apple", "banana", "kiwi", "mango", "peach")
Debug.Print "print out ::"
For i = LBound(var2) To UBound(var2)
  Debug.Print "var2(" & i & ") = " & var2(i)
Next

'sample3 code is array,variant
Debug.Print "--array_variant: sample3 array,variant var3(0-5)----"
var3 = v
Debug.Print "print out ::"
For i = LBound(var3) To UBound(var3)
  Debug.Print "var3(" & i & ") = " & var3(i)
Next

'sample4
Debug.Print "--array_variant: sample4 array,variant,join v----"
Debug.Print "v = " & Join(v, ",")
End Sub

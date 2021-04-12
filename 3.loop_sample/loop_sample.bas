Attribute VB_Name = "loop_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'for num = start_number to end_number
'  Às‚·‚éˆ—
'next
'  for_next:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub for_next()
Dim i As Integer, n As Integer
Dim str As String
'sample1 code is for,next
Debug.Print "--for_next: sample1 for-next i=1`5--------"
Debug.Print "i = ";
For i = 1 To 5
  If 5 > i Then
    Debug.Print i;
  Else
    Debug.Print i
  End If
Next

'sample2 code is for,next
Debug.Print "--for_next: sample2 for-next i=1`5--------"
For i = 1 To 5
  Debug.Print "i = " & i
Next

'sample3 code is for,next
Debug.Print "--for_next: sample3 for-next i=1`5--------"
For i = 5 To 1 Step -1
  Debug.Print "i = " & i
Next

str = ""
'sample4 code is for,next
Debug.Print "--for_next: sample4 for-next i=1`10-------"
For i = 1 To 10
  str = str & i
  If i < 10 Then str = str & "+"
  n = n + i
Next
Debug.Print n & "=" & str
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'do while ğŒ
'  Às‚·‚éˆ—
'loop
'  do_while:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub do_while()
Dim i As Integer, n As Integer
Dim str As String
'sample1 code id do,while
Debug.Print "--do_while: sample1 do while-loop i=1`5----"
Debug.Print "i = ";
i = 1
Do While i < 6
  If i < 5 Then
    Debug.Print i;
  Else
    Debug.Print i
  End If
  i = i + 1
Loop

'sample2 code is do,while
Debug.Print "--do_while: sample2 do while-loop i=1`5----"
i = 1
Do While i < 6
  Debug.Print "i = " & i
  i = i + 1
Loop

'sample3 code is do,while
Debug.Print "--do_while: sample3 do while-loop i=1`5----"
i = 5
Do While i > 0
  Debug.Print "i = " & i
  i = i - 1
Loop

'sample4 code is do,while
Debug.Print "--do_while: sample4 do while-loop i=1`10---"
i = 1: str = ""
Do While i < 11
  str = str & i
  If i < 10 Then str = str & "+"
  i = i + 1
Loop
Debug.Print i & "=" & str
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'Dim v as Varaiant
'for each v in collection,array
'   Às‚·‚éğŒ
'next
'  for_each:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub for_each()
Dim v As Variant, arrays As Variant
Dim i As Integer
arrays = Array(1, 2, 3, 4, 5)
'sample1 code is for_each,in,next
Debug.Print "--for_each: sample1 for each,in-next i=1`5----"
Debug.Print "v = ";
For Each v In arrays
  If v <> 5 Then
    Debug.Print v;
  Else
    Debug.Print v
  End If
Next

'sample2 code is for_each,in,next
Debug.Print "--for_each: sample2 for each,in-next i=1`5----"
For Each v In arrays
  Debug.Print "v = " & v
Next

'sample3 code is for_each,in,next
Debug.Print "--for_each: sample3 for each,in-next i=1`5----"
With ThisWorkbook
  For i = UBound(arrays) To LBound(arrays) Step -1
    .Worksheets.Add.Name = arrays(i)
  Next
  Debug.Print "‘S‚Ä‚ÌƒV[ƒg–¼‚ğ•\¦‚µ‚Ü‚·" & vbLf & "(’Ç‰Á‚µ‚½ƒV[ƒg–¼‚ğŠÜ‚Ş)"
  For Each v In .Worksheets
    Debug.Print v.Name
  Next
End With

End Sub
Sub delsheet()
Dim s As Variant
Application.DisplayAlerts = False
With ThisWorkbook
  For Each s In .Worksheets
    If s.Name Like "[1-5]" Then s.Delete
  Next
End With
Application.DisplayAlerts = True
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  loop_break:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub loop_break()
Dim v As Variant
Dim i As Integer, n As Integer
Dim str As String
'sample1 code is for,break
Debug.Print "--loop_break: sample1 for-break------------"
Debug.Print "i = ";
For i = 1 To 20
  If i < 5 Then
    Debug.Print i;
  Else
    Debug.Print i: Exit For '’†’f‚µ‚Äƒ‹[ƒv‚©‚ç”²‚¯‚Ü‚·
  End If
Next

'sample2 code is do while,break
Debug.Print "--loop_break: sample2 do while-break--------"
Debug.Print "i = ";
i = 1
Do While i < 20
  If i < 5 Then
    Debug.Print i;
  Else
    Debug.Print i: Exit Do '’†’f‚µ‚Äƒ‹[ƒv‚©‚ç”²‚¯‚Ü‚·
  End If
  i = i + 1
Loop
End Sub


Attribute VB_Name = "msg_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  msg_button:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub msg_button()
Dim n As Integer
n = MsgBox("this is vbYesNo", vbYes, "msgbox:")
Debug.Print "vbyes:" & n
n = MsgBox("this is vbYesNo", vbYesNo, "msgbox:")
Debug.Print "vbYesNo:" & n
n = MsgBox("this is vbYesNoCancel", vbYesNoCancel, "msgbox:")
Debug.Print "vbYesNoCancel:" & n

n = MsgBox("this is vbOKOnly", vbOKOnly, "msgbox:")
Debug.Print "vbOKOnly:" & n
n = MsgBox("this is vbOK", vbOKCancel, "msgbox:")
Debug.Print "vbOK:" & n
n = MsgBox("this is vbCancel", vbCancele, "msgbox:")
Debug.Print "vbCancel:" & n

n = MsgBox("this is vbAbortRetryIgnore", vbAbortRetryIgnore, "msgbox:")
Debug.Print "vbAbortRetryIgnore:" & n
n = MsgBox("this is vbRetryCancel", vbRetryCancel, "msgbox:")
Debug.Print "vbRetryCancel:" & n
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  msg_icon:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub msg_icon()
Dim n As Integer
n = MsgBox("this is OK Only" & vbLf & "icon is vbCritical", _
            vbCritical, "msgbox:")
Debug.Print "vbCritical: " & n
n = MsgBox("this is OK Only" & vbLf & "icon is vbInformation", _
            vbInformation, "msgbox:")
Debug.Print "vbInformation: " & n
n = MsgBox("this is OK Only" & vbLf & "icon is vbQuestion", _
            vbQuestion, "msgbox:")
Debug.Print "vbQuestion: " & n
n = MsgBox("this is OK Only" & vbLf & "icon is vbExclamation", _
            vbExclamation, "msgbox:")
Debug.Print "vbExclamation: " & n
'icon vbCritical.
'sanple is vbYesNo,vbOKCancel,vbRetryCancel,vbAbortRetryIgnore
n = MsgBox("this is vbyes" & vbLf & "icon is vbCritical", _
            vbYesNo + vbCritical, "msgbox:")
Debug.Print "vbYesNo + vbCritical: " & n
n = MsgBox("this is vbYesNo" & vbLf & "icon is vbCritical", _
            vbOKCancel + vbCritical, "msgbox:")
Debug.Print "vbOKCancel + vbCritical: " & n
n = MsgBox("this is vbRetryCancel" & vbLf & "icon is vbCritical", _
            vbRetryCancel + vbCritical, "msgbox:")
Debug.Print "vbRetryCancel + vbCritical: " & n
n = MsgBox("this is vbAbortRetryIgnore" & vbLf & "icon is vbCritical", _
            vbAbortRetryIgnore + vbCritical, "msgbox:")
Debug.Print "vbAbortRetryIgnore + vbCritical: " & n

'icon vbInformation.
'sanple is vbYesNo,vbOKCancel,vbRetryCancel,vbAbortRetryIgnore
n = MsgBox("this is vbyes" & vbLf & "icon is vbInformation", _
            vbYesNo + vbInformation, "msgbox:")
Debug.Print "vbYesNo + vbInformation: " & n
n = MsgBox("this is vbYesNo" & vbLf & "icon is vbInformation", _
            vbOKCancel + vbInformation, "msgbox:")
Debug.Print "vbOKCancel + vbInformation: " & n
n = MsgBox("this is vbRetryCancel" & vbLf & "icon is vbInformation", _
            vbRetryCancel + vbInformation, "msgbox:")
Debug.Print "vbRetryCancel + vbInformation: " & n
n = MsgBox("this is vbAbortRetryIgnore" & vbLf & "icon is vbInformation", _
            vbAbortRetryIgnore + vbInformation, "msgbox:")
Debug.Print "vbAbortRetryIgnore + vbInformation: " & n

'icon vbQuestion.
'sanple is vbYesNo,vbOKCancel,vbRetryCancel,vbAbortRetryIgnore
n = MsgBox("this is vbyes" & vbLf & "icon is vbQuestion", _
            vbYesNo + vbQuestion, "msgbox:")
Debug.Print "vbYesNo + vbQuestion: " & n
n = MsgBox("this is vbYesNo" & vbLf & "icon is vbQuestion", _
            vbOKCancel + vbQuestion, "msgbox:")
Debug.Print "vbOKCancel + vbQuestion: " & n
n = MsgBox("this is vbRetryCancel" & vbLf & "icon is vbQuestion", _
            vbRetryCancel + vbQuestion, "msgbox:")
Debug.Print "vbRetryCancel + vbQuestion: " & n
n = MsgBox("this is vbAbortRetryIgnore" & vbLf & "icon is vbQuestion", _
            vbAbortRetryIgnore + vbQuestion, "msgbox:")
Debug.Print "vbAbortRetryIgnore + vbQuestion: " & n

'icon vbExclamation.
'sanple is vbYesNo,vbOKCancel,vbRetryCancel,vbAbortRetryIgnore
n = MsgBox("this is vbyes" & vbLf & "icon is vbExclamation", _
            vbYesNo + vbExclamation, "msgbox:")
Debug.Print "vbYesNo + vbExclamation: " & n
n = MsgBox("this is vbYesNo" & vbLf & "icon is vbExclamation", _
            vbOKCancel + vbExclamation, "msgbox:")
Debug.Print "vbOKCancel + vbExclamation: " & n
n = MsgBox("this is vbRetryCancel" & vbLf & "icon is vbExclamation", _
            vbRetryCancel + vbExclamation, "msgbox:")
Debug.Print "vbRetryCancel + vbExclamation: " & n
n = MsgBox("this is vbAbortRetryIgnore" & vbLf & "icon is vbExclamation", _
            vbAbortRetryIgnore + vbExclamation, "msgbox:")
Debug.Print "vbAbortRetryIgnore + vbExclamation: " & n

End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  msg_sample1:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub msg_sample1()
Dim cat As Variant: cat = Array("", "", "", "", "")
Dim str As Variant, a As Variant
Dim ret As String
Dim i As Integer
str = Array( _
      "�D���ȉʕ�����͂��ĉ�����", _
      "���t�^�̓��͂����肢���܂�", _
      "���ʂ̓��͂����肢���܂�", _
      "�Ō�ɂ�낵����ΔN��̓��͂����肢���܂�")
a = Array( _
    "orange,apple,banana,peach,kiwi", _
    "O,A,AB,B", _
    "man,woman", _
    "11,21,31,41,51,61,?", _
    "��g,���g,���g,�g,��,�勥")
    
For i = LBound(cat) To UBound(cat)
  ret = ""
  If i < UBound(cat) Then
    ret = InputBox(str(i) & vbLf & vbLf & _
          "���͂̊m�F���ł��Ȃ��ꍇ�͎����ǋL����܂�")
  End If
  If ret = "" Then
    a(i) = Split(a(i), ",")
    ret = a(i)(WorksheetFunction.RandBetween(0, UBound(a(i))))
  End If
  cat(i) = ret
Next
a = "�ʁ@��,���t�^,���@��,�N�@��,�^�@��"
a = Split(a, ",")
For i = LBound(a) To UBound(a)
  cat(i) = "          " & a(i) & "�F  " & cat(i)
Next
a = "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<"
ret = ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
MsgBox a & vbLf & Join(cat, vbLf) & vbLf & ret, vbInformation, "info�F"
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  msg_sample2:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub sample2()
Dim cat As Variant: cat = Array("", "", "", "")
Dim a As Variant, b As Variant, str As String
Dim i As Integer, ii As Integer, n As Integer
a = Array( _
  "�����́A�Ƃ�܂������H", _
  "���H�́A�H�ׂ܂������H," & _
    "��؂��Ƃ�܂������H" & vbLf & _
    "(���H��H�ׂ����̎���ł�)", _
  "���C�x�́A�ǂ��ł����H", _
  "���͂��肪�Ƃ��������܂�")
For i = LBound(a) To UBound(a)
  b = Split(a(i), ",")
  For ii = LBound(b) To UBound(b)
    If InStr(b(ii), "�H") Then
      If cat(i) = 7 Then GoTo goto1
      n = MsgBox("����" & i + 1 & "�F" & vbLf & b(ii), _
          vbYesNo + vbQuestion, Left(b(ii), 2) & "�F")
    Else
      n = MsgBox(b(ii), vbYesNo, Left(b(ii), 2) & "�F")
    End If
goto1:
    cat(i) = cat(i) & n
    If ii < UBound(b) Then cat(i) = cat(i) & ","
  Next
  If Right(cat(i), 1) = "," Then _
    cat(i) = Mid(cat(i), 1, Len(cat(i)) - 1)
Next
  cat = Replace(Replace(Join(cat, ";"), 6, "�͂�"), 7, "������")
  cat = Split(cat, ";")
For i = LBound(cat) To UBound(cat)
  b = Split(cat(i), ",")
  str = ""
  For ii = LBound(b) To UBound(b)
    b(ii) = "�@" & Left(Split(a(i), ",")(ii), 2) & "�@�@�F�@�@" & b(ii)
    If ii > LBound(b) Then
      b(ii) = "�@�@" & Replace(b(ii), "�@�@�F", "�F")
    End If
  cat(i) = Join(b, vbLf)
  Next
Next
MsgBox Join(cat, vbLf), vbInformation, "�A���P�[�g"
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  msg_sample3:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub msg_sample3()
Dim f As Object: Set f = WorksheetFunction
Dim sname As String: sname = "���N�Ǘ�"
Dim rc As Variant, cols As Variant
Dim msg As Variant, str As String
Dim rw As Integer, cl As Integer
Dim i As Integer, ret As Integer
rc = make_data: cols = msg_data(1)
'rc = Array(2, 2, 0, 0)
With ThisWorkbook
  With .Worksheets(sname)
    rw = rc(0) + 1
    .Rows(rw).Insert shift:=xlShiftDown
    .Rows(rw).Interior.Color = xlNone
    .Cells(rw, rc(1)) = Date
  
    For i = 1 To UBound(cols)
      str = ""
      msg = Split(cols(i), ";")
      cl = f.Match(msg(0), .Rows(rc(0)), 0)
      ret = vbYesNo + vbQuestion
      If Right(msg(1), 1) <> "�H" Then ret = vbYesNo
      ret = MsgBox(msg(1), ret, msg(0))
      If ret = 6 Then str = "�Z" Else str = "�~"
      If UBound(msg) = 2 Then
        If ret = 6 Then str = InputBox(msg(2), msg(0))
      End If
      If i < UBound(cols) Then .Cells(rw, cl) = str _
        Else .Cells(rw, cl).Select
    Next
  End With
End With
Set f = Nothing
End Sub
Function msg_data()
Dim cols As Variant, col As String
Dim i As Integer
cols = "���t," & _
       "����;�����́A���܂�����?," & _
       "���H;���H�́A���܂�����?," & _
       "����;�����́A����܂������H;" & _
         "�����̓��͂����肢���܂�" & vbLf & _
         "��1�D60-116 " & vbLf & "��2�D66 116 ," & _
       "�����l;�����l�́A����܂������H;" & _
         "�����l�̓��͂����肢���܂�," & _
       "���C�x;�����̌��C�x�́A�ǂ��ł����H," & _
       "���l;��������ꍇ�́A���l���ɋL�����Ă�������"
cols = Split(cols, ",")
For i = 0 To UBound(cols)
  col = col & Split(cols(i), ";")(0)
  If i < UBound(cols) Then col = col & ","
Next
msg_data = Array(col, cols)
End Function
Function make_data()
Dim sname As String: sname = "���N�Ǘ�"
Dim cols As Variant, rc As Variant
Dim i As Integer, str As String
cols = msg_data(0)
cols = Split(cols, ",")
rc = Array(2, 2, 0, 0): rc(2) = rc(1) + UBound(cols)
'Call del_sheet(sname):
If sheets(sname) = 0 Then Call add_sheet(CVar(sname))
With ThisWorkbook
  With .Worksheets(sname)
    For i = 0 To UBound(cols)
      If WorksheetFunction.CountIf(.Rows(rc(0)), cols(i)) = 0 Then
        .Cells(rc(0), rc(1) + i) = cols(i)
      End If
    Next
    With .Range(Cells(rc(0), rc(1)), Cells(rc(0), rc(2)))
      .Interior.Color = RGB(200, 240, 250)
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
    End With
    With .Range(Cells(rc(0), rc(1)), Cells(rc(0) + 1, rc(2)))
      .Borders.LineStyle = xlContinuous
    End With
    str = Split(.Columns(1).Address(False, False), ":")(0)
    str = str & ":" & Split(.Columns(rc(1) - 1).Address(False, False), ":")(0)
    .Columns(.Columns(1).Address(True, False)).ColumnWidth = 2
  End With
End With

make_data = rc
End Function
Function sheets(Optional sname As Variant = "data")
Dim s As Variant
sheets = 0
With ThisWorkbook
  For Each s In .Worksheets
    If s.Name = sname Then sheets = 1: Exit For
  Next
End With
End Function
Sub add_sheet(Optional sname As Variant = "data")
Dim i As Integer
sname = Split(sname, ",")
With ThisWorkbook
  For i = UBound(sname) To LBound(sname) Step -1
    .Worksheets.Add.Name = sname(i)
  Next
End With
End Sub
Sub del_sheet(Optional sname As String = "data")
Dim s As Variant
Application.DisplayAlerts = False
With ThisWorkbook
  For Each s In .Worksheets
    If s.Name = sname Then s.Delete
  Next
End With
Application.DisplayAlerts = False
End Sub

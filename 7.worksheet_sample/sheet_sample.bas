Attribute VB_Name = "sheet_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  newsheet: create new worksheet
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub newsheet()
Dim sname As String: sname = "new"
Call delsheet(CVar(sname))
With ThisWorkbook
  .Worksheets.Add.Name = sname
  With .Worksheets.Add
    .Name = sname & 1
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  copyheet:copy worksheets
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub copyheet()
Dim sname As Variant: sname = "data,new"
Dim s As Variant, d As Variant
Dim i As Integer
sname = Split(sname, ",")
With ThisWorkbook
  Call delsheet(sname, 1)
  Call tab_color
  .Worksheets(1).Name = sname(1)
  .Worksheets(1).Cells.Clear
  .Worksheets.Add.Name = sname(0)
  Call data_sheet(ThisWorkbook)
  s = getsheet
  .Worksheets(sname).Copy after:=.Worksheets(.Worksheets.Count)
  d = getsheet
  d = Split(diff_array(s, d), ",")
  For i = 0 To UBound(d)
    .Worksheets(d(i)).Tab.ColorIndex = 1
  Next
End With
End Sub
'_____________________________________________________________________
Function diff_array(Optional s As Variant, _
    Optional sname As Variant)
Dim d As Variant
Dim i As Integer, ii As Integer
If Not IsArray(s) Then s = Split(s, ",")
If Not IsArray(sname) Then sname = Split(sname, ",")
For i = LBound(s) To UBound(s)
  d = ""
  For ii = LBound(sname) To UBound(sname)
    If s(i) <> sname(ii) Then
      d = d & sname(ii)
      If ii < UBound(sname) Then d = d & ","
    End If
  Next
  sname = Split(d, ",")
Next
diff_array = d
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  copysheet_book1: Move sheet of workbook
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub copysheet_book1()
Dim p As String, sname As Variant
Dim rc As Variant, book As Variant
rc = Array(0, 0, 0, 0)
Set book = ThisWorkbook
sname = data_book
p = sname(0): sname = sname(1)
Call delsheet(sname)
With Application
  .ScreenUpdating = False
  With Workbooks.Open(p)
    .Worksheets(1).Copy before:=book.Worksheets(1)
    .Close savechanges:=False
  End With
  .ScreenUpdating = True
End With
book.Worksheets(1).Tab.ColorIndex = 1
Set book = Nothing
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  1.create excel with hide
'  2.open excel workbook
'  3.add sheet contents to sheet
'  4.copy sheet formats to sheet
'  copysheet_book:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub copysheet_book2()
Dim p As String
Dim rc As Variant, book As Variant
rc = Array(0, 0, 0, 0)
Set book = ThisWorkbook.Worksheets(1)
book.Cells.Clear
p = data_book(0)
With CreateObject("excel.application")
  .Visible = False
  With .Workbooks.Open(p)
    With .Worksheets(1)
      With .UsedRange
        rc(0) = .Item(1, 1).Row
        rc(1) = .Item(1, 1).Column
        rc(2) = .Item(.Rows.Count, 1).Row
        rc(3) = .Item(1, .Columns.Count).Column
      End With
      book.Range(book.Cells(rc(0), rc(1)), book.Cells(rc(2), rc(3))) _
      = .UsedRange.Formula
      .UsedRange.Copy
      book.Range(book.Cells(rc(0), rc(1)), book.Cells(rc(2), rc(3))) _
      .PasteSpecial (xlPasteFormats)
    End With
    .Close savechanges:=False
  End With
  .Quit
End With
book.Tab.ColorIndex = 1
Set book = Nothing
End Sub
'______________________________________________________________________
Function data_book()
Dim book As Variant
Dim cp As String: cp = ThisWorkbook.Path
Dim wname As String: wname = "data"
cp = cp & "\" & wname
If Dir(cp & ".xlsm") <> "" Then Kill cp & ".xlsm"
With CreateObject("excel.application")
  .Visible = False
  Set book = .Workbooks.Add(-4167) 'xlwbatworksheet
  With book
    .SaveAs Filename:=cp, FileFormat:=52
    Call data_sheet(book)
    cp = .FullName
    .Close savechanges:=True
  End With
  .Quit
End With
data_book = Array(cp, wname)
Set book = Nothing
End Function
'______________________________________________________________________
Sub data_sheet(Optional book As Variant, _
    Optional wname As String = "data")
Dim rc As Variant: rc = Array(2, 2, 10, 8)
With book
  With .Worksheets(1)
    .Name = wname
    .Columns(1).ColumnWidth = 2
    .Columns(rc(1)).ColumnWidth = 5
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .Formula = "=RandBetween(1,100)"
      .Columns(1).Formula = "=Row()-2"
      .Rows(1).Formula = _
          "=MID(" & _
          "ADDRESS(ROW(),COLUMN(),2),1," _
          & "FIND(""$"",ADDRESS(ROW(),COLUMN(),2))-1)"
      .Rows(1).Interior.Color = RGB(200, 240, 250)
      .Rows(1).HorizontalAlignment = xlCenter
      .Rows(1).VerticalAlignment = xlCenter
      .Borders.LineStyle = xlContinuous
    End With
    .Cells(rc(0), rc(1)) = "NO"
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  savesheet_book:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub savesheet_book(Optional sname As Variant = "")
Dim cp As String, msg As String
Dim i As Integer, n As Integer
msg = "シートをブックとして保存しました" & vbLf & _
    "閉じますか"
If sname = "" Then sname = "data,new"
sname = Split(sname, ",")
Application.DisplayAlerts = False
With ThisWorkbook
  For i = UBound(sname) To 0 Step -1
    If checksheet(CStr(sname(i))) = 0 Then _
      .Worksheets.Add.Name = sname(i)
  Next
  cp = .Path & "\" & sname(0)
  .Worksheets(sname).Copy
End With
With ActiveWorkbook
  .SaveAs Filename:=cp, FileFormat:=52
  n = MsgBox(msg, vbYesNo + vbQuestion, "確認：")
  If n = 6 Then .Close savechanges:=False
End With

Application.DisplayAlerts = True
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  movesheet: move worksheet. option value is index,sheet name
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub movesheet(Optional sname As Variant = 1)
With ThisWorkbook
  .Worksheets(sname).Move after:=.Worksheets(.Worksheets.Count)
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  delsheet: delete worksheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub delsheet(Optional sname As Variant = "", Optional n As Integer = 0)
Dim s As Variant, str As String
Dim i As Integer
If Not IsArray(sname) Then sname = Split(sname, ",")
Application.DisplayAlerts = False
With ThisWorkbook
  For i = 0 To UBound(sname)
    For Each s In .Worksheets
      If .Worksheets.Count = 1 Then Exit For
      str = sname(i)
      If n = 1 Then str = str & "*"
      If s.Name Like str Then s.Delete
    Next
  Next
End With
Application.DisplayAlerts = True
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  getsheet: get to worksheet name:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function getsheet()
Dim s As Variant
Dim sname As String
With ThisWorkbook
  For Each s In .Worksheets
    sname = sname & s.Name & ","
    Debug.Print s.Name
  Next
End With
getsheet = Mid(sname, 1, Len(sname) - 1)
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  getsheet: get to worksheet name:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function checksheet(Optional sname As String = "data")
Dim s As Variant
checksheet = 0
With ThisWorkbook
  For Each s In .Worksheets
    If s.Name = sname Then checksheet = 1: Exit For
  Next
End With
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  tab_color: change tab color of worksheet
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function tab_color(Optional colors As Integer = 0)
Dim s As Variant
Dim sname As String
If colors = 0 Then colors = xlNone
With ThisWorkbook
  For Each s In .Worksheets
    s.Tab.ColorIndex = colors
  Next
End With
End Function

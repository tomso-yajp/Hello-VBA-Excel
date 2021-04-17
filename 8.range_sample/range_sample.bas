Attribute VB_Name = "range_sample"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_usedrange:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_usedrange()
Dim sname As String: sname = "usedrange"
Call addsheet(sname)
Call data_sheet(ThisWorkbook, sname)
With ThisWorkbook
  With .Worksheets(sname)
    If .Name <> wname Then .Name = sname
    With .UsedRange
      Debug.Print "--cells_usedrange：exsample1------------------------"
      Debug.Print "start row：" & .Rows(1).Row
      Debug.Print "start column：" & .Columns(1).Column
      Debug.Print "last row：" & .Rows(.Rows.Count).Row
      Debug.Print "last column：" & .Columns(.Columns.Count).Column
      Debug.Print "--cells_usedrange：exsample2-------------------------"
      Debug.Print "start row：" & .Item(1, 1).Row
      Debug.Print "start column：" & .Item(1, 1).Column
      Debug.Print "last row：" & .Item(.Rows.Count, 1).Row
      Debug.Print "last column：" & .Item(1, .Columns.Count).Column
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_range:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_range()
Dim sname As String: sname = "range"
Dim rw As Integer, cl As Integer
Dim lrw As Integer, lcl As Integer
Call addsheet(sname)
Call data_sheet(ThisWorkbook, sname)
With ThisWorkbook
  With .Worksheets(sname)
    If .Name <> wname Then .Name = sname
    Do While (1)
      rw = rw + 1:
      If WorksheetFunction.CountA(.Rows(rw)) > 0 Then
        cl = .Cells(rw, 1).End(xlToRight).Column
        Exit Do
      End If
    Loop
    lrw = .Cells(.Rows.Count, cl).End(xlUp).Row
    lcl = .Cells(rw, .Columns.Count).End(xlToLeft).Column
    With .Range(.Cells(rw, cl), .Cells(lrw, lcl))
      Debug.Print "--cells_range：exsample1---------------------------"
      Debug.Print "start row：" & rw
      Debug.Print "start column："; cl
      Debug.Print "last row：" & lrw
      Debug.Print "last column：" & lcl
    End With
  End With
End With
End Sub
'_____________________________________________________________________
Sub data_sheet(Optional book As Variant, _
    Optional wname As String = "data")
Dim rc As Variant: rc = Array(2, 2, 10, 8)
With book
  'If .Worksheets.Count = 0 Then .Worksheets.Add
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
'  cells_offset:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_offset()
Dim sname As String: sname = "offset"
Dim rc As Variant: rc = Array(2, 2, 5, 6)
Dim c(1, 2) As Variant, cc As Variant:
Dim rw As Integer, cl As Integer
rw = 1: cl = 1 'offset(rw,cl)
c(0, 0) = 0: c(0, 1) = 0: c(0, 2) = 255
c(1, 0) = 255: c(1, 1) = 0: c(1, 2) = 0
Call addsheet(sname)
With ThisWorkbook
  With .Worksheets(sname)
    .Columns(1).ColumnWidth = 2
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
      .Interior.Color = RGB(c(0, 0), c(0, 1), c(0, 2))
      Debug.Print "--cells_offset：exsample1---------------------------"
      Debug.Print "address：" & vbLf & .Address(False, False)
      With .Offset(rw, cl)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
        .Interior.Color = RGB(c(1, 0), c(1, 1), c(1, 2))
        Debug.Print "offset(" & rw & "," & cl & ").address：" & _
                    vbLf & .Address(False, False)
      End With
      With Intersect(.Cells, .Offset(rw, cl))
        cc = color_rgb(c)
        .Interior.Color = RGB(cc(0), cc(1), cc(2))
      End With
    End With
  End With
End With
End Sub
'_____________________________________________________________________
Function color_rgb(Optional colors As Variant)
Dim i As Integer, n As Integer
Dim c As Variant: c = Array(0, 0, 0)
n = UBound(colors) + 1
For i = 0 To UBound(colors)
  c(0) = c(0) + Round(colors(i, 0) / n)
  c(1) = c(1) + Round(colors(i, 1) / n)
  c(2) = c(2) + Round(colors(i, 2) / n)
Next
color_rgb = c
'微妙...
End Function

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_range:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_resize()
Dim sname As String: sname = "resize"
Dim rc As Variant: rc = Array(2, 2, 7, 6)
Dim rw As Integer, cl As Integer
rw = 2: cl = 0 'resize(rw,cl)
Call addsheet(sname)
With ThisWorkbook
  With .Worksheets(sname)
    .Columns(1).ColumnWidth = 2
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .Rows(1).Interior.Color = RGB(140, 240, 250)
      .Borders.LineStyle = xlContinuous
      .BorderAround LineStyle:=xlContinuous, Weight:=xlThick
      Debug.Print "--cells_resize：exsample1----------------------------"
      Debug.Print "address：" & vbLf & .Address(False, False)
      With .Resize(.Rows.Count - rw)
        Debug.Print "resize(" & rw & "," & cl & ").address：" & _
                    vbLf & .Address(False, False)
      End With
      With .Offset(.Rows.Count - rw).Resize(rw)
        .BorderAround LineStyle:=xlSlantDashDot, Weight:=xlThick
        .Borders(xlDiagonalUp).LineStyle = xlSlantDashDot
      End With
      .Resize(.Rows.Count - rw).BorderAround _
        LineStyle:=xlContinuous, Weight:=xlThick
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_address:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_address()
Dim sname As String: sname = "address"
Dim rw As Integer, cl As Integer
Dim lrw As Integer, lcl As Integer
Call addsheet(sname)
Call data_sheet(ThisWorkbook, sname)
With ThisWorkbook
  With .Worksheets(sname)
    If .Name <> wname Then .Name = sname
    Do While (1)
      rw = rw + 1:
      If WorksheetFunction.CountA(.Rows(rw)) > 0 Then
        cl = .Cells(rw, 1).End(xlToRight).Column
        Exit Do
      End If
    Loop
    lrw = .Cells(.Rows.Count, cl).End(xlUp).Row
    lcl = .Cells(rw, .Columns.Count).End(xlToLeft).Column
    
    Debug.Print "--cells_address：exsample1---------------------------"
    Debug.Print .Range(.Cells(rw, cl), .Cells(lrw, lcl)) _
                .Address(False, False)
    Debug.Print "left top   ：" & .Cells(rw, cl).Address(False, False)
    Debug.Print "right top  ：" & .Cells(rw, lcl).Address(False, False)
    Debug.Print "left lower ：" & .Cells(lrw, cl).Address(False, False)
    Debug.Print "right lower：" & .Cells(lrw, lcl).Address(False, False)
    
    Debug.Print "--cells_address：exsample2---------------------------"
    With .UsedRange '.Range(.Cells(rw, cl), .Cells(lrw, lcl))
      Debug.Print .Address(False, False)
      Debug.Print .Item(1, 1).Address(False, False)
      Debug.Print .Item(1, .Columns.Count).Address(False, False)
      Debug.Print .Item(.Rows.Count, 1).Address(False, False)
      Debug.Print .Item(.Rows.Count, .Columns.Count).Address(False, False)
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'1pt=1/72inch 1inch=25.4mm
'1pt=25.4mm/72inch a.0.3528mm
'1mm=72inch/25.4mm a.2.8346pt
'  cells_address: 1point 0.3528mm 'ratio = 0.0685 35.4
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_position()
Dim sname As String: sname = "position"
Dim rc As Variant: rc = Array(2, 2, 5, 10)
Dim cm As Variant
Dim a As Double, ratio As Double
Dim pt As Double, mm As Double, inch As Double
pt = 1: mm = 25.4: inch = 72: ratio = 0.0685
a = inch / mm
Call addsheet(sname)
With ThisWorkbook
  With .Worksheets(1)
    cm = Application.CentimetersToPoints(1)
    .Cells.Clear
    .Cells.ColumnWidth = 6
    '.Cells.RowHeight = 25
    a = Round(a, 0)
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .Interior.Color = RGB(200, 240, 250)
      Debug.Print "--cells_position-----------------------------------"
      Debug.Print "左位置：" & .Left & " pt"
      Debug.Print "上位置：" & .Top & " pt"
      Debug.Print "高さ：" & .Height & "pt, " & .Height / a & "mm"
      Debug.Print "横幅：" & .Width & "pt, " & .Width / a & "mm"
      .Borders.LineStyle = xlContinuous
      Debug.Print "面積：" & (.Height / a) * (.Width / a) * 0.01 & "cm"
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'1pt=1/72inch 1inch=25.4mm
'1pt=25.4mm/72inch a.0.3528mm
'1mm=72inch/25.4mm a.2.8346pt
'  cells_graphpaper: 1point 0.3528mm 'ratio = 0.0685
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_graphpaper()
Dim sname As String: sname = "graphpaper"
Dim rc As Variant: rc = Array(2, 2, 5, 10)
Dim cm As Variant: cm = 10 'mm unit
Dim a As Double, ratio As Double
Dim pt As Double, mm As Double, inch As Double
pt = 1: mm = 25.4: inch = 72: ratio = 0.0685
a = (inch / mm) * cm
Call addsheet(sname)
With ThisWorkbook
  With .Worksheets(1)
    'cm = Application.CentimetersToPoints(1)
    .Cells.Clear
    .Cells.ColumnWidth = a * ratio
    .Cells.RowHeight = a
    Debug.Print "セル幅 " & cm & "mm 間隔です"
    a = Round(a)
    With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
      .Interior.Color = RGB(200, 240, 250)
      .Borders.LineStyle = xlContinuous
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_copy:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_copy()
Dim sname As String: sname = "cells_copy"
Dim r As Variant, rc As Variant
Call addsheet(CVar(sname))
Call cols_data(sname)
With ThisWorkbook.Worksheets(sname)
  With .UsedRange
    rc = Array(.Rows(1).Row, _
               .Columns(1).Column, _
               .Rows(.Rows.Count).Row, _
               .Columns(.Columns.Count).Column)
  End With
  .Cells(rc(0) - 1, rc(1)) = "table1"
  Set r = .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
  
  With r.Offset(5)
    r.Copy .Cells
    .Item(0, 1) = "copy table1"
    .Rows(1).Interior.Color = RGB(100, 150, 250)
    .Rows(1).Font.Color = RGB(255, 255, 255)
  End With
  
  With r.Offset(10)
    .Borders.LineStyle = xlContinuous
    .Cells = r.Value
    .Item(0, 1) = "value copy table1"
    With .Rows(1)
      .Interior.Color = RGB(0, 0, 255)
      .Font.Color = RGB(255, 255, 255)
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
    End With
  End With
End With
End Sub
'_____________________________________________________________________
Function cols_data(Optional sname As String = "data", _
    Optional cols As Variant = "", Optional rws As Variant)
Dim rc As Variant:
Call addsheet(CVar(sname))
If cols = "" Then _
  cols = "no,monday,tuesday,wednesday,thurdday,friday,saturdsy,sunday"
If Not IsArray(cols) Then cols = Split(cols, ",")
If rws = "" Then rws = "=row()" & -rc(0)
rc = Array(2, 2, 5, UBound(cols) + 1)
With ThisWorkbook.Worksheets(sname)
  .Cells.Clear
  .Columns(rc(1) - 1).ColumnWidth = 2
  With .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
    .Borders.LineStyle = xlContinuous
    .Cells.Formula = "=RandBetween(1, 100)"
    .Columns(1).Formula = rws
    .Columns(1).ShrinkToFit = True
    With .Rows(1)
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlCenter
      .ShrinkToFit = True
      .Interior.Color = RGB(200, 240, 250)
      .Cells = cols
    End With
  End With
End With
cols_data = rc
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  cells_validation: input rule of cells
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub cells_validation()
Dim sname As String:
Dim r As Variant, rc As Variant:
Dim cols As Variant, rws As Variant
Dim s As Variant, ss As Variant
cols = "日付,出勤確認,出社,退社," & _
       "昼休憩,午後休憩,,残業開始,残業終了"
rws = "=text(today()+(row()-3),""yyyy/mm/dd"")"
sname = "cells_rule": Call addsheet(sname)
v = "出社,遅刻,早退,休日,有給"
rc = cols_data(sname, cols, rws)
With ThisWorkbook.Worksheets(sname)
  Set r = .Range(.Cells(rc(0), rc(1)), .Cells(rc(2), rc(3)))
  
  With r.Offset(1, 1).Resize(r.Rows.Count - 1, r.Columns.Count - 1)
    .Cells.ClearContents
    With .Columns(1).Validation
      .Delete
      .Add Type:=xlValidateList, Formula1:=v
      .IMEMode = xlIMEModeHiragana
    End With
    With .Offset(, 1).Resize(, r.Columns.Count - 2)
      .Cells.NumberFormatLocal = "[h]:mm"
      With .Validation
      .Delete
      .Add Type:=xlValidateDecimal, _
                 Operator:=xlLess, Formula1:=2
      .IMEMode = xlIMEModeAlpha
      End With
    End With
  End With
  
End With
Set r = Nothing
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  addsheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub addsheet(Optional sname As String = "")
Call delsheet(CVar(sname))
With ThisWorkbook
  .Worksheets.Add.Name = sname
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  delsheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub delsheet(Optional sname As Variant)
Dim s As Variant, i As Integer
sname = Split(sname, ",")
Application.DisplayAlerts = False
With ThisWorkbook
  For i = 0 To UBound(sname)
    For Each s In .Worksheets
      If s.Name = sname(i) Then s.Delete
    Next
  Next
End With
Application.DisplayAlerts = True
End Sub

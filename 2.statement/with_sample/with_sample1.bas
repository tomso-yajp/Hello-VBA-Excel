Attribute VB_Name = "with_sample1"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_workbook:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_this()
With ThisWorkbook
  Debug.Print "ブック名：" & .Name
  Debug.Print "ブックの親パス：" & .Path
  Debug.Print "ブックのパス：" & .FullName
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_book:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_book()
Dim wname As String: wname = "hello-vba.xlsm"
With Workbooks(wname)
  Debug.Print "ブック名：" & .Name
  Debug.Print "ブックの親パス：" & .Path
  Debug.Print "ブックのパス：" & .FullName
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_sheet:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_worksheet()
Dim sname As String: sname = ActiveSheet.Name
With ThisWorkbook
  With .Worksheets(sname)
    Debug.Print "表示されているシート名："; .Name
    .Select
    Debug.Print "選択されているセルは、" & _
                Replace(Selection.Item(1).Address, "$", "") & " です"
  End With
End With
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_range: upper left,upper right,lower left,lower right
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_range()
Dim sname As String: sname = ActiveSheet.Name
With ThisWorkbook.Worksheets(sname)
  With .Range(.Cells(2, 2), .Cells(5, 5))
    .Select
    Debug.Print "--------------------------------------------"
    Debug.Print "左上のセル：" & _
                Replace(.Item(1, 1).Address, "$", "")
    Debug.Print "右上のセル：" & _
                Replace(.Item(1, .Columns.Count).Address, "$", "")
    Debug.Print "左下のセル：" & _
                Replace(.Item(.Rows.Count, 1).Address, "$", "")
    Debug.Print "右下のセル：" & _
                Replace(.Item(.Rows.Count, .Columns.Count).Address, "$", "")
    Debug.Print "--------------------------------------------"
    Debug.Print "最初の行：" & .Item(1, 1).Row
    Debug.Print "最初の列：" & .Item(1, .Columns.Count).Column
    Debug.Print "最終の行：" & .Item(.Rows.Count, 1).Row
    Debug.Print "最終の列：" & .Item(.Rows.Count, .Columns.Count).Column
    Debug.Print "--------------------------------------------"
    
    Debug.Print "セルの範囲 " & Replace(.Item(1, 1).Address, "$", "") & ":" & _
                Replace(.Item(.Rows.Count, .Columns.Count).Address, "$", "") & _
                " を選択しました"
    .Interior.Color = RGB(0, 0, 255)
  End With
End With
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_application: excel application in use
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_application()
With Application
  Debug.Print .Name
  Debug.Print .ActiveWorkbook.Name
  Debug.Print .ActiveSheet.Name
End With
End Sub

'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_object: create excel application
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_object()
With CreateObject("excel.application")
  .Visible = True
  With .Workbooks.Add(xlWBATWorksheet)
    .Close SaveChanges:=False
  End With
  .Quit
End With
End Sub





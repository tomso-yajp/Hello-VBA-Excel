Attribute VB_Name = "with_sample1"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_workbook:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_this()
With ThisWorkbook
  Debug.Print "�u�b�N���F" & .Name
  Debug.Print "�u�b�N�̐e�p�X�F" & .Path
  Debug.Print "�u�b�N�̃p�X�F" & .FullName
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_book:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_book()
Dim wname As String: wname = "hello-vba.xlsm"
With Workbooks(wname)
  Debug.Print "�u�b�N���F" & .Name
  Debug.Print "�u�b�N�̐e�p�X�F" & .Path
  Debug.Print "�u�b�N�̃p�X�F" & .FullName
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
    Debug.Print "�\������Ă���V�[�g���F"; .Name
    .Select
    Debug.Print "�I������Ă���Z���́A" & _
                Replace(Selection.Item(1).Address, "$", "") & " �ł�"
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
    Debug.Print "����̃Z���F" & _
                Replace(.Item(1, 1).Address, "$", "")
    Debug.Print "�E��̃Z���F" & _
                Replace(.Item(1, .Columns.Count).Address, "$", "")
    Debug.Print "�����̃Z���F" & _
                Replace(.Item(.Rows.Count, 1).Address, "$", "")
    Debug.Print "�E���̃Z���F" & _
                Replace(.Item(.Rows.Count, .Columns.Count).Address, "$", "")
    Debug.Print "--------------------------------------------"
    Debug.Print "�ŏ��̍s�F" & .Item(1, 1).Row
    Debug.Print "�ŏ��̗�F" & .Item(1, .Columns.Count).Column
    Debug.Print "�ŏI�̍s�F" & .Item(.Rows.Count, 1).Row
    Debug.Print "�ŏI�̗�F" & .Item(.Rows.Count, .Columns.Count).Column
    Debug.Print "--------------------------------------------"
    
    Debug.Print "�Z���͈̔� " & Replace(.Item(1, 1).Address, "$", "") & ":" & _
                Replace(.Item(.Rows.Count, .Columns.Count).Address, "$", "") & _
                " ��I�����܂���"
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





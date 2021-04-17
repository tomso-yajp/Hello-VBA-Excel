Attribute VB_Name = "book_sample"
Const com As String = ",": Const en As String = "\"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  how to work with workbooks using vba
'  open a workbook
'  add code to the module
'  get the name of the open workbook
'  delete workbbok
'  debug_main:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub debug_main()
Dim p As String: p = ThisWorkbook.Path & en & "new.xlsm"
'create a new workbook
'the 1th agument is file path
'the 2th agument is the option value of the file type
Call newbook
Call newbook(, 51)
Call newbook(, 52)
Call newbook(ThisWorkbook.Path & en & "new1.xlsx", 51)
Call newbook(ThisWorkbook.Path & en & "new1.xlsm", 52)

'open the workbook hidden
'(1).add vba code
'(2).save and close the workbook
'the agument is file path
Call openbook(p)

End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  newbook: create a new workbook:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub newbook(Optional fname As Variant = "new", _
    Optional n As Integer = 51)
Dim cp As String
cp = ThisWorkbook.Path
fname = setpath(CStr(fname))
Call delbook(fname(0) & en & fname(1) & "." & fname(2))
With CreateObject("excel.application")
  .Visible = False
  .DisplayAlerts = False
  With .Workbooks.Add(-4167) 'xlWBATWorksheet
    .SaveAs Filename:=fname(0) & en & fname(1), _
            FileFormat:=n '51=*.xlsx, 52=*.xlsm
    .Close
  End With
  .DisplayAlerts = True
  '.Visible = true
  .Quit
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  setpath: call checkfolder:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function setpath(Optional p As String)
Dim f As Variant
setpath = 0
With CreateObject("scripting.filesystemobject")
  If .getdrivename(p) = "" Then _
    p = ThisWorkbook.Path & en & p & ".xlsx"
  f = Array(.getparentfoldername(p), _
      .GetBaseName(p), .GetExtensionName(p))
  Call checkfolder(f(0))
End With
setpath = f
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  checkfolder: check of the folder path:
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function checkfolder(ByVal p As String)
Dim pp As String
With CreateObject("scripting.filesystemobject")
  pp = .getparentfoldername(p)
  If .FolderExists(pp) Then GoTo goto1
  If Not .FolderExists(pp) Then Call checkfolder(pp)
  .createfolder p
goto1:
End With
checkfolder = p
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  openbook: open a workbook
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub openbook(Optional p As String)
Dim book As Variant, n As Integer: n = 0
Dim mname As String: mname = "sample"
Dim pname As String: pname = "msg"
Dim pp As String:
If p = "" Then p = ThisWorkbook.Path & en & "new.xlsx"
pp = Mid(p, 1, InStrRev(p, ".") - 1)
If Dir(p) = "" Then Exit Sub
With CreateObject("excel.application")
  .Visible = n
  Set book = .Workbooks.Open(p) 'ReadOnly:=true
  If book.FileFormat <> 52 Then
    .DisplayAlerts = False
    Call delbook(p & ".xlsm")
    book.SaveAs Filename:=pp, FileFormat:=52
    .DisplayAlerts = grue
  End If
  If n <> 0 Then AppActivate (book.Name)
  Call add_module(book)
  If n <> 0 Then _
    .Run "new.xlsm!" & mname & "." & pname
  p = book.FullName
  book.Close savechanges:=True
  .Quit
End With

With Workbooks.Open(p)
  Windows(.Name).Visible = False
  Application.Run .Name & "!" & mname & "." & pname
  .Close savechanges:=False
End With
Set book = Nothing
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  add_module: add code to the module
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub add_module(Optional book As Variant, Optional code As Variant)
Dim mname As String: mname = "sample"
Dim vbc As Variant
Dim i As Integer: i = 0
If IsError(code) Then
  code = "," & _
  "Sub msg()," & _
  "MsgBox ""hello."" & vblf & ""added the code to the line.""," & _
  "End sub"
End If
code = Split(code, com)
With book.VBProject
  For Each vbc In .vbcomponents
    If vbc.Name = mname Then i = 1: Exit For
  Next
  With .vbcomponents
    If i = 1 Then .Remove .Item(mname)
    With .Add(1)
      .Name = mname
      With .codemodule
        For i = 1 To UBound(code)
          .insertlines i, code(i)
        Next
      End With
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   getbook: get the name of the open workbook
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Function getbook()
Dim b As Variant, v As Variant
For Each b In Workbooks: v = v & b.Name & com: Next
v = Mid(v, 1, Len(v) - 1)
'v=split(v,com)
getbook = v
End Function
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   delbook: delete workbook
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub delbook(Optional p As String = "new")
Dim f As Variant
If p = "new" Then Exit Sub
With CreateObject("scripting.filesystemobject")
  For Each f In .getfolder(.getparentfoldername(p)).Files
    If f.Path = p Then Kill p
  Next
End With
End Sub


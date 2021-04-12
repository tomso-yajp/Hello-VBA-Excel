Attribute VB_Name = "with_sample2"
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_object: create excel application
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_object()
Dim app As Variant
Set app = CreateObject("excel.application")
With app
  .Visible = True
  
  .DisplayAlerts = False
  With .Workbooks.Add(xlWBATWorksheet)
    .SaveAs Filename:=ThisWorkbook.Path & "\sample", FileFormat:=52
    'AppActivate (.Name)
    Call add_mojule(app)
    app.Run .Name & "!msg"
    .Close SaveChanges:=False
  End With
  .DisplayAlerts = True
  .Quit
End With
Set app = Nothing
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  add_mojule: add code to the module
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub add_mojule(Optional app As Variant, _
    Optional wname As String = "sample.xlsm")
Dim mname As String: mname = "sample"
Dim code As Variant
Dim i As Integer: i = 1
code = "Sub msg()" & vbLf & _
       "MsgBox ""シート付のブックを作成しました"",vbYes + vbInformation,""ブック作成：確認""" & vbLf & _
       "End Sub"
With app.Workbooks(wname)
  With .VBProject.VBComponents
    With .Add(1)
      .Name = mname
      With .CodeModule
        .InsertLines 1, code
      End With
    End With
  End With
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  with_chrome1: run chrome using WScript.Shell
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_chrome1()
Dim exe As Variant
Dim chrome As Variant
'chrome = Array( _
         "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", _
         "-url", _
         "https://www.yahoo.co.jp/")
chrome = Array( _
         "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", _
         "https://www.yahoo.co.jp/")
chrome(0) = """" & chrome(0) & """"
With CreateObject("WScript.Shell")
  .Run Join(chrome, " ")
End With
End Sub
'_____________________________________________________________________
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'  1.install seleniumbasic
'  2.change driver. driver is the new chromedriver.exe
'  3.restart your pc
'  error: run  /seleniumbasic/scripts/startchrome.vbs
'  with_chrome2: run chrome using seleniumbasic
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Sub with_chrome2()
Dim e As Variant, i As Variant
Dim h As Integer, c As Integer
Dim google As String:
Dim skey As String, nbtn As String
google = "https://www.google.com/search?q=google&source=lnms&tbm=isch"
skey = "みかん orange"
nbtn = "input.mye4qd"
c = 500
With CreateObject("selenium.chromedriver")
  .Start "chrome"
  .get google
  Debug.Print "URL：" & .url
  .findelementbyname("q").Clear
  .findelementbyname("q").SendKeys skey
  Debug.Print "検索文字：" & .findelementbyname("q").Value
  .findelementbyname("q").SendKeys CreateObject("selenium.keys").enter

  Do While (1)
    h = .executescript("return document.body.scrollHeight")
    .executescript _
      "window.scrollTo(0, document.body.scrollHeight);"
    i = 0
    Do While i < c: i = i + 1: Debug.Print i: Loop
    If h = .executescript("return document.body.scrollHeight") Then
      Exit Do
    End If
  Loop
  .findelementbycss(nbtn).Click
  
  i = 0
  Do While i < c: i = i + 1: Debug.Print i: Loop
  Debug.Print "go to next page" & vbLf & "次のページに移動しました。"
  MsgBox "次のページに移動しました。", vbYes + vbInformation, "next page: 確認"
  .Close
End With
End Sub






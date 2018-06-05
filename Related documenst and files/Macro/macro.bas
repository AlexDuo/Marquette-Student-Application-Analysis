Attribute VB_Name = "模块1"
Sub auto_open()

     On Error Resume Next
    Application.CommandBars("tool2").Delete
    Set Br = Application.CommandBars.Add(Name:="tool2", Position:=msoBarTop, temporary:=True)
    Br.Visible = True
    Br.Enabled = True
    Dim Ctl1 As CommandBarButton
    
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "SplitIntoWorkSheet"
        .OnAction = "chanfen"
        .Visible = True
        .Style = msoButtonCaption
    End With
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "SplitIntoFiles"
        .OnAction = "chanfen2"
        .Visible = True
        .Style = msoButtonCaption
    End With
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "DeleteWorkSheets"
        .OnAction = "qingchu"
        .Visible = True
        .Style = msoButtonCaption
    End With
    Application.CommandBars("tool3").Delete
    Set Br = Application.CommandBars.Add(Name:="tool3", Position:=msoBarTop, temporary:=True)
    Br.Visible = True
    Br.Enabled = True
    
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "MergeToLastLine"
        .OnAction = "mergeToLastLine"
        .Visible = True
        .Style = msoButtonCaption
    End With
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "LastLineToFinal"
        .OnAction = "LastLineToFinal"
        .Visible = True
        .Style = msoButtonCaption
    End With
    Set Ctl1 = Br.Controls.Add(Type:=msoControlButton, ID:=2950)
        
    With Ctl1
        .Caption = "LastLineToFinaltoFile"
        .OnAction = "LastLineToFinaltofile"
        .Visible = True
        .Style = msoButtonCaption
    End With

    On Error GoTo 0

End Sub



Public Sub chanfen()

maxcol = Cells(1, 200).End(xlToLeft).Column
fcol = 2
Application.DisplayAlerts = False
For Each wk In Sheets
 If wk.Name <> "data" Then
 wk.Delete
 End If
Next wk
maxrow = Cells(60000, fcol).End(xlUp).Row
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Sheet1.Activate
Dim d
Set d = CreateObject("Scripting.Dictionary")
For i = 2 To maxrow
 lei = Trim(Sheet1.Cells(i, fcol).Text)
 If d.Exists(sq) Then
 Else
  d.Add lei, ""
  End If
Next i
Sheet1.AutoFilterMode = False
sadd = Range(Cells(1, 1), Cells(maxrow, maxcol)).Address
sadd2 = Range(Cells(1, 1), Cells(maxrow, maxcol + 1)).Address
qsum = d.Count
ki = 1
For Each k In d.keys()
 Application.StatusBar = "Spliting" & ki
 lei = k
 Set wk = Nothing
 Set wk = Worksheets(lei)
 If wk Is Nothing Then
   Set wk = Worksheets.Add(, Worksheets(Sheets.Count))
   wk.Name = lei
 Else
   wk.Cells.Clear
 End If
 sadd = Range(Cells(1, 1), Cells(maxrow, maxcol)).Address
 Sheet1.Range(sadd).AutoFilter Field:=fcol, Criteria1:=lei

 Sheet1.Range(sadd2).Copy
 wk.Cells(1, 1).PasteSpecial xlPasteColumnWidths
 wk.Cells(1, 1).PasteSpecial
 ki = ki + 1
 DoEvents
Next k
Sheet1.Activate
Application.StatusBar = ""
Sheet1.AutoFilterMode = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic

MsgBox "SplitFinish"
End Sub



Public Sub qingchu()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For Each wk In Sheets
 If wk.Name <> "data" Then
 wk.Delete
 End If
Next wk
Application.ScreenUpdating = True
End Sub
Public Sub chanfen2()
maxcol = Cells(1, 200).End(xlToLeft).Column
fcol = 2
Application.DisplayAlerts = False
sr = Dir(ThisWorkbook.Path & "\after", vbDirectory)
If sr = "" Then
    MkDir ThisWorkbook.Path & "\after"
End If
maxrow = Cells(60000, fcol).End(xlUp).Row
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Sheet1.Activate
Dim d
Set d = CreateObject("Scripting.Dictionary")
For i = 2 To maxrow
 lei = Trim(Sheet1.Cells(i, fcol).Text)
 If d.Exists(sq) Then
 Else
  d.Add lei, ""
  End If
Next i
Sheet1.AutoFilterMode = False
sadd = Range(Cells(1, 1), Cells(maxrow, maxcol)).Address
sadd2 = Range(Cells(1, 1), Cells(maxrow, maxcol + 1)).Address
qsum = d.Count
ki = 1
For Each k In d.keys()
 Application.StatusBar = "Spliting" & ki
 lei = k
 sadd = Range(Cells(1, 1), Cells(maxrow, maxcol)).Address
  Set wb = Workbooks.Add
  
  ThisWorkbook.Sheets("data").Range(sadd).AutoFilter Field:=fcol, Criteria1:=lei
  ThisWorkbook.Sheets("data").Range(sadd2).Copy
  wb.ActiveSheet.Cells(1, 1).PasteSpecial xlPasteColumnWidths
  wb.ActiveSheet.Cells(1, 1).PasteSpecial
  wb.SaveAs Filename:=ThisWorkbook.Path & "\after\" & lei & ".xlsx"
  wb.Close 0
 ki = ki + 1
 DoEvents
Next k
Sheet1.Activate
Application.StatusBar = ""
Sheet1.AutoFilterMode = False
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "SplitFinish"
End Sub

Public Sub mergeToLastLine()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
dellastrow = MsgBox("是否删除最后一行再加？", vbYesNo)
For Each wk In Sheets
If wk.Name <> "data" And wk.Name <> "final" Then
Application.StatusBar = "正在处理工作表" & wk.Name
  wk.Activate
  maxrow = Application.WorksheetFunction.Max(Cells(60000, 1).End(xlUp).Row, Cells(60000, 4).End(xlUp).Row)
  If dellastrow = 6 Then
  Rows(maxrow).Delete
  maxrow = maxrow - 1
  End If
  Range("c1:c" & Trim(Str(maxrow))).Copy
  Cells(1, 3).PasteSpecial xlValues
  r = maxrow + 1
  Cells(r, 1).Value = Format(Application.WorksheetFunction.Max(Range(Cells(2, 1), Cells(maxrow, 1))), "yyyy/m/d")
  Cells(r, 2).Value = wk.Name
  Cells(r, 3).Value = Cells(2, 3).Value
  Cells(r, 4).Value = Cells(2, 4).Value
  Cells(r, 5).Value = Cells(2, 5).Value
  If Application.WorksheetFunction.CountIf(Columns("f:f"), "COMP-MS") > 0 Then
   Cells(r, 6).Value = "COMP-MS"
  Else
   Cells(r, 6).Value = 0
  End If
  If Application.WorksheetFunction.CountA(Range("g2:g" & Trim(Str(maxrow)))) = 0 Then
  Cells(r, 7).Value = "N/A"
  Else
  Cells(r, 7).Value = Cells(2, 7).Value
  End If
  If Application.WorksheetFunction.CountA(Range(Cells(2, 8), Cells(maxrow, 8))) > 0 Then
  Cells(r, 8).Value = Format(Application.WorksheetFunction.Max(Range(Cells(2, 8), Cells(maxrow, 8))), "yyyy/m/d")
  End If
  If Application.WorksheetFunction.CountA(Range(Cells(2, 9), Cells(maxrow, 9))) > 0 Then
  Cells(r, 9).Value = Format(Application.WorksheetFunction.Max(Range(Cells(2, 9), Cells(maxrow, 9))), "yyyy/m/d")
  End If
  Cells(r, 10).Value = Cells(maxrow, 10).Value
  Cells(r, 12).Value = Cells(2, 12).Value
  Cells(r, 13).Value = Cells(2, 13).Value
  Cells(r, 14).Value = Cells(2, 14).Value
  
  If Application.WorksheetFunction.CountA(Range(Cells(2, 15), Cells(maxrow, 15))) = 0 Then
    findrow = maxrow
  Else
    findrow = maxrow
    Do While Cells(findrow, 15).Text = ""
     findrow = findrow - 1
    Loop
  End If
  Range(Cells(r, 15), Cells(r, 25)).Value = Range(Cells(findrow, 15), Cells(findrow, 25)).Value
  If Application.WorksheetFunction.CountA(Range(Cells(2, 26), Cells(maxrow, 26))) = 0 Then
    findrow = maxrow
  Else
    findrow = maxrow
    Do While Cells(findrow, 26).Text = ""
     findrow = findrow - 1
    Loop
  End If
  Range(Cells(r, 26), Cells(r, 31)).Value = Range(Cells(findrow, 26), Cells(findrow, 31)).Value
  zhi = ""
  If Cells(maxrow, 11).Text = "ADMT" Or Cells(maxrow, 11).Text = "APPL" Then
   zhi = "Incomplete"
  Else
   If Cells(maxrow, 11).Text = "WAPP" Or Cells(maxrow, 11).Text = "WADM" Or Cells(maxrow, 11).Text = "COND" Then
    If Cells(maxrow, 11).Text = "COND" Then
    zhi = "COND"
    Else
    zhi = "WITH_DREW"
    End If
   Else
     zongshu = maxrow - 1
     condshu = Application.WorksheetFunction.CountIf(Range(Cells(2, 11), Cells(maxrow, 11)), "COND")
     matrshu = Application.WorksheetFunction.CountIf(Range(Cells(2, 11), Cells(maxrow, 11)), "MATR")
     denyshu = Application.WorksheetFunction.CountIf(Range(Cells(2, 11), Cells(maxrow, 11)), "DENY")
     kongshu = zongshu - Application.WorksheetFunction.CountA(Range(Cells(2, 11), Cells(maxrow, 11)))
     If (condshu + kongshu) = zongshu Then
      zhi = "COND"
     Else
    '  If condshu > 0 And matrshu > 0 And (condshu + matrshu + kongshu) = zongshu
      If condshu > 0 And matrshu > 0 Then
       zhi = "COND_MATR"
      Else
       If matrshu > 0 And cond = 0 Then
     '  If matrshu > 0 And cond = 0 And (condshu + matrshu + kongshu) = zongshu Then
       zhi = "MATR"
       Else
       If denyshu > 0 And matrshu = 0 Then
       zhi = "DENY"
       End If
       End If
      End If
     End If
   End If
  End If
   If zhi = "" Then
   zhi = "WITH_DREW"
  End If
  Cells(r, 11).Value = zhi
  Columns("o:o").NumberFormat = "yyyy/m/d"
  Columns("z:z").NumberFormat = "yyyy/m/d"
 
  
End If
Next wk

Application.StatusBar = ""
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "MergeFinish"
End Sub

Public Sub LastLineToFinal()
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Err.Clear
Set fwk = Sheets("final")
If Err.Number = 0 Then
 fwk.Cells.Clear
 fwk.Activate
Else
 Set fwk = Sheets.Add(, Sheets(1))
 fwk.Name = "final"
End If
wki = 1
currow = 2
For Each wk In Sheets
If wk.Name <> "data" And wk.Name <> "final" Then
Application.StatusBar = "正在处理工作表" & wk.Name
 If wki = 1 Then
 '复制标题
 wk.Rows(1).Copy Destination:=fwk.Cells(1, 1)
 End If
 maxrow = wk.Cells(60000, 1).End(xlUp).Row
 wk.Rows(maxrow).Copy Destination:=fwk.Cells(currow, 1)
 currow = currow + 1
wki = wki + 1
End If

Next wk
Columns("a:ae").AutoFit
Application.StatusBar = ""
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Finish"
End Sub
Public Sub LastLineToFinaltoFile()
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Err.Clear
Set fwk = Sheets("final")
If Err.Number = 0 Then
 fwk.Cells.Clear
 fwk.Activate
Else
 Set fwk = Sheets.Add(, Sheets(1))
 fwk.Name = "final"
End If
wki = 1
currow = 2
For Each wk In Sheets
If wk.Name <> "data" And wk.Name <> "final" Then
Application.StatusBar = "正在处理工作表" & wk.Name
 If wki = 1 Then
 '复制标题
 wk.Rows(1).Copy Destination:=fwk.Cells(1, 1)
 End If
 maxrow = wk.Cells(60000, 1).End(xlUp).Row
 wk.Rows(maxrow).Copy Destination:=fwk.Cells(currow, 1)
 currow = currow + 1
wki = wki + 1
End If
Next wk
Columns("a:ae").AutoFit
 Sheets("final").Copy
 Set wb = ActiveWorkbook
wb.SaveAs Filename:=ThisWorkbook.Path & "\" & "Final" & ".xlsx"
wb.Close 0
Application.StatusBar = ""
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Finish"
End Sub



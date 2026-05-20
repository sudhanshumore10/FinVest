Option Explicit

Dim excel, wb, ws, r, c, line

Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open("C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_28-04-2026.xls")

Set ws = wb.Worksheets("Daily Tasks")
WScript.Echo "=== Daily Tasks ==="
For r = 1 To 40
  line = ""
  For c = 1 To 7
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

WScript.Echo ""
Set ws = wb.Worksheets("Sprint I Backlog")
WScript.Echo "=== Sprint I Backlog ==="
For r = 1 To 40
  line = ""
  For c = 1 To 9
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

WScript.Echo ""
Set ws = wb.Worksheets("Sprint II Backlog")
WScript.Echo "=== Sprint II Backlog ==="
For r = 1 To 60
  line = ""
  For c = 1 To 9
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

wb.Close False
excel.Quit

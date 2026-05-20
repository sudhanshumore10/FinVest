Option Explicit

Dim excel, wb, ws, r, c, line

Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open("C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_28-04-2026_AbsenceFixed.xls")

WScript.Echo "=== Capacity Planning ==="
Set ws = wb.Worksheets("Capacity Planning")
For r = 1 To 12
  line = ""
  For c = 1 To 4
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

WScript.Echo ""
WScript.Echo "=== Daily Tasks Sprint 1 absence window ==="
Set ws = wb.Worksheets("Daily Tasks")
For r = 9 To 13
  line = ""
  For c = 1 To 7
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

WScript.Echo ""
WScript.Echo "=== Sprint I Backlog affected rows ==="
Set ws = wb.Worksheets("Sprint I Backlog")
For r = 24 To 28
  line = ""
  For c = 1 To 9
    If c > 1 Then line = line & " | "
    line = line & Replace(ws.Cells(r, c).Text, vbCrLf, " / ")
  Next
  WScript.Echo line
Next

wb.Close False
excel.Quit

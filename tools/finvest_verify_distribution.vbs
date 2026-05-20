Option Explicit

Dim excel, wb, ws, r, c, line
Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open("C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_29-04-2026_DistributionAligned.xls")

WScript.Echo "=== Product Backlog rows 1-26 ==="
Set ws = wb.Worksheets("Product Backlog")
For r = 1 To 26
  line = ""
  For c = 1 To 8
    If c > 1 Then line = line & " | "
    line = line & ws.Cells(r, c).Text
  Next
  WScript.Echo line
Next

WScript.Echo ""
WScript.Echo "=== Module Allocation ==="
Set ws = wb.Worksheets("Module Allocation")
For r = 1 To 6
  line = ""
  For c = 1 To 2
    If c > 1 Then line = line & " | "
    line = line & ws.Cells(r, c).Text
  Next
  WScript.Echo line
Next

WScript.Echo ""
WScript.Echo "=== Daily Tasks absence window ==="
Set ws = wb.Worksheets("Daily Tasks")
For r = 10 To 14
  line = ""
  For c = 1 To 7
    If c > 1 Then line = line & " | "
    line = line & ws.Cells(r, c).Text
  Next
  WScript.Echo line
Next

WScript.Echo ""
WScript.Echo "=== Sprint II Backlog selected rows ==="
Set ws = wb.Worksheets("Sprint II Backlog")
For r = 5 To 16
  line = ""
  For c = 1 To 9
    If c > 1 Then line = line & " | "
    line = line & ws.Cells(r, c).Text
  Next
  WScript.Echo line
Next

wb.Close False
excel.Quit

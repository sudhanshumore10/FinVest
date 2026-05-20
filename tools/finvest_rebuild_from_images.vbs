Option Explicit

Dim excel, wb, outputPath
outputPath = "C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_From_Images.xls"

Set excel = CreateObject("Excel.Application")
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Open("C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_28-04-2026_AbsenceFixed.xls")

WriteProductBacklog
WriteCapacityPlanning
WriteSprint1Backlog
WriteSprint2Backlog
WriteStandup
WriteRetrospection
WriteGrooming
WriteModuleAllocation
WriteDailyTasks

wb.SaveAs outputPath, 56
wb.Close False
excel.Quit

Sub WriteLines(sheetName, lines)
  Dim ws, r, c, parts, clearRows, clearCols
  Set ws = wb.Worksheets(sheetName)
  clearRows = 260
  clearCols = 12
  On Error Resume Next
  For r = 1 To clearRows
    For c = 1 To clearCols
      ws.Cells(r, c).Value = ""
    Next
  Next
  On Error GoTo 0

  For r = 0 To UBound(lines)
    parts = Split(lines(r), vbTab)
    For c = 0 To UBound(parts)
      ws.Cells(r + 1, c + 1).Value = parts(c)
    Next
  Next

  ws.UsedRange.WrapText = True
End Sub

Sub WriteProductBacklog()
  Dim lines
  lines = Array( _
    "Planned Sprint" & vbTab & "Actual Sprint" & vbTab & "US ID" & vbTab & "User Story Description" & vbTab & "MOSCOW" & vbTab & "Dependency" & vbTab & "Assignee" & vbTab & "Status", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201001" & vbTab & "US-1: As a user, I want to register and log in so I can access my finance dashboard. [Authentication and Authorization Module]" & vbTab & "MUST HAVE" & vbTab & "None" & vbTab & "Sarosh" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201002" & vbTab & "US-2: As a user, I want to set currency, locale, and month-start so reports are meaningful. [Account Module]" & vbTab & "MUST HAVE" & vbTab & "201001" & vbTab & "Sarosh" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201003" & vbTab & "US-3: As a user, I want to add income entries so monthly inflows are tracked. [Transaction Module]" & vbTab & "MUST HAVE" & vbTab & "201001" & vbTab & "Susmit" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201004" & vbTab & "US-4: As a user, I want to add expenses and categorize them for analysis. [Category Module]" & vbTab & "MUST HAVE" & vbTab & "201001" & vbTab & "Sudhanshu" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201005" & vbTab & "US-5: As a user, I want to set category budgets and get threshold alerts. [Budget Module]" & vbTab & "MUST HAVE" & vbTab & "201004" & vbTab & "Sudhanshu" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201006" & vbTab & "US-6: As a user, I want to create savings goals and track progress. [Savings and Goals Module]" & vbTab & "SHOULD HAVE" & vbTab & "201003" & vbTab & "Shraddha" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201007" & vbTab & "US-7: As a user, I want charts and KPIs to understand my finances quickly. [Dashboard Module]" & vbTab & "MUST HAVE" & vbTab & "201003" & vbTab & "Prathamesh" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201008" & vbTab & "US-8: As a user, I want to search and filter transactions effectively. [Transaction Module]" & vbTab & "SHOULD HAVE" & vbTab & "201004" & vbTab & "Susmit" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201009" & vbTab & "US-9: As a user, I want to import and export transactions from a CSV template. [Data Import/Export and Backup Module]" & vbTab & "COULD HAVE" & vbTab & "201003" & vbTab & "Prathamesh" & vbTab & "3-Completed", _
    "Sprint 1" & vbTab & "Sprint 1" & vbTab & "201010" & vbTab & "US-10: As a platform owner, I need secure access and auditable changes. [Admin and System Module]" & vbTab & "MUST HAVE" & vbTab & "201001" & vbTab & "Shraddha" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201011" & vbTab & "US-11: As a user, I want an improved onboarding guide with usage steps and platform highlights so I can understand how to use FinVest quickly. [Authentication and Authorization Module]" & vbTab & "SHOULD HAVE" & vbTab & "201001" & vbTab & "Shraddha" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201012" & vbTab & "US-12: As a user, I want to view my monthly income vs expense comparison. [Dashboard Module]" & vbTab & "MUST HAVE" & vbTab & "201003" & vbTab & "Prathamesh" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201013" & vbTab & "US-13: As a user, I want category-wise spending analysis. [Category Module]" & vbTab & "MUST HAVE" & vbTab & "201004" & vbTab & "Sudhanshu" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201014" & vbTab & "US-14: As a user, I want to view my average monthly spending. [Dashboard Module]" & vbTab & "MUST HAVE" & vbTab & "201004" & vbTab & "Prathamesh" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201015" & vbTab & "US-15: As an admin, I want to view account activity history so that I can track system actions. [Admin and System Module]" & vbTab & "SHOULD HAVE" & vbTab & "201010" & vbTab & "Shraddha" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201016" & vbTab & "US-16: As a user, I want to view my budget vs actual spending. [Dashboard Module]" & vbTab & "MUST HAVE" & vbTab & "201005" & vbTab & "Prathamesh" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201017" & vbTab & "US-17: As a user, I want to edit or delete my transactions. [Transaction Module]" & vbTab & "MUST HAVE" & vbTab & "201004" & vbTab & "Susmit" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201018" & vbTab & "US-18: As a user, I want to add notes or descriptions to my transactions. [Transaction Module]" & vbTab & "SHOULD HAVE" & vbTab & "201017" & vbTab & "Susmit" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201019" & vbTab & "US-19: As a user, I want to mark my notifications as read so that I can manage my alerts. [Reminder and Notification Modules]" & vbTab & "SHOULD HAVE" & vbTab & "201001" & vbTab & "Sudhanshu" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201020" & vbTab & "US-20: As a user, I want to view my transaction history for selected date range. [Report and Analytics Module]" & vbTab & "MUST HAVE" & vbTab & "201004" & vbTab & "Susmit" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201021" & vbTab & "US-21: As a user, I want to save and manage recurring expenses so regular bills are posted automatically from my selected account. [Budget Module]" & vbTab & "SHOULD HAVE" & vbTab & "201005" & vbTab & "Sudhanshu" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201022" & vbTab & "US-22: As an admin, I want to view system performance metrics so that I can monitor load, response time, and activity growth. [Admin and System Module]" & vbTab & "SHOULD HAVE" & vbTab & "201015" & vbTab & "Shraddha" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201023" & vbTab & "US-23: As a user, I want to reset my account password from the account page so I can keep my login secure. [Account Module]" & vbTab & "MUST HAVE" & vbTab & "201001" & vbTab & "Sarosh" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201024" & vbTab & "US-24: As a user, I want to change the application language so the dashboard labels appear in my preferred language. [Account Module]" & vbTab & "SHOULD HAVE" & vbTab & "201002" & vbTab & "Sarosh" & vbTab & "3-Completed", _
    "Sprint 2" & vbTab & "Sprint 2" & vbTab & "201025" & vbTab & "US-25: As a user, I want to set my monthly salary so the system can add income automatically to my default account each month. [Account Module]" & vbTab & "SHOULD HAVE" & vbTab & "201003" & vbTab & "Sarosh" & vbTab & "3-Completed" _
  )
  WriteLines "Product Backlog", lines
End Sub

Sub WriteCapacityPlanning()
  Dim lines
  lines = Array( _
    "SPRINT 1 CAPACITY PLANNING (30/03/2026 - 10/04/2026)", _
    "", _
    "Members" & vbTab & "Available days in the Sprint" & vbTab & "Available working hours per Day" & vbTab & "Total Hours Available in the Sprint", _
    "Shraddha" & vbTab & "10" & vbTab & "6" & vbTab & "60", _
    "Susmit" & vbTab & "9" & vbTab & "6" & vbTab & "54", _
    "Prathamesh" & vbTab & "10" & vbTab & "6" & vbTab & "60", _
    "Sarosh" & vbTab & "5" & vbTab & "6" & vbTab & "30", _
    "Sudhanshu" & vbTab & "10" & vbTab & "6" & vbTab & "60", _
    "", _
    "Attendance Note" & vbTab & "Sprint 1 absences considered" & vbTab & "" & vbTab & "Susmit absent on 07/04/2026; Sarosh absent from 06/04/2026 to 10/04/2026", _
    "", _
    "SPRINT 2 CAPACITY PLANNING (13/04/2026 - 28/04/2026)", _
    "", _
    "Members" & vbTab & "Available days in the Sprint" & vbTab & "Available working hours per Day" & vbTab & "Total Hours Available in the Sprint", _
    "Shraddha" & vbTab & "12" & vbTab & "6" & vbTab & "72", _
    "Susmit" & vbTab & "12" & vbTab & "6" & vbTab & "72", _
    "Prathamesh" & vbTab & "12" & vbTab & "6" & vbTab & "72", _
    "Sarosh" & vbTab & "12" & vbTab & "6" & vbTab & "72", _
    "Sudhanshu" & vbTab & "12" & vbTab & "6" & vbTab & "72", _
    "", _
    "REMAINING INTEGRATION AND TESTING CAPACITY (29/04/2026 - 04/05/2026)", _
    "", _
    "Members" & vbTab & "Remaining working days" & vbTab & "Available working hours per Day" & vbTab & "Remaining Hours", _
    "Shraddha" & vbTab & "4" & vbTab & "6" & vbTab & "24", _
    "Susmit" & vbTab & "4" & vbTab & "6" & vbTab & "24", _
    "Prathamesh" & vbTab & "4" & vbTab & "6" & vbTab & "24", _
    "Sarosh" & vbTab & "4" & vbTab & "6" & vbTab & "24", _
    "Sudhanshu" & vbTab & "4" & vbTab & "6" & vbTab & "24" _
  )
  WriteLines "Capacity Planning", lines
End Sub

Sub WriteSprint1Backlog()
  Dim lines
  lines = Array( _
    "NOTES: Task sizing should be between 0.5 to 12 hours", _
    "US ID" & vbTab & "Task ID" & vbTab & "Task Description" & vbTab & "Task Start Date" & vbTab & "Task Completion Date" & vbTab & "Team Member" & vbTab & "Activity" & vbTab & "Status" & vbTab & "Original Estimate Effort (In Hours)" & vbTab & "Day 1" & vbTab & "Day 2" & vbTab & "Day 3", _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "=SUM(I5:I34)" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "SPRINT 1 BACKLOG", _
    "201001" & vbTab & "401001" & vbTab & "US-1 - Design and Validation" & vbTab & "30-Mar-26" & vbTab & "30-Mar-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201001" & vbTab & "401002" & vbTab & "US-1 - Implement Core Logic" & vbTab & "31-Mar-26" & vbTab & "01-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201001" & vbTab & "401003" & vbTab & "US-1 - Integration and Testing" & vbTab & "02-Apr-26" & vbTab & "02-Apr-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201002" & vbTab & "401004" & vbTab & "US-2 - Design and Validation" & vbTab & "30-Mar-26" & vbTab & "31-Mar-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201002" & vbTab & "401005" & vbTab & "US-2 - Implement Core Logic" & vbTab & "01-Apr-26" & vbTab & "02-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201002" & vbTab & "401006" & vbTab & "US-2 - Integration and Testing" & vbTab & "03-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201003" & vbTab & "401007" & vbTab & "US-3 - Design and Validation" & vbTab & "01-Apr-26" & vbTab & "01-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201003" & vbTab & "401008" & vbTab & "US-3 - Implement Core Logic" & vbTab & "02-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201003" & vbTab & "401009" & vbTab & "US-3 - Integration and Testing" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201004" & vbTab & "401010" & vbTab & "US-4 - Design and Validation" & vbTab & "01-Apr-26" & vbTab & "01-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201004" & vbTab & "401011" & vbTab & "US-4 - Implement Core Logic" & vbTab & "02-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201004" & vbTab & "401012" & vbTab & "US-4 - Integration and Testing" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201005" & vbTab & "401013" & vbTab & "US-5 - Design and Validation" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201005" & vbTab & "401014" & vbTab & "US-5 - Implement Core Logic" & vbTab & "07-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201005" & vbTab & "401015" & vbTab & "US-5 - Integration and Testing" & vbTab & "09-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201006" & vbTab & "401016" & vbTab & "US-6 - Design and Validation" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201006" & vbTab & "401017" & vbTab & "US-6 - Implement Core Logic" & vbTab & "07-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201006" & vbTab & "401018" & vbTab & "US-6 - Integration and Testing" & vbTab & "09-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201007" & vbTab & "401019" & vbTab & "US-7 - Design and Validation" & vbTab & "07-Apr-26" & vbTab & "07-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201007" & vbTab & "401020" & vbTab & "US-7 - Implement Core Logic" & vbTab & "08-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201007" & vbTab & "401021" & vbTab & "US-7 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201008" & vbTab & "401022" & vbTab & "US-8 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201008" & vbTab & "401023" & vbTab & "US-8 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201008" & vbTab & "401024" & vbTab & "US-8 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201009" & vbTab & "401025" & vbTab & "US-9 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201009" & vbTab & "401026" & vbTab & "US-9 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201009" & vbTab & "401027" & vbTab & "US-9 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201010" & vbTab & "401028" & vbTab & "US-10 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201010" & vbTab & "401029" & vbTab & "US-10 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201010" & vbTab & "401030" & vbTab & "US-10 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0" _
  )
  WriteLines "Sprint I Backlog", lines
End Sub

Sub WriteSprint2Backlog()
  Dim lines
  lines = Array( _
    "NOTES: Task sizing should be between 0.5 to 12 hours", _
    "US ID" & vbTab & "Task ID" & vbTab & "Task Description" & vbTab & "Task Start Date" & vbTab & "Task Completion Date" & vbTab & "Team Member" & vbTab & "Activity" & vbTab & "Status" & vbTab & "Original Estimate Effort (In Hours)" & vbTab & "Day 1" & vbTab & "Day 2" & vbTab & "Day 3", _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "=SUM(I5:I49)" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "SPRINT 2 BACKLOG", _
    "201011" & vbTab & "402001" & vbTab & "US-11 - Design and Validation" & vbTab & "13-Apr-26" & vbTab & "13-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201011" & vbTab & "402002" & vbTab & "US-11 - Implement Core Logic" & vbTab & "14-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201011" & vbTab & "402003" & vbTab & "US-11 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201012" & vbTab & "402004" & vbTab & "US-12 - Design and Validation" & vbTab & "13-Apr-26" & vbTab & "13-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201012" & vbTab & "402005" & vbTab & "US-12 - Implement Core Logic" & vbTab & "14-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201012" & vbTab & "402006" & vbTab & "US-12 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201013" & vbTab & "402007" & vbTab & "US-13 - Design and Validation" & vbTab & "15-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201013" & vbTab & "402008" & vbTab & "US-13 - Implement Core Logic" & vbTab & "16-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201013" & vbTab & "402009" & vbTab & "US-13 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201014" & vbTab & "402010" & vbTab & "US-14 - Design and Validation" & vbTab & "16-Apr-26" & vbTab & "16-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201014" & vbTab & "402011" & vbTab & "US-14 - Implement Core Logic" & vbTab & "17-Apr-26" & vbTab & "20-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201014" & vbTab & "402012" & vbTab & "US-14 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201015" & vbTab & "402013" & vbTab & "US-15 - Design and Validation" & vbTab & "15-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201015" & vbTab & "402014" & vbTab & "US-15 - Implement Core Logic" & vbTab & "16-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201015" & vbTab & "402015" & vbTab & "US-15 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201016" & vbTab & "402016" & vbTab & "US-16 - Design and Validation" & vbTab & "17-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201016" & vbTab & "402017" & vbTab & "US-16 - Implement Core Logic" & vbTab & "20-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201016" & vbTab & "402018" & vbTab & "US-16 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201017" & vbTab & "402019" & vbTab & "US-17 - Design and Validation" & vbTab & "17-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201017" & vbTab & "402020" & vbTab & "US-17 - Implement Core Logic" & vbTab & "20-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201017" & vbTab & "402021" & vbTab & "US-17 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201018" & vbTab & "402022" & vbTab & "US-18 - Design and Validation" & vbTab & "20-Apr-26" & vbTab & "20-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201018" & vbTab & "402023" & vbTab & "US-18 - Implement Core Logic" & vbTab & "21-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201018" & vbTab & "402024" & vbTab & "US-18 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201019" & vbTab & "402025" & vbTab & "US-19 - Design and Validation" & vbTab & "21-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201019" & vbTab & "402026" & vbTab & "US-19 - Implement Core Logic" & vbTab & "22-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201019" & vbTab & "402027" & vbTab & "US-19 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201020" & vbTab & "402028" & vbTab & "US-20 - Design and Validation" & vbTab & "21-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201020" & vbTab & "402029" & vbTab & "US-20 - Implement Core Logic" & vbTab & "22-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201020" & vbTab & "402030" & vbTab & "US-20 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201021" & vbTab & "402031" & vbTab & "US-21 - Design and Validation" & vbTab & "22-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201021" & vbTab & "402032" & vbTab & "US-21 - Implement Core Logic" & vbTab & "23-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201021" & vbTab & "402033" & vbTab & "US-21 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201022" & vbTab & "402034" & vbTab & "US-22 - Design and Validation" & vbTab & "23-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201022" & vbTab & "402035" & vbTab & "US-22 - Implement Core Logic" & vbTab & "24-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201022" & vbTab & "402036" & vbTab & "US-22 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201023" & vbTab & "402037" & vbTab & "US-23 - Design and Validation" & vbTab & "22-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201023" & vbTab & "402038" & vbTab & "US-23 - Implement Core Logic" & vbTab & "23-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201023" & vbTab & "402039" & vbTab & "US-23 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201024" & vbTab & "402040" & vbTab & "US-24 - Design and Validation" & vbTab & "24-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201024" & vbTab & "402041" & vbTab & "US-24 - Implement Core Logic" & vbTab & "27-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201024" & vbTab & "402042" & vbTab & "US-24 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201025" & vbTab & "402043" & vbTab & "US-25 - Design and Validation" & vbTab & "27-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201025" & vbTab & "402044" & vbTab & "US-25 - Implement Core Logic" & vbTab & "28-Apr-26" & vbTab & "28-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5" & vbTab & "0" & vbTab & "0" & vbTab & "0", _
    "201025" & vbTab & "402045" & vbTab & "US-25 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" & vbTab & "0" & vbTab & "0" & vbTab & "0" _
  )
  WriteLines "Sprint II Backlog", lines
End Sub

Sub WriteStandup()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "Day" & vbTab & "Impediments" & vbTab & "Action Taken", _
    "Sprint 1" & vbTab & "Day 1" & vbTab & "Authentication and account boundaries needed agreement between the shared owners." & vbTab & "Sarosh led the core auth/account setup while Shraddha aligned shared access and admin constraints.", _
    "Sprint 1" & vbTab & "Day 2" & vbTab & "Category, budget, and dashboard work depended on stable transaction inputs." & vbTab & "Mapped transaction output fields first so dashboard, budget, and category work could proceed cleanly.", _
    "Sprint 1" & vbTab & "Day 3" & vbTab & "Sarosh and Susmit absences reduced Sprint 1 availability." & vbTab & "Shifted work away from the leave window and updated capacity planning to reflect the real attendance.", _
    "Sprint 1" & vbTab & "Day 4" & vbTab & "Backup/export tasks needed shared schema understanding with dashboard work." & vbTab & "Prathamesh aligned export structure and dashboard data contracts together.", _
    "Sprint 1" & vbTab & "Day 5" & vbTab & "Shared admin/auth ownership needed clearer separation." & vbTab & "Shraddha focused on admin-system concerns while Sarosh focused on user auth/account flow delivery.", _
    "Sprint 2" & vbTab & "Day 1" & vbTab & "Improved onboarding content required alignment with auth flow and first-use guidance." & vbTab & "Assigned the onboarding improvement under the shared authentication module and completed the content updates.", _
    "Sprint 2" & vbTab & "Day 2" & vbTab & "Admin activity and system-performance views depended on consistent logging." & vbTab & "Shraddha finalized both the audit history and performance metrics stories with the same logging source.", _
    "Sprint 2" & vbTab & "Day 3" & vbTab & "Budget, category, and dashboard stories overlapped in user-facing spending analysis." & vbTab & "Kept recurring expense and notification work with Sudhanshu while Prathamesh handled dashboard views.", _
    "Sprint 2" & vbTab & "Day 4" & vbTab & "Account settings changes touched language, password, and salary automation together." & vbTab & "Sarosh finished the account-module stories and queued testing/integration into the post-28/04/2026 window.", _
    "Sprint 2" & vbTab & "Day 5" & vbTab & "Code was complete but integration and testing still remained." & vbTab & "Kept integration and testing tasks in progress up to 04/05/2026 without using weekends as working days." _
  )
  WriteLines "Stand up Meeting", lines
End Sub

Sub WriteRetrospection()
  Dim lines
  lines = Array( _
    "SL #" & vbTab & "Sprint #" & vbTab & "Sprint start date" & vbTab & "Sprint end date" & vbTab & "Team member name" & vbTab & "Start Doing" & vbTab & "Stop Doing" & vbTab & "Continue Doing" & vbTab & "Action taken", _
    "1" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Shraddha" & vbTab & "Start locking shared-module boundaries earlier." & vbTab & "Stop letting admin and auth assumptions stay implicit." & vbTab & "Continue driving admin-system decisions and goals delivery." & vbTab & "Shared ownership with Sarosh was made explicit in the module allocation and backlog.", _
    "2" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Susmit" & vbTab & "Start sequencing report work directly after transaction changes." & vbTab & "Stop scattering analytics fixes across unrelated tasks." & vbTab & "Continue owning transaction and reporting depth." & vbTab & "Transaction Module and Report and Analytics Module remained fully aligned under Susmit.", _
    "3" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Prathamesh" & vbTab & "Start pairing dashboard and export planning from the start." & vbTab & "Stop waiting until the end to verify shared data views." & vbTab & "Continue clean delivery on presentation-heavy modules." & vbTab & "Dashboard, import/export, and backup work stayed grouped under Prathamesh.", _
    "4" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Sarosh" & vbTab & "Start surfacing leave impact early in the sprint plan." & vbTab & "Stop assuming shared auth/admin coverage happens automatically." & vbTab & "Continue strong ownership of account and user-entry flows." & vbTab & "The absence-aware plan now reflects Sarosh's Sprint 1 leave correctly.", _
    "5" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Sudhanshu" & vbTab & "Start reviewing budget and category touchpoints together." & vbTab & "Stop treating notification behavior as separate from user budget flow." & vbTab & "Continue consistent ownership of spending-control features." & vbTab & "Budget, Notification, and Category responsibilities now match the image-based allocation." _
  )
  WriteLines "Retrospection", lines
End Sub

Sub WriteGrooming()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "US ID" & vbTab & "Points Discussed", _
    "Sprint 1" & vbTab & "US 1" & vbTab & "Module: Authentication and Authorization Module; Priority: Must Have; Assignee: Sarosh; Shared ownership note: module is shared with Shraddha for security alignment.", _
    "Sprint 1" & vbTab & "US 2" & vbTab & "Module: Account Module; Priority: Must Have; Assignee: Sarosh; Acceptance: Save currency, locale, and month-start values with validation.", _
    "Sprint 1" & vbTab & "US 3" & vbTab & "Module: Transaction Module; Priority: Must Have; Assignee: Susmit; Acceptance: Add income entries that update balances and summaries.", _
    "Sprint 1" & vbTab & "US 4" & vbTab & "Module: Category Module; Priority: Must Have; Assignee: Sudhanshu; Acceptance: Support category-based expense capture and later analysis.", _
    "Sprint 1" & vbTab & "US 5" & vbTab & "Module: Budget Module; Priority: Must Have; Assignee: Sudhanshu; Acceptance: Monthly category budgets with threshold visibility.", _
    "Sprint 1" & vbTab & "US 6" & vbTab & "Module: Savings and Goals Module; Priority: Should Have; Assignee: Shraddha; Acceptance: Create goals and track progress.", _
    "Sprint 1" & vbTab & "US 7" & vbTab & "Module: Dashboard Module; Priority: Must Have; Assignee: Prathamesh; Acceptance: KPI cards and core finance charts.", _
    "Sprint 1" & vbTab & "US 8" & vbTab & "Module: Transaction Module; Priority: Should Have; Assignee: Susmit; Acceptance: Search and filter transaction history.", _
    "Sprint 1" & vbTab & "US 9" & vbTab & "Module: Data Import/Export and Backup Module; Priority: Could Have; Assignee: Prathamesh; Acceptance: Import/export transaction data by CSV template.", _
    "Sprint 1" & vbTab & "US 10" & vbTab & "Module: Admin and System Module; Priority: Must Have; Assignee: Shraddha; Shared ownership note: module is shared with Sarosh for access alignment.", _
    "Sprint 2" & vbTab & "US 11" & vbTab & "Module: Authentication and Authorization Module; Priority: Should Have; Assignee: Shraddha; Acceptance: Better onboarding page with platform steps, highlights, and guidance.", _
    "Sprint 2" & vbTab & "US 12" & vbTab & "Module: Dashboard Module; Priority: Must Have; Assignee: Prathamesh; Acceptance: Monthly income-versus-expense comparison.", _
    "Sprint 2" & vbTab & "US 13" & vbTab & "Module: Category Module; Priority: Must Have; Assignee: Sudhanshu; Acceptance: Category-wise spending analysis.", _
    "Sprint 2" & vbTab & "US 14" & vbTab & "Module: Dashboard Module; Priority: Must Have; Assignee: Prathamesh; Acceptance: Average monthly spending view.", _
    "Sprint 2" & vbTab & "US 15" & vbTab & "Module: Admin and System Module; Priority: Should Have; Assignee: Shraddha; Acceptance: Admin activity history from logs.", _
    "Sprint 2" & vbTab & "US 16" & vbTab & "Module: Dashboard Module; Priority: Must Have; Assignee: Prathamesh; Acceptance: Budget-versus-actual dashboard comparison.", _
    "Sprint 2" & vbTab & "US 17" & vbTab & "Module: Transaction Module; Priority: Must Have; Assignee: Susmit; Acceptance: Edit and delete transactions.", _
    "Sprint 2" & vbTab & "US 18" & vbTab & "Module: Transaction Module; Priority: Should Have; Assignee: Susmit; Acceptance: Notes and descriptions on transactions.", _
    "Sprint 2" & vbTab & "US 19" & vbTab & "Module: Reminder and Notification Modules; Priority: Should Have; Assignee: Sudhanshu; Acceptance: Mark notifications as read.", _
    "Sprint 2" & vbTab & "US 20" & vbTab & "Module: Report and Analytics Module; Priority: Must Have; Assignee: Susmit; Acceptance: Date-range transaction history view.", _
    "Sprint 2" & vbTab & "US 21" & vbTab & "Module: Budget Module; Priority: Should Have; Assignee: Sudhanshu; Acceptance: Save/manage recurring expenses with schedule and status.", _
    "Sprint 2" & vbTab & "US 22" & vbTab & "Module: Admin and System Module; Priority: Should Have; Assignee: Shraddha; Acceptance: System performance metrics for admins.", _
    "Sprint 2" & vbTab & "US 23" & vbTab & "Module: Account Module; Priority: Must Have; Assignee: Sarosh; Acceptance: Reset password from account page.", _
    "Sprint 2" & vbTab & "US 24" & vbTab & "Module: Account Module; Priority: Should Have; Assignee: Sarosh; Acceptance: Language switch for dashboard labels.", _
    "Sprint 2" & vbTab & "US 25" & vbTab & "Module: Account Module; Priority: Should Have; Assignee: Sarosh; Acceptance: Monthly salary setting with automatic monthly income addition." _
  )
  WriteLines "Product Backlog Grooming", lines
End Sub

Sub WriteModuleAllocation()
  Dim lines
  lines = Array( _
    "Owner" & vbTab & "Assigned Modules", _
    "Shraddha" & vbTab & "Admin and System Module(Shared), Saving and Goals Module, Authentication and Authorization Module(Shared)", _
    "Susmit" & vbTab & "Transaction Module, Report and Analytics Module", _
    "Prathamesh" & vbTab & "Dashboard Module, Data Import/Export and Backup Module", _
    "Sarosh" & vbTab & "Authentication and Authorization Module(Shared), Admin and System Module(Shared), Account Module", _
    "Sudhanshu" & vbTab & "Budget Module, Notification Module, Category Module" _
  )
  WriteLines "Module Allocation", lines
End Sub

Sub WriteDailyTasks()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "Date" & vbTab & "Susmit" & vbTab & "Sudhanshu" & vbTab & "Shraddha (SCRUM Master)" & vbTab & "Prathamesh" & vbTab & "Sarosh", _
    "0" & vbTab & "25/03/2026" & vbTab & "Studied SRS for transaction/report scope" & vbTab & "Studied SRS for budget/category/notification scope" & vbTab & "Studied SRS for shared admin/auth and goals scope" & vbTab & "Studied SRS for dashboard and backup scope" & vbTab & "Studied SRS for shared auth/admin and account scope", _
    "0" & vbTab & "26/03/2026" & vbTab & "Reviewed transaction and analytics requirements" & vbTab & "Reviewed budget, category, and notification requirements" & vbTab & "Reviewed admin, auth, and goals dependencies" & vbTab & "Reviewed dashboard and import/export dependencies" & vbTab & "Reviewed authentication and account journeys", _
    "0" & vbTab & "27/03/2026" & vbTab & "Helped size transaction/report stories" & vbTab & "Helped size budget/category stories" & vbTab & "Prepared sprint structure and shared-module plan" & vbTab & "Prepared dashboard/export plan" & vbTab & "Prepared account flow and auth notes", _
    "1" & vbTab & "30/03/2026" & vbTab & "Created transaction/report structure" & vbTab & "Created budget/category structure" & vbTab & "Set up goals and admin-system structure" & vbTab & "Set up dashboard and backup structure" & vbTab & "Set up auth/account structure", _
    "1" & vbTab & "31/03/2026" & vbTab & "Defined transaction entities and routes" & vbTab & "Defined budget and category data needs" & vbTab & "Aligned shared auth/admin rules and goal entities" & vbTab & "Planned dashboard widgets and export flow" & vbTab & "Built register/login/account flow structure", _
    "1" & vbTab & "01/04/2026" & vbTab & "Started income entry logic" & vbTab & "Started expense/category handling" & vbTab & "Started goals module logic" & vbTab & "Started dashboard KPI planning" & vbTab & "Started onboarding and auth screens", _
    "1" & vbTab & "02/04/2026" & vbTab & "Implemented transaction create flows" & vbTab & "Implemented category mapping and budget create flow" & vbTab & "Implemented shared security checks and goal CRUD" & vbTab & "Implemented dashboard base and backup plan" & vbTab & "Implemented login/register and settings flow", _
    "1" & vbTab & "03/04/2026" & vbTab & "Added transaction validations" & vbTab & "Added budget thresholds and category checks" & vbTab & "Added admin access rules and goal progress rules" & vbTab & "Connected dashboard data cards and import/export structure" & vbTab & "Improved onboarding copy and account setup", _
    "1" & vbTab & "06/04/2026" & vbTab & "Added transaction listing and filters" & vbTab & "Added budget views" & vbTab & "Added goal summary cards and shared auth reviews" & vbTab & "Built dashboard KPI layout" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "07/04/2026" & vbTab & "Absent (sick leave)" & vbTab & "Added category refinement and overspending indicators" & vbTab & "Added admin access plan and shared auth support" & vbTab & "Added dashboard chart logic" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "08/04/2026" & vbTab & "Implemented transaction filters" & vbTab & "Implemented budget alert states" & vbTab & "Implemented admin dashboard base" & vbTab & "Implemented dashboard comparison base" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "09/04/2026" & vbTab & "Implemented report export structure" & vbTab & "Implemented notification trigger hooks" & vbTab & "Implemented audit logging hooks" & vbTab & "Implemented backup/export wiring" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "10/04/2026" & vbTab & "Completed sprint 1 transaction/report tasks" & vbTab & "Completed sprint 1 budget/category tasks" & vbTab & "Completed sprint 1 admin/goals/shared-auth tasks" & vbTab & "Completed sprint 1 dashboard/backup tasks" & vbTab & "Absent (planned leave)", _
    "2" & vbTab & "13/04/2026" & vbTab & "Started transaction history improvements" & vbTab & "Started notification module tasks" & vbTab & "Started improved onboarding and admin refinements" & vbTab & "Started advanced dashboard work" & vbTab & "Started account module refinements", _
    "2" & vbTab & "14/04/2026" & vbTab & "Worked on notes and delete flows" & vbTab & "Worked on notification read states" & vbTab & "Worked on onboarding experience and admin controls" & vbTab & "Worked on dashboard layout cleanup" & vbTab & "Worked on account settings UX", _
    "2" & vbTab & "15/04/2026" & vbTab & "Worked on report/date-range flow" & vbTab & "Worked on category-wise spending support" & vbTab & "Worked on activity history view" & vbTab & "Worked on category-linked dashboard insights" & vbTab & "Improved account flows and reset-password planning", _
    "2" & vbTab & "16/04/2026" & vbTab & "Added report filtering refinements" & vbTab & "Added scheduled notification behavior" & vbTab & "Added admin logs and system controls" & vbTab & "Added dashboard analytics widgets" & vbTab & "Added login security and account updates", _
    "2" & vbTab & "17/04/2026" & vbTab & "Hardened transaction update/delete rules" & vbTab & "Connected budget and notification alerts" & vbTab & "Added system review checkpoints" & vbTab & "Built budget-vs-actual dashboard handling" & vbTab & "Added account reset-password flow", _
    "2" & vbTab & "20/04/2026" & vbTab & "Finalized analytics and transaction validations" & vbTab & "Finalized budget recurring-expense and notification work" & vbTab & "Finalized admin audit refinements" & vbTab & "Finalized dashboard/category visual summaries" & vbTab & "Finalized account settings validations", _
    "2" & vbTab & "21/04/2026" & vbTab & "Added report history response cleanup" & vbTab & "Added notification actions" & vbTab & "Added admin data review" & vbTab & "Added category and budget-vs-actual analysis" & vbTab & "Added language switching support", _
    "2" & vbTab & "22/04/2026" & vbTab & "Reviewed report and transaction bugs" & vbTab & "Reviewed budget/category edge cases" & vbTab & "Reviewed system performance dependencies" & vbTab & "Reviewed dashboard/export dependencies" & vbTab & "Reviewed account module edge cases", _
    "2" & vbTab & "23/04/2026" & vbTab & "Fixed transaction/report defects" & vbTab & "Fixed budget/category/notification defects" & vbTab & "Fixed admin/goals defects" & vbTab & "Fixed dashboard and backup defects" & vbTab & "Fixed onboarding/account defects", _
    "2" & vbTab & "24/04/2026" & vbTab & "Completed transaction and analytics coding" & vbTab & "Completed budget/category/notification coding" & vbTab & "Completed admin and goals coding" & vbTab & "Completed dashboard and backup coding" & vbTab & "Completed account coding", _
    "2" & vbTab & "27/04/2026" & vbTab & "Prepared report integration checklist" & vbTab & "Prepared budget/category integration checklist" & vbTab & "Prepared admin/system integration checklist" & vbTab & "Prepared dashboard/backup integration checklist" & vbTab & "Prepared account integration checklist", _
    "2" & vbTab & "28/04/2026" & vbTab & "Reviewed cross-module transaction/report integration issues" & vbTab & "Reviewed budget, category, and notification full flow" & vbTab & "Reviewed onboarding, admin, and system metrics links" & vbTab & "Reviewed dashboard and backup full flow" & vbTab & "Reviewed auth/account full flow" _
  )
  WriteLines "Daily Tasks", lines
End Sub

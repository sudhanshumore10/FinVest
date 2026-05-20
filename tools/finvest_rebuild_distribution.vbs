Option Explicit

Dim excel, wb, outputPath
outputPath = "C:\Users\susmi\Downloads\AgilesheetFinVest_Updated_29-04-2026_DistributionAligned.xls"

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

Sub WriteLines(sheetName, lines, clearRows, clearCols)
  Dim ws, r, c, parts
  Set ws = wb.Worksheets(sheetName)
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
  WriteLines "Product Backlog", lines, 140, 10
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
  WriteLines "Capacity Planning", lines, 120, 12
End Sub

Sub WriteSprint1Backlog()
  Dim lines
  lines = Array( _
    "NOTES: Task sizing should be between 0.5 to 12 hours", _
    "US ID" & vbTab & "Task ID" & vbTab & "Task Description" & vbTab & "Task Start Date" & vbTab & "Task Completion Date" & vbTab & "Team Member" & vbTab & "Activity" & vbTab & "Status" & vbTab & "Original Estimate Effort (In Hours)", _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "141", _
    "SPRINT 1 BACKLOG", _
    "201001" & vbTab & "401001" & vbTab & "US-1 - Design and Validation" & vbTab & "30-Mar-26" & vbTab & "30-Mar-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "5", _
    "201001" & vbTab & "401002" & vbTab & "US-1 - Implement Core Logic" & vbTab & "31-Mar-26" & vbTab & "01-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201001" & vbTab & "401003" & vbTab & "US-1 - Integration and Testing" & vbTab & "02-Apr-26" & vbTab & "02-Apr-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201002" & vbTab & "401004" & vbTab & "US-2 - Design and Validation" & vbTab & "30-Mar-26" & vbTab & "31-Mar-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201002" & vbTab & "401005" & vbTab & "US-2 - Implement Core Logic" & vbTab & "01-Apr-26" & vbTab & "02-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201002" & vbTab & "401006" & vbTab & "US-2 - Integration and Testing" & vbTab & "03-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201003" & vbTab & "401007" & vbTab & "US-3 - Design and Validation" & vbTab & "01-Apr-26" & vbTab & "01-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201003" & vbTab & "401008" & vbTab & "US-3 - Implement Core Logic" & vbTab & "02-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201003" & vbTab & "401009" & vbTab & "US-3 - Integration and Testing" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201004" & vbTab & "401010" & vbTab & "US-4 - Design and Validation" & vbTab & "01-Apr-26" & vbTab & "01-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201004" & vbTab & "401011" & vbTab & "US-4 - Implement Core Logic" & vbTab & "02-Apr-26" & vbTab & "03-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201004" & vbTab & "401012" & vbTab & "US-4 - Integration and Testing" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201005" & vbTab & "401013" & vbTab & "US-5 - Design and Validation" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201005" & vbTab & "401014" & vbTab & "US-5 - Implement Core Logic" & vbTab & "07-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201005" & vbTab & "401015" & vbTab & "US-5 - Integration and Testing" & vbTab & "09-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201006" & vbTab & "401016" & vbTab & "US-6 - Design and Validation" & vbTab & "06-Apr-26" & vbTab & "06-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201006" & vbTab & "401017" & vbTab & "US-6 - Implement Core Logic" & vbTab & "07-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201006" & vbTab & "401018" & vbTab & "US-6 - Integration and Testing" & vbTab & "09-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201007" & vbTab & "401019" & vbTab & "US-7 - Design and Validation" & vbTab & "07-Apr-26" & vbTab & "07-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201007" & vbTab & "401020" & vbTab & "US-7 - Implement Core Logic" & vbTab & "08-Apr-26" & vbTab & "09-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201007" & vbTab & "401021" & vbTab & "US-7 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201008" & vbTab & "401022" & vbTab & "US-8 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201008" & vbTab & "401023" & vbTab & "US-8 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201008" & vbTab & "401024" & vbTab & "US-8 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201009" & vbTab & "401025" & vbTab & "US-9 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201009" & vbTab & "401026" & vbTab & "US-9 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201009" & vbTab & "401027" & vbTab & "US-9 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4", _
    "201010" & vbTab & "401028" & vbTab & "US-10 - Design and Validation" & vbTab & "08-Apr-26" & vbTab & "08-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201010" & vbTab & "401029" & vbTab & "US-10 - Implement Core Logic" & vbTab & "09-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201010" & vbTab & "401030" & vbTab & "US-10 - Integration and Testing" & vbTab & "10-Apr-26" & vbTab & "10-Apr-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "3-Completed" & vbTab & "4" _
  )
  WriteLines "Sprint I Backlog", lines, 180, 12
End Sub

Sub WriteSprint2Backlog()
  Dim lines
  lines = Array( _
    "NOTES: Task sizing should be between 0.5 to 12 hours", _
    "US ID" & vbTab & "Task ID" & vbTab & "Task Description" & vbTab & "Task Start Date" & vbTab & "Task Completion Date" & vbTab & "Team Member" & vbTab & "Activity" & vbTab & "Status" & vbTab & "Original Estimate Effort (In Hours)", _
    vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "205", _
    "SPRINT 2 BACKLOG", _
    "201011" & vbTab & "402001" & vbTab & "US-11 - Design and Validation" & vbTab & "13-Apr-26" & vbTab & "13-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201011" & vbTab & "402002" & vbTab & "US-11 - Implement Core Logic" & vbTab & "14-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201011" & vbTab & "402003" & vbTab & "US-11 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201012" & vbTab & "402004" & vbTab & "US-12 - Design and Validation" & vbTab & "13-Apr-26" & vbTab & "13-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201012" & vbTab & "402005" & vbTab & "US-12 - Implement Core Logic" & vbTab & "14-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201012" & vbTab & "402006" & vbTab & "US-12 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201013" & vbTab & "402007" & vbTab & "US-13 - Design and Validation" & vbTab & "15-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201013" & vbTab & "402008" & vbTab & "US-13 - Implement Core Logic" & vbTab & "16-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201013" & vbTab & "402009" & vbTab & "US-13 - Testing and Integration" & vbTab & "29-Apr-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201014" & vbTab & "402010" & vbTab & "US-14 - Design and Validation" & vbTab & "16-Apr-26" & vbTab & "16-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201014" & vbTab & "402011" & vbTab & "US-14 - Implement Core Logic" & vbTab & "17-Apr-26" & vbTab & "20-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201014" & vbTab & "402012" & vbTab & "US-14 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201015" & vbTab & "402013" & vbTab & "US-15 - Design and Validation" & vbTab & "15-Apr-26" & vbTab & "15-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201015" & vbTab & "402014" & vbTab & "US-15 - Implement Core Logic" & vbTab & "16-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201015" & vbTab & "402015" & vbTab & "US-15 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201016" & vbTab & "402016" & vbTab & "US-16 - Design and Validation" & vbTab & "17-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Prathamesh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201016" & vbTab & "402017" & vbTab & "US-16 - Implement Core Logic" & vbTab & "20-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Prathamesh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201016" & vbTab & "402018" & vbTab & "US-16 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Prathamesh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201017" & vbTab & "402019" & vbTab & "US-17 - Design and Validation" & vbTab & "17-Apr-26" & vbTab & "17-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201017" & vbTab & "402020" & vbTab & "US-17 - Implement Core Logic" & vbTab & "20-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201017" & vbTab & "402021" & vbTab & "US-17 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201018" & vbTab & "402022" & vbTab & "US-18 - Design and Validation" & vbTab & "20-Apr-26" & vbTab & "20-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201018" & vbTab & "402023" & vbTab & "US-18 - Implement Core Logic" & vbTab & "21-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5", _
    "201018" & vbTab & "402024" & vbTab & "US-18 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201019" & vbTab & "402025" & vbTab & "US-19 - Design and Validation" & vbTab & "21-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201019" & vbTab & "402026" & vbTab & "US-19 - Implement Core Logic" & vbTab & "22-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5", _
    "201019" & vbTab & "402027" & vbTab & "US-19 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201020" & vbTab & "402028" & vbTab & "US-20 - Design and Validation" & vbTab & "21-Apr-26" & vbTab & "21-Apr-26" & vbTab & "Susmit" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201020" & vbTab & "402029" & vbTab & "US-20 - Implement Core Logic" & vbTab & "22-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Susmit" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201020" & vbTab & "402030" & vbTab & "US-20 - Testing and Integration" & vbTab & "30-Apr-26" & vbTab & "04-May-26" & vbTab & "Susmit" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201021" & vbTab & "402031" & vbTab & "US-21 - Design and Validation" & vbTab & "22-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201021" & vbTab & "402032" & vbTab & "US-21 - Implement Core Logic" & vbTab & "23-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sudhanshu" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201021" & vbTab & "402033" & vbTab & "US-21 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sudhanshu" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201022" & vbTab & "402034" & vbTab & "US-22 - Design and Validation" & vbTab & "23-Apr-26" & vbTab & "23-Apr-26" & vbTab & "Shraddha" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201022" & vbTab & "402035" & vbTab & "US-22 - Implement Core Logic" & vbTab & "24-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Shraddha" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "6", _
    "201022" & vbTab & "402036" & vbTab & "US-22 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Shraddha" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201023" & vbTab & "402037" & vbTab & "US-23 - Design and Validation" & vbTab & "22-Apr-26" & vbTab & "22-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201023" & vbTab & "402038" & vbTab & "US-23 - Implement Core Logic" & vbTab & "23-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5", _
    "201023" & vbTab & "402039" & vbTab & "US-23 - Testing and Integration" & vbTab & "01-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201024" & vbTab & "402040" & vbTab & "US-24 - Design and Validation" & vbTab & "24-Apr-26" & vbTab & "24-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201024" & vbTab & "402041" & vbTab & "US-24 - Implement Core Logic" & vbTab & "27-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5", _
    "201024" & vbTab & "402042" & vbTab & "US-24 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4", _
    "201025" & vbTab & "402043" & vbTab & "US-25 - Design and Validation" & vbTab & "27-Apr-26" & vbTab & "27-Apr-26" & vbTab & "Sarosh" & vbTab & "Design" & vbTab & "3-Completed" & vbTab & "4", _
    "201025" & vbTab & "402044" & vbTab & "US-25 - Implement Core Logic" & vbTab & "28-Apr-26" & vbTab & "28-Apr-26" & vbTab & "Sarosh" & vbTab & "Coding" & vbTab & "3-Completed" & vbTab & "5", _
    "201025" & vbTab & "402045" & vbTab & "US-25 - Testing and Integration" & vbTab & "04-May-26" & vbTab & "04-May-26" & vbTab & "Sarosh" & vbTab & "Unit Test" & vbTab & "2-In Progress" & vbTab & "4" _
  )
  WriteLines "Sprint II Backlog", lines, 240, 12
End Sub

Sub WriteStandup()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "Day" & vbTab & "Impediments" & vbTab & "Action Taken", _
    "Sprint 1" & vbTab & "Day 1" & vbTab & "Authentication, account, and admin boundaries needed cleaner ownership." & vbTab & "Split auth/account/admin work exactly as shared in the module table and aligned dependencies.", _
    "Sprint 1" & vbTab & "Day 2" & vbTab & "Category ownership had to stay with the same person handling budgets and notifications." & vbTab & "Kept Category and Budget stories with Sudhanshu to reduce context switching.", _
    "Sprint 1" & vbTab & "Day 3" & vbTab & "Dashboard and import/export outputs depended on the same reporting data shape." & vbTab & "Grouped dashboard and data import/export work under Prathamesh for consistency.", _
    "Sprint 1" & vbTab & "Day 4" & vbTab & "Transaction and analytics features were tightly coupled." & vbTab & "Kept report and transaction stories with Susmit so filters and history stayed aligned.", _
    "Sprint 1" & vbTab & "Day 5" & vbTab & "Sprint 1 capacity was reduced by leave." & vbTab & "Adjusted work around Susmit's sick leave on 07/04/2026 and Sarosh's leave from 06/04/2026 to 10/04/2026.", _
    "Sprint 2" & vbTab & "Day 1" & vbTab & "Onboarding improvements required both product guidance and shared auth understanding." & vbTab & "Assigned onboarding enhancement to Shraddha under shared authentication responsibility.", _
    "Sprint 2" & vbTab & "Day 2" & vbTab & "System performance work needed admin metrics and activity history alignment." & vbTab & "Kept both admin activity and system performance stories under Shraddha.", _
    "Sprint 2" & vbTab & "Day 3" & vbTab & "Budget and recurring expense automation crossed reminder and category flows." & vbTab & "Handled these under Sudhanshu's budget/notification/category ownership.", _
    "Sprint 2" & vbTab & "Day 4" & vbTab & "Dashboard spending comparison depended on existing budget data rather than budget ownership." & vbTab & "Assigned the budget-vs-actual dashboard story to Prathamesh as shown in the story table.", _
    "Sprint 2" & vbTab & "Day 5" & vbTab & "Language, password reset, and salary settings all touched the account surface." & vbTab & "Kept those Sprint 2 account stories together under Sarosh.", _
    "Sprint 2" & vbTab & "Day 6" & vbTab & "Cross-module integration was still open on 27/04/2026." & vbTab & "Reserved the remaining working window from 29/04/2026 to 04/05/2026 for integration and regression testing.", _
    "Sprint 2" & vbTab & "Day 7" & vbTab & "Regression and integration testing were still pending on 28/04/2026." & vbTab & "Kept testing tasks in progress through 04/05/2026 while showing completed story ownership in the backlog." _
  )
  WriteLines "Stand up Meeting", lines, 120, 10
End Sub

Sub WriteRetrospection()
  Dim lines
  lines = Array( _
    "SL #" & vbTab & "Sprint #" & vbTab & "Sprint start date" & vbTab & "Sprint end date" & vbTab & "Team member name" & vbTab & "Start Doing" & vbTab & "Stop Doing" & vbTab & "Continue Doing" & vbTab & "Action taken", _
    "1" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Shraddha" & vbTab & "Start defining shared-module boundaries earlier." & vbTab & "Stop letting onboarding and admin ownership blur together informally." & vbTab & "Continue supporting admin and goals delivery." & vbTab & "Shraddha handled Admin and System Module (Shared), Saving and Goals Module, and shared authentication tasks such as onboarding.", _
    "2" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Susmit" & vbTab & "Start pairing report filters with transaction validation from day one." & vbTab & "Stop treating analytics and transaction fixes as separate streams." & vbTab & "Continue strong delivery on transaction-heavy work." & vbTab & "Susmit owned Transaction Module and Report and Analytics Module together.", _
    "3" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Prathamesh" & vbTab & "Start planning dashboard outputs with import/export needs together." & vbTab & "Stop waiting until late sprint to validate export-facing data views." & vbTab & "Continue shaping presentation-heavy modules well." & vbTab & "Prathamesh handled Dashboard Module and Data Import/Export and Backup Module, including budget-vs-actual dashboard work.", _
    "4" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Sarosh" & vbTab & "Start documenting shared auth responsibilities alongside account tasks." & vbTab & "Stop separating account decisions from authentication discussions." & vbTab & "Continue improving user-facing flows." & vbTab & "Sarosh handled Authentication and Authorization Module (Shared), Admin and System Module (Shared), and Account Module stories.", _
    "5" & vbTab & "Sprint 1 & 2" & vbTab & "30/03/2026" & vbTab & "28/04/2026" & vbTab & "Sudhanshu" & vbTab & "Start reviewing category effects whenever budget changes are introduced." & vbTab & "Stop postponing reminder-state review until the end of the sprint." & vbTab & "Continue maintaining finance-control consistency." & vbTab & "Sudhanshu handled Budget Module, Notification Module, and Category Module together." _
  )
  WriteLines "Retrospection", lines, 120, 12
End Sub

Sub WriteGrooming()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "US ID" & vbTab & "Points Discussed", _
    "Sprint 1" & vbTab & "US 1" & vbTab & "Module: Authentication and Authorization Module; Priority: Must Have; Acceptance: Register with unique email, secure password hash, login session, logout, and protected route access.", _
    "Sprint 1" & vbTab & "US 2" & vbTab & "Module: Account Module; Priority: Must Have; Acceptance: Save currency, locale, and month-start day and reflect them across reports and account settings.", _
    "Sprint 1" & vbTab & "US 3" & vbTab & "Module: Transaction Module; Priority: Must Have; Acceptance: Add income entries with amount, category, date, and description; update balances and summaries.", _
    "Sprint 1" & vbTab & "US 4" & vbTab & "Module: Category Module; Priority: Must Have; Acceptance: Add expenses with category-wise analysis support and clean category ownership.", _
    "Sprint 1" & vbTab & "US 5" & vbTab & "Module: Budget Module; Priority: Must Have; Acceptance: Set category budgets by month and show threshold indicators.", _
    "Sprint 1" & vbTab & "US 6" & vbTab & "Module: Savings and Goals Module; Priority: Should Have; Acceptance: Create goals with target, deadline, and progress tracking.", _
    "Sprint 1" & vbTab & "US 7" & vbTab & "Module: Dashboard Module; Priority: Must Have; Acceptance: Show KPI cards and charts for quick financial understanding.", _
    "Sprint 1" & vbTab & "US 8" & vbTab & "Module: Transaction Module; Priority: Should Have; Acceptance: Filter and search transactions by date, category, and amount.", _
    "Sprint 1" & vbTab & "US 9" & vbTab & "Module: Data Import/Export and Backup Module; Priority: Could Have; Acceptance: Import and export transactions using a CSV-ready structure.", _
    "Sprint 1" & vbTab & "US 10" & vbTab & "Module: Admin and System Module; Priority: Must Have; Acceptance: Restrict admin routes, maintain audit visibility, and support secure system actions.", _
    "Sprint 2" & vbTab & "US 11" & vbTab & "Module: Authentication and Authorization Module; Priority: Should Have; Acceptance: Show an improved onboarding guide with usage steps and platform highlights.", _
    "Sprint 2" & vbTab & "US 12" & vbTab & "Module: Dashboard Module; Priority: Must Have; Acceptance: Display monthly income-versus-expense comparison using dynamic visual summaries.", _
    "Sprint 2" & vbTab & "US 13" & vbTab & "Module: Category Module; Priority: Must Have; Acceptance: Show category-wise spending analysis and highlight major spending areas.", _
    "Sprint 2" & vbTab & "US 14" & vbTab & "Module: Dashboard Module; Priority: Must Have; Acceptance: Calculate average monthly spending from past data for better insight.", _
    "Sprint 2" & vbTab & "US 15" & vbTab & "Module: Admin and System Module; Priority: Should Have; Acceptance: Show login, logout, and action history so admins can track activity.", _
    "Sprint 2" & vbTab & "US 16" & vbTab & "Module: Dashboard Module; Priority: Must Have; Acceptance: Display budget versus actual spending through dashboard comparison visuals.", _
    "Sprint 2" & vbTab & "US 17" & vbTab & "Module: Transaction Module; Priority: Must Have; Acceptance: Edit and delete transaction entries with validation and balance consistency.", _
    "Sprint 2" & vbTab & "US 18" & vbTab & "Module: Transaction Module; Priority: Should Have; Acceptance: Save notes and descriptions against transactions for better traceability.", _
    "Sprint 2" & vbTab & "US 19" & vbTab & "Module: Reminder and Notification Modules; Priority: Should Have; Acceptance: Mark notifications read or unread and manage alert state clearly.", _
    "Sprint 2" & vbTab & "US 20" & vbTab & "Module: Report and Analytics Module; Priority: Must Have; Acceptance: Filter transaction history by selected date range and present the result set clearly.", _
    "Sprint 2" & vbTab & "US 21" & vbTab & "Module: Budget Module; Priority: Should Have; Acceptance: Save recurring expenses with account, category, amount, frequency, start date, and status.", _
    "Sprint 2" & vbTab & "US 22" & vbTab & "Module: Admin and System Module; Priority: Should Have; Acceptance: Show active logins, DB response time, last request time, and system volume.", _
    "Sprint 2" & vbTab & "US 23" & vbTab & "Module: Account Module; Priority: Must Have; Acceptance: Verify current password, validate new password, confirm match, and store the new hash securely.", _
    "Sprint 2" & vbTab & "US 24" & vbTab & "Module: Account Module; Priority: Should Have; Acceptance: Let users select English, Hindi, or Marathi and apply translated labels across dashboard pages.", _
    "Sprint 2" & vbTab & "US 25" & vbTab & "Module: Account Module; Priority: Should Have; Acceptance: Save monthly salary and auto-add it once per month to the default account." _
  )
  WriteLines "Product Backlog Grooming", lines, 140, 12
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
  WriteLines "Module Allocation", lines, 80, 8
End Sub

Sub WriteDailyTasks()
  Dim lines
  lines = Array( _
    "Sprint" & vbTab & "Date" & vbTab & "Susmit" & vbTab & "Sudhanshu" & vbTab & "Shraddha (SCRUM Master)" & vbTab & "Prathamesh" & vbTab & "Sarosh", _
    "0" & vbTab & "25/03/2026" & vbTab & "Studied SRS for transaction and analytics scope" & vbTab & "Studied SRS for budget, notification, and category scope" & vbTab & "Studied SRS for shared admin/auth and goals scope" & vbTab & "Studied SRS for dashboard and import/export scope" & vbTab & "Studied SRS for shared auth/admin and account scope", _
    "0" & vbTab & "26/03/2026" & vbTab & "Reviewed transaction and report dependencies" & vbTab & "Reviewed category-budget-notification links" & vbTab & "Reviewed shared-module dependencies" & vbTab & "Reviewed dashboard and backup needs" & vbTab & "Reviewed authentication and account journey mapping", _
    "0" & vbTab & "27/03/2026" & vbTab & "Helped finalize transaction/report estimates" & vbTab & "Helped finalize category and budget links" & vbTab & "Prepared sprint plan and shared-module split" & vbTab & "Drafted dashboard/data-flow map" & vbTab & "Drafted auth/account flow notes", _
    "1" & vbTab & "30/03/2026" & vbTab & "Created transaction/report file structure" & vbTab & "Created budget/category structure" & vbTab & "Set up goals and admin backend structure" & vbTab & "Set up dashboard and import/export structure" & vbTab & "Set up shared auth and account module structure", _
    "1" & vbTab & "31/03/2026" & vbTab & "Defined transaction entities and routes" & vbTab & "Defined category and budget data needs" & vbTab & "Finalized sprint board and goal entities" & vbTab & "Planned dashboard widgets and data import/export needs" & vbTab & "Built register/login and account flow structure", _
    "1" & vbTab & "01/04/2026" & vbTab & "Started income and expense logic" & vbTab & "Started category expense flow" & vbTab & "Started goals module logic and shared auth support" & vbTab & "Started dashboard KPI layout" & vbTab & "Started onboarding and auth screens", _
    "1" & vbTab & "02/04/2026" & vbTab & "Implemented transaction create flows" & vbTab & "Implemented category handling and budget setup" & vbTab & "Implemented goal create/update flows" & vbTab & "Implemented dashboard sections and CSV I/O planning" & vbTab & "Implemented login, register, and session handling", _
    "1" & vbTab & "03/04/2026" & vbTab & "Added validation for transaction inputs" & vbTab & "Added category and budget rules" & vbTab & "Added goal progress rules" & vbTab & "Connected dashboard cards with source data" & vbTab & "Improved onboarding page content", _
    "1" & vbTab & "06/04/2026" & vbTab & "Added transaction listing and filters" & vbTab & "Added monthly budget views" & vbTab & "Added goal summary cards" & vbTab & "Built dashboard KPI layout and export planning" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "07/04/2026" & vbTab & "Absent (sick leave)" & vbTab & "Added overspending indicators" & vbTab & "Added admin access plan" & vbTab & "Added dashboard chart logic" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "08/04/2026" & vbTab & "Implemented transaction filters" & vbTab & "Implemented budget alert states" & vbTab & "Implemented admin dashboard base" & vbTab & "Implemented CSV import/export flow base" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "09/04/2026" & vbTab & "Implemented report export flow" & vbTab & "Implemented notification trigger hooks" & vbTab & "Implemented audit logging hooks" & vbTab & "Implemented dashboard comparison charts" & vbTab & "Absent (planned leave)", _
    "1" & vbTab & "10/04/2026" & vbTab & "Completed sprint 1 transaction/report tasks" & vbTab & "Completed sprint 1 category/budget tasks" & vbTab & "Completed sprint 1 goals/admin tasks" & vbTab & "Completed sprint 1 dashboard/import-export tasks" & vbTab & "Absent (planned leave)", _
    "2" & vbTab & "13/04/2026" & vbTab & "Started transaction history improvements" & vbTab & "Started notification module tasks" & vbTab & "Started onboarding enhancement and shared auth refinements" & vbTab & "Started advanced dashboard work" & vbTab & "Started account polish and shared auth review", _
    "2" & vbTab & "14/04/2026" & vbTab & "Worked on notes and delete flows" & vbTab & "Worked on notification read states" & vbTab & "Worked on onboarding flow and admin user controls" & vbTab & "Worked on dashboard layout cleanup" & vbTab & "Worked on account settings UX", _
    "2" & vbTab & "15/04/2026" & vbTab & "Worked on report/date-range flow" & vbTab & "Worked on category-wise spending analysis" & vbTab & "Worked on activity history view" & vbTab & "Worked on category-linked dashboard metrics and backup mapping" & vbTab & "Improved onboarding support and account flow polish", _
    "2" & vbTab & "16/04/2026" & vbTab & "Added report filtering refinements" & vbTab & "Added scheduled reminder behavior" & vbTab & "Added admin logs and system controls" & vbTab & "Added dashboard analytics widgets" & vbTab & "Added login security and account updates", _
    "2" & vbTab & "17/04/2026" & vbTab & "Hardened transaction update/delete rules" & vbTab & "Connected budget and reminder alerts" & vbTab & "Added user block and unblock controls" & vbTab & "Built data backup/export handling and budget-vs-actual dashboard view" & vbTab & "Added account reset-password flow", _
    "2" & vbTab & "20/04/2026" & vbTab & "Finalized analytics and transaction validations" & vbTab & "Finalized category, budget, and notification behaviors" & vbTab & "Finalized admin audit refinements" & vbTab & "Finalized dashboard visual summaries and backup flow" & vbTab & "Finalized account settings validations", _
    "2" & vbTab & "21/04/2026" & vbTab & "Added report history response cleanup" & vbTab & "Added notification actions" & vbTab & "Added admin performance review" & vbTab & "Added dashboard budget-comparison refinements" & vbTab & "Added language switching support", _
    "2" & vbTab & "22/04/2026" & vbTab & "Reviewed report and transaction bugs" & vbTab & "Reviewed budget/category/reminder edge cases" & vbTab & "Reviewed admin metrics dependencies" & vbTab & "Reviewed dashboard/export dependencies" & vbTab & "Reviewed account module edge cases", _
    "2" & vbTab & "23/04/2026" & vbTab & "Fixed transaction/report defects" & vbTab & "Fixed budget/category/reminder defects" & vbTab & "Fixed admin/goal/shared-auth defects" & vbTab & "Fixed dashboard/import-export defects" & vbTab & "Fixed onboarding/account defects", _
    "2" & vbTab & "24/04/2026" & vbTab & "Completed transaction and analytics coding" & vbTab & "Completed category, budget, and notification coding" & vbTab & "Completed admin, goals, and shared-auth coding" & vbTab & "Completed dashboard and import/export coding" & vbTab & "Completed account coding", _
    "2" & vbTab & "27/04/2026" & vbTab & "Prepared report integration checklist" & vbTab & "Prepared budget/category/reminder integration checklist" & vbTab & "Prepared admin/goals/shared-auth integration checklist" & vbTab & "Prepared dashboard/import-export integration checklist" & vbTab & "Prepared account integration checklist", _
    "2" & vbTab & "28/04/2026" & vbTab & "Reviewed cross-module transaction/report integration issues" & vbTab & "Reviewed category-budget-notification full flow" & vbTab & "Reviewed onboarding, admin metrics, and goal links" & vbTab & "Reviewed dashboard/import-export full flow" & vbTab & "Reviewed auth/account final flow" _
  )
  WriteLines "Daily Tasks", lines, 180, 10
End Sub

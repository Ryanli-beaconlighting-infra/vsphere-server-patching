Option Explicit

' ============================================
' VSPHERE SERVER PATCHING AUTOMATION MACROS
' Updated: February 2026
' ============================================
' COLUMN STRUCTURE (NextDC M1):
'   A = Server Name
'   B = Priority
'   C = Next Scheduled Date
'   D = Last Patch Date
'   E = Status (OK/OVERDUE/UNSCHEDULED)
'   F = Days Until Due
'   G = Master Status
'   H = Notes
'   I onwards = Patch History
' ============================================

' --------------------------------------------------------------------------
' Module-level constants
' --------------------------------------------------------------------------
Private Const TARGET_SHEET_NAME As String = "NextDC M1"
Private Const MAX_COLUMNS As Long = 500

' --------------------------------------------------------------------------
' Helper: BackupSheet
' Creates a timestamped copy of the given worksheet before modifications.
' --------------------------------------------------------------------------
Private Sub BackupSheet(ws As Worksheet)
    Dim backupName As String
    backupName = ws.Name & "_backup_" & Format(Now, "yyyymmdd_hhnnss")
    ws.Copy After:=ws
    ActiveSheet.Name = backupName
End Sub

' --------------------------------------------------------------------------
' Helper: GetLastDataRow
' Returns the last row containing data in column 1 of the given worksheet.
' --------------------------------------------------------------------------
Private Function GetLastDataRow(ws As Worksheet) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

Sub RecordPatchDate()
    ' Records today's date as a patch for the selected server
    ' Usage: Select a server row on NextDC M1, then run this macro

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim nextCol As Long
    Dim lastCol As Long
    Dim serverName As String
    Dim response As VbMsgBoxResult

    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    ' Check we're on the right sheet
    If ActiveSheet.Name <> TARGET_SHEET_NAME Then
        MsgBox "Please go to the '" & TARGET_SHEET_NAME & "' sheet and select a server row first.", vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    selectedRow = ActiveCell.Row

    ' Validate row selection
    If selectedRow < 2 Then
        MsgBox "Please select a valid server row (row 2 or below).", vbExclamation, "Invalid Selection"
        Exit Sub
    End If

    serverName = ws.Cells(selectedRow, 1).Value
    If serverName = "" Then
        MsgBox "No server found in the selected row.", vbExclamation, "No Server"
        Exit Sub
    End If

    ' Confirm with user
    response = MsgBox("Record today's date (" & Format(Date, "DD/MM/YYYY") & ") as patch date for:" & vbCrLf & vbCrLf & _
                      "SERVER: " & serverName & vbCrLf & vbCrLf & _
                      "Click Yes to confirm.", vbYesNo + vbQuestion, "Confirm Patch Record")

    If response = vbYes Then
        ' Create backup before modifying data
        BackupSheet ws
        ' Re-set ws reference since BackupSheet changes ActiveSheet
        Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

        ' Find next empty column starting from I (column 9)
        nextCol = 9
        Do While ws.Cells(selectedRow, nextCol).Value <> ""
            nextCol = nextCol + 1
            If nextCol > MAX_COLUMNS Then Exit Do  ' Safety limit
        Loop

        ' Warn if approaching column limit
        lastCol = nextCol
        If lastCol > MAX_COLUMNS * 0.8 Then
            MsgBox "Warning: Column count (" & lastCol & ") approaching limit (" & MAX_COLUMNS & "). Consider archiving old data.", vbExclamation
        End If

        ' Record the date
        ws.Cells(selectedRow, nextCol).Value = Date
        ws.Cells(selectedRow, nextCol).NumberFormat = "DD/MM/YYYY"

        MsgBox "Patch date recorded successfully!" & vbCrLf & vbCrLf & _
               "Server: " & serverName & vbCrLf & _
               "Date: " & Format(Date, "DD/MM/YYYY"), vbInformation, "Success"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error in RecordPatchDate: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Exit Sub
End Sub

Sub QuickPatchMultiple()
    ' Records today's date for multiple selected servers at once
    ' Usage: Select multiple server rows on NextDC M1, then run this macro

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim serverList As String
    Dim serverCount As Integer
    Dim response As VbMsgBoxResult
    Dim nextCol As Long
    Dim lastCol As Long
    Dim serverName As String

    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    ' Check we're on the right sheet
    If ActiveSheet.Name <> TARGET_SHEET_NAME Then
        MsgBox "Please go to the '" & TARGET_SHEET_NAME & "' sheet and select server rows first.", vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    Set rng = Selection
    serverCount = 0
    serverList = ""

    ' Build list of selected servers
    For Each cell In rng.Rows
        If cell.Row >= 2 Then
            serverName = ws.Cells(cell.Row, 1).Value
            If serverName <> "" Then
                serverList = serverList & "  - " & serverName & vbCrLf
                serverCount = serverCount + 1
            End If
        End If
    Next cell

    If serverCount = 0 Then
        MsgBox "No valid servers selected." & vbCrLf & vbCrLf & _
               "Tip: Select cells in rows 2 or below that contain server names.", vbExclamation, "No Servers"
        Exit Sub
    End If

    ' Confirm with user
    response = MsgBox("Record today's date (" & Format(Date, "DD/MM/YYYY") & ") for " & serverCount & " server(s):" & vbCrLf & vbCrLf & _
                      serverList & vbCrLf & _
                      "Click Yes to confirm.", vbYesNo + vbQuestion, "Confirm Bulk Patch")

    If response = vbYes Then
        ' Create backup before modifying data
        BackupSheet ws
        ' Re-set ws reference since BackupSheet changes ActiveSheet
        Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

        ' Process each selected row
        For Each cell In rng.Rows
            If cell.Row >= 2 Then
                If ws.Cells(cell.Row, 1).Value <> "" Then
                    ' Find next empty column starting from I (column 9)
                    nextCol = 9
                    Do While ws.Cells(cell.Row, nextCol).Value <> ""
                        nextCol = nextCol + 1
                        If nextCol > MAX_COLUMNS Then Exit Do
                    Loop

                    ' Warn if approaching column limit
                    lastCol = nextCol
                    If lastCol > MAX_COLUMNS * 0.8 Then
                        MsgBox "Warning: Column count (" & lastCol & ") approaching limit (" & MAX_COLUMNS & "). Consider archiving old data.", vbExclamation
                    End If

                    ws.Cells(cell.Row, nextCol).Value = Date
                    ws.Cells(cell.Row, nextCol).NumberFormat = "DD/MM/YYYY"
                End If
            End If
        Next cell

        MsgBox serverCount & " server(s) updated successfully!" & vbCrLf & vbCrLf & _
               "Date recorded: " & Format(Date, "DD/MM/YYYY"), vbInformation, "Bulk Update Complete"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error in QuickPatchMultiple: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Exit Sub
End Sub

Sub ShowOverdueServers()
    ' Displays a list of all servers with OVERDUE status
    ' Text is automatically copied to clipboard for pasting
    ' Usage: Run from any sheet

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim overdueList As String
    Dim overdueCount As Integer
    Dim nextDueDate As Variant
    Dim daysOverdue As Long
    Dim clipboardText As String
    Dim DataObj As Object

    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    overdueList = ""
    overdueCount = 0

    ' Dynamic last row detection
    lastRow = GetLastDataRow(ws)

    ' Scan all servers
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            ' Check Status column (E)
            If ws.Cells(i, 5).Value = "OVERDUE" Then
                overdueCount = overdueCount + 1
                nextDueDate = ws.Cells(i, 3).Value

                If IsDate(nextDueDate) Then
                    daysOverdue = DateDiff("d", nextDueDate, Date)
                    overdueList = overdueList & overdueCount & ". " & ws.Cells(i, 1).Value & _
                                  " (Due: " & Format(nextDueDate, "DD/MM/YYYY") & _
                                  ", " & daysOverdue & " days overdue)" & vbCrLf
                Else
                    overdueList = overdueList & overdueCount & ". " & ws.Cells(i, 1).Value & vbCrLf
                End If
            End If
        End If
    Next i

    ' Display results
    If overdueCount = 0 Then
        MsgBox "No overdue servers found!" & vbCrLf & vbCrLf & _
               "All patching is on schedule.", vbInformation, "Overdue Check - All Clear"
    Else
        ' Build clipboard text
        clipboardText = "OVERDUE SERVERS (" & overdueCount & "):" & vbCrLf & vbCrLf & overdueList

        ' Copy to clipboard
        On Error Resume Next
        Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        DataObj.SetText clipboardText
        DataObj.PutInClipboard
        Set DataObj = Nothing
        On Error GoTo ErrorHandler

        MsgBox "OVERDUE SERVERS (" & overdueCount & "):" & vbCrLf & vbCrLf & _
               overdueList & vbCrLf & _
               ">>> TEXT COPIED TO CLIPBOARD - Press Ctrl+V to paste <<<", _
               vbExclamation, "Overdue Servers Found"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Error in ShowOverdueServers: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Exit Sub
End Sub

Sub GenerateEmailList()
    ' Creates a list of selected servers for email notification
    ' Text is automatically copied to clipboard for pasting
    ' Usage: Select server rows on NextDC M1, then run this macro

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim masterSheet As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim serverList As String
    Dim serverName As String
    Dim serverCount As Integer
    Dim i As Long
    Dim lastRow As Long
    Dim team As String
    Dim teamList As String
    Dim clipboardText As String
    Dim DataObj As Object

    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    On Error Resume Next
    Set masterSheet = ThisWorkbook.Sheets("Master Servers")
    On Error GoTo ErrorHandler

    If ActiveSheet.Name <> TARGET_SHEET_NAME Then
        MsgBox "Please go to the '" & TARGET_SHEET_NAME & "' sheet and select server rows first.", vbExclamation, "Wrong Sheet"
        Exit Sub
    End If

    Set rng = Selection
    serverList = ""
    teamList = ""
    serverCount = 0

    ' Build server list
    For Each cell In rng.Rows
        If cell.Row >= 2 Then
            serverName = ws.Cells(cell.Row, 1).Value
            If serverName <> "" Then
                serverCount = serverCount + 1
                serverList = serverList & "  - " & serverName & vbCrLf

                ' Try to get team from Master Servers (dynamic last row)
                If Not masterSheet Is Nothing Then
                    lastRow = GetLastDataRow(masterSheet)
                    For i = 5 To lastRow
                        If masterSheet.Cells(i, 1).Value = serverName Then
                            team = masterSheet.Cells(i, 4).Value
                            If team <> "" Then
                                If InStr(teamList, team) = 0 Then
                                    teamList = teamList & team & ", "
                                End If
                            End If
                            Exit For
                        End If
                    Next i
                End If
            End If
        End If
    Next cell

    If serverCount = 0 Then
        MsgBox "No servers selected.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' Remove trailing comma from team list
    If Len(teamList) > 2 Then
        teamList = Left(teamList, Len(teamList) - 2)
    End If

    ' Build the text for display and clipboard
    Dim msg As String
    msg = "SERVERS TO BE PATCHED (" & serverCount & "):" & vbCrLf & vbCrLf
    msg = msg & serverList
    If teamList <> "" Then
        msg = msg & vbCrLf & "TEAMS TO NOTIFY: " & teamList
    End If

    ' Copy to clipboard
    clipboardText = msg
    On Error Resume Next
    Set DataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    DataObj.SetText clipboardText
    DataObj.PutInClipboard
    Set DataObj = Nothing
    On Error GoTo ErrorHandler

    ' Display with clipboard confirmation
    MsgBox msg & vbCrLf & vbCrLf & _
           ">>> TEXT COPIED TO CLIPBOARD - Press Ctrl+V to paste <<<", _
           vbInformation, "Server List for Email"

    Exit Sub

ErrorHandler:
    MsgBox "Error in GenerateEmailList: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Exit Sub
End Sub

Sub ExportPatchReport()
    ' Generates a printable patch status report on a new sheet
    ' Usage: Run from any sheet

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim reportSheet As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim reportRow As Long
    Dim response As VbMsgBoxResult
    Dim statusValue As String

    Set ws = ThisWorkbook.Sheets(TARGET_SHEET_NAME)

    response = MsgBox("This will create a new 'Patch Report' sheet with the current patching status." & vbCrLf & vbCrLf & _
                      "Any existing report will be replaced." & vbCrLf & vbCrLf & _
                      "Continue?", vbYesNo + vbQuestion, "Generate Patch Report")

    If response <> vbYes Then Exit Sub

    ' Delete existing report if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Patch Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrorHandler

    ' Create new report sheet
    Set reportSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    reportSheet.Name = "Patch Report"

    ' Header
    reportSheet.Cells(1, 1) = "VSPHERE SERVER PATCHING STATUS REPORT"
    reportSheet.Cells(1, 1).Font.Bold = True
    reportSheet.Cells(1, 1).Font.Size = 16
    reportSheet.Range("A1:F1").Merge

    reportSheet.Cells(2, 1) = "Generated: " & Format(Now, "DD/MM/YYYY HH:MM") & " by " & Environ("USERNAME")
    reportSheet.Cells(2, 1).Font.Italic = True

    ' Column headers
    reportSheet.Cells(4, 1) = "Server Name"
    reportSheet.Cells(4, 2) = "Priority"
    reportSheet.Cells(4, 3) = "Next Due Date"
    reportSheet.Cells(4, 4) = "Last Patched"
    reportSheet.Cells(4, 5) = "Status"
    reportSheet.Cells(4, 6) = "Days Until Due"

    For i = 1 To 6
        reportSheet.Cells(4, i).Font.Bold = True
        reportSheet.Cells(4, i).Interior.Color = RGB(68, 114, 196)
        reportSheet.Cells(4, i).Font.Color = RGB(255, 255, 255)
    Next i

    ' Dynamic last row detection
    lastRow = GetLastDataRow(ws)

    ' Data rows
    reportRow = 5
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> "" Then
            reportSheet.Cells(reportRow, 1) = ws.Cells(i, 1).Value  ' Server Name (A)
            reportSheet.Cells(reportRow, 2) = ws.Cells(i, 2).Value  ' Priority (B)
            reportSheet.Cells(reportRow, 3) = ws.Cells(i, 3).Value  ' Next Due Date (C)
            reportSheet.Cells(reportRow, 4) = ws.Cells(i, 4).Value  ' Last Patched (D)
            reportSheet.Cells(reportRow, 5) = ws.Cells(i, 5).Value  ' Status (E)
            reportSheet.Cells(reportRow, 6) = ws.Cells(i, 6).Value  ' Days Until Due (F)

            ' Format dates
            reportSheet.Cells(reportRow, 3).NumberFormat = "DD/MM/YYYY"
            reportSheet.Cells(reportRow, 4).NumberFormat = "DD/MM/YYYY"

            ' Color code status
            statusValue = ws.Cells(i, 5).Value
            Select Case statusValue
                Case "OVERDUE"
                    reportSheet.Cells(reportRow, 5).Interior.Color = RGB(255, 107, 107)
                    reportSheet.Cells(reportRow, 5).Font.Bold = True
                Case "OK"
                    reportSheet.Cells(reportRow, 5).Interior.Color = RGB(144, 238, 144)
                Case "UNSCHEDULED"
                    reportSheet.Cells(reportRow, 5).Interior.Color = RGB(211, 211, 211)
            End Select

            reportRow = reportRow + 1
        End If
    Next i

    ' Auto-fit columns
    reportSheet.Columns("A:F").AutoFit

    ' Add summary at bottom
    reportRow = reportRow + 2
    reportSheet.Cells(reportRow, 1) = "SUMMARY:"
    reportSheet.Cells(reportRow, 1).Font.Bold = True
    reportSheet.Cells(reportRow + 1, 1) = "Total Servers: " & Application.WorksheetFunction.CountA(ws.Range("A:A")) - 1
    reportSheet.Cells(reportRow + 2, 1) = "Overdue: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "OVERDUE")
    reportSheet.Cells(reportRow + 3, 1) = "OK: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "OK")
    reportSheet.Cells(reportRow + 4, 1) = "Unscheduled: " & Application.WorksheetFunction.CountIf(ws.Range("E:E"), "UNSCHEDULED")

    reportSheet.Activate

    MsgBox "Patch report generated successfully!" & vbCrLf & vbCrLf & _
           "The report is ready to print or export.", vbInformation, "Report Complete"

    Exit Sub

ErrorHandler:
    MsgBox "Error in ExportPatchReport: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    ' Clean up display alerts if error occurred during sheet deletion
    Application.DisplayAlerts = True
    Exit Sub
End Sub

Sub RefreshDashboard()
    ' Recalculates all formulas and shows the Dashboard
    ' Usage: Run from any sheet

    On Error GoTo ErrorHandler

    Application.Calculate
    ThisWorkbook.Sheets("Dashboard").Activate
    MsgBox "Dashboard refreshed!" & vbCrLf & vbCrLf & _
           "All formulas have been recalculated.", vbInformation, "Refresh Complete"

    Exit Sub

ErrorHandler:
    MsgBox "Error in RefreshDashboard: " & Err.Description & " (Error " & Err.Number & ")", vbCritical
    Exit Sub
End Sub

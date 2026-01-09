# vSphere Server Patching Schedule - Complete User Guide

**Beacon Lighting Infrastructure Team**
**Last Updated:** January 2026

---

## Table of Contents

### Getting Started
- [1. Introduction](#1-introduction)
  - [1.1 Purpose of This Workbook](#11-purpose-of-this-workbook)
  - [1.2 Key Features](#12-key-features)
  - [1.3 File Location](#13-file-location)

### Understanding the Workbook
- [2. Workbook Overview](#2-workbook-overview)
  - [2.1 Sheet Structure](#21-sheet-structure)
  - [2.2 How the Sheets Work Together](#22-how-the-sheets-work-together)
  - [2.3 Color Coding Guide](#23-color-coding-guide)

### Sheet Guides
- [3. Dashboard Sheet](#3-dashboard-sheet)
  - [3.1 Purpose](#31-purpose)
  - [3.2 Sections Explained](#32-sections-explained)
  - [3.3 How to Use](#33-how-to-use)
- [4. Settings Sheet](#4-settings-sheet)
  - [4.1 Purpose](#41-purpose)
  - [4.2 Priority Table](#42-priority-table)
  - [4.3 Cluster List](#43-cluster-list)
  - [4.4 Team List](#44-team-list)
  - [4.5 How to Modify Settings](#45-how-to-modify-settings)
- [5. Master Servers Sheet](#5-master-servers-sheet)
  - [5.1 Purpose](#51-purpose)
  - [5.2 Column Descriptions](#52-column-descriptions)
  - [5.3 Adding a New Server](#53-adding-a-new-server)
  - [5.4 Decommissioning a Server](#54-decommissioning-a-server)
  - [5.5 Changing Server Details](#55-changing-server-details)
- [6. NextDC M1 Sheet (Main Patching Schedule)](#6-nextdc-m1-sheet-main-patching-schedule)
  - [6.1 Purpose](#61-purpose)
  - [6.2 Column Descriptions](#62-column-descriptions)
  - [6.3 Understanding Status Values](#63-understanding-status-values)
  - [6.4 Recording a Patch (Manual Method)](#64-recording-a-patch-manual-method)
  - [6.5 Using Dropdowns](#65-using-dropdowns)

### Automation
- [7. Macros - Automation Tools](#7-macros---automation-tools)
  - [7.1 What Are Macros?](#71-what-are-macros)
  - [7.2 One-Time Setup (Installing Macros)](#72-one-time-setup-installing-macros)
  - [7.3 How to Run Macros](#73-how-to-run-macros)
  - [7.4 Available Macros](#74-available-macros)

### How-To Guides
- [8. Common Tasks - Step by Step](#8-common-tasks---step-by-step)
  - [8.1 Recording a Single Patch](#81-recording-a-single-patch)
  - [8.2 Recording Multiple Patches (Bulk)](#82-recording-multiple-patches-bulk)
  - [8.3 Adding a New Server to the System](#83-adding-a-new-server-to-the-system)
  - [8.4 Checking Which Servers Are Overdue](#84-checking-which-servers-are-overdue)
  - [8.5 Changing a Server's Priority](#85-changing-a-servers-priority)
  - [8.6 Generating a Report](#86-generating-a-report)
  - [8.7 Sending Patch Notifications](#87-sending-patch-notifications)

### Help & Reference
- [9. Troubleshooting](#9-troubleshooting)
  - [9.1 Common Issues and Solutions](#91-common-issues-and-solutions)
  - [9.2 Formula Errors](#92-formula-errors)
  - [9.3 Dropdown Issues](#93-dropdown-issues)
  - [9.4 Macro Issues](#94-macro-issues)
- [10. Appendix](#10-appendix)
  - [10.1 Named Ranges Reference](#101-named-ranges-reference)
  - [10.2 Formula Reference](#102-formula-reference)
  - [10.3 Keyboard Shortcuts](#103-keyboard-shortcuts)

---

## 1. Introduction

### 1.1 Purpose of This Workbook

This workbook is the central system for tracking server patching schedules across our vSphere infrastructure. It helps the team:

- Know which servers need patching and when
- Track patch history for compliance and auditing
- Identify overdue servers at a glance
- Maintain consistent patching cycles based on server priority
- Generate reports for management and compliance

[Back to Top](#table-of-contents)

### 1.2 Key Features

- **Automatic status calculation** (OK, OVERDUE, UNSCHEDULED)
- **Priority-based scheduling** (configurable days per priority)
- **Master server list** as single source of truth
- **Dropdown menus** for consistent data entry
- **Clickable macro buttons** for one-click automation
- **Automation macros** for quick updates
- **Dashboard** for at-a-glance status overview
- **Unlimited patch history** tracking

[Back to Top](#table-of-contents)

### 1.3 File Location

**Primary File:**
```
OneDrive - Beacon Lighting\projects\vsphere-server-patching\
Vsphere_Server_Patching_Schedule.xlsx
```

**Macro-Enabled Version:**
```
Vsphere_Server_Patching_Schedule with macros.xlsm
```

**Supporting Files:**
- `Vsphere_Patching_Macros_UPDATED.bas` - VBA code for macros
- `Vsphere_Patching_Quick_Reference_Card.txt` - Printable quick reference
- This guide (`Vsphere_Patching_Complete_User_Guide.md`)

[Back to Top](#table-of-contents)

---

## 2. Workbook Overview

### 2.1 Sheet Structure

The workbook contains the following sheets (tabs at bottom of Excel):

| Sheet Name | Purpose |
|------------|---------|
| **Dashboard** | Overview of patching status, quick stats |
| **Settings** | Configure priorities, clusters, teams |
| **Master Servers** | Central list of all servers (source of truth) |
| **NextDC M1** | Main patching schedule and history tracking |
| **Server list** | Legacy/reference list (may contain app owners) |
| **Applications** | Application reference data |

[Back to Top](#table-of-contents)

### 2.2 How the Sheets Work Together

```
                    +------------+
                    |  Settings  |
                    +------------+
                         |
          +--------------+--------------+
          |              |              |
          v              v              v
    +-----------+  +-----------+  +-----------+
    | Priority  |  | Cluster   |  | Team      |
    | Table     |  | List      |  | List      |
    +-----------+  +-----------+  +-----------+
          |              |              |
          |              v              |
          |     +---------------+       |
          +---->| Master Servers|<------+
                +---------------+
                       |
                       | (Server names, priorities, status)
                       v
                +---------------+
                |   NextDC M1   |  <-- Main working sheet
                +---------------+
                       |
                       | (Statistics & counts)
                       v
                +---------------+
                |   Dashboard   |
                +---------------+
```

**Flow explanation:**
1. Settings defines the priority cycles (30/60/90 days)
2. Master Servers holds all server details
3. NextDC M1 pulls server info from Master and calculates due dates
4. Dashboard summarizes the status from NextDC M1

[Back to Top](#table-of-contents)

### 2.3 Color Coding Guide

#### Status Colors (NextDC M1 Column E)

| Color | Status | Meaning |
|-------|--------|---------|
| 🔴 **RED** | OVERDUE | Server is past its scheduled patch date. **ACTION REQUIRED: Patch ASAP** |
| 🟢 **GREEN** | OK | Server is within its patching schedule. No action needed |
| ⚪ **GREY** | UNSCHEDULED | Server has no priority set or no patch history. Needs to be configured |

#### Settings Sheet Colors

| Color | Meaning |
|-------|---------|
| **Blue Header** | Section title - do not edit |
| **Light Blue** | Column headers - do not edit |
| **Green cells** | Editable cells - you can change these values |
| **White cells** | Data cells or calculated values |

#### Master Servers Status

| Status | Meaning |
|--------|---------|
| **Active** | Server is in production, requires patching |
| **Decommissioned** | Server is retired, shown with strikethrough |
| **Pending** | Server is being set up, not yet in schedule |

[Back to Top](#table-of-contents)

---

## 3. Dashboard Sheet

### 3.1 Purpose

The Dashboard provides a quick overview of the entire patching status. Use this sheet to:
- See total server counts and status breakdown
- Check how many servers are overdue
- View current patching cycle settings
- Access quick reference for macros

[Back to Top](#table-of-contents)

### 3.2 Sections Explained

#### Patching Status Overview (Rows 4-8)
Shows counts for each status type:
- **Total Servers:** All servers in NextDC M1
- **OVERDUE:** Servers needing immediate attention (red)
- **OK:** Servers within schedule (green)
- **UNSCHEDULED:** Servers without priority/history (grey)

#### By Priority (Rows 11-15)
Breaks down server counts by priority level:
- How many servers in each priority
- How many are overdue per priority
- Helps identify which priority group needs most attention

#### Current Patching Cycles (Rows 17-21)
Shows the current day settings for each priority:
- Priority 1 (Critical): XX days
- Priority 2 (Medium): XX days
- Priority 3 (Low/Dev): XX days

These values come from the Settings sheet.

#### Quick Actions (Rows 23-29)
Reference list of available macros and what they do. Press `Alt+F8` to access these macros.

#### Master Server Inventory (Rows 32-36)
Shows counts from the Master Servers sheet:
- **Total Servers:** All servers in master list
- **Active:** Servers requiring patching
- **Decommissioned:** Retired servers
- **Pending:** Servers being set up

#### Priority System (Rows 38-40)
Shows number of active priority levels and indicates where to add new priorities.

[Back to Top](#table-of-contents)

### 3.3 How to Use

The Dashboard is primarily for **viewing, not editing**. All values update automatically when data changes in other sheets.

**To refresh the Dashboard:**
1. Press `Alt+F8`
2. Select "RefreshDashboard"
3. Click Run

Or simply press `F9` to recalculate all formulas.

[Back to Top](#table-of-contents)

---

## 4. Settings Sheet

### 4.1 Purpose

The Settings sheet is where you configure the core parameters that control how the patching schedule works. **Changes here affect the entire workbook.**

[Back to Top](#table-of-contents)

### 4.2 Priority Table

This table defines how often each priority level requires patching.

| Priority | Days | Description | Server Types |
|----------|------|-------------|--------------|
| 1 | 30 | Critical | Production, DC |
| 2 | 60 | Medium | Standard servers |
| 3 | 90 | Low/Dev | Dev, Test, DR |
| 4 | 120 | Extended | Archive, Backup |
| 5 | 180 | Long-term | Rarely changed |

**Column explanations:**
- **Column A:** Priority number (1, 2, 3, etc.)
- **Column B:** Days between patches (EDITABLE - green cells)
- **Column C:** Description of the priority level
- **Column D:** Typical server types for this priority

#### To Change Patching Frequency:
1. Find the priority row you want to change
2. Edit the "Days" column (green cell)
3. All servers with that priority will automatically recalculate

#### To Add a New Priority:
1. Add a new row below existing priorities (e.g., row 10)
2. Enter: Priority number, Days, Description, Server Types
3. Update the PriorityList named range if needed (Formulas > Name Manager > PriorityList)

[Back to Top](#table-of-contents)

### 4.3 Cluster List

Defines the available clusters for the dropdown in Master Servers.

**Current clusters:**
- NextDC M1 (Primary datacenter)
- Cluster 2 (Secondary)
- Cluster 3 (Development)
- Cluster 4 (DR/Backup)

#### To Add a Cluster:
1. Add cluster name in the next empty row
2. Update ClusterList named range if needed

[Back to Top](#table-of-contents)

### 4.4 Team List

Defines the team names for assignment in Master Servers.

**Current teams:**
- Team 1
- Team 2
- Team 3
- Infrastructure

#### To Add a Team:
1. Add team name in the next empty row
2. Update TeamList named range if needed

[Back to Top](#table-of-contents)

### 4.5 How to Modify Settings

> ⚠️ **IMPORTANT:** Only edit the GREEN cells. Other cells contain formulas or are used by the system.

Changes take effect immediately - all formulas recalculate automatically.

**Example: Changing Priority 1 from 30 days to 21 days:**
1. Go to Settings sheet
2. Find Priority 1 row (row 5)
3. Change cell B5 from "30" to "21"
4. All Priority 1 servers now have 21-day cycles

[Back to Top](#table-of-contents)

---

## 5. Master Servers Sheet

### 5.1 Purpose

The Master Servers sheet is the **SINGLE SOURCE OF TRUTH** for all server information. When you add, modify, or decommission a server, do it here.

All other sheets reference this data.

[Back to Top](#table-of-contents)

### 5.2 Column Descriptions

| Column | Header | Description |
|--------|--------|-------------|
| A | Server Name | The server/VM name (unique identifier) |
| B | Cluster | Which cluster the server belongs to (dropdown from Settings) |
| C | Priority | Patching priority 1-5 (dropdown) |
| D | Team | Responsible team (dropdown from Settings) |
| E | Status | Active, Decommissioned, or Pending |
| F | Notes | Any additional information |

- **Row 1-4:** Headers and instructions
- **Row 5+:** Server data

[Back to Top](#table-of-contents)

### 5.3 Adding a New Server

> **ALWAYS add new servers here FIRST**, then they appear in NextDC M1.

**Step-by-step:**
1. Go to Master Servers sheet
2. Find the first empty row (after existing servers)
3. Enter the server details:
   - **Column A:** Server name (type exactly as it appears in vSphere)
   - **Column B:** Select cluster from dropdown
   - **Column C:** Select priority (1-5) from dropdown
   - **Column D:** Select team from dropdown
   - **Column E:** Set to "Active"
   - **Column F:** Add any notes (optional)
4. Go to NextDC M1 sheet
5. In the next empty row, Column A, use dropdown to select new server
6. Priority and other fields auto-populate

[Back to Top](#table-of-contents)

### 5.4 Decommissioning a Server

When a server is retired, **don't delete it** - mark it as decommissioned. This preserves history for auditing.

**Step-by-step:**
1. Go to Master Servers sheet
2. Find the server row
3. Change Column E (Status) to "Decommissioned"
4. The server will show with strikethrough formatting
5. It remains in NextDC M1 but is visually marked as inactive

[Back to Top](#table-of-contents)

### 5.5 Changing Server Details

To change cluster, priority, or team:
1. Go to Master Servers sheet
2. Find the server
3. Use the dropdown to select new value
4. Changes automatically reflect in NextDC M1

[Back to Top](#table-of-contents)

---

## 6. NextDC M1 Sheet (Main Patching Schedule)

### 6.1 Purpose

This is your **PRIMARY WORKING SHEET** for day-to-day patching operations. Here you:
- See which servers need patching
- Record patch dates after patching
- Track patch history
- Monitor status of all servers

[Back to Top](#table-of-contents)

### 6.2 Column Descriptions

| Column | Header | Description |
|--------|--------|-------------|
| A | Server/VM Name | Server name (dropdown from Master) |
| B | Priority | Auto-filled from Master Servers |
| C | Next Scheduled Date | Calculated: Last Patch + Priority Days |
| D | Last Patch Date | Calculated: Most recent date from history |
| E | Status | Calculated: OK, OVERDUE, or UNSCHEDULED |
| F | Days Until Due | Calculated: Days until next patch due |
| G | Master Status | Shows Active/Decommissioned from Master |
| H | Notes | Manual notes field |
| I+ | Patch History | Date columns - add new dates here |

- **Columns A-G:** Mostly automatic (formulas) - don't edit directly
- **Column H:** Manual notes - you can edit
- **Columns I onwards:** PATCH HISTORY - this is where you record dates

[Back to Top](#table-of-contents)

### 6.3 Understanding Status Values

#### OVERDUE (Red background)
- The "Next Scheduled Date" is in the past
- **Action:** This server needs patching immediately
- The "Days Until Due" shows negative number (days overdue)

#### OK (Green background)
- The server was patched and next date is in the future
- **Action:** No action needed until due date approaches
- The "Days Until Due" shows positive number

#### UNSCHEDULED (Grey background)
- Either: No priority assigned, OR no patch history recorded
- **Action:** Either assign priority in Master Servers, or record first patch
- The "Days Until Due" shows "N/A"

[Back to Top](#table-of-contents)

### 6.4 Recording a Patch (Manual Method)

After patching a server:

1. Find the server's row in NextDC M1
2. Scroll right to find the first empty column after column H
3. Enter today's date in format: `DD/MM/YYYY`
4. Press Enter
5. The formulas automatically:
   - Update "Last Patch Date" to your new date
   - Recalculate "Next Scheduled Date"
   - Change "Status" to OK (if not overdue)
   - Update "Days Until Due"

> 💡 **TIP:** Use the RecordPatchDate macro instead - see [Section 7.4](#74-available-macros)

[Back to Top](#table-of-contents)

### 6.5 Using Dropdowns

#### Server Name (Column A)
- Click the cell, then click the dropdown arrow
- Select from list of servers defined in Master Servers
- Only "Active" servers appear in main list

#### Priority (Column B)
- Automatically filled based on Master Servers
- If you need to override, use dropdown (1-5)
- Better to change in Master Servers for consistency

[Back to Top](#table-of-contents)

---

## 7. Macros - Automation Tools

### 7.1 What Are Macros?

Macros are automated scripts that perform tasks with a single click. Instead of manually finding columns and entering dates, macros do it for you instantly.

**Benefits:**
- ⚡ Faster data entry
- ✅ Fewer errors
- 📋 Consistent formatting
- 📦 Bulk operations

[Back to Top](#table-of-contents)

### 7.2 One-Time Setup (Installing Macros)

Each team member needs to do this **ONCE** to enable macros.

#### Option A: Use the Macro-Enabled Workbook (Recommended)
1. Open: `Vsphere_Server_Patching_Schedule with macros.xlsm`
2. If prompted, click "Enable Macros"
3. Done! Macros are ready to use.

#### Option B: Import Macros Manually (if using .xlsx version)
1. Open the patching workbook (.xlsx)
2. Open file: `Vsphere_Patching_Macros_UPDATED.bas` with Notepad
3. Select all text (`Ctrl+A`) and copy (`Ctrl+C`)
4. In Excel, press `Alt+F11` to open VBA Editor
5. Click: Insert menu > Module
6. Paste the code (`Ctrl+V`) into the white area
7. Press `Alt+Q` to close VBA Editor
8. Save As: Excel Macro-Enabled Workbook (.xlsm)

#### Enabling Macros in Excel Settings
If macros won't run, check Excel's security settings:
1. File > Options > Trust Center > Trust Center Settings
2. Click "Macro Settings"
3. Select "Enable all macros" or "Disable with notification"
4. Click OK and restart Excel

[Back to Top](#table-of-contents)

### 7.3 How to Run Macros

#### Method 1: Clickable Buttons (Recommended)
The workbook includes clickable buttons at the top of each sheet for quick access to the most common macros.

**Dashboard Buttons:**
| Button | Action |
|--------|--------|
| Refresh Dashboard | Recalculate all data |
| Show Overdue Servers | List overdue servers (copies to clipboard) |
| Export Report | Generate printable status report |

**NextDC M1 Buttons:**
| Button | Action |
|--------|--------|
| Record Patch Date | Record today's date for selected server |
| Quick Patch Multiple | Bulk update multiple servers |
| Generate Email List | Create server list for notifications (copies to clipboard) |
| Show Overdue | List all overdue servers |
| Refresh Dashboard | Recalculate and go to Dashboard |

Simply click the button to run the macro!

#### Method 2: Macro Dialog (Alt+F8)
1. Press `Alt+F8`
2. Select the macro name from the list
3. Click "Run"

#### Method 3: Developer Tab (if enabled)
1. Click Developer tab > Macros
2. Select macro and click Run

#### Method 4: Keyboard Shortcut (if configured)
- Can assign custom shortcuts via Options

[Back to Top](#table-of-contents)

### 7.4 Available Macros

#### 7.4.1 RecordPatchDate

**PURPOSE:** Record today's date as a patch for ONE server

**HOW TO USE:**
1. Go to NextDC M1 sheet
2. Click anywhere on the row of the server you just patched
3. Press `Alt+F8` > select "RecordPatchDate" > Run
4. A confirmation dialog appears showing the server name
5. Click "Yes" to confirm
6. The date is recorded in the next available history column

**WHAT IT DOES:**
- Finds the selected server row
- Finds the first empty column in patch history (column I onwards)
- Enters today's date with DD/MM/YYYY format
- Displays confirmation message

**WHEN TO USE:**
- After patching a single server
- Fastest way to record one patch

[Back to Top](#table-of-contents)

---

#### 7.4.2 QuickPatchMultiple

**PURPOSE:** Record today's date for MULTIPLE servers at once

**HOW TO USE:**
1. Go to NextDC M1 sheet
2. Select multiple rows:
   - Hold `Ctrl` and click each row, OR
   - Click first row, hold `Shift`, click last row (for range)
3. Press `Alt+F8` > select "QuickPatchMultiple" > Run
4. Review the list of servers in the confirmation dialog
5. Click "Yes" to confirm
6. All selected servers get today's date recorded

**WHAT IT DOES:**
- Loops through all selected rows
- For each valid server, records today's date
- Shows count of servers updated

**WHEN TO USE:**
- After a bulk patching session
- When you've patched multiple servers in one maintenance window

[Back to Top](#table-of-contents)

---

#### 7.4.3 ShowOverdueServers

**PURPOSE:** Display a list of all servers that are currently OVERDUE

**HOW TO USE:**
1. Press `Alt+F8` > select "ShowOverdueServers" > Run
2. A message box appears listing all overdue servers
3. Shows server name, due date, and days overdue
4. Click OK - **text is automatically copied to clipboard**
5. Press `Ctrl+V` to paste the list anywhere (email, document, etc.)

**WHAT IT DOES:**
- Scans all servers in NextDC M1
- Finds those with "OVERDUE" status
- Displays formatted list with details
- **AUTO-COPIES text to clipboard** for easy pasting

**WHEN TO USE:**
- At the start of your day to plan patching
- Before maintenance windows
- Weekly status checks

[Back to Top](#table-of-contents)

---

#### 7.4.4 GenerateEmailList

**PURPOSE:** Create a list of servers for email notifications

**HOW TO USE:**
1. Go to NextDC M1 sheet
2. Select the rows of servers you're about to patch
3. Press `Alt+F8` > select "GenerateEmailList" > Run
4. A message box appears with the server list
5. Click OK - **text is automatically copied to clipboard**
6. Press `Ctrl+V` to paste directly into your email

**WHAT IT DOES:**
- Lists all selected servers
- Shows associated teams (from Master Servers)
- Formats for easy copying
- **AUTO-COPIES text to clipboard** for easy pasting

**WHEN TO USE:**
- Before sending maintenance notifications
- When preparing change requests

[Back to Top](#table-of-contents)

---

#### 7.4.5 ExportPatchReport

**PURPOSE:** Generate a printable status report

**HOW TO USE:**
1. Press `Alt+F8` > select "ExportPatchReport" > Run
2. Click "Yes" to confirm
3. A new "Patch Report" sheet is created
4. The report is formatted and ready to print

**WHAT IT DOES:**
- Creates a new sheet called "Patch Report"
- Lists all servers with current status
- Color-codes by status
- Includes summary statistics
- Adds generation timestamp

**WHEN TO USE:**
- Weekly/monthly reporting
- Management updates
- Compliance audits
- Printing for meetings

[Back to Top](#table-of-contents)

---

#### 7.4.6 RefreshDashboard

**PURPOSE:** Recalculate all formulas and show Dashboard

**HOW TO USE:**
1. Press `Alt+F8` > select "RefreshDashboard" > Run
2. Dashboard sheet activates with updated values

**WHAT IT DOES:**
- Forces recalculation of all formulas
- Switches to Dashboard sheet

**WHEN TO USE:**
- If values seem stale or incorrect
- After bulk changes
- Before generating reports

[Back to Top](#table-of-contents)

---

## 8. Common Tasks - Step by Step

### 8.1 Recording a Single Patch

**Scenario:** You just finished patching server "APPSERVER01"

#### Using Macro (Recommended):
1. Open the workbook
2. Go to NextDC M1 sheet
3. Click on any cell in APPSERVER01's row
4. Press `Alt+F8`
5. Select "RecordPatchDate"
6. Click Run
7. Click Yes to confirm
8. ✅ Done!

#### Manual Method:
1. Go to NextDC M1 sheet
2. Find APPSERVER01's row
3. Scroll right to first empty column after H
4. Type today's date: `09/01/2026`
5. Press Enter
6. ✅ Done!

[Back to Top](#table-of-contents)

### 8.2 Recording Multiple Patches (Bulk)

**Scenario:** You patched 10 servers in tonight's maintenance window

1. Go to NextDC M1 sheet
2. Find the first server you patched
3. Hold `Ctrl` and click on each server's row (column A works well)
4. After selecting all 10, press `Alt+F8`
5. Select "QuickPatchMultiple"
6. Click Run
7. Verify the list shows your 10 servers
8. Click Yes
9. ✅ Done! All 10 servers updated.

[Back to Top](#table-of-contents)

### 8.3 Adding a New Server to the System

**Scenario:** A new server "NEWAPP01" needs to be added for patching

#### Step 1 - Add to Master Servers:
1. Go to Master Servers sheet
2. Scroll to first empty row
3. Enter:
   - **A:** NEWAPP01
   - **B:** (select cluster from dropdown)
   - **C:** (select priority from dropdown)
   - **D:** (select team from dropdown)
   - **E:** Active
   - **F:** (optional notes)

#### Step 2 - Add to Patching Schedule:
1. Go to NextDC M1 sheet
2. Scroll to first empty row
3. In column A, click dropdown and select "NEWAPP01"
4. Priority and other fields auto-populate
5. Status will show "UNSCHEDULED" until first patch

#### Step 3 - Record Initial Patch (if already patched):
1. Click on the NEWAPP01 row
2. Press `Alt+F8` > RecordPatchDate > Run
3. Status changes to "OK"

[Back to Top](#table-of-contents)

### 8.4 Checking Which Servers Are Overdue

#### Method 1 - Dashboard:
1. Go to Dashboard sheet
2. Look at "OVERDUE" count in Status Overview
3. Check "By Priority" section to see breakdown

#### Method 2 - Macro:
1. Press `Alt+F8`
2. Select "ShowOverdueServers"
3. Click Run
4. View detailed list with days overdue
5. Press `Ctrl+V` to paste the list anywhere

#### Method 3 - Filter in NextDC M1:
1. Go to NextDC M1 sheet
2. Click column E header (Status)
3. Click Filter icon (or Data > Filter)
4. Uncheck all except "OVERDUE"
5. Only overdue servers shown

[Back to Top](#table-of-contents)

### 8.5 Changing a Server's Priority

**Scenario:** DBSERVER01 needs to change from Priority 3 to Priority 1

1. Go to Master Servers sheet
2. Find DBSERVER01 row
3. Click Priority cell (column C)
4. Select "1" from dropdown
5. Go to NextDC M1 sheet
6. DBSERVER01 now shows Priority 1
7. Next Scheduled Date automatically recalculates

[Back to Top](#table-of-contents)

### 8.6 Generating a Report

**Scenario:** Weekly report needed for management meeting

1. Press `Alt+F8`
2. Select "ExportPatchReport"
3. Click Run
4. Click Yes to confirm
5. New "Patch Report" sheet appears
6. Review the report
7. To print: File > Print
8. To save as PDF: File > Save As > PDF format

[Back to Top](#table-of-contents)

### 8.7 Sending Patch Notifications

**Scenario:** You're about to patch 5 servers and need to notify teams

1. Go to NextDC M1 sheet
2. Select the 5 server rows (`Ctrl+Click` each)
3. Press `Alt+F8`
4. Select "GenerateEmailList"
5. Click Run
6. Review the server list in the message box
7. Click OK (text is automatically copied to clipboard)
8. Open your email and press `Ctrl+V` to paste the list
9. Add date/time of maintenance
10. Send to appropriate teams

[Back to Top](#table-of-contents)

---

## 9. Troubleshooting

### 9.1 Common Issues and Solutions

#### "Macros are disabled" warning
**Solution:**
- File > Options > Trust Center > Trust Center Settings
- Macro Settings > Enable all macros
- Restart Excel

#### Macro list shows empty when pressing Alt+F8
**Solution:**
- Macros haven't been imported yet
- Follow [Section 7.2](#72-one-time-setup-installing-macros) to import macros
- Or use the .xlsm version of the file

#### Can't save the file (file is locked)
**Solution:**
- Another user has the file open
- Ask them to close it, OR
- Use OneDrive's co-authoring feature

#### Dropdown shows wrong values (like clusters in priority)
**Solution:**
- Named range is pointing to wrong cells
- Go to Formulas > Name Manager
- Check and correct the named range

[Back to Top](#table-of-contents)

### 9.2 Formula Errors

#### #REF! ERROR
- A reference is broken (deleted row/column)
- Check if named ranges still exist in Name Manager
- May need to recreate the named range

#### #VALUE! ERROR
- Wrong data type in a cell
- Check that dates are formatted as dates
- Check that priority is a number

#### #N/A ERROR
- VLOOKUP can't find the value
- Check server name matches exactly in Master Servers
- Check for extra spaces in names

#### CIRCULAR REFERENCE WARNING
- A formula refers to itself
- Don't put formulas in patch history columns (I onwards)
- Those columns should only contain dates

[Back to Top](#table-of-contents)

### 9.3 Dropdown Issues

#### Dropdown Shows Wrong Values
1. Go to Formulas > Name Manager
2. Find the relevant named range (PriorityList, ClusterList, etc.)
3. Click Edit
4. Verify it points to the correct cells
5. Adjust if needed

#### Dropdown Doesn't Appear
1. Click the cell
2. Go to Data > Data Validation
3. Check if validation is applied
4. If not, add validation with Source: `=PriorityList` (or relevant name)

[Back to Top](#table-of-contents)

### 9.4 Macro Issues

#### Macro gives error "Subscript out of range"
- Sheet name doesn't match what macro expects
- Check that sheets are named: "NextDC M1", "Master Servers", "Settings"

#### Macro doesn't record date
- Make sure you're on NextDC M1 sheet
- Make sure you've selected a valid server row (row 2+)
- Check that the server has a name in column A

#### Macro runs but nothing happens
- Enable macros in Trust Center
- Check that you selected a row before running
- Try running RefreshDashboard to ensure macros work

[Back to Top](#table-of-contents)

---

## 10. Appendix

### 10.1 Named Ranges Reference

| Name | Location | Purpose |
|------|----------|---------|
| PriorityList | Settings!$A$5:$A$9 | Priority dropdown values |
| PriorityTable | Settings!$A$5:$D$9 | Priority lookup table |
| ClusterList | Settings!$A$15:$A$18 | Cluster dropdown values |
| TeamList | Settings!$A$23:$A$26 | Team dropdown values |
| MasterServerList | Master Servers!$A$5:$A$200 | Server name dropdown |

**To view/edit:** Formulas > Name Manager

[Back to Top](#table-of-contents)

### 10.2 Formula Reference

#### Next Scheduled Date (Column C)
```
=IFERROR(VLOOKUP(B2,PriorityTable,2,FALSE)+D2,"No date")
```
*Meaning: Look up priority days, add to last patch date*

#### Last Patch Date (Column D)
```
=IF(COUNT(I2:ZZ2)=0,"No Date",MAX(I2:ZZ2))
```
*Meaning: Find the most recent date in patch history*

#### Status (Column E)
```
=IFERROR(IF(OR(B2="",C2="No date"),"UNSCHEDULED",
  IF(C2<TODAY(),"OVERDUE","OK")),"UNSCHEDULED")
```
*Meaning: Calculate status based on due date vs today*

#### Days Until Due (Column F)
```
=IF(AND(ISNUMBER(C2),C2<>0),C2-TODAY(),"N/A")
```
*Meaning: Subtract today from due date*

[Back to Top](#table-of-contents)

### 10.3 Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Alt+F8` | Open Macro dialog |
| `F9` | Recalculate all formulas |
| `Alt+F11` | Open VBA Editor |
| `Ctrl+Home` | Go to cell A1 |
| `Ctrl+End` | Go to last used cell |
| `Ctrl+Down` | Jump to end of data in column |
| `Ctrl+Shift+L` | Toggle filters |
| `Ctrl+S` | Save workbook |
| `Ctrl+P` | Print |

[Back to Top](#table-of-contents)

---

## Document Information

| | |
|---|---|
| **Version** | 2.1 |
| **Last Updated** | January 2026 |
| **Author** | Infrastructure Team |

For questions or updates to this guide, contact your team lead.

---

[Back to Top](#table-of-contents)

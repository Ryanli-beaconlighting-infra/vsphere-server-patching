# vSphere Server Patching Schedule

**Last Updated:** February 2026
**Author:** Ryan Li / Infrastructure Team
**Classification:** Internal Use - Beacon Lighting

A comprehensive Excel-based system for tracking and managing server patching schedules across vSphere infrastructure.

## Overview

This workbook helps infrastructure teams:
- Track which servers need patching and when
- Record patch history for compliance and auditing
- Identify overdue servers at a glance
- Maintain consistent patching cycles based on server priority
- Generate reports for management and compliance

## Features

- **Automatic status calculation** - OK, OVERDUE, or UNSCHEDULED
- **Priority-based scheduling** - Configurable days per priority level (30/60/90/120/180 days)
- **Master server list** - Single source of truth for all server information
- **Dropdown menus** - Consistent data entry with validation
- **Clickable macro buttons** - One-click access to common actions
- **VBA Macros** - Automation for quick updates and reporting
- **Dashboard** - At-a-glance status overview

## Files

| File | Description |
|------|-------------|
| `Vsphere_Server_Patching_Schedule with macros.xlsm` | **Primary file** - Macro-enabled workbook with VBA code embedded. Use this for day-to-day operations. |
| `Vsphere_Patching_Macros_UPDATED.bas` | VBA source code for macro import. Use this to re-import macros into a fresh .xlsx or to review code changes. |
| [`Vsphere_Patching_Complete_User_Guide.md`](Vsphere_Patching_Complete_User_Guide.md) | Comprehensive documentation with clickable navigation |
| `Vsphere_Patching_Quick_Reference_Card.txt` | Printable desk reference |
| [`README.md`](README.md) | This file |

### xlsm vs xlsx

- **`.xlsm` (macro-enabled)**: Contains embedded VBA macros. This is the version you should open for normal use. Macros run directly without any setup.
- **`.xlsx` (standard)**: Does not support macros. If you only have an `.xlsx` file, you must manually import the VBA code from `Vsphere_Patching_Macros_UPDATED.bas` via the VBA Editor (`Alt+F11` > Insert > Module > paste code), then Save As `.xlsm`.

### The `old/` Directory

The `old/` directory contains previous versions of the workbook and macros kept for reference and rollback purposes. These files are **not** actively used:

| File | Purpose |
|------|---------|
| `Vsphere Server Patching schedule (TBR).xlsx` | Original workbook before enhancements |
| `PatchingMacros_VBA_Import.bas` | Earlier version of the VBA macros |
| `PatchingMacros_CLEAN.txt` | Plain-text copy of an older macro version |
| `Vsphere_Server_Patching_Schedule_Enhanced_BACKUP.xlsx` | Backup snapshot of the enhanced workbook |
| `Vsphere_Server_Patching_Schedule_Enhanced with macros.xlsm` | Previous enhanced version with macros |
| `OLd_ Vsphere_Server_Patching_Schedule with macros.xlsx/.xlsm` | Legacy versions |

Do not delete these files without confirming no one relies on them.

## Quick Start

1. Open `Vsphere_Server_Patching_Schedule with macros.xlsm`
2. Enable macros when prompted
3. Go to **NextDC M1** sheet to view/update patching schedule
4. Use the **clickable buttons** at the top of each sheet, or press `Alt+F8` for all macros

## Macro Descriptions

The workbook includes the following VBA macros (defined in `Vsphere_Patching_Macros_UPDATED.bas`):

| Macro | Purpose | Sheet Required |
|-------|---------|----------------|
| `RecordPatchDate` | Record today's date as a patch for the selected server row. Creates a backup sheet before writing. | NextDC M1 |
| `QuickPatchMultiple` | Bulk-record today's date for multiple selected server rows at once. Creates a backup sheet before writing. | NextDC M1 |
| `ShowOverdueServers` | Scan all servers and display a list of those with OVERDUE status. Auto-copies the list to clipboard. | Any sheet |
| `GenerateEmailList` | Build a formatted server list (with team names from Master Servers) for pasting into emails. Auto-copies to clipboard. | NextDC M1 |
| `ExportPatchReport` | Generate a new "Patch Report" sheet with colour-coded status, formatted for printing. | Any sheet |
| `RefreshDashboard` | Force-recalculate all formulas and navigate to the Dashboard sheet. | Any sheet |

### Macro Safety Features (February 2026 Update)

- **Structured error handling**: Every Sub/Function has `On Error GoTo ErrorHandler` with descriptive error messages.
- **BackupSheet**: `RecordPatchDate` and `QuickPatchMultiple` automatically create a timestamped backup of the sheet before modifying data.
- **Dynamic row scanning**: Row limits use `ws.Cells(ws.Rows.Count, 1).End(xlUp).Row` instead of hardcoded values, so the macros scale with any number of servers.
- **Increased column limit**: The patch-history column safety limit is set to 500 (up from 100) with a warning at 80% capacity (400 columns).
- **Sheet name constant**: All sheet references use `Private Const TARGET_SHEET_NAME` so the target sheet can be renamed in one place.

## Macro Buttons

The workbook includes clickable buttons for quick access to common macros:

**Dashboard Sheet:**
| Button | Action |
|--------|--------|
| Refresh Dashboard | Recalculate all data |
| Show Overdue Servers | List overdue servers (copies to clipboard) |
| Export Report | Generate printable status report |

**NextDC M1 Sheet:**
| Button | Action |
|--------|--------|
| Record Patch Date | Record today's date for selected server |
| Quick Patch Multiple | Bulk update multiple servers |
| Generate Email List | Create server list for notifications (copies to clipboard) |
| Show Overdue | List all overdue servers |
| Refresh Dashboard | Recalculate and go to Dashboard |

## Sheet Structure

- **Dashboard** - Overview of patching status and quick stats
- **Settings** - Configure priorities, clusters, and teams
- **Master Servers** - Central list of all servers (source of truth)
- **NextDC M1** - Main patching schedule and history tracking

## Status Colors

| Status | Color | Meaning |
|--------|-------|---------|
| OVERDUE | Red | Past scheduled date - patch immediately |
| OK | Green | Within schedule - no action needed |
| UNSCHEDULED | Grey | No priority set or no patch history |

## Documentation

See [`Vsphere_Patching_Complete_User_Guide.md`](Vsphere_Patching_Complete_User_Guide.md) for full documentation including:
- Detailed sheet explanations
- Step-by-step task guides
- Macro setup and usage
- Troubleshooting guide

## Requirements

- Microsoft Excel 2016 or later (desktop app recommended)
- Macros enabled for automation features

## Excel Online (Web Browser) Compatibility

| Feature | Works in Browser? |
|---------|-------------------|
| Viewing data | Yes |
| Editing cells | Yes |
| Formulas & calculations | Yes |
| Dropdowns | Yes |
| Status colors | Yes |
| **VBA Macros** | No - requires desktop Excel |

For full functionality including macros, use the desktop Excel application.

## License

Internal use - Beacon Lighting Infrastructure Team

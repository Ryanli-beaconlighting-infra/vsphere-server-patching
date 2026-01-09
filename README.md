# vSphere Server Patching Schedule

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
- **VBA Macros** - Automation for quick updates and reporting
- **Dashboard** - At-a-glance status overview

## Files

| File | Description |
|------|-------------|
| `Vsphere_Server_Patching_Schedule.xlsx` | Main workbook |
| `Vsphere_Server_Patching_Schedule with macros.xlsm` | Macro-enabled version |
| `Vsphere_Patching_Complete_User_Guide.txt` | Comprehensive documentation |
| `Vsphere_Patching_Quick_Reference_Card.txt` | Printable desk reference |
| `Vsphere_Patching_Macros_UPDATED.bas` | VBA code for macro import |

## Quick Start

1. Open `Vsphere_Server_Patching_Schedule with macros.xlsm`
2. Enable macros when prompted
3. Go to **NextDC M1** sheet to view/update patching schedule
4. Press `Alt+F8` to access automation macros

## Available Macros

| Macro | Purpose |
|-------|---------|
| `RecordPatchDate` | Record today's date for selected server |
| `QuickPatchMultiple` | Bulk update multiple servers at once |
| `ShowOverdueServers` | List all servers needing patching |
| `GenerateEmailList` | Create server list for notifications |
| `ExportPatchReport` | Generate printable status report |
| `RefreshDashboard` | Recalculate all data |

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

See `Vsphere_Patching_Complete_User_Guide.txt` for full documentation including:
- Detailed sheet explanations
- Step-by-step task guides
- Macro setup and usage
- Troubleshooting guide

## Requirements

- Microsoft Excel 2016 or later
- Macros enabled for automation features

## License

Internal use - Beacon Lighting Infrastructure Team

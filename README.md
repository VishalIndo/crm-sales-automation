# Excel VBA CRM App

This repo now contains a working first-pass CRM app for Excel/VBA. The app is sheet-driven so it can run immediately inside a macro-enabled workbook without depending on custom add-ins or external libraries.

## What is included

- `vba_crm/modules/modConstants.bas`
- `vba_crm/modules/modUtils.bas`
- `vba_crm/modules/modSchema.bas`
- `vba_crm/modules/modCRM.bas`
- `vba_crm/modules/modDashboard.bas`
- `main.py` helper commands for seed CSVs and setup guidance

## App scope

The VBA app supports:

- lead database
- lead search
- follow-up reminders
- activity history logging
- quotation tracking
- stage progression
- dashboard metrics
- simple home screen buttons

## Setup

1. Create a new Excel macro-enabled workbook: `.xlsm`
2. Open the VBA editor with `Alt+F11`
3. Import every `.bas` file from `vba_crm/modules`
4. Run `SetupCRMApp`
5. Save, close, and reopen the workbook

## Python helper

The Python helper is optional.

```bash
python main.py guide
python main.py seed
python main.py manifest
```

`seed` creates sample CSV files in `generated/` for `Leads`, `Activities`, and `Quotations`.

## Workbook layout

`SetupCRMApp` creates:

- `Home`
- `Leads`
- `Activities`
- `Quotations`
- `Dashboard`
- `Lists`
- `Settings`

## Main macros

- `SetupCRMApp`
- `AddLead`
- `SearchLead`
- `AddActivityForSelectedLead`
- `AddQuotationForSelectedLead`
- `AdvanceSelectedLeadStage`
- `RefreshDashboard`
- `ShowTodayFollowUps`



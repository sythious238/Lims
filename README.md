# EnviroLIMS — Environmental LIMS for Microsoft Access

EnviroLIMS is a Laboratory Information Management System for environmental testing labs (drinking water, groundwater, wastewater, soil). It runs in Microsoft Access. Everything in it (tables, relationships, reference data, queries, forms and reports) is generated from the VBA source in `src/`, so the database can be rebuilt and kept under version control.

## Features

| Area | What it does |
|---|---|
| **Sample receiving** | Work orders with auto-numbering (`WO2611-001`), cooler temperature, ice and custody-seal checks, and a grid for logging in samples. Lab sample IDs are assigned automatically (`WO2611-001-01`, `-02`, …). |
| **Chain of custody** | Custody events for each work order (relinquished, received, storage, disposal) and a printable custody report. |
| **Test assignment** | Add a method to every sample on a work order in one click. Only samples whose matrix class (Aqueous/Solid) fits the method get it. Result rows are created from the method's analyte list, with that method's MDL, RL and units. |
| **Result entry** | One test at a time, with a results grid. **U**/**J** qualifiers are set automatically from the MDL and RL. The current regulatory limit appears next to each result, and exceedances are flagged in red. |
| **Holding times** | Holding-time expiry is calculated from the collection time and the method's hold time. Each test is flagged `OK`, `DUE <24H`, `EXPIRED`, `MET` or `OUT OF HOLD`. |
| **Review** | "Mark Reviewed" signs off a test (who and when are stamped on each result) and rolls the status up to the work order. Reviews are written to an audit log. |
| **QC batches** | Method blanks, LCS/LCSD, MS/MSD, duplicates and ICV/CCV. % recovery, RPD and PASS/FAIL against the method's acceptance limits are calculated for you. "Add QC Rows" fills in each analyte and its limits. |
| **Regulatory limits** | Limits are grouped into named limit sets (for example `EPA Primary MCL`). A project picks a limit set, and every result is compared against it. US EPA drinking-water MCLs and SMCLs are preloaded. |
| **Dashboard** | Live counts of open tests, overdue tests, tests at holding-time risk, exceedances, QC failures and instrument calibrations due. Double-click a count to open the matching list. |
| **Reports** | Certificate of Analysis, Exceedance Report, Holding Time/Backlog, QC Summary, Chain of Custody, Work Order Status. |

Preloaded methods: EPA 200.8, 245.1, 300.0, 524.2, 6010D, 8260D; SM 4500-H+ B, 2540 D, 5210 B, 9223 B.

## Building the database

**You need:** Windows and Microsoft Access (Microsoft 365, 2016 or later, 32- or 64-bit).

### Option A: one click (recommended)
1. Download or clone this repository.
2. Double-click **`Build-EnviroLIMS.cmd`**.
3. Access opens and builds `EnviroLIMS.accdb` in the same folder, then shows the main menu.

To build without the demo data, run `Build-EnviroLIMS.cmd -NoDemoData`.

### Option B: manually
1. In Access, create a new **Blank database**.
2. Press **Alt+F11** to open the VBA editor. Choose **File → Import File…** and import all four files from `src/`.
3. Press **Ctrl+G** to open the Immediate window. Type `BuildLIMS` and press Enter. (`BuildLIMS False` skips the demo data.)

When you open the database later, click **Enable Content**, or put the file in a Trusted Location. The forms use VBA functions and won't work with content disabled.

## First-time setup
1. **Setup → Settings:** enter your lab name, address, phone, accreditation number and director. These print on every report.
2. **Setup → Methods & Limits:** ⚠️ **replace the example MDL, RL and QC acceptance values** with the values from your own MDL studies and SOPs.
3. **Setup → Regulatory Limits:** check the preloaded EPA limits. Add limit sets for your state, permits or soil cleanup levels.
4. **Setup → Analysts / Instruments / Clients / Projects & Locations:** enter your own records. You can delete the demo client.

## Daily workflow
1. **Work Orders:** create a work order, choose the client and project, and record the cooler details.
2. Enter samples in the grid (client ID, location, matrix, collection time), then click **Assign Lab IDs**.
3. Choose a method in **Add method to all matching samples** and click **Add Method**. Repeat for each method. Tests and blank result rows are created.
4. **Result Entry:** enter results. Qualifiers, analysis dates and test status update as you type.
5. **QC Batches:** create a batch, add QC rows, enter QC results and assign tests to the batch.
6. Click **Mark Reviewed** on each test, then **Print COA** from the work order.

## Data model

```
tblClients ─┬─< tblProjects ─< tblLocations
            │        │               │
            └─< tblWorkOrders >──────┘ (project)
                     │
                     ├─< tblCustodyLog
                     └─< tblSamples >── tblMatrices ──< tblRegulatoryLimits >── tblAnalytes
                              │                                                     │
                              └─< tblSampleTests >── tblMethods ─< tblMethodAnalytes >┘
                                       │      └── tblQCBatches ─< tblQCResults
                                       └─< tblResults >── tblAnalytes
```

| Table | Purpose |
|---|---|
| `tblClients`, `tblProjects`, `tblLocations` | Who the samples are for, which project and limit set, and the sampling points |
| `tblWorkOrders`, `tblSamples`, `tblCustodyLog` | Receiving, sample login and custody |
| `tblMethods`, `tblMethodAnalytes` | Methods with hold time, preservation and turnaround. Each analyte has MDL, RL, %recovery limits and RPD limit |
| `tblSampleTests`, `tblResults` | Tests on each sample and the result for each analyte |
| `tblQCBatches`, `tblQCResults` | Analytical batches and their QC samples |
| `tblMatrices`, `tblAnalytes`, `tblRegulatoryLimits` | Reference data |
| `tblAnalysts`, `tblInstruments`, `tblSettings`, `tblAuditLog` | Staff, equipment, configuration and audit trail |

Deleting a work order cascades to its samples, tests, results and custody records. A sample that a QC result names as its parent sample can't be deleted.

Key queries: `qryResultsWithLimits` (each result with its limit and exceedance flag), `qryHoldTimeStatus`, `qryOpenTests`, `qryOverdueTests`, `qryQCEvaluation`, `qryExceedances`, `qryWorkOrderProgress`, `qryLocationTrends`.

## Customizing
- **Change forms or reports:** edit `src/modBuildUI.bas`, import it again, and run `RebuildUI` in the Immediate window. `RebuildUI` regenerates queries, forms and reports and leaves your data alone. Changes made by hand in Design view are overwritten by the next `RebuildUI`.
- **Business rules** such as qualifiers, QC pass/fail, numbering and status roll-up are in `src/modLIMS.bas`.
- **Multi-user:** use Access's **Database Tools → Access Database** splitter to put the tables in a shared back-end file, and give each user a copy of the front end.
- `BuildLIMS True, True` drops **all** EnviroLIMS tables and rebuilds from scratch. Without the second `True`, BuildLIMS refuses to run on a database that already has data.

## Limitations
- No user log-in or role-based security. The audit log records the Windows user name.
- On its own this does not meet ISO/IEC 17025, TNI/NELAP or 21 CFR Part 11 requirements. Your QA program has to validate it and control access to it.
- Units are not converted. Make sure each limit's units match the method's reporting units.
- Results with a dilution factor are entered as final (already corrected) values. Scale the MDL and RL on the result row if your SOP requires it.

## Repository layout
```
Build-EnviroLIMS.cmd / .ps1   One-click builder (Windows + Access)
src/modBuild.bas              Entry points, tables, indexes, relationships, queries
src/modBuildUI.bas            Form and report generator
src/modSeed.bas               Reference data (methods, analytes, EPA limits) and demo data
src/modLIMS.bas               Runtime logic called by the forms
```

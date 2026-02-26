# CleanGuard

Production department labor protection and locker management system built with .NET Framework 4.6.2 + WinForms + SQLite. Designed for factory locker room scenarios, providing end-to-end management of employees, lockers, and labor protection items.

## Features

### Employee Management (FrmMain)

| Action | Description |
|--------|-------------|
| Add Employee | Enter employee number, name, process; assign 1F/2F clothing and shoe lockers |
| Edit Employee | Modify info and lockers; old lockers auto-released, new lockers auto-occupied |
| Resign | Release all lockers and mark resigned (high-risk, requires confirmation) |
| Restore | Restore resigned employee to active, reassign lockers |
| Delete | Permanently delete and release lockers (high-risk, irreversible) |
| Search | Fuzzy search by name / pinyin initials / process |

### Labor Protection Items (FrmEditor)

- Dynamic add/remove rows per category (cleanroom suit, safety shoe, canvas shoe, clean cap)
- Enter size, code (suit/cap), new/old status (shoes), issue date
- Batch paste, CSV template download and import with pre-check validation
- Category limits configurable online (persisted in T_SystemConfig)

### Label Printing (Printer)

- Print preview and direct printing modes, batch multi-person printing
- Label content: name/ID/process + four locker blocks + QR code (QRCoder optional, text fallback)
- Print preset persistence (default printer, paper, margins, orientation, label size)
- Batch printing with retry/skip/abort error handling and missing-field warnings

### Data Import (FrmImport + ImportHelper)

- CSV / XLSX bulk employee import
- XLSX template preferred; auto-fallback to CSV when SharpZipLib is missing
- Import results: success/failure counts, failure detail preview (first 20), one-click error copy
- Export backfill template: failed rows exported as-is for correction and re-import
- Error code system (CG-IMP-001 ~ CG-IMP-012): precise to row and column with fix suggestions

### Locker Management

| Module | Description |
|--------|-------------|
| Locker Maintenance (FrmLockerManage) | Filter by floor/type/anomaly, maintain anomaly remarks, bulk import lockers |
| Locker Chart (FrmLockerChart) | Occupancy pie chart + area heatmap blocks (groups of 10) + summary, PNG export |
| Occupancy Trend (FrmLockerTrend) | Line chart of four locker types based on T_LockerSnapshot |

### Process Dictionary (FrmProcessManage)

- Add, delete, rename (auto-sync assigned employees)
- CSV/TXT bulk import
- Audit view (FrmProcessAudit): filter by operation type/date/keyword, row color coding, CSV export

### System Log (FrmSystemLog)

- Filter by type (Import / Backup / Print / Employee) and CSV export
- All key operations automatically logged

### Operations and Fault Tolerance

- **Auto Backup**: Backup CleanGuard.db to Backup/ on exit, auto-clean backups older than 7 days
- **Global Exception Handling**: ThreadException + UnhandledException + key-path try-catch
- **Database Auto Migration**: CREATE TABLE IF NOT EXISTS + EnsureColumnExists

## Tech Stack

| Item | Description |
|------|-------------|
| Framework | .NET Framework 4.6.2, Windows Forms |
| Database | SQLite (System.Data.SQLite 1.0.119.0) |
| Excel | NPOI 2.5.6 |
| Compression | SharpZipLib 1.4.2 (NPOI dependency, CSV fallback when missing) |
| QR Code | QRCoder (optional, reflection-loaded) |
| Charts | System.Windows.Forms.DataVisualization |
| Pinyin | Built-in GB2312 zone code first-letter (PinYin.cs) |
| Build Target | x86 (compatible with ARM64 Windows) |

## Database

| Table | Purpose |
|-------|---------|
| T_Employee | Employee master (ID, name, pinyin, process, four lockers, status) |
| T_Emp_Items | Labor items sub-table (category, slot, size, code, new/old, issue date) |
| T_Lockers | Locker master data (locker ID, floor, type, occupancy, remark) |
| T_Process | Process dictionary |
| T_SystemLog | Operation log |
| T_SystemConfig | Key-value config (item limits, print presets, etc.) |
| T_LockerSnapshot | Locker occupancy snapshots (trend chart data source) |

Default initialization: 60 clothing + 60 shoe lockers per floor (1F/2F), 6 preset processes.

## Project Structure

```text
Src/CleanGuard_App/
  Program.cs                  # Entry point
  Forms/
    FrmMain.cs                # Main form: employee list, search, action buttons
    FrmEditor.cs              # Employee editor: basic info + lockers + items tabs
    FrmImport.cs              # Import wizard: template download, import, failure preview
    FrmLockerChart.cs         # Locker chart: pie + heatmap + summary
    FrmLockerTrend.cs         # Locker trend: line chart
    FrmLockerManage.cs        # Locker maintenance: filter, remark, bulk import
    FrmProcessManage.cs       # Process dictionary: CRUD + bulk import
    FrmProcessAudit.cs        # Process audit: filter, color coding, export
    FrmSystemLog.cs           # System log
  Utils/
    SQLiteHelper.cs           # Data access layer
    ImportHelper.cs           # CSV/XLSX import/export engine
    Printer.cs                # Label printing engine
    PinYin.cs                 # Chinese pinyin first-letter
    UiTheme.cs                # UI theme
Src/Database/
  init.sql                    # Reference schema script
Output/                       # Build output
```

## Build and Run

### Requirements

- Windows 7+ (.NET Framework 4.6.2)
- Visual Studio 2017+ / MSBuild 15.0+

### Build

```bash
nuget restore Src/CleanGuard_App/packages.config -PackagesDirectory packages
msbuild Src/CleanGuard_App/CleanGuard_App.csproj /t:Build /p:Configuration=Debug
```

Build output goes to `Output/` directory.

### Run

Run `Output/CleanGuard_App.exe` directly. On first launch it auto-creates `CleanGuard.db` and initializes all tables with default data.

## Startup Flow

```text
Program.Main()
  |-- Register global exception handlers (ThreadException + UnhandledException)
  |-- SQLiteHelper.InitializeDatabase()
  |     |-- CREATE TABLE IF NOT EXISTS (7 tables)
  |     |-- EnsureColumnExists (forward compat)
  |     |-- SeedDefaultConfig
  |     +-- CaptureLockerSnapshot("Startup")
  |-- new FrmMain() -> InitializeLayout() -> LoadEmployeeData()
  |-- Application.Run(mainForm)
  +-- FormClosing -> BackupDatabase()
```

## License

[MIT](LICENSE)

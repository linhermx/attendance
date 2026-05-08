# Changelog

All notable changes to this project will be documented in this file.

This project follows **Semantic Versioning** (`MAJOR.MINOR.PATCH`).

---

## [v1.2.0] - Range Reporting & Historical View

### Added
- Attendance analysis in `Rango` mode using the same personnel database and a multi-day export from ZKAccess
- Historical Excel view with employees in horizontal layout and daily history in vertical blocks
- Period-level summaries for alerts, overtime, and consolidated attendance detail

### Implemented
- GUI separation between `Diario` and `Rango` workflows
- Date-range expansion from the minimum to maximum date in the export, limited to labor days (`Monday` to `Saturday`)
- Detection of partial cutoff on the last day of the range
- Freeze panes in the historical weekly view to keep worker identity visible during vertical review
- Clear validation messages when the user loads `Personal` and `Eventos` files in the wrong fields
- Removal of `run_log.txt` export in both daily and range modes

### Business Rules
- All valid workers appear in the range report, even if they have no punches on a given labor day
- Labor days without global records are treated as `Sin operacion`, not as mass absences
- Historical overtime is reported both as employee summary and per-day detail
- Three-punch inference now prioritizes plausible slot windows to avoid misclassifying very late entries as lunch punches

### Notes
- This release extends the system from single-day control to weekly historical review without changing the daily workflow
- The visual style and status colors remain aligned with the existing daily attendance report

---

## [v1.1.0] - Smart Punch Inference

### Added
- Intelligent inference for one missing punch when exactly `3` punches exist in a coherent daily pattern
- Visual marker `~` for inferred entry times in the report and quick view

### Implemented
- Three-punch pattern matching against expected schedule slots (`entry`, `lunch out`, `lunch return`, `exit`)
- Protection against misleading worked-hours totals on incomplete punch sequences
- Lunch-overrun details now show only the minutes above the allowed maximum

### Business Rules
- When exactly one punch is missing and the remaining three match the expected sequence, the system infers the missing checkpoint
- Inferred entries can be used to estimate worked hours
- Inferred entries do not generate a tardiness penalty

### Notes
- This release improves robustness for real-world missed punches without changing standard schedules
- Fully compatible with the reporting and Windows distribution flow introduced in `v1.0.0`

---

## [v1.0.0] - First Official Release

### Added
- Windows desktop application for daily attendance review
- Command-line interface for technical usage
- Windows launcher with update flow through GitHub Releases
- Direct reading of `Personal_*.xls` and `Eventos de hoy_*.xls`
- Excel export with:
  - `Resumen`
  - `Vista rápida`
  - `Faltas`
  - `Retardos`
  - `Incidencias`
  - `Detalle diario`
- Separate overtime report generation when payable overtime exists
- Execution log with summary and global observations
- Worked-hours calculation in the main attendance report and quick view

### Implemented
- Daily attendance detection
- Tardy detection from `08:01`
- Absence detection when no punch exists in the day
- Missing punch detection for:
  - lunch out
  - lunch return
  - final exit
- Early-exit detection
- Duplicate punch cleanup keeping the earliest valid record
- Automatic exclusion of invalid personnel records without usable names
- Quick-view layout designed for sharing with supervisors
- Five-minute grace period for workday completion without removing tardy status

### Business Rules
- Monday to Friday schedule: `08:00` to `17:00`
- Saturday schedule: `08:00` to `14:00`
- Lunch rule:
  - Monday to Friday: `30` to `45` minutes
  - Saturday: `30` minutes
- Payable overtime only counts after completing the full workday
- A five-minute grace period applies to workday completion for payable overtime calculation
- Overtime is counted only in full hours

### Notes
- This release establishes the first formal baseline for the attendance control system
- The project follows the same Windows distribution pattern used in other LINHER internal tools

---

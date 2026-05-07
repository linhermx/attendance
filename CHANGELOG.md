# Changelog

All notable changes to this project will be documented in this file.

This project follows **Semantic Versioning** (`MAJOR.MINOR.PATCH`).

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

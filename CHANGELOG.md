# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [v2.1.1] - 2026-06-15

### Fixed

- Assigned isolated meal-period punches to the nearest flexible meal event when the distance is decisive, instead of leaving them fully ambiguous.
- Prevented isolated return-from-meal punches from being reported as missing both meal events when only the meal start is actually missing.

## [v2.1.0] - 2026-06-13

### Changed

- Classification now uses the punch state and source device as auditable evidence while still validating assignments against schedule, sequence, duplicate normalization, and lunch duration.
- Explicit meal-state punches are treated as strong classification hints without making lunch schedules rigid.
- Complete four-punch sequences can reinforce a valid entry, lunch-out, lunch-return, and exit hypothesis without reverting to positional-only assignment.

### Fixed

- Improved late-entry and lunch-pair classification when source punch states are generic or incorrectly selected by the user.
- Kept source state and device details in the technical audit instead of operational report detail.

## [v2.0.3] - 2026-06-10

### Fixed

- Classified a late first punch before lunch as a contextual late entry when followed by a coherent lunch pair, even if the final exit punch is missing.
- Avoided inventing an entry for isolated pre-lunch punches without later evidence to support a started shift.

## [v2.0.2] - 2026-06-09

### Fixed

- Prevented nearby duplicate punches from being used as separate attendance events.
- Kept duplicate-punch normalization in technical audit without adding operational incidents.

## [v2.0.1] - 2026-06-08

### Fixed

- Prevented meal-period punches from being incorrectly classified as early departures in incomplete attendance sequences.

## [v2.0.0] - 2026-06-06

### Changed

- Entry and exit remain rigid events evaluated against the configured schedule.
- Lunch punches are classified as flexible chronological pairs using duration and sequence coherence.
- Lunch duration is evaluated with real seconds and a dynamic allowed return based on the work schedule.
- Operational report detail is separated from technical classification audit.
- Missing lunch punches in partial cutoffs no longer depend on theoretical lunch times.
- Worked hours require a complete sequence of real punches.
- Saturday lunch has a 30-minute maximum; weekday lunch retains a 45-minute maximum.
- Sunday is treated as a non-working day; all Sunday punches are retained only for review without attendance incidents, deduplication, or schedule evaluation.
- Quick-view personnel is validated against the loaded personnel source.
- A first punch before the lunch reference is prioritized as contextual late entry when a plausible final exit exists.
- Isolated lunch-time punches, late entries with later exits, severe tardies, incomplete records, and ambiguous sequences are covered by the contextual classification rules.

### Added

- Business evaluation layer for operational incidents, status, and visible detail.
- Explainable classification audit with score, reference, assigned event, confidence, and reason.

### Removed

- Overtime calculation and overtime report generation.

## [v1.3.0]

- Added Windows installer and launcher workflows.
- Added historical range view and aligned daily/range report categories.

## [v1.2.0]

- Added range analysis, partial-cutoff detection, and consolidated historical reporting.

## [v1.1.0]

- Added initial smart punch inference and incomplete-sequence protections.

## [v1.0.0]

- Added the first daily attendance report, GUI, CLI, launcher, and Excel export.

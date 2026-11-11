# Changelog

## v0.1 — reviewed starter

Based on the SOLIDWORKS 2024 VBA macro in “Solidworks Drawing API Availability.”

- Corrected `Get6` document-property fallback to use `swCustomInfoGetResult_NotPresent` rather than `0` (cached-value success).
- Avoided copying unresolved property expressions into manufacturing notes.
- Replaced obsolete `ISheet.GetProperties` with `GetProperties2`, retaining the documented projection/width/height indices.
- Simplified configuration-member line continuation and replaced object-return expressions with explicit `Set` assignments.
- Added failure checks for every view and verified its referenced configuration.
- Activated the front view for marked-dimension import, requested duplicate elimination, and activated the sheet before inserting the note.
- Made required note insertion/position failures visible rather than silently skipping them.
- Checked save Boolean and error code; surfaced warning codes and distinguished PDF failure after drawing save.
- Added an importable VBA module attribute and documented import/compile steps, automatic scaling behavior, partial outputs, and runtime limitations.

Not compiled or executed in SOLIDWORKS; see `docs/VALIDATION.md`.

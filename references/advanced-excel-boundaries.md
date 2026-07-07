# Advanced Excel Boundaries

This skill can preserve many workbook parts, but some Excel features need explicit caution.

## Merged Cells

- Read merged ranges before editing.
- Translate/write only the top-left anchor cell unless the range is intentionally unmerged.
- Repair width and height around merged headings after translation.

## Pivot Tables

- Formula-backed pivot-style analysis is supported.
- Native PivotTable cache regeneration is not guaranteed after structural edits.
- If the workbook depends on native PivotTables, preserve package parts and tell user to refresh in Excel unless automation refresh succeeds.

## VBA and Macros

- Do not rewrite macro modules.
- Preserve `.xlsm` package parts through safe-copy paths when possible.
- Do not claim VBA logic was tested unless the macro code was inspected and executed in a suitable environment.

## External Connections

- Preserve connection/package parts where possible.
- Do not refresh Power Query, external DB connections, or web queries unless Excel/LibreOffice automation succeeds.
- Surface refresh status in delivery notes.

## Protected Workbooks

- Do not bypass protection.
- Ask for password or operate on accessible sheets/cells only.

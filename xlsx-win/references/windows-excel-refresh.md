# Windows Excel refresh notes

This skill is designed for native Excel refresh on Windows machines with Microsoft 365 Excel installed.

## Why native Excel is required

Some workbooks depend on Excel-specific behavior that LibreOffice does not reproduce reliably:

- workbook connections
- Power Query refresh
- PivotTables and data-model refresh behavior
- async query completion
- cached formula values expected by downstream consumers

## Recommended execution pattern

1. Edit the workbook with `openpyxl`, `pandas`, `scripts/power_query_excel.ps1`, or direct Excel COM as appropriate.
2. Run `scripts/refresh_excel.ps1`.
3. Run `scripts/check_formula_errors.ps1`.
4. If Power Query load settings changed, verify the expected worksheet table or model connection exists after refresh.
5. Use `scripts/self_test_xlsx_win.ps1` on new machines or after script changes to validate the local Excel environment.

## Contracts

### `scripts/refresh_excel.ps1`

Purpose:
- refresh workbook connections and async queries
- force full recalculation
- save the workbook
- emit a JSON status artifact

Key arguments:
- `-WorkbookPath` required
- `-LogPath` optional; defaults to a unique temp path
- `-JsonPath` optional; defaults to a unique temp path
- `-EnableMacros` optional opt-in; macros are disabled by default
- `-TimeoutSeconds` optional

JSON payload:
- `status`
- `message`
- `workbook`
- `macro_policy`
- `log_path`
- `json_path`
- `connection_count`
- `duration_seconds`
- `timestamp`
- `exit_code`
- `error_kind` on failure

Exit codes:
- `0` success
- `2` operational failure

### `scripts/check_formula_errors.ps1`

Purpose:
- run `check_formula_errors.py` through a Python interpreter with `openpyxl`
- report visible Excel error cells and formula counts as JSON on stdout

Exit codes:
- `0` clean workbook
- `2` findings present
- `1` validator failure

Important validator behavior:
- supports `.xlsx`, `.xlsm`, `.xltx`, and `.xltm`
- returns `python_not_found` when no usable Python + `openpyxl` environment is available
- caps sampled error locations for payload size, but the reported per-error `count` is the true occurrence count

### `scripts/self_test_xlsx_win.ps1`

Purpose:
- create temporary workbooks
- validate refresh, validation, path handling, macro policy, and Power Query helper behavior
- emit a JSON summary with `passed`, `failed`, and `skipped`

Use it:
- when setting up a new Windows machine
- after changing these scripts
- when Excel COM or Power Query behavior is suspect in the current session

## Example commands

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\refresh_excel.ps1 -WorkbookPath .\model.xlsx
powershell -ExecutionPolicy Bypass -File .\scripts\check_formula_errors.ps1 -WorkbookPath .\model.xlsx
powershell -ExecutionPolicy Bypass -File .\scripts\self_test_xlsx_win.ps1
```

## Failure patterns

### Excel cannot be created through COM

Likely causes:

- desktop Excel is not installed
- Office installation is damaged
- execution is not actually happening on Windows
- the current Codex session is sandboxed and cannot access an interactive desktop COM session

When the script reports Excel COM is unavailable in the current session, rerun it from an interactive Windows desktop session or outside the Codex sandbox.

### Workbook opens read-only or cannot be saved

Likely causes:

- another Excel instance is holding the file
- the file is marked read-only
- OneDrive or another sync agent has the file locked transiently

Inspect the JSON `error_kind` and close other Excel instances before retrying.

### Workbook opens but refresh hangs

Likely causes:

- blocked credentials or prompts
- slow external query
- another Excel instance is holding the file

Try a larger timeout and inspect the generated log and JSON output.

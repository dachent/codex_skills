# Power Query Excel COM workflow

Use native Excel COM for Power Query `Workbook.Queries` work on Windows. Prefer the bundled helper `scripts/power_query_excel.ps1` over ad hoc COM snippets for routine query lifecycle operations.

## Why the helper exists

Excel Power Query work is fragile because each query is a combination of:

- the M definition in `Workbook.Queries`
- the workbook connection
- the load target

`Queries.Add(name, mFormula)` creates only the query definition. It does not create a worksheet load or a Data Model load.

## Helper contract

`scripts/power_query_excel.ps1` supports these actions:

- `upsert-query`
  - create or update the M definition only
  - preserves existing load settings
- `load-worksheet`
  - ensure the query exists
  - create or reuse a worksheet table load at the requested sheet and start cell
  - refresh and verify the resulting row count
- `load-model`
  - ensure the query exists
  - create or reuse the required workbook connection
  - add the connection to the Data Model
  - refresh and verify the resulting model table
- `delete-query`
  - remove worksheet loads
  - remove related workbook connections
  - remove the query definition
  - verify the query and related artifacts are gone

## Important arguments

- `-WorkbookPath` required
- `-Action` required
- `-QueryName` required
- `-MFormula` optional inline M definition
- `-MFormulaPath` optional file path for M definition
- `-WorksheetName` required for `load-worksheet`
- `-StartCell` optional for `load-worksheet`, defaults to `A1`
- `-EnableMacros` optional opt-in; macros are disabled by default
- `-LogPath` optional
- `-JsonPath` optional

If the M definition is nontrivial or multiline, prefer `-MFormulaPath`. That avoids shell-quoting issues and is the default-safe contract for Codex.

Do not pass both `-MFormula` and `-MFormulaPath`.

## JSON payload

Successful runs emit structured JSON with:

- `status`
- `action`
- `message`
- `workbook`
- `query_name`
- `macro_policy`
- `log_path`
- `json_path`
- `duration_seconds`
- `timestamp`
- `exit_code`
- action-specific fields such as:
  - `query_created`
  - `query_updated`
  - `worksheet_load_created`
  - `worksheet_load_reused`
  - `model_load_created`
  - `deleted_worksheet_loads`
  - `deleted_connections`
  - `query_deleted`
  - `verification`
  - `refresh_log_path`
  - `refresh_json_path`
  - `refresh_status`

Failure payloads include `error_kind`.

Exit codes:

- `0` success
- `1` operational failure

## Idempotency and expectations

- `upsert-query` updates the query in place and preserves existing load targets.
- `load-worksheet` creates a worksheet table if missing, or reuses the matching load at the requested sheet and cell.
- `load-model` reuses an existing model load when present; otherwise it creates the needed connection and adds it to the Data Model.
- `delete-query` is safe to rerun; it succeeds even if the query definition is already gone, as long as related artifacts are fully removed by the end.

The helper verifies the final workbook state after refresh. Treat that verification as authoritative.

## Example commands

Use a file for nontrivial M:

```powershell
$formulaPath = ".\query.m"
powershell -ExecutionPolicy Bypass -File .\scripts\power_query_excel.ps1 `
  -WorkbookPath .\model.xlsx `
  -Action upsert-query `
  -QueryName SalesQuery `
  -MFormulaPath $formulaPath
```

Load a query to a worksheet table:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\power_query_excel.ps1 `
  -WorkbookPath .\model.xlsx `
  -Action load-worksheet `
  -QueryName SalesQuery `
  -WorksheetName Output `
  -StartCell B3
```

Load a query into the Data Model:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\power_query_excel.ps1 `
  -WorkbookPath .\model.xlsx `
  -Action load-model `
  -QueryName SalesQuery `
  -MFormulaPath .\query.m
```

Delete a query and its related workbook artifacts:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\power_query_excel.ps1 `
  -WorkbookPath .\model.xlsx `
  -Action delete-query `
  -QueryName SalesQuery
```

# SilverShop Import — Excel/CSV connector

The spreadsheet source connector for [`silvershop-suite/import`](https://github.com/silvershop-suite/silvershop-import),
and its **export → correct → re-apply roundtrip**. The MVP connector: it exercises the whole engine
(stage → map → upsert → dry-run → audit) with **no external system**.

## What it adds

- **`excel` import source** — reads an uploaded/dropped `.xlsx`, `.csv` or `.ods` workbook (via
  `phpoffice/phpspreadsheet`), using the header row as the source field names and one `SourceRecord` per
  data row. The upsert key column is the one your profile maps to `InternalItemID` (or `GTIN`), falling
  back to the first column.
- **The roundtrip** — the spreadsheet workflow:
  - **Result/error workbook** — after a run (or a `--dry-run`), export the staged rows with `_Status`
    and `_Error` columns appended. The client fixes the flagged rows in place and re-uploads; unchanged
    rows no-op (the engine's `PayloadHash`), failed rows retry.
  - **Full catalogue export** — dump products to a workbook as an editing template / pre-change backup.

## Usage

Create an `ImportProfile` in the CMS (Imports section), choose the **Excel / CSV spreadsheet** source,
set **SourceLocation** to the workbook path, and add field mappings (e.g. `SKU → InternalItemID`,
`Name → Title`, `Price → BasePrice`, `Qty → Quantity` for the Stock domain). Then:

```
# preview (writes nothing) — then export the result workbook to see per-row status/errors
sake tasks:SilverShop-Import-Run --profile=<id|key> --dry-run
sake tasks:SilverShop-Import-Excel-Export --type=result --profile=<id|key> --path=/tmp/result.xlsx

# apply for real
sake tasks:SilverShop-Import-Run --profile=<id|key>

# export the current catalogue as an editing template
sake tasks:SilverShop-Import-Excel-Export --type=catalogue --path=/tmp/catalogue.xlsx
```

## Requirements

- `silvershop-suite/import` (the core engine) — resolved from the composer `repositories` VCS entry
  until it is on Packagist.
- `phpoffice/phpspreadsheet` ^2 || ^3 (the only dependency this connector adds).

## Credits

Authored by Remy Vaartjes and the SilverShop-Suite Contributors. Licensed BSD-3-Clause.

# SilverShop Import — Excel / CSV connector

The spreadsheet source connector for [`silvershop-suite/import`](https://github.com/silvershop-suite/silvershop-import),
and its **export → correct → re-apply roundtrip**. It is the suite's MVP connector: it drives the whole
import engine (stage → map → upsert → dry-run → audit) with **no external system** — just a workbook.

Use it to:

- **Bulk-load** a catalogue from a supplier spreadsheet.
- **Bulk-edit** the live catalogue: export it, change it in Excel / LibreOffice / Numbers, re-import it.
- **Update stock and prices** on a schedule from a dropped file.
- **Fix-and-retry**: a run flags the bad rows in a result workbook; you correct those rows and re-upload.

It reads and writes `.xlsx`, `.csv` and `.ods`.

---

## Contents

- [For developers](#for-developers) — install, configure, extend
  - [Install](#install)
  - [How it fits the engine](#how-it-fits-the-engine)
  - [Configure a profile](#configure-a-profile)
  - [Field mappings](#field-mappings)
  - [The column schema](#the-column-schema)
  - [Domains & key strategy](#domains--key-strategy)
  - [record_type routing](#record_type-routing)
  - [Transforms](#transforms)
  - [Exporters & tasks](#exporters--tasks)
  - [YAML configuration](#yaml-configuration)
  - [Security](#security)
  - [Extending](#extending)
  - [Requirements](#requirements)
- [For shop managers / editors](#for-shop-managers--editors) — the day-to-day spreadsheet workflow
  - [The workflow](#the-workflow)
  - [Understanding the spreadsheet](#understanding-the-spreadsheet)
  - [Product rows vs. variant rows](#product-rows-vs-variant-rows)
  - [How to make common edits](#how-to-make-common-edits)
  - [The golden rules](#the-golden-rules)
  - [Re-importing & fixing errors](#re-importing--fixing-errors)
  - [FAQ](#faq)

---

# For developers

## Install

```bash
composer require silvershop-suite/import-excel
```

This pulls in the core engine (`silvershop-suite/import`) and `phpoffice/phpspreadsheet`. Run
`vendor/bin/sake db:build flush=all` afterwards.

## How it fits the engine

The core `silvershop-suite/import` is a two-phase **stage → apply** pipeline that knows nothing about
spreadsheets. This module registers one `ImportSource` under the key **`excel`**
(`SilverShop\Import\Excel\ExcelImportSource`). Read the [core README](https://github.com/silvershop-suite/silvershop-import)
for the engine model (profiles, staging, runs, ownership, scheduling); this README covers only the
spreadsheet specifics.

The `ExcelImportSource` reads the workbook's active sheet, uses the **header row** as source field names,
and yields one `SourceRecord` per data row. The profile's `FieldMapping`s then turn those source columns
into silvershop target fields during the apply phase.

## Configure a profile

An `ImportProfile` (managed in **ImportAdmin → Imports**) is the unit of configuration:

| Field              | For Excel, set to…                                                                        |
|--------------------|-------------------------------------------------------------------------------------------|
| **SourceKey**      | `excel`                                                                                    |
| **SourceLocation** | Absolute path to the workbook (e.g. `/var/imports/catalogue.xlsx`).                        |
| **Domains**        | Comma list of what to sync: `Product,Variation,Stock` (and/or `Price`).                    |
| **WriteMode**      | `upsert` (create + update), `add-only`, or `update-only`.                                   |
| **MissingPolicy**  | `none` (leave missing records alone — safest), `deactivate`, or `delete`.                   |
| **KeyStrategy**    | `sku,gtin` — resolve by SKU, then GTIN. (`sku` alone for product-only sheets.)              |
| **OwnershipPreset**| `full-erp` (sheet wins on every field), `pos-price`, `stock-only`, or `project-default`.    |
| **Enabled**        | `true`                                                                                      |

> **Tip:** for a bulk-edit roundtrip you almost always want `WriteMode=upsert`,
> `MissingPolicy=none`, `KeyStrategy=sku,gtin`, `OwnershipPreset=full-erp`.

## Field mappings

Each `FieldMapping` row (per profile) maps one **source column → one target field**, scoped to a
**domain**, with a `Transform`, optional `ValidationRules`, and an `Ownership` override. For the grouped
catalogue shape a typical set is:

| Domain      | Source column    | Target field     | Transform | Note                                        |
|-------------|------------------|------------------|-----------|---------------------------------------------|
| Product     | `InternalItemID` | `InternalItemID` | `trim`    | the SKU — the Product key                   |
| Product     | `Title`          | `Title`          | `trim`    |                                             |
| Product     | `BasePrice`      | `BasePrice`      | `number`  |                                             |
| Product     | `GTIN`           | `GTIN`           | `trim`    | a simple product's own barcode (optional)   |
| Product     | `Category`       | `Category`       | `trim`    | pseudo-target → category resolver           |
| Product     | `Brand`          | `Brand`          | `trim`    | pseudo-target → brand resolver              |
| Variation   | `InternalItemID` | `ProductSKU`     | `trim`    | the **parent** SKU → resolves the product   |
| Variation   | `GTIN`           | `GTIN`           | `trim`    | the variation key                           |
| Variation   | `Price`          | `Price`          | `number`  |                                             |
| Variation   | `Option1Name`    | `Option1Name`    | `trim`    | axis label (→ `Option2/3` for more)         |
| Variation   | `Option1Value`   | `Option1Value`   | `trim`    | axis value                                  |
| Stock       | `GTIN`           | `GTIN`           | `trim`    | the stock key                               |
| Stock       | `Quantity`       | `Quantity`       | `integer` | pseudo-target → stock writer                |

**Pseudo-targets** (`Category`, `Brand`, `Quantity`/`Stock`, `ProductSKU`, `Option{N}Name/Value`) aren't
database fields — the engine routes them to the category/brand/stock/attribute resolvers. All other target
fields are written straight onto the `Product` / `Variation`.

## The column schema

The grouped export produces these columns (order configurable, see [YAML](#yaml-configuration)):

```
record_type, InternalItemID, Title, BasePrice, GTIN, Price, Quantity,
Category, Brand, Option1Name, Option1Value, Option2Name, Option2Value, Option3Name, Option3Value
```

- **Category** paths use ` > ` between levels and `|` between multiple paths (configurable on
  `CategoryResolver`). The first path's leaf becomes the product's primary parent (`ParentID`); every
  leaf is added as a `ProductCategories` membership. Missing categories are created.
- **Brand** matches an existing brand by `Title` (creates it if missing). Requires
  `silvershop-suite/brands`; without it, the `Brand` column is ignored.
- **Option{N}** axes map to silvershop variation attributes (`AttributeType` + `AttributeValue`), up to
  `AttributeResolver.max_axes` (default 3).
- **Quantity** writes through `silvershop/stock` (`setStockLevel()`); requires that module and a
  configured warehouse. Without it, the column is ignored.
- **GTIN** keying/validation requires `silvershop-suite/product-identity`.

## Domains & key strategy

A single profile can sync several domains; each run processes one domain. List the **Product** domain
before **Variation** so parent products exist before their variations are linked (a missing parent is
created minimally regardless).

| Domain        | Rows it reads (by `record_type`) | Keyed by            | Resolves/writes                               |
|---------------|----------------------------------|---------------------|-----------------------------------------------|
| **Product**   | `product` / blank                | SKU (`InternalItemID`) | create/update products + category + brand  |
| **Variation** | `variant`                        | GTIN                | create/update variations under their parent SKU |
| **Stock**     | every row with a key             | GTIN (then SKU)     | `setStockLevel()` on the matched buyable      |
| **Price**     | `product` / blank                | SKU                 | folds `Price`/`BasePrice` onto the product    |

The upsert key column is chosen per domain from the mappings (a field mapped to `InternalItemID`, else
`GTIN`, else the first column) — so for the Variation/Stock domains, mapping `GTIN` (and *not*
`InternalItemID`) makes the rows key on GTIN. A varianted product's parent row has a blank GTIN and is
therefore skipped by the Stock domain (its stock lives on the variants).

## record_type routing

The reader splits rows by the **`record_type`** column (configurable name):

- `product`, `simple`, or **blank** → applied on the **Product** domain.
- `variant` or `variation` → applied on the **Variation** domain.
- The **Stock** / **Price** domains take every keyed row regardless of `record_type`.

A sheet with **no `record_type` column at all** is a flat product-only catalogue — every row is a product
(handy for a quick price/stock update). It imports unchanged on the Product domain and yields nothing on
the Variation domain.

## Transforms

Set per `FieldMapping.Transform`; unknown transforms (and `null` values) pass through unchanged:

`trim` · `number` (float) · `integer` · `boolean` (`1/true/yes/y`) · `uppercase` · `lowercase` ·
`date` (normalises to `Y-m-d H:i:s`).

## Exporters & tasks

Two writers back the roundtrip, both exposed through one task:

- **`ExcelCatalogueExporter`** — the full catalogue in the grouped shape (product + variant rows). An
  editing template / pre-change backup.
- **`ExcelResultWriter`** — a profile+domain's staged rows written back with `_Status` and `_Error`
  columns appended (the fix-and-re-upload workbook).

```bash
# Export the current catalogue as an editing template
sake tasks:SilverShop-Import-Excel-Export --type=catalogue --path=/tmp/catalogue.xlsx

# After a run (or a --dry-run), export the result workbook with per-row status/errors
sake tasks:SilverShop-Import-Excel-Export --type=result --profile=<id|key> --path=/tmp/result.xlsx

# Preview (stage + simulate, writes nothing), then apply — from the core engine:
sake tasks:SilverShop-Import-Run --profile=<id|key> --dry-run
sake tasks:SilverShop-Import-Run --profile=<id|key>
```

`--type` is `catalogue` (default) or `result`; `--domain` narrows a result export; `--path` is optional
(auto-named under `import-exports/` if omitted).

All exported cells are written **formula-injection-safe** (see [Security](#security)).

## YAML configuration

```yaml
# This module — the record_type column name and the export column set.
SilverShop\Import\Excel\ExcelImportSource:
  record_type_column: 'record_type'

SilverShop\Import\Excel\Export\ExcelCatalogueExporter:
  columns:
    - record_type
    - InternalItemID
    - Title
    - BasePrice
    - GTIN
    - Price
    - Quantity
    - Category
    - Brand
    - Option1Name
    - Option1Value

# Core resolvers that shape how the sheet's Category / Option columns are interpreted.
SilverShop\Import\Service\CategoryResolver:
  path_separator: ' > '
  paths_separator: '|'
  category_parent_id: 0      # create new categories under this ProductCategory (0 = site root)
  create_missing: true

SilverShop\Import\Service\AttributeResolver:
  max_axes: 3

# Project default for who-wins-on-a-field (see the core README).
SilverShop\Import\Service\OwnershipResolver:
  default_ownership:
    BasePrice: erp-authoritative
    Title: erp-authoritative
```

## Security

- **Spreadsheet formula injection (export).** Every exported cell whose text starts with `=`, `+`, `-`,
  `@`, tab or CR is written as an explicit literal string, so a crafted field value can't execute as a
  formula when the workbook is opened. This is a fixed policy, not a config toggle.
- **XXE (import).** Workbook XML is parsed by phpspreadsheet, which disables external-entity loading by
  default — keep it patched (`^2 || ^3`).
- **`SourceLocation` is an arbitrary local path** read with the web user's permissions. It is set in
  `ImportAdmin`, which requires CMS access — treat who can edit profiles as a trust boundary.

## Extending

- **Change the detected columns** — override `ExcelCatalogueExporter.columns` in YAML, or subclass it.
- **Rename the type column** — set `ExcelImportSource.record_type_column`.
- **A different spreadsheet layout** — write your own `ImportSource` (implement the interface, register it
  under a new key); you don't have to use this reader to use the engine.
- **Multiple supplier layouts** — one profile per layout, each with its own mappings, all using the
  `excel` source.

## Requirements

- PHP `^8.3` · `silvershop/core` `^6.2` · `silvershop-suite/import`
- `phpoffice/phpspreadsheet` `^2 || ^3` (the only dependency this connector adds)
- Optional (guarded, per column): `silvershop/stock` (Quantity), `silvershop-suite/product-identity`
  (GTIN), `silvershop-suite/brands` (Brand)

---

# For shop managers / editors

You don't need to know anything about code to use this. You work in a spreadsheet; the shop imports it.

## The workflow

```
   ┌─────────────┐   export    ┌──────────────┐   edit in    ┌──────────────┐  re-import  ┌─────────────┐
   │  Your shop  │ ──────────► │  catalogue   │ ───────────► │   edited     │ ──────────► │  Your shop  │
   │ (products)  │             │   .xlsx      │  Excel/Calc  │   .xlsx      │             │ (updated)   │
   └─────────────┘             └──────────────┘              └──────────────┘             └─────────────┘
```

1. **Export** the current catalogue to a spreadsheet (your developer sets this up, or runs it for you —
   see the [tasks](#exporters--tasks)). You get an `.xlsx` with one row per product and one row per
   product variation.
2. **Edit** it in Excel, LibreOffice Calc or Numbers. Change prices, titles, stock, categories, brands —
   whatever you need. Save it.
3. **Re-import** it. The shop reads your changes back in. Rows you didn't touch stay exactly as they were.

The import is **safe to repeat**: importing the same unchanged file twice does nothing the second time,
and a run never deletes products just because they're missing from your sheet (unless your developer
turned that on deliberately).

## Understanding the spreadsheet

The first row is the **header**. Every row after it is one product or one variation. Here's a small
example — a trail-running shoe that comes in two sizes, and a one-size sock:

| record_type | InternalItemID | Title        | BasePrice | GTIN          | Price | Quantity | Category          | Brand | Option1Name | Option1Value |
|-------------|----------------|--------------|-----------|---------------|-------|----------|-------------------|-------|-------------|--------------|
| product     | SHOE-01        | Trail Runner | 89.95     |               |       |          | Footwear > Trail  | Acme  |             |              |
| variant     | SHOE-01        |              |           | 4006381333931 | 89.95 | 12       |                   |       | Size        | 42           |
| variant     | SHOE-01        |              |           | 4012345678901 | 94.95 | 7        |                   |       | Size        | 43           |
| product     | SOCK-01        | Merino Sock  | 14.95     | 5901234123457 |       | 40       | Accessories       | Acme  |             |              |

What each column means:

| Column                    | What it is                                                                                      |
|---------------------------|-------------------------------------------------------------------------------------------------|
| **record_type**           | `product` for a product row, `variant` for one of its variations. (Blank counts as `product`.)  |
| **InternalItemID**        | The product code / **SKU**. On a `variant` row this is **the parent product's SKU** (repeated).  |
| **Title**                 | The product name. (Only on `product` rows.)                                                      |
| **BasePrice**             | The product's price. (Only on `product` rows.)                                                   |
| **GTIN**                  | The barcode / **EAN** — the unique identity of a *variation*. A simple product may have its own. |
| **Price**                 | A variation's price. (Only on `variant` rows.)                                                   |
| **Quantity**              | Stock level. On a `variant` row = that variation's stock; on a simple product = the product's.   |
| **Category**              | Where the product lives. `Parent > Child` for a path; `A > B \| C` to put it in several.          |
| **Brand**                 | The brand name.                                                                                  |
| **Option1Name / Value**   | A variation's choice — e.g. `Size` = `42`, or `Colour` = `Red`. Up to `Option3` for two more axes. |

## Product rows vs. variant rows

A product that has **variations** (sizes, colours…) is written as **one `product` row followed by one
`variant` row per variation**:

- The **`product` row** carries the shared stuff: title, base price, category, brand.
- Each **`variant` row** repeats the parent's SKU in `InternalItemID`, and carries its own **GTIN**
  (barcode), price, stock, and its option values (`Size = 42`).

A product with **no variations** is a single `product` row — put its stock, and its own barcode, right
on that row (like `SOCK-01` above).

The **GTIN is how the shop recognises a variation.** Two variations of the same product share the SKU but
have different GTINs — that's what tells them apart.

## How to make common edits

- **Change a price** — edit `BasePrice` on a product row, or `Price` on a variant row.
- **Change a title** — edit `Title` on the product row.
- **Change stock** — edit `Quantity` (on the variant row for a varied product, on the product row for a
  simple one).
- **Move / set a category** — edit `Category`. Use `Men > Shoes` for a nested path. A category that
  doesn't exist yet is created for you. Put a product in two categories with a pipe: `Men > Shoes | Sale`.
- **Change the brand** — edit `Brand` (the brand name). A new brand name creates/assigns that brand.
- **Add a new variation to an existing product** — add a new `variant` row with the parent's SKU in
  `InternalItemID`, a **new unique GTIN**, its price, stock and option values.
- **Add a whole new product** — add a `product` row (and, if it has variations, the `variant` rows under
  it). Give it a new SKU.

## The golden rules

1. **Don't change a GTIN or a SKU to "rename" something.** They are *identities*. Change a SKU and the
   shop thinks it's a different product; change a GTIN and it's a different variation. To rename the
   *display name*, edit `Title`, not the codes.
2. **Keep each variant row's `InternalItemID` equal to its parent product's SKU.** That's the link
   between a variation and its product.
3. **A GTIN must be a valid barcode** (8–14 digits with a correct check digit). An invalid one is
   rejected and that row is flagged — the rest of the import still goes through.
4. **Leaving a cell blank doesn't delete anything** — a blank price/stock is simply "no change intended"
   where the column doesn't apply (e.g. `Price` on a product row). Deleting a product is not done by
   removing its row; ask your developer about the missing-row policy.
5. **Don't rename or reorder the header row.** The shop matches columns by their header names.

## Re-importing & fixing errors

When the file is re-imported, you (or your developer) can run a **preview** first — it checks every row
and writes nothing. If some rows have problems (a bad barcode, a price that isn't a number…), you get a
**result workbook**: the same rows with two extra columns, **`_Status`** and **`_Error`**, telling you
exactly which rows failed and why.

Fix just those rows in the result workbook, save, and re-import. Rows you already got right are recognised
as unchanged and skipped — only the fixed rows are retried. Repeat until every row says `Applied`.

## FAQ

- **Do I have to include every product every time?** No. You can import a sheet with only the rows you
  want to change (e.g. a price-and-stock update for ten products). Untouched products are left alone.
- **Can I just update stock?** Yes — a flat sheet with `InternalItemID`/`GTIN` + `Quantity` columns works;
  you don't need the variation structure for a stock-only update.
- **What happens to a product that's missing from my sheet?** By default, nothing (it's left as-is). Your
  developer can configure a run to deactivate or remove missing products, but that's off unless chosen.
- **I changed a variation's colour from Red to Green and nothing happened.** Changing an *option value*
  on an existing variation isn't an update in this version — it would be a *new* variation. Price, stock
  and the product-level fields update fine; option changes mean adding a new variant row.

---

## Credits

Authored by Remy Vaartjes and the SilverShop-Suite Contributors. Licensed BSD-3-Clause.

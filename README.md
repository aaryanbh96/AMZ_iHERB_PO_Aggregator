# The Tao of Tea — AMZ/iHerb PO Aggregator
## Factbook & Reference Guide — v0.0.7
**Built by Aryan Bhardwaj for The Tao of Tea**
**Code:** github.com/aaryanbh96/AMZ_iHERB_PO_Aggregator

---

## What Is This

A native desktop app that processes Purchase Orders from Amazon Vendor Central and iHerb, calculates costs per SKU/ASIN, and generates formatted Excel reports. Built with Tauri (Rust backend, HTML/JS frontend). Replaced manual Jupyter notebook analysis entirely.

---

## How to Run

**Dev mode** (requires Node.js, Rust, Tauri CLI):
```
cd tao-of-tea-po
cargo tauri dev
```
First compile takes 5–10 min. After that ~20 seconds.

**Build for distribution:**
```
cargo tauri build
```
Output is in `src-tauri/target/release/bundle/` — distribute the `.msi` (Windows) or `.dmg` (Mac).

---

## Project Structure

| File / Folder | Location | Purpose |
|---|---|---|
| `main.rs` | src-tauri/src/ | All Rust logic + Tauri commands |
| `lib.rs` | src-tauri/src/ | Stub only — `pub fn run() {}` |
| `Cargo.toml` | src-tauri/ | Rust dependencies |
| `tauri.conf.json` | src-tauri/ | Window size, title, plugins |
| `capabilities/default.json` | src-tauri/ | Permissions: dialog open/save |
| `index.html` | src/ | Full frontend: HTML + CSS + JS |
| `vite.config.js` | root | Vite dev server (root: src) |
| `package.json` | root | npm scripts and JS dependencies |

> ⚠️ All `#[tauri::command]` functions live in `main.rs` only. Adding commands to `lib.rs` causes a duplicate macro error on Windows — this was a painful lesson learned.

---

## Tab 1 — PO Aggregator

Processes a single Amazon or iHerb PO and shows a cost summary grouped by SKU/ASIN.

### Amazon PO
- File: `.xlsx` from Amazon Vendor Central (EditLineItems export)
- Columns used: `ASIN`, `Title`, `Quantity Requested`, `Case Cost`
- Output: ASIN | Title | Qty Ordered | Total Cost
- Total Cost = Case Cost × Qty Requested, grouped per ASIN

### iHerb PO
- File: `.csv` from iHerb portal
- Columns used: `Buyers Catalog or Stock Keeping #`, `Product/Item Description`, `Qty Ordered`, `Unit Price`
- Output: SKU | Title | Case Qty | Total Cost
- Total Cost = Unit Price × Qty Ordered (direct)
- Case Qty = Qty Ordered ÷ Units Per Case

### Premier SKU Rule
These 7 SKUs have **6 units/case**. Everything else has **12 units/case**:
```
TOT91451  TOT91471  TOT91461  TOT91411  TOT91421  TOT91491  TOT91441
```

### Save Excel
Exports to `.xlsx` — Times New Roman 14pt, no colors, grand total row at bottom.

---

## Tab 2 — Combined Reports

Generates two Excel reports by combining Amazon and iHerb POs through a mapping file.

### The Mapping File
An Excel file that links iHerb SKUs to Amazon ASINs for products on both platforms.

| Column | Notes |
|---|---|
| `Buyers Catalog or Stock Keeping #` | iHerb SKU |
| `ASIN` | Amazon ASIN — blank if iHerb-only |
| `Product Description` | Canonical product name used in reports |
| `Pack?` | Write "pack exists" if Amazon sells a pack variant |

The mapping file path is **saved automatically** after first selection. Only needs updating when new products are added.

### Report 1: Combined Summary
One row per product — Amazon and iHerb quantities/costs side by side, sorted by Total Cost.

- Products in mapping → merged into one row
- Amazon-only ASINs (not in mapping) → own row, iHerb columns blank
- iHerb-only SKUs (not in mapping) → own row, Amazon columns blank
- Pack rows → grey highlight on ASIN/SKU and Title cells, no combined total calculated
- Grand total row sums all 6 numeric columns

Designed for legal paper, landscape mode.

### Report 2: Big Ticket Breakdown
Only products where **combined Amazon + iHerb total ≥ $2,000**.

Each product shows:
- Header row: ASIN/SKU | Product Name | Combined Total
- Amazon section (if applicable): PO# | Expected Date | Ship To | Qty | Cost
- iHerb section (if applicable): PO# | Delivery Date | Ship To State | Cases | Cost

Both reports use Times New Roman 14pt, no colors, black borders.

---

## Known Business Rules

| Rule | Detail |
|---|---|
| Premier SKUs | 6 units/case (see list above) |
| All other SKUs | 12 units/case |
| Big ticket threshold | $2,000 combined total |
| iHerb Ship To State | Only filled on PO header row in CSV — forward-filled per PO group |
| Pack products | Shown but not combined with iHerb units (different pack sizes) |
| Mapping purpose | Bridge only for overlap — not a whitelist |

---

## Common Issues & Fixes

**App shows blank / 404 on launch**
Vite root must be set to `src/` in `vite.config.js`. Check: `root: "src"` is present.

**Duplicate command error when compiling**
All Tauri commands must be in `main.rs` only. If `lib.rs` has any `#[tauri::command]` functions, move them all to `main.rs` and make `lib.rs` just `pub fn run() {}`.

**iHerb Ship To State showing blank**
The CSV only fills this column on the first row of each PO. The app forward-fills it per PO group using `groupby('PO Number')`. If blank, check if the CSV has the column at all.

**Combined report missing rows / wrong totals**
The mapping file is a bridge, not a whitelist. Amazon ASINs not in the mapping still appear — they show Amazon data only. Run the diagnostic script (`diagnose_combined.py` in the Python version) to see exactly what's matched and what's not.

**Font too small in app**
In `src/index.html`, set `html { font-size: 125%; }` in the CSS. This scales everything proportionally in the Chromium webview.

---

## Dependencies

**Rust crates (Cargo.toml):**
- `tauri` v2 — app framework
- `tauri-plugin-dialog` v2 — file open/save dialogs
- `calamine` v0.26 — reads .xlsx files
- `csv` v1.3 — reads .csv files
- `rust_xlsxwriter` v0.80 — writes formatted .xlsx files
- `indexmap` v2 — ordered hashmaps for deterministic output
- `dirs` v5 — finds home directory for config storage
- `serde` / `serde_json` — serialization between Rust and JS

**JS (package.json):**
- `@tauri-apps/api` v2 — invoke Rust commands from JS
- `@tauri-apps/plugin-dialog` v2 — file picker in JS
- `vite` v6 — dev server and bundler

---

## Config Storage

The mapping file path is saved to `~/.taoftea_po_config.txt` on the user's machine. This persists between sessions so it only needs to be set once.

---
## Screenshots

![Screenshot](Screenshots/Screenshot%202026-03-25%20185431.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185511.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185547.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185625.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185656.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185723.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185826.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185859.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20185920.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20190019.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20190116.png)

![Screenshot](Screenshots/Screenshot%202026-03-25%20190228.png)
*Last updated: June 2026*

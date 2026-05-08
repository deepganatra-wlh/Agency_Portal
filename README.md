# Agency Grid Processor — README

A web portal for processing **Special Motor Matrix** commission grid Excel files
into structured CSV outputs. Fully dynamic and config-driven — no code changes
needed for new grids, new sheets, or layout variations.

---

## Quick Start

### 1. Install dependencies

```bash
pip install flask pandas openpyxl pyxlsb
```

### 2. Run the server

```bash
python app.py
```

Open **http://localhost:5051** in your browser.

---

## File Structure

```
agency_portal/
├── app.py              # Flask backend — all processing logic
├── templates/
│   └── index.html      # Full portal UI (single-page, no build step)
├── uploads/            # Temp uploaded files (auto-created)
├── outputs/            # Generated CSV files (auto-created)
└── README.md           # This file
```

---

## Supported Input Files

| File Type | Notes |
|---|---|
| `.xlsx` | Standard Excel — primary format |
| `.xlsb` | Binary Excel — requires `pyxlsb` |
| `.xls` | Legacy Excel |

Two Excel files are used:

- **Matrix Grid File** *(required)* — The Special Motor Matrix commission grid.
  Contains agent metadata, multi-row column headers, and per-cluster rate data.
- **RTO vs Cluster File** *(optional)* — Maps UW Cluster codes to RTO codes.
  Required only if you want `Rto Code*` populated. Must contain columns:
  `RTO CODE`, `PRODUCT CATEGORY`, `UW CLUSTER (25-26)`.

---

## Two Processing Modes

Toggle between modes using the **Special Mode / Normal Mode** buttons at the top.

### ⚡ Special Mode

For sheets where blocked/special values still need to produce an output row
(e.g. to record an IRDA commission instruction in the system).

| Cell Value | Behaviour |
|---|---|
| Empty | Skip this cell — no output row |
| `Block`, `NA`, `IRDA`, `MISP`, `SYSTEM COMMISSION` | **Produce output row** with `Span Outgo = IRDA` and `Span Prct = -0.1` |
| Any numeric rate | Produce output row with `Span Outgo = GWP` and `Span Prct = rate` |

**Important:** In both modes, the skip/IRDA behaviour applies to the **individual
cell only**, not the whole row. Other columns on the same data row are still
processed normally.

### 📋 Normal Mode

For Std Grid sheets where blocked cells should simply be excluded.

| Cell Value | Behaviour |
|---|---|
| Empty | Skip this cell |
| `Block`, `NA`, `IRDA`, `MISP`, `SYSTEM COMMISSION` | **Skip this cell only** — row continues for other columns |
| Any numeric rate | Produce output row with `Span Outgo = GWP` and `Span Prct = rate` |

---

## 7-Step Workflow

### Step 1 — Upload Files

Upload the Matrix Grid Excel and optionally the RTO vs Cluster file.
Drag-and-drop or click to browse. Both files are stored temporarily in `uploads/`.

---

### Step 2 — Sheet & Mode Configuration

#### Sheet & Header Configuration

| Field | What it means |
|---|---|
| **Data Sheet** | Which sheet in the Excel to process |
| **Header Row 1** | Row number (1-indexed) containing the BIZ MIX group label (e.g. `PVT CAR(1+1)`) |
| **Header Row 2** | Row with the sub-category label (e.g. `DIESEL & NCB`) |
| **Header Row 3** | Row with the data label (e.g. `GWP`) |
| **Data Start Row** | First row containing actual rate data |
| **Rate Columns Start At Column #** | Column index where rate data begins (meta cols are to the left) |

**Typical values for Special Matrix sheets:** Header rows 2, 3, 4 · Data start row 5 · Rate cols start at 13

**Typical values for Std Grid sheets** (Agency-Comp-Expt, SATP etc.): Same header rows · Data start row 5 · Rate cols start at **4**

#### Meta Column Positions

These tell the portal which column holds each piece of row metadata.
Use **1-indexed** column numbers. Set to **0** if the column does not exist in this sheet.

| Field | Description |
|---|---|
| **IMD Code** | Column containing the IMD / agent code |
| **IMD Name** | Column containing the agent name |
| **Relationship Code** | Column containing the relationship code |
| **IMD Type** | Column containing IMD Type (`Agency`, `Key Broking`, `Prime Broking`) |
| **Vol Lower** | Column with the GWP lower bound value |
| **Vol Upper** | Column with the GWP upper bound value |
| **Vol Consideration** | Column with the Volume Consideration label (e.g. `PCV-3W`, `TRACTOR`) |
| **UW Cluster / Outgo On** | Column with the UW Cluster code (e.g. `MH1`, `AP`) |

> **Std Grid sheets have no IMD columns** — set IMD Code, IMD Name, Relationship Code,
> and IMD Type all to **0**. The portal will automatically use UW Cluster and Vol Lower
> as the row identity instead.

#### Value Logic Configuration

Configure the list of cell values that are treated as "blocked" or "special":

- **Values to ignore / trigger IRDA** — one per line. In Normal mode these skip the cell.
  In Special mode these trigger IRDA output.
- **IRDA Outgo Value** — the string written to `Span Outgo*` for IRDA rows (default: `IRDA`)
- **IRDA Prct Value** — the rate written to `Span Prct*` for IRDA rows (default: `-0.1`)
- **Normal/GWP Outgo Value** — the string written to `Span Outgo*` for regular rows (default: `GWP`)
- **Skip row if Vol Lower equals** — comma-separated values that indicate a header-bleed row
  (e.g. if `Agency` appears in the Vol Lower column, it's a repeated header, not a data row)

After filling in Sheet & Header Configuration, click **Inspect Sheet Structure** to
auto-detect columns, preview data, and populate Step 3 with default values.

---

### Step 3 — Column Configuration

Each rate column detected in the sheet is listed in a table. The portal
auto-matches columns to the built-in default configuration (54 pre-configured
column types — see [Default Column Mappings](#default-column-mappings) below).

| Table Column | Description |
|---|---|
| **#** | The Excel column index for this rate column |
| **Sheet Column** | Display name from the multi-row headers (e.g. `PVT CAR(1+1) \| DIESEL & NCB`) |
| **Biz Mix Output** | Value written to the `Biz Mix*` output column for rows from this column |
| **RTO Category** | Category name used to look up RTO codes from the RTO vs Cluster file |
| **Extra Output Fields** | Additional fixed values written to output — one `Key=Value` per line |
| **Enable** | Checkbox to include/exclude this column from processing |

**Column colour coding:**
- 🟢 Green border = auto-matched to a built-in default
- 🟡 Amber border = not matched — requires manual configuration

**Extra fields format:**
```
Fuel Type*=Diesel
Gross Vehicle Weight Ll*=12000.001
Gross Vehicle Weight Ul*=20000.001
Vehicle Age Ul*=5
```

You can add new columns using **+ Add Row** for columns the auto-inspect may have missed.

---

### Step 4 — GWP Columns & Agent Group Map

#### Volume Consideration → GWP Column Mapping

Maps each unique **Volume Consideration** value from the source sheet to the
correct GWP Lower Limit and Upper Limit **output column names** in the target schema.

When a data row is processed, the portal reads its Vol Lower and Vol Upper values,
then writes them to the correct output columns based on this mapping.

| Vol Consideration Value | GWP LL Column | GWP UL Column |
|---|---|---|
| `PCV-3W` | `PCV 3W GWP Ll*` | `PCV 3W GWP Ul*` |
| `TRACTOR` | `Totalgwp Tractor Ll*` | `Totalgwp Tractor Ul*` |
| `GCV 7.5T - 12T` | `Totalgwp 7500 12000 Ll*` | `Totalgwp 7500 12000 Ul*` |
| *(any unmatched)* | Default GWP LL Col | Default GWP UL Col |

**Special case — `std-grid`:** When Vol Consideration is `std-grid`, the portal
automatically uses **Prime Broking GWP columns** for Prime Broking IMD types and
**Agency/Key Broking GWP columns** for all others.

The detected Vol Consideration values from the uploaded sheet are shown as chips
above the mapping table so you can verify all values are covered.

#### IMD Type → Agent Group Code Map

Maps each IMD Type to the Agent Group Code written to the output:

| IMD Type | Agent Group Code |
|---|---|
| `Agency` | `Agency` |
| `Key Broking` | `Key_Broking` |
| `Prime Broking` | `Prime_Broking` |

Add rows for any additional IMD Types detected in your sheet.

---

### Step 5 — RTO Configuration

#### RTO File Structure

Configure how the RTO vs Cluster file is read:

| Field | Default |
|---|---|
| RTO Sheet Name | `RTO Vs Cluster (New)` |
| Header Row | `2` |
| RTO Code Column Header | `RTO CODE` |
| Cluster Column Header | `UW CLUSTER (25-26)` |
| Category Column Header | `PRODUCT CATEGORY` |

#### Category Normalization

When an RTO lookup for a category returns no results, the portal retries using a
normalized category name before falling back to ALL RTO codes.

Pre-configured normalizations:

| Source Category | Lookup As |
|---|---|
| `PCV 3W` | `PCV` |
| `PCV-BUS` | `PCV` |
| `PCV-TAXI` | `PCV` |
| `MISD` | `MISC-D` |
| `MISD GARBAGE` | `MISC-D` |
| `TRACTOR` | `MISC-D` |

#### Output Column Names

All output column names are configurable here to match your target system schema:

| Setting | Default Value |
|---|---|
| Span Outgo Column | `Span Outgo*` |
| Span Prct Column | `Span Prct*` |
| RTO Code Column | `Rto Code*` |
| RTO Cluster Column | `Rto Cluster*` |
| Parent Agent Code | `Parent Agent Code*` |
| Primary Agent Code | `Primary Agent Code*` |
| Agent Group Code | `Agent Group Code*` |
| Biz Mix Column | `Biz Mix*` |

---

### Step 6 — Output Settings

#### Static Output Fields

Fields written with a fixed value to every output row regardless of source data.
Typical use: `Version Id*` for grid versioning.

```
Version Id*  =  agency_common_sep_Grid_3
Oem Type*    =  AGENCY
```

#### Output File Name

Sets the filename of the downloaded CSV (without extension).

---

### Step 7 — Process & Export

Shows a configuration summary, then runs the transformation.

The result summary shows:
- **Output Rows** — total rows generated
- **Columns** — number of output columns
- **Blank Rows Skipped** — source rows skipped because all identity columns were empty or it was a header-bleed row
- **Cells Skipped** — individual cells skipped due to empty value or ignore-list match (Normal mode)
- **Mode** — which mode was used

A 5-row preview table is shown after processing. Click **Download CSV** to save.

---

## Output Row Structure

Each output row contains:

| Output Column | Source |
|---|---|
| Static fields | From Step 6 configuration |
| `Agent Group Code*` | Mapped from IMD Type via agent group map |
| GWP LL/UL column(s) | Vol Lower / Vol Upper from source row, written to the correct GWP col based on Vol Consideration |
| `Parent Agent Code*` | IMD Code from source row |
| `Primary Agent Code*` | Relationship Code from source row |
| `Rto Cluster*` | UW Cluster from source row |
| `Span Outgo*` | `GWP` (Normal) / `IRDA` (Special trigger) |
| `Span Prct*` | Rate value (Normal) / `-0.1` (Special IRDA trigger) |
| `Biz Mix*` | From column configuration |
| Extra fields | Fixed values from column's extra fields config |
| `Rto Code*` | Comma-separated RTO codes from RTO lookup |

---

## Default Column Mappings

The portal ships with 54 pre-configured column mappings keyed by `(Parent BIZ, Sub-label)`
from the sheet headers. Auto-matching is case-insensitive.

### GCV (Goods Carrying Vehicle)

| Column | Biz Mix Output | GVW Range |
|---|---|---|
| GCV <=2.5 T | `GCV <=2.5 T` | 0 – 2,500 kg |
| GCV 2.5T - 2.8T | `GCV 2.5T - 2.8T` | 2,500 – 2,800 kg |
| GCV 2.8T - 3.5T | `GCV 2.8T - 3.5T` | 2,800 – 3,500 kg |
| GCV 3.5T - 7.5T | `GCV 3.5T - 7.5T` | 3,500 – 7,500 kg |
| GCV 7.5T - 12T | `GCV 7.5T - 12T` | 7,500 – 12,000 kg |
| GCV 12T-20T AGE<5 | `GCV 12T - 20T` | 12,000 – 20,000 kg, Vehicle Age < 5 |
| GCV 12T-20T AGE>=5 | `GCV 12T - 20T` | 12,000 – 20,000 kg, Vehicle Age ≥ 5 |
| GCV 20T-40T AGE<5 | `GCV 20T - 40T` | 20,000 – 40,000 kg, Vehicle Age < 5 |
| GCV 20T-40T AGE>=5 | `GCV 20T - 40T` | 20,000 – 40,000 kg, Vehicle Age ≥ 5 |
| GCV > 40T | `GCV > 40T` | > 40,000 kg |
| GCV-3W | `GCV-3W` | Three-wheeler GCV |

### PCV (Passenger Carrying Vehicle)

| Column | Biz Mix Output | Extra Fields |
|---|---|---|
| PCV 3W ELECTRIC | `PCV-3W` | Fuel Type = Electric |
| PCV 3W NEW | `PCV-3W` | Type Of Business = New, Fuel Type = Petrol, Diesel |
| PCV 3W OLD | `PCV-3W` | Type Of Business = Renewal, Roll Over |
| PCV-BUS_OTHER | `PCV-BUS` | Bus Type = Other Bus |
| PCV-BUS_SCHOOL | `PCV-BUS` | Bus Type = School Bus |
| PCV-TAXI | `PCV-TAXI` | — |

### Specialised Vehicles

| Column | Biz Mix Output | Extra Fields |
|---|---|---|
| TRACTOR NEW | `TRACTOR` | Type Of Business = New |
| TRACTOR OLD | `TRACTOR` | Type Of Business = Renewal, Roll Over |
| CE-CONSTRUCTION EQ | `CE-CONSTRUCTION EQ` | — |
| MISD GARBAGE | `MISD GARBAGE` | — |
| HARVESTER NEW | `HARVESTER NEW` | Type Of Business = New |
| HARVESTER OLD | `HARVESTER OLD` | Type Of Business = Renewal, Roll Over |

### 2W / 2W(1+1) / 2W(1+5) — per CC range

Each 2W variant has 5 sub-columns:

| Sub-column | Extra Fields |
|---|---|
| `<75CC` | Cubic Capacity Ul = 76 |
| `75-150CC` | Cubic Capacity Ll = 76, Ul = 151 |
| `150-350CC` | Cubic Capacity Ll = 151, Ul = 351 |
| `>350CC` | Cubic Capacity Ll = 351 |
| `SCOOTER` | Two Wheeler Category = Scooter |

### PVT CAR variants

| Column | Biz Mix Output | Extra Fields |
|---|---|---|
| PVT CAR(1+1) DIESEL & NCB | `PVT CAR(1+1)` | Fuel Type = Diesel, Ncb Ll = 1 |
| PVT CAR(1+1) DIESEL & ZERO NCB | `PVT CAR(1+1)` | Fuel Type = Diesel, Ncb Ul = 1 |
| PVT CAR(1+1) PETROL & NCB | `PVT CAR(1+1)` | Fuel Type = Petrol, Ncb Ll = 1 |
| PVT CAR(1+1) PETROL & ZERO NCB | `PVT CAR(1+1)` | Fuel Type = Petrol, Ncb Ul = 1 |
| PVT CAR(1+3) DIESEL | `PVT CAR(1+3)` | Fuel Type = Diesel |
| PVT CAR(1+3) PETROL | `PVT CAR(1+3)` | Fuel Type = Petrol |
| PVT CAR(3+3) ROLLOVER DIESEL & NCB | `PVT CAR(3+3)` | Type = Roll Over, Fuel = Diesel, Ncb Ll = 1 |
| PVT CAR(3+3) ROLLOVER DIESEL & ZERO NCB | `PVT CAR(3+3)` | Type = Roll Over, Fuel = Diesel, Ncb Ul = 1 |
| PVT CAR(3+3) ROLLOVER PETROL & NCB | `PVT CAR(3+3)` | Type = Roll Over, Fuel = Petrol, Ncb Ll = 1 |
| PVT CAR(3+3) ROLLOVER PETROL & ZERO NCB | `PVT CAR(3+3)` | Type = Roll Over, Fuel = Petrol, Ncb Ul = 1 |
| PVT CAR(3+3) NEW DIESEL | `PVT CAR(3+3)` | Type = New, Fuel = Diesel |
| PVT CAR(3+3) NEW PETROL | `PVT CAR(3+3)` | Type = New, Fuel = Petrol |
| PVT CAR DIESEL & NCB | `PVT CAR` | Fuel = Diesel, Ncb Ll = 1 |
| PVT CAR DIESEL & ZERO NCB | `PVT CAR` | Fuel = Diesel, Ncb Ul = 1 |
| PVT CAR PETROL & NCB | `PVT CAR` | Fuel = Petrol, Ncb Ll = 1 |
| PVT CAR PETROL & ZERO NCB | `PVT CAR` | Fuel = Petrol, Ncb Ul = 1 |

---

## RTO Lookup Logic

1. Look up `(UW Cluster, Product Category)` in the RTO vs Cluster index.
2. If not found, try again with the **normalized** category name (from Step 5 normalization map).
3. If still not found, fall back to **all RTO codes** in the file as a comma-separated string.
4. If no RTO file was uploaded, write `ANY`.

The fallback to all codes ensures no row is left with a missing RTO — but you should
verify the normalization map covers all categories in your data.

---

## Known Sheet Layouts

### Special Matrix (Special mode) — e.g. `Special Matrix-Comp.-1-31`

```
Row 1:   (blank / title)
Row 2:   Parent BIZ group labels  → Header Row 1
Row 3:   Sub-category labels       → Header Row 2
Row 4:   "GWP" labels              → Header Row 3
Row 5+:  Data rows
Cols 1-12:  Meta columns (IMD Code, Name, Rel Code, Type, Vol Ll/Ul, Consideration, Cluster etc.)
Cols 13+:   Rate columns (one per LOB sub-type)
```

**Mode:** Special | **Rate cols start:** 13 | **Data start:** 5

---

### Std Grid (Normal mode) — e.g. `Agency-Comp-Expt Pvt.Car.-8-30`

```
Row 1:   (blank / title)
Row 2:   Parent BIZ group labels  → Header Row 1
Row 3:   Sub-category labels       → Header Row 2
Row 4:   "GWP" labels              → Header Row 3
Row 5+:  Data rows
Col 1:   Vol Lower (GWP lower bound)
Col 2:   Vol Upper (GWP upper bound)
Col 3:   UW Cluster
Cols 4+: Rate columns
```

No IMD Code / Name / Type / Relationship Code columns exist. Set all IMD meta cols to **0**.

**Mode:** Normal | **Rate cols start:** 4 | **Data start:** 5
**IMD Code/Name/Rel/Type:** all 0

---

## API Reference

All routes accept and return JSON unless noted.

### `POST /api/upload`
Upload files. Form data with `file` (required) and `rto_file` (optional).
Returns `{ session_id, filepath, rto_filepath, sheets, filename }`.

### `POST /api/inspect`
Inspect sheet structure.
```json
{
  "filepath": "...",
  "sheet_name": "...",
  "header_rows": [2, 3, 4],
  "data_start_row": 5,
  "start_col": 13
}
```
Returns `{ col_defs, meta_defs, preview, vol_values, imd_values, max_row, max_col }`.
Each `col_def` includes `matched_default: true/false` indicating whether a built-in
default configuration was found for that column.

### `POST /api/process`
Run the full transformation.
```json
{
  "filepath": "...",
  "rto_filepath": "...",
  "session_id": "...",
  "output_name": "agency_grid_output",
  "rto_sheet": "RTO Vs Cluster (New)",
  "rto_header_row": 2,
  "rto_col": "RTO CODE",
  "rto_cluster_col": "UW CLUSTER (25-26)",
  "rto_cat_col": "PRODUCT CATEGORY",
  "config": { ... }
}
```
Returns `{ success, output_path, output_filename, rows, cols, skipped_rows, skipped_cells, preview, columns }`.

### `GET /api/download/<filename>`
Download a previously generated output CSV.

---

## Adding a New Column Type

If a new column type appears in the grid that isn't auto-matched, you have two options:

**Option A — Configure in the UI (Step 3):** Find the unmatched column (amber indicator),
fill in Biz Mix Output, RTO Category, and Extra Fields manually. This is per-run.

**Option B — Add to default config in `app.py` (permanent):** Add a new entry to
`DEFAULT_COL_CONFIG_KEYED` in `app.py`:

```python
DEFAULT_COL_CONFIG_KEYED = {
    # ...existing entries...

    # New column: e.g. GCV 40T - 55T
    ('GCV 40T - 55T', 'GCV 40T - 55T'): {
        'biz_mix_output': 'GCV 40T - 55T',
        'rto_category':   'GCV',
        'extra_fields': {
            'Gross Vehicle Weight Ll*': '40000.001',
            'Gross Vehicle Weight Ul*': '55000.001',
        }
    },
}
```

The key is `(PARENT_BIZ_UPPER, SUB_LABEL_UPPER)` matching exactly what appears in the
Excel header rows (case-insensitive).

---

## Adding a New GWP Column Mapping

When a new Volume Consideration value appears (e.g. a new LOB category):

1. Go to **Step 4 → GWP Columns & Agent Group Map**
2. Click **+ Add** in the mapping table
3. Enter the Vol Consideration value exactly as it appears in the sheet
4. Enter the LL and UL output column names for that category
5. Proceed with processing

To make it permanent, add to `VOL_GWP_DEFAULTS` in `index.html`:

```javascript
const VOL_GWP_DEFAULTS = {
    // ...existing entries...
    'gcv 40t - 55t': { ll: 'GCV 40 55t GWP Ll*', ul: 'GCV 40 55t GWP Ul*' },
};
```

---

## Troubleshooting

| Problem | Likely Cause | Fix |
|---|---|---|
| "No output rows generated" | IMD Code col set to wrong number, or wrong Data Start Row | Check Step 2 meta column positions. For Std Grid set IMD cols to 0. |
| All rows skipped | Header bleed guard matching real data | Remove or update "Skip row if Vol Lower equals" in Step 2 |
| `0` skipped cells, `0` output rows | Data start row is too high (skipping all data) | Decrease Data Start Row in Step 2 |
| RTO codes all `ANY` with file uploaded | Sheet name or column header mismatch | Check RTO sheet name and column headers in Step 5 |
| Column shows amber "manual" tag | Header text doesn't match any built-in default | Fill in Biz Mix and RTO Category manually in Step 3 |
| Sheet not found error | Wrong sheet name typed | Use the sheet dropdown which is auto-populated from the file |
| File too large | Over 200 MB limit | Split the grid or increase `MAX_CONTENT_LENGTH` in `app.py` |

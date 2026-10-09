# Grid Checker — Agency Grid Processor

Checks whether a converted grid CSV is correct **before upload**, including the manual
post-portal steps. Works on three layers; each layer is optional except the first.

| Layer | Needs | Catches |
|---|---|---|
| **OUTPUT** (S-checks) | the CSV | leftover `-`/blanks, wrong/missing columns, LL ≥ UL, volume in two columns or LL=999999999, un-scaled % (0.4 instead of 40.001), IRDA ≠ -0.1, rates < 1 (×100 missing), float artefacts, Biz Mix + Fuel/NCB/CC/Bus/ToB/Age combos that match no grid column, RTO list inconsistent per cluster / vs master, duplicates & conflicting rows |
| **SOURCE** (R-checks) | `--source` + `--sheet` | rebuilds the expected grid **from the Excel using only the rules file** (not the portal config) and compares every cell; missing/extra rows; unknown cell values; Volume Considerations / IMD Types / Biz-Mix labels with no rule; duplicated source blocks. Every mismatch points to the source cell and CSV line. |
| **CONFIG** (C-checks) | `--config` | portal config vs the actual sheet: header/data rows, shifted Step-3 columns, missing columns, Biz Mix/extra fields vs rules, unmapped Volume Considerations, LL/UL target pairs, extra-meta columns pointing at the wrong source column, defaults, ×100 transform, IRDA value, OD-basis columns, Version Id month |

## Run

```bash
pip install pandas openpyxl
python grid_checker.py \
  --csv    Special_Matrix-Comp_-1-14.csv \
  --rules  rules/special_comp_rules.json \
  --source 202609_Special_Motor_Matrix_Comp_Sep_26.xlsx --sheet "Special Matrix-Comp.-1-14" \
  --config agency_grid_config_special_comp_v1.json \
  --version-id agency_spl_sep26_Grid_3 \
  --report check_report.xlsx
```

Exit code `0` = no errors, `1` = errors found (do not upload), `2` = checker failed.
Runtime ≈ 30 s for 110k rows. Without `--source` only output checks run (≈ 10 s).

Recommended routine: run with `--config` **before** processing (catches setup mistakes),
then run on the final CSV **after** the manual steps.

## Inside the portal

The checker is built into the portal as **Step 9 Grid Checker** (upload the final CSV after manual changes) and
**Step 10 Checker Rules** (edit rules, switch checks on/off, change severity). See the portal README.

## The rules file is the source of truth

`rules/special_comp_rules.json` describes what a correct grid looks like, independent of the
portal config — otherwise a wrong config would validate itself. Edit it when the business rule
changes (new LOB column, new Volume Consideration, new default), not when the config changes.

Main sections: `std_grid_volume_rules` (ordered: Agent Group + Biz Mix condition → volume column, first match wins),
`checks` (per check: `{"enabled": false}` or `{"severity": "WARN"}`), `expected_columns`, `defaults`, `column_map` (header → Biz Mix + fields),
`imd_type_map` (Agent Group + STD-GRID volume pair), `volume_consideration_map`,
`bizmix_consideration_map` (which Prct Vol pair a "… on Overall Motor Biz" label feeds),
`rate_cell` (IRDA triggers, ×100, outgo per header label), `percent_band` (×100, UL +0.001),
`blocked_rtos`.

Entries marked `_note` / CONFIRM are assumptions to verify with the grid owner.

A different grid type (e.g. STP/TP) needs its own rules file — copy this one and change
`column_map`, `volume_consideration_map` and `allowed_values`. Without one, output checks still
run; the Biz-Mix signature check is skipped with a warning.

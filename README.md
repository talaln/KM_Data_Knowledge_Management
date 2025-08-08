# Knowledge Management System – Helpdesk Efficiency (Reproducible Demo)

This repository recreates the core analysis from the knowledge management (KMS) helpdesk study using **synthetic data**.

> Publication: *Using knowledge management tools in the Saudi National Mental Health Survey helpdesk: pre and post study* (2019) – https://doi.org/10.1186/s13033-019-0288-5

## What's inside
- `data/KM_Data.xlsx` — Synthetic dataset (sheet: `Combined`).
- `scripts/km_tables.sas` — SAS script to generate Tables 1–4.
- `out/` — Folder where outputs will be saved.
- `visuals/` — (Optional) screenshots of outputs.

## How to run
1. Open SAS (Base SAS/EG/Studio).
2. Ensure the working directory is the repo root.
3. Run: `scripts/km_tables.sas`

Outputs:
- `out/KM_Data_Tables.xlsx` — Excel with tables 1–4.
- `out/KM_Data_Tables.rtf` — (optional) RTF output.

## Data schema detected from your XLSX
- **Sheet used**: `Sheet1`
- **Total rows**: 991
- **Column mapping (auto-detected)**:
  - Category → `Category`
  - Priority → `Priority`
  - Phase → `Phase`
  - Response_Time_Hours → `Response_Time_Hours`
  - Support_level → `Support_level`

> The synthetic data preserves **structure and approximate distributions** only. No original records are included.

## License
Scripts: MIT. Data: synthetic (CC BY 4.0).

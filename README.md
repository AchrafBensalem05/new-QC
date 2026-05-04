# Daily Report Cleaner

This project cleans daily Excel production reports and combines them into a single table.

## Setup

1. Create and activate a Python environment.
2. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Input files

- Place Excel files anywhere under `./data` (year/month subfolders supported) or update the notebook path to your folder.
- The notebook scans subdirectories recursively.
- Each file should have a sheet named `Main Prod`.
- The date is parsed from the filename (example: `01-01-2023.xlsx`).
- In the Main Prod sheet, `CONC.` is treated as the field name.
- Total rows are excluded from the combined output.
- The notebook keeps `capacity`, `actual`, `gravity`, `gas prod`, `oil prod`, and `water` in the final CSV.
- `actual` comes from the workbook when present and falls back to `oil prod` for other days.
- The notebook handles `OIL`/`GAS` sheets and the `EXPORT`/`PRODUCED` October/November-style variants.

## Run

Open the notebook and run all cells. The output will be a single CSV file with these columns:

- date
- field
- station
- capacity
- oil prod
- actual
- gravity
- gas prod
- water

If the notebook cannot match a source column in a future workbook variant, update the column mapping cell with the exact headers from your Excel sheet.

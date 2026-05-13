import pandas as pd
from pathlib import Path

# ==========================================
# Configuration
# ==========================================
INPUT_ROOT = Path("./PARP")
OUTPUT_FILE = Path("./clean_data/combined_parp_reports.csv")
POSSIBLE_SHEET_NAMES = ["OAPARP61-B"]
HEADER_FILTER_COL_INDEX = 3  # Column4 in Power Query (0-based index).
SKIP_ROWS_AFTER_FILTER = 3
STANDARD_COLUMNS = [
    "source",
    "FIELD",
    "RESVR",
    "RESERVOIR",
    "WELL",
    "TEST DATE",
    "TEST LENGHT IN HOURS",
    "WELL TYPE",
    "PCT WTR",
    "GRAVITY",
    "64S",
    "PRES HEAD",
    "PSIG SEP",
    "DAILY OIL BBLS",
    "DAILY WATER BBLS",
    "GOR",
]
GROUP_HEADERS = {
    "WELL TEST DATE",
    "WELL STATUS",
    "WELL TEST DATA",
    "WELL TEST",
}
SUBHEADER_TOKENS = {
    "MTH",
    "DAY",
    "YR",
    "HRS",
    "CPR",
    "TPR",
    "SEP",
    "BOPD",
    "BWPD",
    "GOR",
}
COLUMN_ALIASES = {
    "FIELD NAME": "FIELD",
    "RESVR CODE": "RESVR",
    "RESERVOIR NAME": "RESERVOIR",
    "RESERVOIR DETAILED NAME": "RESERVOIR",
    "WELL NO": "WELL",
    "WELL": "WELL",
    "WELL STATUS": "WELL_1",
    "TEST DAY": "DAY",
    "WELL TEST DAY": "DAY",
    "TEST MONTH": "MTH",
    "TEST YEAR": "YR",
    "TEST LENGTH IN HOURS": "TEST LENGHT IN HOURS",
    "PERC BS&W": "PERC",
    "GTY & 60 F": "GTY &",
    "GRAV API": "GRAVITY",
}


def list_input_files():
    files = list(INPUT_ROOT.rglob("*.xls")) + list(INPUT_ROOT.rglob("*.xlsx"))
    return sorted([p for p in files if not p.name.startswith("~$")])


def choose_sheet(xl):
    if POSSIBLE_SHEET_NAMES:
        for name in POSSIBLE_SHEET_NAMES:
            if name in xl.sheet_names:
                return name
    return xl.sheet_names[0]


def build_headers(header_row):
    headers = []
    for idx, value in enumerate(header_row):
        name = "" if pd.isna(value) else str(value).strip()
        if not name or name.lower() == "nan":
            name = f"Column{idx + 1}"
        headers.append(name)
    return headers


def make_unique_columns(columns):
    counts = {}
    unique_cols = []
    for col in columns:
        base = str(col)
        if base in counts:
            counts[base] += 1
            unique_cols.append(f"{base}.{counts[base]}")
        else:
            counts[base] = 0
            unique_cols.append(base)
    return unique_cols


def normalize_col_name(value):
    text = "" if pd.isna(value) else str(value)
    text = " ".join(text.replace("\n", " ").split())
    return text.strip()


def is_subheader_row(row):
    values = {
        normalize_col_name(v).upper()
        for v in row.tolist()
        if pd.notna(v) and normalize_col_name(v)
    }
    return bool(values & SUBHEADER_TOKENS)


def merge_headers(primary, secondary):
    merged = list(primary)
    for idx, sub in enumerate(secondary):
        sub_clean = normalize_col_name(sub)
        if not sub_clean or sub_clean.lower().startswith("column"):
            continue
        prim_clean = normalize_col_name(primary[idx]).upper()
        if not prim_clean or prim_clean.startswith("COLUMN") or prim_clean in GROUP_HEADERS:
            merged[idx] = sub_clean
    return merged


def find_date_columns(columns):
    candidates = [
        ("MTH", "DAY", "YR"),
        ("MONTH", "DAY", "YEAR"),
    ]
    for month_col, day_col, year_col in candidates:
        if all(col in columns for col in [month_col, day_col, year_col]):
            return month_col, day_col, year_col
    return None


def locate_header_index(df):
    for idx, row in df.iterrows():
        values = {
            str(value).strip().upper()
            for value in row.tolist()
            if pd.notna(value) and str(value).strip()
        }
        has_field = "FIELD" in values or "FIELD NAME" in values
        has_resvr = "RESVR" in values or "RESVR CODE" in values
        has_date = (
            "TEST YEAR" in values
            or "YR" in values
            or "TEST MONTH" in values
            or "WELL TEST DATE" in values
            or "WELL TEST DAY" in values
        )
        if has_field and has_resvr and has_date:
            return idx
    return None


def ensure_source_column(df, file_name):
    source_col = None
    for col in df.columns:
        col_text = str(col).strip().lower()
        if col_text in {file_name.lower(), "source.name", "source"}:
            source_col = col
            break

    if source_col and source_col != "source":
        df = df.rename(columns={source_col: "source"})

    if "source" not in df.columns:
        df.insert(0, "source", file_name)
    else:
        df["source"] = df["source"].fillna(file_name)

    return df


def normalize_year_part(value):
    if pd.isna(value):
        return value
    text = str(value).strip()
    if not text:
        return text
    try:
        number = int(float(text))
    except ValueError:
        return text
    if number < 100:
        if number <= 29:
            return str(2000 + number)
        return str(1900 + number)
    return str(number)


def transform_file(path):
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    with pd.ExcelFile(path, engine=engine) as xl:
        sheet = choose_sheet(xl)
        df = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=str)

    if df.shape[1] <= HEADER_FILTER_COL_INDEX:
        return pd.DataFrame()

    df = df[df.iloc[:, HEADER_FILTER_COL_INDEX].notna()]
    if df.empty:
        return pd.DataFrame()

    header_idx = locate_header_index(df)
    if header_idx is not None:
        df = df.loc[header_idx:].copy()
    else:
        if len(df) <= SKIP_ROWS_AFTER_FILTER:
            return pd.DataFrame()
        df = df.iloc[SKIP_ROWS_AFTER_FILTER:].copy()

    header_row = df.iloc[0]
    subheader_row = df.iloc[1] if len(df) > 1 else None

    headers = build_headers(header_row)
    if subheader_row is not None and is_subheader_row(subheader_row):
        subheaders = build_headers(subheader_row)
        headers = merge_headers(headers, subheaders)
        df = df.iloc[2:].copy()
    else:
        df = df.iloc[1:].copy()

    normalized = [normalize_col_name(c) for c in headers]
    mapped = [COLUMN_ALIASES.get(name, name) for name in normalized]
    df.columns = make_unique_columns(mapped)

    df = ensure_source_column(df, path.name)

    drop_cols = [
        "****************W E L L  T E S T  D A T A ****************",
        "Column20",
    ]
    df = df.drop(columns=[c for c in drop_cols if c in df.columns])

    rename_map = {
        "GTY &": "GRAVITY",
        "WELL_1": "WELL TYPE",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    date_cols = find_date_columns(df.columns)
    if date_cols:
        month_col, day_col, year_col = date_cols
        df[year_col] = df[year_col].apply(normalize_year_part)
        df = df[df[day_col].notna()]
        df = df[df[day_col].astype(str).str.strip().str.upper() != "DAY"]
        date_text = (
            df[month_col].fillna("").astype(str).str.strip()
            + "/"
            + df[day_col].fillna("").astype(str).str.strip()
            + "/"
            + df[year_col].fillna("").astype(str).str.strip()
        )
        df["TEST DATE"] = date_text.str.strip("/")
        df = df.drop(columns=[month_col, day_col, year_col])
    elif "WELL TEST DATE" in df.columns:
        df["TEST DATE"] = pd.to_datetime(
            df["WELL TEST DATE"], errors="coerce", dayfirst=True
        ).dt.date
        df = df.drop(columns=["WELL TEST DATE"])

    rename_map2 = {
        "Column9": "TEST LENGHT IN HOURS",
        "PERC": "PCT WTR",
        "Column15": "PRES HEAD",
        "Column16": "PSIG SEP",
        "Column17": "DAILY OIL BBLS",
        "Column18": "DAILY WATER BBLS",
        "Column19": "GOR",
    }
    df = df.rename(columns={k: v for k, v in rename_map2.items() if k in df.columns})

    if "TEST DATE" in df.columns:
        df["TEST DATE"] = pd.to_datetime(df["TEST DATE"], errors="coerce", dayfirst=True).dt.date

    if all(c in df.columns for c in ["FIELD", "WELL", "TEST DATE"]):
        df = df.drop_duplicates(subset=["FIELD", "WELL", "TEST DATE"])

    for col in STANDARD_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
    df = df[STANDARD_COLUMNS]

    return df


def main():
    files = list_input_files()
    if not files:
        print(f"No Excel files found in {INPUT_ROOT}")
        return

    print(f"Found {len(files)} file(s). Processing...")

    all_data = []
    for path in files:
        try:
            cleaned = transform_file(path)
        except Exception as exc:
            print(f"Error reading {path.name}: {exc}")
            continue

        if cleaned.empty:
            print(f"No data extracted from {path.name}")
            continue

        all_data.append(cleaned)

    if not all_data:
        print("No data processed.")
        return

    combined = pd.concat(all_data, ignore_index=True)
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(OUTPUT_FILE, index=False)

    print(f"Wrote {len(combined)} rows to {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

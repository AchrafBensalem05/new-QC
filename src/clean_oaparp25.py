import pandas as pd
import re
from pathlib import Path

# ==============================
# Config
# ==============================
INPUT_ROOT = Path("./PARP")
OUTPUT_CSV = Path("./clean_data/parp_cleand_new.csv")
SHEET_NAME_25 = "OAPARP25"

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
HEADER_ALIASES = {
    "TEST MON TH": "TEST MONTH",
    "TEST MON": "TEST MONTH",
    "TEST YR": "TEST YEAR",
    "TEST DY": "TEST DAY",
    "WELL NO": "WELL",
}

def get_year_from_filename(filename):
    """
    Extracts a 4-digit year from the filename (e.g., '02 FEB 2000...') 
    to use as a reliable truth for missing report years.
    """
    match = re.search(r'\b(19[89]\d|20[0-2]\d)\b', str(filename))
    if match:
        return int(match.group(1))
    return None

def normalize_col_name(value):
    text = "" if pd.isna(value) else str(value)
    text = " ".join(text.replace("\n", " ").split())
    return text.strip()

def build_headers(row):
    headers = []
    for idx, value in enumerate(row):
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

def locate_header_index(df):
    for idx, row in df.iterrows():
        values = {
            normalize_col_name(v).upper()
            for v in row.tolist()
            if pd.notna(v) and normalize_col_name(v)
        }
        has_field = "FIELD" in values or "FIELD NAME" in values
        has_resvr = "RESVR" in values or "RESVR CODE" in values
        has_date = (
            "WELL TEST DATE" in values
            or "TEST DATE" in values
            or "TEST YEAR" in values
            or "YR" in values
            or "TEST MONTH" in values
            or "TEST DAY" in values
            or "WELL TEST DAY" in values
            )
        if has_field and has_resvr and has_date:
            return idx
    return None

def normalize_year_value(value, file_year=None):
    """
    Converts 2-digit years into 4-digit years. Treats missing values, empty strings, 
    and '0' as the file's year (falling back to 2000 if not found).
    """
    default_y = 2000
    
    if pd.isna(value):
        return default_y
        
    text = str(value).strip().lower()
    if not text or text == "nan":
        return default_y
        
    try:
        number = int(float(text))
    except ValueError:
        return value
        
    if number == 0:
        return default_y  
        
    if number == 1:
        return 2001
    if number >= 1900:
        return number
    if number > 60:
        return 1900 + number
    if number < 27:
        return 2000 + number
    return number

def clean_date_part(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    try:
        num = float(text)
    except ValueError:
        return text
    if num.is_integer():
        return str(int(num))
    return text

def find_column(columns, candidates):
    for name in candidates:
        if name in columns:
            return name
    return None

def fix_daily_volume_headers(columns):
    # Some files duplicate "DAILY OIL BBLS" where water should be; fix by order.
    cols = list(columns)
    def base_name(value):
        text = normalize_col_name(value)
        return re.sub(r"\.\d+$", "", text).strip()

    normalized = [base_name(c).upper() for c in cols]
    expected = [
        "DAILY OIL BBLS",
        "DAILY WATER BBLS",
        "DAILY NETGAS SMCF",
        "DAILY GLGAS SMCF",
    ]

    anchor_index = None
    for anchor in ("GRAV API", "GRAVITY"):
        if anchor in normalized:
            anchor_index = normalized.index(anchor)
            break

    if anchor_index is not None:
        start = anchor_index + 1
        end = start + len(expected)
        if end <= len(cols):
            window = normalized[start:end]
            if window != expected:
                for offset, name in enumerate(expected):
                    cols[start + offset] = name
        return cols

    oil_indices = [i for i, v in enumerate(normalized) if v == "DAILY OIL BBLS"]
    if len(oil_indices) >= 2 and "DAILY WATER BBLS" not in normalized:
        cols[oil_indices[1]] = "DAILY WATER BBLS"
    return cols

def parse_sheet(df):
    header_idx = locate_header_index(df)
    if header_idx is None:
        return pd.DataFrame()

    start = df.loc[header_idx:].copy()
    header_row = start.iloc[0]
    subheader_row = start.iloc[1] if len(start) > 1 else None

    headers = build_headers(header_row)
    if subheader_row is not None and is_subheader_row(subheader_row):
        subheaders = build_headers(subheader_row)
        headers = merge_headers(headers, subheaders)
        data = start.iloc[2:].copy()
    else:
        data = start.iloc[1:].copy()

    headers = [normalize_col_name(h) for h in headers]
    headers = [HEADER_ALIASES.get(h, h) for h in headers]
    data.columns = make_unique_columns(headers)
    data = data.dropna(how="all")
    return data

def clean_sheet(path):
    # Extract year from filename to fill missing years
    file_year = get_year_from_filename(path.name)
    
    engine = "xlrd" if path.suffix.lower() == ".xls" else "openpyxl"
    with pd.ExcelFile(path, engine=engine) as xl:
        if SHEET_NAME_25 not in xl.sheet_names:
            return pd.DataFrame()
        df = pd.read_excel(xl, sheet_name=SHEET_NAME_25, header=None, dtype=str)

    data = parse_sheet(df)
    if data.empty:
        return data

    if "WELL" in data.columns:
        well_text = data["WELL"].astype(str).str.strip()
        data = data[data["WELL"].notna() & (well_text != "") & (well_text.str.lower() != "nan")]

    year_col = find_column(data.columns, ["TEST YEAR", "YR", "YEAR"])
    hours_col = find_column(
        data.columns,
        ["TEST LENGTH IN HOURS", "TEST LENGTH IN HOUR", "TEST LENGTH HRS", "HRS"],
    )

    # 1. Drop rows where hours are empty, missing, or zero
    if hours_col:
        hours_text = data[hours_col].astype(str).str.strip().str.lower()
        invalid_hours = (
            data[hours_col].isna() | 
            (hours_text == "") | 
            (hours_text == "nan") | 
            (hours_text == "0") | 
            (hours_text == "0.0")
        )
        if invalid_hours.any():
            data = data[~invalid_hours].copy()

    # 2. Normalize year values (replaces empty/0/NaN with file_year automatically)
    if year_col:
        data[year_col] = data[year_col].apply(lambda x: normalize_year_value(x, file_year))

    # Format Date
    date_parts = ["TEST DAY", "TEST MONTH", "TEST YEAR"]
    if all(part in data.columns for part in date_parts):
        day = data["TEST DAY"].apply(clean_date_part)
        month = data["TEST MONTH"].apply(clean_date_part)
        year = data["TEST YEAR"].apply(clean_date_part)
        date_text = day + "/" + month + "/" + year
        data["TEST DATE"] = pd.to_datetime(date_text, errors="coerce", dayfirst=True).dt.strftime("%m/%d/%Y")

    for col in data.columns:
        if col == "TEST DATE":
            continue
        data[col] = data[col].replace(r"^\s*$", 0, regex=True).fillna(0)

    rename_map = {
        "WELL.1": "WELL TYPE",
        "GTY &": "GRAVITY",
        "BOPD": "DAILY OIL PRODUCTION",
        "BWPD": "WATER PRODUCTION",
        "PERC": "BSW",
    }
    data = data.rename(columns={k: v for k, v in rename_map.items() if k in data.columns})

    data.columns = make_unique_columns(fix_daily_volume_headers(data.columns))

    drop_cols = ["...", "WELL TEST DATA", "Column19", "Column1", "Column26"]
    data = data.drop(columns=[c for c in drop_cols if c in data.columns], errors="ignore")

    for col in data.select_dtypes(include="object").columns:
        data[col] = data[col].astype(str).str.strip()

    data.insert(0, "sheet", SHEET_NAME_25)
    data.insert(0, "source", path.name)

    return data


def main():
    files = sorted([
        *INPUT_ROOT.rglob("*.xls"), 
        *INPUT_ROOT.rglob("*.xlsx"), 
        *INPUT_ROOT.rglob("*.xlsm")
    ])
    
    if not files:
        print(f"No Excel files found in {INPUT_ROOT}")
        return

    all_data = []
    for path in files:
        cleaned = clean_sheet(path)
        if cleaned.empty:
            print(f"No data found in {path.name}")
            continue
        all_data.append(cleaned)

    combined = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    OUTPUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(OUTPUT_CSV, index=False)

    print(f"Wrote {len(combined)} rows to {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
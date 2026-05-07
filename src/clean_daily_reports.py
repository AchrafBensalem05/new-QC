import re
import pandas as pd
from pathlib import Path
from datetime import datetime, date
import numpy as np

# ==========================================
# Configuration
# ==========================================
INPUT_ROOT = Path("./data")
POSSIBLE_SHEET_NAMES = ["Main Prod", "Main Production"]
HEADER_ROW = 11
OUTPUT_CSV = Path("./clean_data/combined_daily_reports_20_25.csv")

# Add the years you want to process here. 
TARGET_YEARS = [2016] 

# Column Mapping Strategy
COLUMN_MAP = {
    "field": ["conc", "field", "area"],
    "station": ["stations", "station", "stn"],
    "capacity": ["capacity", "cap"],
    "oil_prod": ["export", "oil", "metered"], 
    "actual": ["produced", "actual", "corrected"], 
    "gravity": ["gravity", "api"],
    "bsw": ["bs w", "bsw", "bs&w"],
    "water": ["water", "bwpd"],
    "gas": ["gas", "mmscfd"],
}

def normalize(text):
    """Normalize text for robust column matching."""
    if pd.isna(text):
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(text).lower()).strip()

def parse_date_from_filename(name):
    """Extract date from filename."""
    match = re.search(r"(\d{4}[-_ ]\d{1,2}[-_ ]\d{1,2}|\d{1,2}[-_ ]\d{1,2}[-_ ]\d{2,4})", name)
    if not match:
        return None
    
    raw = match.group(1).replace("_", "-").replace(" ", "-")
    
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%m-%d-%Y", "%d-%m-%y", "%m-%d-%y"):
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None

def normalize_field_name(value):
    """Map raw field/area names to standard names."""
    text = normalize(value)
    if not text or "pipe line" in text or "pipeline" in text:
        return None
    if "gialo" in text: return "GIALO"
    if "waha" in text: return "WAHA"
    if "dahra" in text: return "DAHRA"
    if "samah" in text: return "SAMAH"
    if "belhedan" in text: return "BELHEDAN"
    if "woc" in text: return "WOC"
    return None

def pick_column(available_columns, candidates):
    """Find the best matching column from available names."""
    norm_map = {normalize(c): c for c in available_columns}
    for cand in candidates:
        norm_cand = normalize(cand)
        if norm_cand in norm_map:
            return norm_map[norm_cand]
    return None

def process_files():
    # ---------------------------------------------------------
    # MODIFIED SECTION: Pre-filter files based on TARGET_YEARS
    # ---------------------------------------------------------
    raw_files = INPUT_ROOT.rglob("*.xlsx")
    
    if TARGET_YEARS:
        files = [
            p for p in raw_files 
            if not p.name.startswith("~$") 
            and any(str(year) in p.name for year in TARGET_YEARS)
        ]
    else:
        files = [p for p in raw_files if not p.name.startswith("~$")]
        
    files = sorted(files)
    
    print(f"Found {len(files)} files matching the target years. Starting processing...")
    # ---------------------------------------------------------

    all_data = []
    found_dates = set() 
    
    for path in files:
        report_date = parse_date_from_filename(path.name)
        if not report_date:
            print(f"Warning: Could not parse date from {path.name}. Skipping.")
            continue
            
        # Keep this check as a safeguard in case a filename happens to contain 
        # a year-like number (e.g., "report_2020_10-12-2019.xlsx")
        if TARGET_YEARS and report_date.year not in TARGET_YEARS:
            continue
            
        try:
            # Using pd.ExcelFile as a context manager for efficient sheet discovery
            with pd.ExcelFile(path, engine='openpyxl') as xl:
                target_sheet = None
                # Check for sheet name variations
                for s in xl.sheet_names:
                    if s.strip() in POSSIBLE_SHEET_NAMES:
                        target_sheet = s
                        break
                
                if not target_sheet:
                    print(f"Warning: Target sheet not found in {path.name}. Sheets: {xl.sheet_names}")
                    continue
                
                # Load the data from the identified sheet
                df = pd.read_excel(
                    xl, 
                    sheet_name=target_sheet, 
                    header=HEADER_ROW, 
                    usecols="A:J",
                    engine='openpyxl'
                )
        except Exception as e:
            print(f"Error reading {path.name}: {e}")
            continue
            
        # Clean column names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Detect source columns
        sources = {}
        for key, candidates in COLUMN_MAP.items():
            sources[key] = pick_column(df.columns, candidates)
            
        # Validate required columns
        required = ["field", "station", "oil_prod"]
        missing = [r for r in required if not sources[r]]
        if missing:
            print(f"Warning: Missing required columns {missing} in {path.name}. Skipping.")
            continue
            
        # Create processed dataframe
        processed = pd.DataFrame()
        processed["date"] = [report_date] * len(df)
        
        # Field and Station
        processed["field"] = df[sources["field"]].ffill().map(normalize_field_name)
        processed["station"] = df[sources["station"]].astype(str).str.strip()
        
        # Metrics
        processed["capacity"] = pd.to_numeric(df[sources["capacity"]], errors="coerce") if sources["capacity"] else np.nan
        processed["oil_prod"] = pd.to_numeric(df[sources["oil_prod"]], errors="coerce")
        
        if sources["actual"]:
            processed["actual"] = pd.to_numeric(df[sources["actual"]], errors="coerce")
        else:
            processed["actual"] = processed["oil_prod"]
            
        # Other metrics
        for key in ["gravity", "bsw", "water", "gas"]:
            if sources[key]:
                processed[key] = pd.to_numeric(df[sources[key]], errors="coerce")
            else:
                processed[key] = np.nan
        
        # Filter out rows that are totals or have empty field/station
        total_mask = (
            processed["field"].isna() |
            processed["station"].str.contains("total", case=False, na=False) |
            (processed["station"] == "nan") |
            (processed["station"] == "")
        )
        
        clean_df = processed[~total_mask].copy()
        clean_df["actual"] = clean_df["actual"].fillna(clean_df["oil_prod"])
        
        all_data.append(clean_df)
        found_dates.add(report_date)
        
    if not all_data:
        print("No data processed!")
        return
        
    combined = pd.concat(all_data, ignore_index=True)
    combined = combined.sort_values(["date", "field", "station"])
    combined.to_csv(OUTPUT_CSV, index=False)
    
    print(f"\nSuccess! Wrote {len(combined)} rows to {OUTPUT_CSV}")
    
    print("\n" + "="*40)
    print("📅 DATA COMPLETENESS REPORT")
    print("="*40)
    
    combined['year'] = pd.to_datetime(combined['date']).dt.year
    years_processed = combined['year'].unique()
    print(f"Total Rows by Year:\n{combined.groupby('year').size()}\n")
    
    years_to_check = TARGET_YEARS if TARGET_YEARS else years_processed
    
    for year in sorted(years_to_check):
        start_date = date(year, 1, 1)
        end_date = date(year, 12, 31)
        
        if year == datetime.today().year:
            end_date = min(end_date, datetime.today().date())
            
        full_year_dates = set(pd.date_range(start_date, end_date).date)
        found_this_year = {d for d in found_dates if d.year == year}
        missing_dates = sorted(full_year_dates - found_this_year)
        
        if not missing_dates:
            print(f"✅ {year}: Perfect! 0 missing dates (Out of {len(full_year_dates)} days).")
        else:
            print(f"❌ {year}: {len(missing_dates)} missing dates (Out of {len(full_year_dates)} days).")
            missing_strs = [d.strftime("%Y-%m-%d") for d in missing_dates]
            
            for i in range(0, len(missing_strs), 5):
                print(f"    {', '.join(missing_strs[i:i+5])}")
        print("-" * 40)

if __name__ == "__main__":
    process_files()
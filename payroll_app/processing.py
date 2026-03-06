# import io
# import pandas as pd
# from rapidfuzz import fuzz, process

# def read_flexible_file(content, filename):
#     """Reads CSV or Excel files from bytes."""
#     try:
#         if filename.lower().endswith(".csv"):
#             return pd.read_csv(io.BytesIO(content))
#         if filename.lower().endswith((".xlsx", ".xls")):
#             return pd.read_excel(io.BytesIO(content))
#     except Exception as e:
#         print(f"Error reading {filename}: {e}")
#         return None
#     return None

# def get_fuzzy_name_mapper(target_names, threshold=70):
#     """
#     Creates a helper function to map names to a master list.
#     Standardizes names to Upper Case and Stripped whitespace.
#     """
#     def _match(name):
#         if pd.isna(name) or str(name).strip() == "":
#             return name
#         name_str = str(name).upper().strip()
#         # Find best match in the ADP master names
#         candidate = process.extractOne(name_str, target_names, scorer=fuzz.token_sort_ratio)
#         if candidate and candidate[1] >= threshold:
#             return candidate[0]
#         return name_str # Return original if no good match found
#     return _match

# def process_adp(contents):
#     """Processes ADP CSV files into a Daily Hours Summary."""
#     dfs = []
#     for raw in contents:
#         df = pd.read_csv(io.BytesIO(raw))
#         # Ensure required columns exist
#         if "Payroll Name" in df.columns:
#             dfs.append(df)
            
#     if not dfs:
#         return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None
        
#     df = pd.concat(dfs)
#     # Clean ADP Names (Standard: LASTNAME FIRSTNAME)
#     df["Driver"] = df["Payroll Name"].astype(str).str.replace(",", "").str.strip().str.upper()
#     df["Date"] = pd.to_datetime(df["Pay Date"]).dt.date
#     df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)
    
#     grouped = df.groupby(["Driver", "Date"], as_index=False)["Hours"].sum()
#     return grouped, None

# def process_relay(contents):
#     """Processes Relay CSV files and calculates trip duration."""
#     dfs = []
#     for raw in contents:
#         dfs.append(pd.read_csv(io.BytesIO(raw)))
    
#     if not dfs:
#         return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None
        
#     df = pd.concat(dfs)
    
#     # Cleaning time columns
#     date_cols = ["Stop 1  Actual Arrival Date", "Stop 2  Actual Arrival Date", "Stop 1 Planned Arrival Date"]
#     for col in date_cols:
#         if col in df.columns:
#             df[col] = df[col].astype(str).str.strip()

#     # Convert to datetime
#     df["Stop1_Actual"] = pd.to_datetime(df[date_cols[0]] + " " + df["Stop 1  Actual Arrival Time"], errors="coerce")
#     df["Stop2_Actual"] = pd.to_datetime(df[date_cols[1]] + " " + df["Stop 2  Actual Arrival Time"], errors="coerce")
#     df["Stop1_Planned"] = pd.to_datetime(df[date_cols[2]] + " " + df["Stop 1 Planned Arrival Time"], errors="coerce")
    
#     df = df.dropna(subset=["Stop1_Actual", "Stop2_Actual"])
    
#     rows = []
#     for _, gp in df.groupby("Trip ID"):
#         gp = gp.sort_values("Stop1_Actual")
#         start_act, end_act = gp.iloc[0]["Stop1_Actual"], gp.iloc[-1]["Stop2_Actual"]
        
#         # Handle overnight trips
#         if end_act < start_act: 
#             end_act += pd.Timedelta(days=1)
        
#         raw_hours = (end_act - start_act).total_seconds() / 3600
#         # If trip is suspiciously long, fallback to planned start
#         start = gp.iloc[0]["Stop1_Planned"] if raw_hours > 20 and pd.notna(gp.iloc[0]["Stop1_Planned"]) else start_act
        
#         rows.append([
#             gp["Driver Name"].str.upper().iloc[0].strip(), 
#             start.date(), 
#             round((end_act - start).total_seconds() / 3600, 2)
#         ])

#     return pd.DataFrame(rows, columns=["Driver", "Date", "Hours"]), None

# def build_final_dataset(adp_df, relay_df, dp_df):
#     """Merges all datasets using ADP as the Master List for fuzzy matching."""
#     # 1. Establish Master List from ADP
#     adp_names = adp_df["Driver"].unique().tolist()
#     mapper = get_fuzzy_name_mapper(adp_names)

#     # 2. Apply Fuzzy Mapping to Relay and DriverPay
#     relay_df["Driver"] = relay_df["Driver"].apply(mapper)
#     dp_df["Driver"] = dp_df["Driver"].astype(str).str.upper().str.strip().apply(mapper)
    
#     # 3. Pivot Tables
#     relay_pivot = relay_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
#     relay_pivot.columns = [f"R_{col}" for col in relay_pivot.columns]
    
#     adp_pivot = adp_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
#     adp_pivot.columns = [f"A_{col}" for col in adp_pivot.columns]

#     # 4. Merge Data
#     final_df = pd.concat([relay_pivot, adp_pivot], axis=1).fillna(0).reset_index()
#     final_df = final_df.merge(dp_df, on="Driver", how="left")
    
#     # Organize columns
#     relay_cols = sorted([c for c in final_df.columns if c.startswith("R_")])
#     adp_cols = sorted([c for c in final_df.columns if c.startswith("A_")])
#     base_cols = [c for c in final_df.columns if not c.startswith(("R_", "A_"))]
    
#     return final_df[base_cols + relay_cols + adp_cols], relay_cols, adp_cols

# def build_override_dict(override_df, start_date, end_date, adp_df):
#     """Filters overrides by date and applies fuzzy matching to map to ADP drivers."""
#     if override_df is None or override_df.empty: 
#         return {}
    
#     # Get master names for fuzzy matching
#     adp_names = adp_df["Driver"].unique().tolist()
#     mapper = get_fuzzy_name_mapper(adp_names)
    
#     # Clean Override data
#     override_df["Driver"] = override_df["Driver"].astype(str).apply(mapper)
#     override_df["Date"] = pd.to_datetime(override_df["Date"]).dt.date
    
#     # Date range filtering
#     mask = (override_df["Date"] >= start_date) & (override_df["Date"] <= end_date)
#     filtered_df = override_df.loc[mask]
    
#     # Return as a lookup dictionary: {(Driver, Date): Price}
#     return {(row["Driver"], row["Date"]): row["Override Price"] for _, row in filtered_df.iterrows()}

import io
import pandas as pd
from rapidfuzz import fuzz, process

def read_flexible_file(content, filename):
    """Reads CSV or Excel files from bytes."""
    try:
        if filename.lower().endswith(".csv"):
            return pd.read_csv(io.BytesIO(content))
        if filename.lower().endswith((".xlsx", ".xls")):
            return pd.read_excel(io.BytesIO(content))
    except Exception as e:
        print(f"Error reading {filename}: {e}")
        return None
    return None

def get_fuzzy_name_mapper(target_names, threshold=70):
    """Maps names to ADP Master list using fuzzy matching."""
    def _match(name):
        if pd.isna(name) or str(name).strip() == "":
            return name
        name_str = str(name).upper().strip()
        candidate = process.extractOne(name_str, target_names, scorer=fuzz.token_sort_ratio)
        if candidate and candidate[1] >= threshold:
            return candidate[0]
        return name_str 
    return _match

def process_adp(contents):
    """Processes ADP CSV files into a Daily Hours Summary."""
    dfs = []
    for raw in contents:
        df = pd.read_csv(io.BytesIO(raw))
        if "Payroll Name" in df.columns:
            dfs.append(df)
            
    if not dfs:
        return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None
        
    df = pd.concat(dfs)
    df["Driver"] = df["Payroll Name"].astype(str).str.replace(",", "").str.strip().str.upper()
    df["Date"] = pd.to_datetime(df["Pay Date"]).dt.date
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)
    
    grouped = df.groupby(["Driver", "Date"], as_index=False)["Hours"].sum()
    return grouped, None

def process_relay(contents):
    """Processes Relay CSV files and calculates trip duration."""
    all_rows = []
    for raw in contents:
        df = pd.read_csv(io.BytesIO(raw))
        
        # Cleaning time columns
        date_cols = ["Stop 1  Actual Arrival Date", "Stop 2  Actual Arrival Date", "Stop 1 Planned Arrival Date"]
        # Standardizing headers (handling potential extra spaces)
        df.columns = [c.strip() for c in df.columns]
        
        # Mapping specific time/date columns
        # Using exact strings based on your previous provided code
        s1_date = "Stop 1  Actual Arrival Date"
        s1_time = "Stop 1  Actual Arrival Time"
        s2_date = "Stop 2  Actual Arrival Date"
        s2_time = "Stop 2  Actual Arrival Time"
        p1_date = "Stop 1 Planned Arrival Date"
        p1_time = "Stop 1 Planned Arrival Time"

        # Convert to datetime
        df["Stop1_Actual"] = pd.to_datetime(df[s1_date] + " " + df[s1_time], errors="coerce")
        df["Stop2_Actual"] = pd.to_datetime(df[s2_date] + " " + df[s2_time], errors="coerce")
        df["Stop1_Planned"] = pd.to_datetime(df[p1_date] + " " + df[p1_time], errors="coerce")
        
        df = df.dropna(subset=["Stop1_Actual", "Stop2_Actual"])
        
        for _, gp in df.groupby("Trip ID"):
            gp = gp.sort_values("Stop1_Actual")
            start_act, end_act = gp.iloc[0]["Stop1_Actual"], gp.iloc[-1]["Stop2_Actual"]
            
            if end_act < start_act: 
                end_act += pd.Timedelta(days=1)
            
            raw_hours = (end_act - start_act).total_seconds() / 3600
            start = gp.iloc[0]["Stop1_Planned"] if raw_hours > 20 and pd.notna(gp.iloc[0]["Stop1_Planned"]) else start_act
            
            all_rows.append([
                str(gp["Driver Name"].iloc[0]).upper().strip(), 
                start.date(), 
                round((end_act - start).total_seconds() / 3600, 2)
            ])

    if not all_rows:
        return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None
        
    result_df = pd.DataFrame(all_rows, columns=["Driver", "Date", "Hours"])
    return result_df, None

def build_final_dataset(adp_df, relay_df, dp_df):
    """
    Merges datasets and handles the Relay Date Drop requirement.
    1. Establish Master List from ADP.
    2. Drops the first chronological date column from Relay pivot.
    3. Sorts columns to ensure 1st 7 days = Week 1, next 7 days = Week 2.
    """
    adp_names = adp_df["Driver"].unique().tolist()
    mapper = get_fuzzy_name_mapper(adp_names)

    # Apply Fuzzy Mapping
    relay_df["Driver"] = relay_df["Driver"].apply(mapper)
    dp_df["Driver"] = dp_df["Driver"].astype(str).str.upper().str.strip().apply(mapper)
    
    # --- Pivot Relay Data ---
    relay_pivot = relay_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
    
    # Sort Relay dates chronologically
    sorted_dates = sorted(relay_pivot.columns)
    
    # REQUIREMENT: Drop the first chronological date column (index 0).
    # The relay data spans 15 days; dropping the first leaves exactly 14
    # which split cleanly into W1 (days 1-7) and W2 (days 8-14).
    if len(sorted_dates) > 0:
        first_date = sorted_dates[0]
        relay_pivot = relay_pivot.drop(columns=[first_date])
        remaining_dates = sorted_dates[1:]  # all dates after the dropped first
    else:
        remaining_dates = []

    # Apply R_ prefix for the remaining dates
    relay_pivot.columns = [f"R_{col}" for col in relay_pivot.columns]
    
    # Pivot ADP Data
    adp_pivot = adp_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
    adp_pivot = adp_pivot.reindex(sorted(adp_pivot.columns), axis=1)
    adp_pivot.columns = [f"A_{col}" for col in adp_pivot.columns]

    # Merge Data
    final_df = pd.concat([relay_pivot, adp_pivot], axis=1).fillna(0).reset_index()
    final_df = final_df.merge(dp_df, on="Driver", how="left")
    
    # Define Column Order (Base info -> Relay Dates -> ADP Dates)
    relay_cols = [f"R_{d}" for d in remaining_dates]
    adp_cols = [c for c in final_df.columns if c.startswith("A_")]
    base_cols = [c for c in final_df.columns if not c.startswith(("R_", "A_"))]
    
    return final_df[base_cols + relay_cols + adp_cols], relay_cols, adp_cols

def build_override_dict(override_df, start_date, end_date, adp_df):
    """Filters overrides and maps to ADP master drivers."""
    if override_df is None or override_df.empty: 
        return {}
    
    adp_names = adp_df["Driver"].unique().tolist()
    mapper = get_fuzzy_name_mapper(adp_names)
    
    override_df["Driver"] = override_df["Driver"].astype(str).apply(mapper)
    override_df["Date"] = pd.to_datetime(override_df["Date"]).dt.date
    
    mask = (override_df["Date"] >= start_date) & (override_df["Date"] <= end_date)
    filtered_df = override_df.loc[mask]
    
    return {(row["Driver"], row["Date"]): row["Override Price"] for _, row in filtered_df.iterrows()}
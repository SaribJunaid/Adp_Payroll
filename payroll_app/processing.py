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
#     """Maps names to ADP Master list using fuzzy matching."""
#     def _match(name):
#         if pd.isna(name) or str(name).strip() == "":
#             return name
#         name_str = str(name).upper().strip()
#         candidate = process.extractOne(name_str, target_names, scorer=fuzz.token_sort_ratio)
#         if candidate and candidate[1] >= threshold:
#             return candidate[0]
#         return name_str 
#     return _match

# def process_adp(contents):
#     """Processes ADP CSV files into a Daily Hours Summary."""
#     dfs = []
#     for raw in contents:
#         df = pd.read_csv(io.BytesIO(raw))
#         if "Payroll Name" in df.columns:
#             dfs.append(df)
            
#     if not dfs:
#         return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None
        
#     df = pd.concat(dfs)
#     df["Driver"] = df["Payroll Name"].astype(str).str.replace(",", "").str.strip().str.upper()
#     df["Date"] = pd.to_datetime(df["Pay Date"]).dt.date
#     df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce").fillna(0)
    
#     grouped = df.groupby(["Driver", "Date"], as_index=False)["Hours"].sum()
#     return grouped, None

# def process_relay(contents):
#     """Processes Relay CSV files and calculates trip duration."""
#     all_rows = []
#     for raw in contents:
#         df = pd.read_csv(io.BytesIO(raw))

#         # Standardize headers (strip extra spaces)
#         df.columns = [c.strip() for c in df.columns]

#         s1_act_date = "Stop 1  Actual Arrival Date"
#         s1_act_time = "Stop 1  Actual Arrival Time"
#         s2_arr_date = "Stop 2  Actual Arrival Date"
#         s2_arr_time = "Stop 2  Actual Arrival Time"
#         s2_dep_date = "Stop 2 Actual Departure Date"
#         s2_dep_time = "Stop 2 Actual Departure Time"
#         s1_dep_date = "Stop 1 Actual Departure Date"
#         s1_dep_time = "Stop 1 Actual Departure Time"
#         p1_date     = "Stop 1 Planned Arrival Date"
#         p1_time     = "Stop 1 Planned Arrival Time"

#         df["Stop1_Actual"]  = pd.to_datetime(df[s1_act_date].astype(str) + " " + df[s1_act_time].astype(str), errors="coerce")
#         df["Stop1_Planned"] = pd.to_datetime(df[p1_date].astype(str)     + " " + df[p1_time].astype(str),     errors="coerce")

#         # Best end time per row — cascading fallback:
#         # 1st choice: Stop 2 Actual Arrival (most accurate)
#         # 2nd choice: Stop 2 Actual Departure (if arrival missing)
#         # 3rd choice: Stop 1 Actual Departure (last resort)
#         df["Stop2_Arrival"] = pd.to_datetime(df[s2_arr_date].astype(str) + " " + df[s2_arr_time].astype(str), errors="coerce")
#         df["Stop2_Depart"]  = pd.to_datetime(df[s2_dep_date].astype(str) + " " + df[s2_dep_time].astype(str), errors="coerce")
#         df["Stop1_Depart"]  = pd.to_datetime(df[s1_dep_date].astype(str) + " " + df[s1_dep_time].astype(str), errors="coerce")
#         df["Best_End"]      = df["Stop2_Arrival"].fillna(df["Stop2_Depart"]).fillna(df["Stop1_Depart"])

#         # Only count trips with status "Completed" — Cancelled and Not Started are excluded
#         if "Load Execution Status" in df.columns:
#             df = df[df["Load Execution Status"].astype(str).str.strip() == "Completed"]

#         # Only drop rows where we have no start OR no end — not just missing Stop2_Actual
#         df = df.dropna(subset=["Stop1_Actual", "Best_End"])

#         for _, gp in df.groupby("Trip ID"):
#             gp = gp.sort_values("Stop1_Actual")
#             start_act = gp.iloc[0]["Stop1_Actual"]
#             end_act   = gp.iloc[-1]["Best_End"]

#             # Handle overnight trips (midnight crossing)
#             if end_act < start_act:
#                 end_act += pd.Timedelta(days=1)

#             raw_hours = (end_act - start_act).total_seconds() / 3600

#             # If trip duration is suspiciously long (>20h), use planned start
#             # for the HOUR CALCULATION only — date is always from end_act
#             start_for_hours = (
#                 gp.iloc[0]["Stop1_Planned"]
#                 if raw_hours > 20 and pd.notna(gp.iloc[0]["Stop1_Planned"])
#                 else start_act
#             )

#             # Date assigned = day the trip ENDS (not starts).
#             # Overnight trips starting Feb 13 and ending Feb 14 count under Feb 14.
#             assigned_date = end_act.date()
#             trip_hours    = round((end_act - start_for_hours).total_seconds() / 3600, 2)

#             # Most frequent driver name wins ALL hours for the whole trip/block.
#             # Clean semicolon-duplicated names (e.g. "JOHN DOE;JOHN DOE" → "JOHN DOE")
#             # so they don't split the vote against the true dominant driver.
#             dominant_driver = (
#                 gp["Driver Name"]
#                 .astype(str)
#                 .str.split(";").str[0]   # take the part before any semicolon
#                 .str.upper()
#                 .str.strip()
#                 .value_counts()
#                 .idxmax()
#             )

#             all_rows.append([dominant_driver, assigned_date, trip_hours])

#     if not all_rows:
#         return pd.DataFrame(columns=["Driver", "Date", "Hours"]), None

#     result_df = pd.DataFrame(all_rows, columns=["Driver", "Date", "Hours"])
#     return result_df, None

# def build_final_dataset(adp_df, relay_df, dp_df):
#     """
#     Merges datasets and handles the Relay Date Drop requirement.
#     1. Establish Master List from ADP.
#     2. Drops the first chronological date column from Relay pivot.
#     3. Sorts columns to ensure 1st 7 days = Week 1, next 7 days = Week 2.
#     """
#     adp_names = adp_df["Driver"].unique().tolist()
#     mapper = get_fuzzy_name_mapper(adp_names)

#     # Apply Fuzzy Mapping
#     relay_df["Driver"] = relay_df["Driver"].apply(mapper)
#     dp_df["Driver"] = dp_df["Driver"].astype(str).str.upper().str.strip().apply(mapper)
    
#     # --- Pivot Relay Data ---
#     relay_pivot = relay_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
    
#     # Sort Relay dates chronologically
#     sorted_dates = sorted(relay_pivot.columns)
    
#     # REQUIREMENT: Drop the first chronological date column (index 0).
#     # The relay data spans 15 days; dropping the first leaves exactly 14
#     # which split cleanly into W1 (days 1-7) and W2 (days 8-14).
#     if len(sorted_dates) > 0:
#         first_date = sorted_dates[0]
#         relay_pivot = relay_pivot.drop(columns=[first_date])
#         remaining_dates = sorted_dates[1:]  # all dates after the dropped first
#     else:
#         remaining_dates = []

#     # Apply R_ prefix for the remaining dates
#     relay_pivot.columns = [f"R_{col}" for col in relay_pivot.columns]
    
#     # Pivot ADP Data
#     adp_pivot = adp_df.pivot_table(index="Driver", columns="Date", values="Hours", aggfunc="sum").fillna(0)
#     adp_pivot = adp_pivot.reindex(sorted(adp_pivot.columns), axis=1)
#     adp_pivot.columns = [f"A_{col}" for col in adp_pivot.columns]

#     # Merge Data
#     final_df = pd.concat([relay_pivot, adp_pivot], axis=1).fillna(0).reset_index()
#     final_df = final_df.merge(dp_df, on="Driver", how="left")
    
#     # Define Column Order (Base info -> Relay Dates -> ADP Dates)
#     relay_cols = [f"R_{d}" for d in remaining_dates]
#     adp_cols = [c for c in final_df.columns if c.startswith("A_")]
#     base_cols = [c for c in final_df.columns if not c.startswith(("R_", "A_"))]
    
#     return final_df[base_cols + relay_cols + adp_cols], relay_cols, adp_cols

# def build_override_dict(override_df, start_date, end_date, adp_df):
#     """Filters overrides and maps to ADP master drivers."""
#     if override_df is None or override_df.empty: 
#         return {}
    
#     adp_names = adp_df["Driver"].unique().tolist()
#     mapper = get_fuzzy_name_mapper(adp_names)
    
#     override_df["Driver"] = override_df["Driver"].astype(str).apply(mapper)
#     override_df["Date"] = pd.to_datetime(override_df["Date"]).dt.date
    
#     mask = (override_df["Date"] >= start_date) & (override_df["Date"] <= end_date)
#     filtered_df = override_df.loc[mask]
    
#     return {(row["Driver"], row["Date"]): row["Override Price"] for _, row in filtered_df.iterrows()}

import io
import pandas as pd
from difflib import SequenceMatcher

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
    """Maps names to ADP Master list using fuzzy matching (without rapidfuzz)."""
    def _match(name):
        if pd.isna(name) or str(name).strip() == "":
            return name
        name_str = str(name).upper().strip()
        
        best_match = None
        best_ratio = 0
        
        for target in target_names:
            ratio = SequenceMatcher(None, name_str, target).ratio()
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = target
        
        # Convert threshold from 70 to 0.70 scale
        threshold_ratio = threshold / 100.0
        if best_match and best_ratio >= threshold_ratio:
            return best_match
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
    """
    Processes Relay CSV files and calculates trip duration.
    
    DEDUPLICATION LOGIC:
    - Combines all relay files with source tracking
    - Removes duplicate Trip IDs (keeps first occurrence from first file)
    - This prevents double-counting overlapping trips across multiple relay files
    """
    all_dfs = []
    
    # Load and mark each file with source index
    for file_idx, raw in enumerate(contents):
        df = pd.read_csv(io.BytesIO(raw))
        df['_source_file'] = file_idx  # Track which file this came from
        all_dfs.append(df)
    
    # Combine all relay data
    combined_df = pd.concat(all_dfs, ignore_index=True)
    
    print(f"\n=== RELAY DEDUPLICATION REPORT ===")
    print(f"Total trip records before dedup: {len(combined_df)}")
    print(f"Unique Trip IDs: {combined_df['Trip ID'].nunique()}")
    
    # Count duplicate Trip IDs
    dup_trips = combined_df[combined_df['Trip ID'].duplicated(keep=False)]['Trip ID'].unique()
    print(f"Trip IDs appearing in multiple files: {len(dup_trips)}")
    
    # DEDUPLICATION: Keep only FIRST occurrence of each Trip ID
    # This means keeping from Trips__2_ if it appears in both files
    combined_df = combined_df.drop_duplicates(subset=['Trip ID'], keep='first')
    
    print(f"Total trip records after dedup: {len(combined_df)}")
    print(f"Trip records REMOVED: {len(all_dfs[0]) + (len(all_dfs[1]) if len(all_dfs) > 1 else 0) - len(combined_df)}")
    
    # Remove the source file tracking column
    combined_df = combined_df.drop(columns=['_source_file'])
    
    all_rows = []

    # Standardize headers (strip extra spaces)
    combined_df.columns = [c.strip() for c in combined_df.columns]

    s1_act_date = "Stop 1  Actual Arrival Date"
    s1_act_time = "Stop 1  Actual Arrival Time"
    s2_arr_date = "Stop 2  Actual Arrival Date"
    s2_arr_time = "Stop 2  Actual Arrival Time"
    s2_dep_date = "Stop 2 Actual Departure Date"
    s2_dep_time = "Stop 2 Actual Departure Time"
    s1_dep_date = "Stop 1 Actual Departure Date"
    s1_dep_time = "Stop 1 Actual Departure Time"
    p1_date     = "Stop 1 Planned Arrival Date"
    p1_time     = "Stop 1 Planned Arrival Time"

    combined_df["Stop1_Actual"]  = pd.to_datetime(combined_df[s1_act_date].astype(str) + " " + combined_df[s1_act_time].astype(str), errors="coerce")
    combined_df["Stop1_Planned"] = pd.to_datetime(combined_df[p1_date].astype(str)     + " " + combined_df[p1_time].astype(str),     errors="coerce")

    # Best end time per row — cascading fallback:
    # 1st choice: Stop 2 Actual Arrival (most accurate)
    # 2nd choice: Stop 2 Actual Departure (if arrival missing)
    # 3rd choice: Stop 1 Actual Departure (last resort)
    combined_df["Stop2_Arrival"] = pd.to_datetime(combined_df[s2_arr_date].astype(str) + " " + combined_df[s2_arr_time].astype(str), errors="coerce")
    combined_df["Stop2_Depart"]  = pd.to_datetime(combined_df[s2_dep_date].astype(str) + " " + combined_df[s2_dep_time].astype(str), errors="coerce")
    combined_df["Stop1_Depart"]  = pd.to_datetime(combined_df[s1_dep_date].astype(str) + " " + combined_df[s1_dep_time].astype(str), errors="coerce")
    combined_df["Best_End"]      = combined_df["Stop2_Arrival"].fillna(combined_df["Stop2_Depart"]).fillna(combined_df["Stop1_Depart"])

    # Only count trips with status "Completed" — Cancelled and Not Started are excluded
    if "Load Execution Status" in combined_df.columns:
        combined_df = combined_df[combined_df["Load Execution Status"].astype(str).str.strip() == "Completed"]

    # Only drop rows where we have no start OR no end — not just missing Stop2_Actual
    combined_df = combined_df.dropna(subset=["Stop1_Actual", "Best_End"])

    for trip_id, gp in combined_df.groupby("Trip ID"):
        gp = gp.sort_values("Stop1_Actual")
        start_act = gp.iloc[0]["Stop1_Actual"]
        end_act   = gp.iloc[-1]["Best_End"]

        # Handle overnight trips (midnight crossing)
        if end_act < start_act:
            end_act += pd.Timedelta(days=1)

        raw_hours = (end_act - start_act).total_seconds() / 3600

        # If trip duration is suspiciously long (>20h), use planned start
        # for the HOUR CALCULATION only — date is always from end_act
        start_for_hours = (
            gp.iloc[0]["Stop1_Planned"]
            if raw_hours > 20 and pd.notna(gp.iloc[0]["Stop1_Planned"])
            else start_act
        )

        # Date assigned = day the trip ENDS (not starts).
        # Overnight trips starting Feb 13 and ending Feb 14 count under Feb 14.
        assigned_date = end_act.date()
        trip_hours    = round((end_act - start_for_hours).total_seconds() / 3600, 2)

        # Most frequent driver name wins ALL hours for the whole trip/block.
        # Clean semicolon-duplicated names (e.g. "JOHN DOE;JOHN DOE" → "JOHN DOE")
        # so they don't split the vote against the true dominant driver.
        dominant_driver = (
            gp["Driver Name"]
            .astype(str)
            .str.split(";").str[0]   # take the part before any semicolon
            .str.upper()
            .str.strip()
            .value_counts()
            .idxmax()
        )

        all_rows.append([dominant_driver, assigned_date, trip_hours])

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
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from rapidfuzz import process, fuzz
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from collections import Counter

# ============ PAGE CONFIGURATION ============
st.set_page_config(
    page_title="Payroll Processor",
    page_icon="💰",
    layout="wide"
)

# ============ CONSTANTS ============
REG_RATE = 24
OT_RATE = 36

# Excel Styles
HEADER_FILL = PatternFill(start_color='D9D9D9', end_color='D9D9D9', fill_type='solid')
RELAY_FILL = PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid') 
ADP_FILL = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')  
DRIVER_FILL = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
RED_FILL = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
YELLOW_FILL = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
OVERRIDE_FILL = PatternFill(start_color='FF6B00', end_color='FF6B00', fill_type='solid')  # Bright Orange
ANOMALY_FILL = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')  # Red for anomalies
THIN_BORDER = Border(
    left=Side(style='thin'), 
    right=Side(style='thin'), 
    top=Side(style='thin'), 
    bottom=Side(style='thin')
)

# ============ HEADER ============
st.title("🚛 Driver Payroll Processing System")
st.markdown("---")

# ============ SIDEBAR FILE UPLOADS ============
st.sidebar.header("📁 File Uploads")

st.sidebar.markdown("### ADP Files")
st.sidebar.caption("Upload multiple CSV files with payroll data")
adp_files = st.sidebar.file_uploader(
    "ADP CSVs",
    type=['csv'],
    accept_multiple_files=True,
    key="adp",
    help="Required columns: Payroll Name, Pay Date, Hours"
)

st.sidebar.markdown("### Relay Files")
st.sidebar.caption("Upload multiple CSV files with trip data")
relay_files = st.sidebar.file_uploader(
    "Relay CSVs",
    type=['csv'],
    accept_multiple_files=True,
    key="relay",
    help="Required columns: Driver Name, Trip ID, Stop 1/2 dates and times"
)

st.sidebar.markdown("### DriverPay File")
st.sidebar.caption("Upload single CSV with driver pay information")
driverpay_file = st.sidebar.file_uploader(
    "DriverPay CSV",
    type=['csv'],
    key="driverpay",
    help="Required columns: Drivers, Fixed Pay (optional)"
)

st.sidebar.markdown("### Override Payment File (Optional)")
st.sidebar.caption("🔶 Special pricing for specific drivers/dates")
override_file = st.sidebar.file_uploader(
    "Override Payment CSV",
    type=['csv'],
    key="override",
    help="Required columns: Driver, Date, Override Price"
)

if override_file:
    st.sidebar.success("✅ Override file loaded - special pricing will be applied")

# ============ PROCESSING FUNCTIONS ============

def format_date_column(date_obj):
    """Format date as '4-Jan' instead of full date"""
    if isinstance(date_obj, str):
        date_obj = datetime.strptime(date_obj, "%Y-%m-%d").date()
    return date_obj.strftime("%-d-%b")  # For Unix/Mac, use %#d for Windows


def process_adp_files(files):
    """
    Process ADP files: Extract Payroll Name, Pay Date, and Hours.
    Clean names, aggregate by Driver and Date.
    """
    if not files:
        return None, "No ADP files uploaded"
    
    all_data = []
    
    for file in files:
        try:
            df = pd.read_csv(file)
            
            # Validate required columns
            required_cols = ['Payroll Name', 'Pay Date', 'Hours']
            missing = [col for col in required_cols if col not in df.columns]
            if missing:
                return None, f"Missing columns in {file.name}: {missing}"
            
            # Extract and clean
            df = df[required_cols].copy()
            
            # Clean driver names
            df['Driver'] = (
                df['Payroll Name']
                .astype(str)
                .str.replace(',', '', regex=False)
                .str.strip()
                .str.upper()
            )
            
            # Parse dates and hours
            df['Date'] = pd.to_datetime(df['Pay Date']).dt.date
            df['Hours'] = pd.to_numeric(df['Hours'], errors='coerce').fillna(0)
            
            all_data.append(df[['Driver', 'Date', 'Hours']])
            
        except Exception as e:
            return None, f"Error processing {file.name}: {str(e)}"
    
    if not all_data:
        return None, "No valid data found in ADP files"
    
    # Combine and aggregate
    combined = pd.concat(all_data, ignore_index=True)
    daily_hours = (
        combined.groupby(['Driver', 'Date'], as_index=False)['Hours']
        .sum()
        .sort_values(['Driver', 'Date'])
    )
    
    return daily_hours, None


def process_relay_files(files):
    """
    Process Relay files: Construct timestamps, calculate trip durations,
    handle 24-hour rollover, aggregate by Driver and Date.
    """
    if not files:
        return None, "No Relay files uploaded"
    
    all_data = []
    
    for file in files:
        try:
            df = pd.read_csv(file)
            all_data.append(df)
        except Exception as e:
            return None, f"Error reading {file.name}: {str(e)}"
    
    if not all_data:
        return None, "No valid data in Relay files"
    
    relay_df = pd.concat(all_data, ignore_index=True)
    
    # Validate columns
    required_cols = [
        'Driver Name', 'Trip ID',
        'Stop 1 Planned Arrival Date', 'Stop 1 Planned Arrival Time',
        'Stop 1  Actual Arrival Date', 'Stop 1  Actual Arrival Time',
        'Stop 2  Actual Arrival Date', 'Stop 2  Actual Arrival Time'
    ]
    missing = [col for col in required_cols if col not in relay_df.columns]
    if missing:
        return None, f"Missing columns in Relay files: {missing}"
    
    try:
        # Construct timestamps
        relay_df['Stop1_Planned'] = pd.to_datetime(
            relay_df['Stop 1 Planned Arrival Date'] + " " +
            relay_df['Stop 1 Planned Arrival Time'],
            errors='coerce'
        )
        
        relay_df['Stop1_Actual'] = pd.to_datetime(
            relay_df['Stop 1  Actual Arrival Date'] + " " +
            relay_df['Stop 1  Actual Arrival Time'],
            errors='coerce'
        )
        
        relay_df['Stop2_Actual'] = pd.to_datetime(
            relay_df['Stop 2  Actual Arrival Date'] + " " +
            relay_df['Stop 2  Actual Arrival Time'],
            errors='coerce'
        )
        
        # Drop rows with missing timestamps
        relay_df = relay_df.dropna(subset=['Stop1_Actual', 'Stop2_Actual'])
        
        # Clean driver names
        relay_df['Driver Name'] = relay_df['Driver Name'].str.upper()
        
        # Calculate trip durations
        relay_rows = []
        
        for trip_id, group in relay_df.groupby('Trip ID'):
            group = group.sort_values('Stop1_Actual')
            
            start_actual = group.iloc[0]['Stop1_Actual']
            start_planned = group.iloc[0]['Stop1_Planned']
            end_actual = group.iloc[-1]['Stop2_Actual']
            
            # Handle day rollover
            if end_actual < start_actual:
                end_actual += pd.Timedelta(days=1)
            
            raw_hours = (end_actual - start_actual).total_seconds() / 3600
            
            # Use planned if trip > 20 hours
            start_dt = start_planned if raw_hours > 20 and pd.notna(start_planned) else start_actual
            hours = (end_actual - start_dt).total_seconds() / 3600
            
            # Get most common driver
            driver = Counter(group['Driver Name']).most_common(1)[0][0]
            work_date = start_dt.date()
            
            relay_rows.append([driver, work_date, round(hours, 2)])
        
        relay_daily = pd.DataFrame(
            relay_rows, columns=['Driver', 'Date', 'Hours']
        )
        
        return relay_daily, None
        
    except Exception as e:
        return None, f"Error processing Relay data: {str(e)}"


def process_override_file(file):
    """
    Process Override Payment file: Driver, Date, Override Price
    Returns dictionary: {(driver, date): override_price}
    """
    if not file:
        return {}, None
    
    try:
        df = pd.read_csv(file)
        
        # Validate columns
        required_cols = ['Driver', 'Date', 'Override Price']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"Missing columns in Override file: {missing}"
        
        # Clean data
        df['Driver'] = df['Driver'].astype(str).str.upper().str.strip()
        df['Date'] = pd.to_datetime(df['Date']).dt.date
        df['Override Price'] = pd.to_numeric(
            df['Override Price'].astype(str).str.replace(r'[\$,]', '', regex=True),
            errors='coerce'
        )
        
        # Create lookup dictionary
        override_dict = {}
        for _, row in df.iterrows():
            key = (row['Driver'], row['Date'])
            override_dict[key] = row['Override Price']
        
        return override_dict, None
        
    except Exception as e:
        return None, f"Error processing Override file: {str(e)}"


def create_excel_with_formulas(final_df, r_cols, a_cols, override_dict):
    """
    Create Excel file with clean formatting, formulas, and anomaly detection
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Payroll"
    
    # Define calculated column headers
    calc_headers = [
        "Total Relay Hours", "Relay_Loads", 
        "W1 Hours", "W1 Regular", "W1 OT", 
        "W2 Hours", "W2 Regular", "W2 OT",
        "Total ADP Hours", "Total Regular", "Total OT",
        "Override Pay", "Final Pay", "Equivalent Hours", "Hour Adjustment"
    ]
    
    # Format date columns with clean names (4-Jan format)
    r_cols_formatted = []
    for col in r_cols:
        date_str = col.replace("R_", "")
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        formatted = date_obj.strftime("%d-%b")
        r_cols_formatted.append(formatted)
    
    a_cols_formatted = []
    for col in a_cols:
        date_str = col.replace("A_", "")
        date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        formatted = date_obj.strftime("%d-%b")
        a_cols_formatted.append(formatted)
    
    # Build final headers with formatted date names
    base_headers = [col for col in final_df.columns if not col.startswith('R_') and not col.startswith('A_')]
    all_headers = base_headers + r_cols_formatted + a_cols_formatted + calc_headers
    
    # Write headers
    ws.append(all_headers)
    
    # Create column letter mapping using FORMATTED names
    col_map = {name: get_column_letter(i) for i, name in enumerate(all_headers, 1)}
    
    # Also keep original column mapping for data access
    original_to_formatted = {}
    for i, col in enumerate(r_cols):
        original_to_formatted[col] = r_cols_formatted[i]
    for i, col in enumerate(a_cols):
        original_to_formatted[col] = a_cols_formatted[i]
    
    # Track cells for highlighting
    override_cells = []
    anomaly_rows = []
    
    # Write data rows with formulas
    for r_idx, row_data in enumerate(final_df.values, start=2):
        row_num = str(r_idx)
        driver_name = row_data[0]  # First column is Driver
        
        # Write base data (non-date columns)
        col_idx = 1
        for i, col_name in enumerate(final_df.columns):
            if not col_name.startswith('R_') and not col_name.startswith('A_'):
                ws.cell(row=r_idx, column=col_idx, value=row_data[i])
                col_idx += 1
        
        # Write Relay data
        has_relay = False
        for orig_col in r_cols:
            val = final_df[orig_col].iloc[r_idx - 2]
            ws.cell(row=r_idx, column=col_idx, value=val)
            if val > 0:
                has_relay = True
            col_idx += 1
        
        # Write ADP data and check for anomalies
        has_adp = False
        for orig_col in a_cols:
            val = final_df[orig_col].iloc[r_idx - 2]
            ws.cell(row=r_idx, column=col_idx, value=val)
            if val > 0:
                has_adp = True
            
            # Check for override dates
            date_str = orig_col.replace("A_", "")
            try:
                work_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                if (driver_name, work_date) in override_dict:
                    # Mark this cell for orange highlighting
                    formatted_col = original_to_formatted[orig_col]
                    col_letter = col_map[formatted_col]
                    override_cells.append(f"{col_letter}{r_idx}")
            except:
                pass
            
            col_idx += 1
        
        # ANOMALY DETECTION: Relay hours but no ADP hours
        if has_relay and not has_adp:
            anomaly_rows.append(r_idx)
        
        # Now write formulas starting from current col_idx
        
        # 1. Total Relay Hours
        if r_cols_formatted:
            first_r = col_map[r_cols_formatted[0]]
            last_r = col_map[r_cols_formatted[-1]]
            ws.cell(row=r_idx, column=col_idx).value = f"=SUM({first_r}{row_num}:{last_r}{row_num})"
            col_idx += 1
        
        # 2. Relay Loads
        if r_cols_formatted:
            first_r = col_map[r_cols_formatted[0]]
            last_r = col_map[r_cols_formatted[-1]]
            ws.cell(row=r_idx, column=col_idx).value = f'=COUNTIF({first_r}{row_num}:{last_r}{row_num}, ">0")'
            col_idx += 1
        
        # 3. W1 Hours (first 7 days)
        if len(a_cols_formatted) >= 7:
            first_a = col_map[a_cols_formatted[0]]
            seventh_a = col_map[a_cols_formatted[6]]
            ws.cell(row=r_idx, column=col_idx).value = f"=SUM({first_a}{row_num}:{seventh_a}{row_num})"
            col_idx += 1
            
            # W1 Regular
            w1_col = col_map["W1 Hours"]
            ws.cell(row=r_idx, column=col_idx).value = f"=MIN({w1_col}{row_num}, 40)"
            col_idx += 1
            
            # W1 OT
            ws.cell(row=r_idx, column=col_idx).value = f"=MAX(0, {w1_col}{row_num}-40)"
            col_idx += 1
        else:
            col_idx += 3
        
        # 4. W2 Hours (last 7 days)
        if len(a_cols_formatted) >= 14:
            eighth_a = col_map[a_cols_formatted[7]]
            fourteenth_a = col_map[a_cols_formatted[13]]
            ws.cell(row=r_idx, column=col_idx).value = f"=SUM({eighth_a}{row_num}:{fourteenth_a}{row_num})"
            col_idx += 1
            
            # W2 Regular
            w2_col = col_map["W2 Hours"]
            ws.cell(row=r_idx, column=col_idx).value = f"=MIN({w2_col}{row_num}, 40)"
            col_idx += 1
            
            # W2 OT
            ws.cell(row=r_idx, column=col_idx).value = f"=MAX(0, {w2_col}{row_num}-40)"
            col_idx += 1
        else:
            col_idx += 3
        
        # 5. Total ADP Hours
        w1_col = col_map["W1 Hours"]
        w2_col = col_map["W2 Hours"]
        ws.cell(row=r_idx, column=col_idx).value = f"={w1_col}{row_num}+{w2_col}{row_num}"
        col_idx += 1
        
        # Total Regular
        w1_reg = col_map["W1 Regular"]
        w2_reg = col_map["W2 Regular"]
        ws.cell(row=r_idx, column=col_idx).value = f"={w1_reg}{row_num}+{w2_reg}{row_num}"
        col_idx += 1
        
        # Total OT
        w1_ot = col_map["W1 OT"]
        w2_ot = col_map["W2 OT"]
        ws.cell(row=r_idx, column=col_idx).value = f"={w1_ot}{row_num}+{w2_ot}{row_num}"
        col_idx += 1
        
        # 6. Override Pay
        driver_overrides = [price for (drv, dt), price in override_dict.items() if drv == driver_name]
        override_total = sum(driver_overrides) if driver_overrides else 0
        ws.cell(row=r_idx, column=col_idx).value = override_total
        col_idx += 1
        
        # 7. Final Pay
        pt = col_map["Pay Type"]
        fixed = col_map["Fixed Pay"]
        target_lds = col_map.get("DriverPay_Target_Loads", "1")
        actual_lds = col_map["Relay_Loads"]
        tot_adp = col_map["Total ADP Hours"]
        override_pay = col_map["Override Pay"]
        
        ws.cell(row=r_idx, column=col_idx).value = (
            f"=IF({pt}{row_num}=\"FIXED\", "
            f"({fixed}{row_num}/MAX(1,{target_lds}{row_num}))*{actual_lds}{row_num}, "
            f"({tot_adp}{row_num}*{REG_RATE})+"
            f"({w1_ot}{row_num}*({OT_RATE-REG_RATE}))+"
            f"({w2_ot}{row_num}*({OT_RATE-REG_RATE})))"
            f"+{override_pay}{row_num}"
        )
        col_idx += 1
        
        # 8. Equivalent Hours
        f_pay = col_map["Final Pay"]
        ws.cell(row=r_idx, column=col_idx).value = f"={f_pay}{row_num}/{REG_RATE}"
        col_idx += 1
        
        # 9. Hour Adjustment
        equiv = col_map["Equivalent Hours"]
        ws.cell(row=r_idx, column=col_idx).value = f"={equiv}{row_num}-{tot_adp}{row_num}"
    
    # ============ STYLING ============
    
    # Conditional formatting for Hour Adjustment
    adj_letter = col_map["Hour Adjustment"]
    ws.conditional_formatting.add(
        f'{adj_letter}2:{adj_letter}{ws.max_row}',
        CellIsRule(operator='lessThan', formula=['0'], fill=RED_FILL)
    )
    ws.conditional_formatting.add(
        f'{adj_letter}2:{adj_letter}{ws.max_row}',
        CellIsRule(operator='greaterThan', formula=['0'], fill=YELLOW_FILL)
    )
    
    # Apply cell formatting
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), start=1):
        for cell in row:
            header = all_headers[cell.column-1]
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal='center')
            
            # Hide zeros (except Driver and Pay Type)
            if r_idx > 1 and header not in ["Driver", "Pay Type"]:
                cell.number_format = '[=0]"";#,##0.00'
            
            # Apply colors
            if r_idx == 1:
                # Header row
                cell.fill = HEADER_FILL
                cell.font = Font(bold=True)
            else:
                # ANOMALY DETECTION: Red row for Relay without ADP
                if r_idx in anomaly_rows and header == "Driver":
                    cell.fill = ANOMALY_FILL
                    cell.font = Font(bold=True, color="FFFFFF")
                # Check if this is an override cell
                elif f"{get_column_letter(cell.column)}{cell.row}" in override_cells:
                    cell.fill = OVERRIDE_FILL
                    cell.font = Font(bold=True, color="FFFFFF")
                # Data rows
                elif header == "Driver":
                    cell.fill = DRIVER_FILL
                elif header in r_cols_formatted or "Relay" in header or header == "Relay_Loads":
                    cell.fill = RELAY_FILL
                elif (header in a_cols_formatted or "ADP" in header or 
                      "W1" in header or "W2" in header or 
                      "Regular" in header or "OT" in header or "Total" in header):
                    cell.fill = ADP_FILL
                elif header == "Override Pay":
                    if ws.cell(row=r_idx, column=cell.column).value and ws.cell(row=r_idx, column=cell.column).value > 0:
                        cell.fill = OVERRIDE_FILL
                        cell.font = Font(bold=True, color="FFFFFF")
    
    # Set column widths
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 13
    
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output, anomaly_rows


# ============ MAIN PROCESSING ============

if st.button("🔄 Process Payroll", type="primary", use_container_width=True):
    
    # Validation
    if not adp_files:
        st.error("❌ Please upload ADP files")
        st.stop()
    
    if not relay_files:
        st.error("❌ Please upload Relay files")
        st.stop()
    
    if not driverpay_file:
        st.error("❌ Please upload DriverPay file")
        st.stop()
    
    with st.spinner("Processing payroll data..."):
        
        # ============ STEP 1: Process ADP ============
        st.info("📊 Processing ADP files...")
        adp, error = process_adp_files(adp_files)
        if error:
            st.error(f"ADP Error: {error}")
            st.stop()
        st.success(f"✅ Processed {len(adp)} ADP records from {len(adp_files)} files")
        
        # ============ STEP 2: Process Relay ============
        st.info("🚚 Processing Relay files...")
        relay, error = process_relay_files(relay_files)
        if error:
            st.error(f"Relay Error: {error}")
            st.stop()
        st.success(f"✅ Processed {len(relay)} Relay records from {len(relay_files)} files")
        
        # ============ STEP 3: Process Override (Optional) ============
        override_dict = {}
        if override_file:
            st.info("🔶 Processing Override Payments...")
            override_dict, error = process_override_file(override_file)
            if error:
                st.error(f"Override Error: {error}")
                st.stop()
            st.success(f"✅ Loaded {len(override_dict)} override payments")
            
            # Display override summary
            if override_dict:
                st.info("📋 Override Summary:")
                override_summary = pd.DataFrame([
                    {"Driver": drv, "Date": dt, "Override Price": f"${price:,.2f}"}
                    for (drv, dt), price in override_dict.items()
                ])
                st.dataframe(override_summary, use_container_width=True)
        
        # ============ STEP 4: Load DriverPay ============
        st.info("💵 Loading DriverPay data...")
        try:
            driverpay = pd.read_csv(driverpay_file)
            
            # Clean driver names
            if 'Drivers' not in driverpay.columns:
                st.error("❌ DriverPay file must have 'Drivers' column")
                st.stop()
            
            driverpay['Drivers'] = driverpay['Drivers'].astype(str).str.upper().str.strip()
            
            # Clean Fixed Pay
            if 'Fixed Pay' in driverpay.columns:
                driverpay['Fixed Pay'] = (
                    driverpay['Fixed Pay']
                    .astype(str)
                    .str.replace(r'[\$,]', '', regex=True)
                    .apply(pd.to_numeric, errors='coerce')
                )
            
            # Rename Total Loads to avoid confusion
            if 'Total Loads' in driverpay.columns:
                driverpay = driverpay.rename(columns={'Total Loads': 'DriverPay_Target_Loads'})
            
            st.success(f"✅ Loaded {len(driverpay)} driver pay records")
            
        except Exception as e:
            st.error(f"DriverPay Error: {str(e)}")
            st.stop()
        
        # ============ STEP 5: Fuzzy Matching ============
        st.info("🔍 Matching driver names (fuzzy matching)...")
        
        adp_drivers = adp['Driver'].unique().tolist()
        
        def match_driver(name):
            result = process.extractOne(name, adp_drivers, scorer=fuzz.token_sort_ratio)
            if result and result[1] >= 60:
                return result[0]
            return name
        
        relay['Driver'] = relay['Driver'].apply(match_driver)
        
        # ============ STEP 6: Pivot Tables ============
        st.info("📋 Creating pivot tables...")
        
        # Relay pivot
        relay_pivot = relay.pivot_table(
            index='Driver', 
            columns='Date', 
            values='Hours', 
            aggfunc='sum'
        ).fillna(0)
        
        # Remove first column if multiple columns (skip first date)
        if len(relay_pivot.columns) > 1:
            relay_pivot = relay_pivot.iloc[:, 1:]
        
        relay_pivot.columns = [f"R_{d}" for d in relay_pivot.columns]
        
        # ADP pivot
        adp_pivot = adp.pivot_table(
            index='Driver',
            columns='Date',
            values='Hours',
            aggfunc='sum'
        ).fillna(0)
        
        adp_pivot.columns = [f"A_{d}" for d in adp_pivot.columns]
        
        # ============ STEP 7: Merge Everything ============
        st.info("🔗 Merging datasets...")
        
        final = pd.concat([relay_pivot, adp_pivot], axis=1).fillna(0).reset_index()
        final = final.merge(driverpay, how='left', left_on="Driver", right_on="Drivers")
        
        # Cleanup helper columns
        drops = ['Drivers', 'Unnamed: 3', 'Unnamed: 4']
        final = final.drop(columns=[c for c in drops if c in final.columns])
        
        # Add Pay Type
        final["Pay Type"] = final["Fixed Pay"].apply(
            lambda x: "FIXED" if pd.notna(x) and x > 0 else "HOURLY"
        )
        
        # Reorder columns: Driver, Pay Type, Fixed Pay, Target Loads, then dates
        base_cols = ['Driver', 'Pay Type', 'Fixed Pay']
        if 'DriverPay_Target_Loads' in final.columns:
            base_cols.append('DriverPay_Target_Loads')
        
        other_cols = [c for c in final.columns if c not in base_cols and not c.startswith('R_') and not c.startswith('A_')]
        base_cols.extend(other_cols)
        
        # Get column lists for Excel formulas
        r_cols = [c for c in final.columns if c.startswith("R_")]
        a_cols = [c for c in final.columns if c.startswith("A_")]
        
        # Reorder final dataframe
        final = final[base_cols + r_cols + a_cols]
        
        st.success("✅ Data merged successfully!")
        
        # ============ STEP 8: Display Preview ============
        st.markdown("---")
        st.subheader("📊 Data Preview")
        
        # Display metrics
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Drivers", len(final))
        with col2:
            st.metric("Date Range (Relay)", f"{len(r_cols)} days")
        with col3:
            st.metric("Date Range (ADP)", f"{len(a_cols)} days")
        with col4:
            st.metric("Override Payments", len(override_dict))
        
        # Display dataframe
        st.dataframe(final, use_container_width=True, height=400)
        
        # ============ STEP 9: Generate Excel ============
        st.info("📄 Generating Excel file with clean formatting...")
        
        try:
            excel_file, anomaly_rows = create_excel_with_formulas(final, r_cols, a_cols, override_dict)
            
            st.success("✅ Excel file generated successfully!")
            
            # Display warnings
            if anomaly_rows:
                st.error(f"🚨 {len(anomaly_rows)} anomalies detected: Drivers with Relay hours but no ADP hours (marked in RED)")
            
            if override_dict:
                st.warning("🔶 Orange highlighted cells indicate override payments applied")
            
            # Download button
            st.download_button(
                label="⬇️ Download Payroll Excel",
                data=excel_file,
                file_name=f"payroll_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Excel Generation Error: {str(e)}")
            st.exception(e)

# ============ FOOTER ============
st.markdown("---")
st.caption("💡 Tip: Upload all files, then click 'Process Payroll' to generate the Excel report.")
st.caption("🔶 Override payments highlighted in orange | 🔴 Anomalies (Relay without ADP) highlighted in red")
st.caption("First 14 dates are indicate relay dates | next 14 dates indicate ADP dates")
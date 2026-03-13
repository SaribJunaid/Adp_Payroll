# import io
# from datetime import datetime, timedelta
# import streamlit as st
# from openpyxl import load_workbook

# # Internal Imports
# from payroll_app.config import load_azure_credentials
# from payroll_app.excel_builder import create_excel
# from processing import (
#     build_final_dataset,
#     build_override_dict,
#     process_adp,
#     process_relay,
#     read_flexible_file,
# )
# from payroll_app.sharepoint import (
#     add_sheet_to_workbook,
#     check_workbook_exists,
#     download_sharepoint_file,
#     exchange_code_for_token,
#     get_auth_url,
#     get_site_from_url,
#     get_user_info,
#     list_sharepoint_files,
#     upload_to_sharepoint,
# )
# from payroll_app.state import init_session_state

# def _render_sharepoint_file_picker(
#     title,
#     file_types,
#     state_key_data,
#     state_key_name=None,
#     select_mode="single",
#     confirm_label="Confirm File",
#     current_path_key="input_path",
# ):
#     key_pref = title.lower().replace(" ", "_")
    
#     with st.expander(title, expanded=True):
#         source_key = f"{key_pref}_source"
#         source = st.radio(
#             f"{title} Source:",
#             ["Desktop", "SharePoint"],
#             key=source_key,
#             horizontal=True,
#         )

#         if source == "Desktop":
#             if select_mode == "multiple":
#                 uploads = st.file_uploader(
#                     f"Upload {title}",
#                     type=file_types,
#                     accept_multiple_files=True,
#                     key=f"{key_pref}_upload",
#                 )
#                 if uploads:
#                     st.session_state[state_key_data] = [f.read() for f in uploads]
#                     st.success(f"Loaded {len(uploads)} files")
#             else:
#                 upload = st.file_uploader(
#                     f"Upload {title}",
#                     type=file_types,
#                     key=f"{key_pref}_upload",
#                 )
#                 if upload:
#                     st.session_state[state_key_data] = upload.read()
#                     if state_key_name:
#                         st.session_state[state_key_name] = upload.name
#                     st.success(f"Loaded {upload.name}")
#             return

#         if not st.session_state.get("access_token"):
#             st.warning("Connect to SharePoint first")
#             return

#         items, _ = list_sharepoint_files(
#             st.session_state.access_token,
#             st.session_state.site_info["id"],
#             st.session_state[current_path_key],
#         )
#         if not items:
#             st.info("No files found in current folder")
#             return

#         files = [item for item in items if item["name"].lower().endswith(tuple(file_types))]
#         if not files:
#             st.info("No matching files in current folder")
#             return

#         if select_mode == "multiple":
#             selected = st.multiselect(
#                 f"Select {title}:",
#                 [f["name"] for f in files],
#                 key=f"{key_pref}_select",
#             )
#             if st.button(confirm_label, key=f"{key_pref}_confirm"):
#                 with st.spinner(f"Downloading {title}..."):
#                     st.session_state[state_key_data] = []
#                     for filename in selected:
#                         file_obj = next(f for f in files if f["name"] == filename)
#                         content, _ = download_sharepoint_file(
#                             st.session_state.access_token,
#                             st.session_state.site_info["id"],
#                             file_obj["id"],
#                         )
#                         if content:
#                             st.session_state[state_key_data].append(content)
#                     st.success(f"Downloaded {len(st.session_state[state_key_data])} files")
#         else:
#             selected = st.selectbox(
#                 f"Select {title}:",
#                 ["-- None --"] + [f["name"] for f in files],
#                 key=f"{key_pref}_select",
#             )
#             if selected != "-- None --" and st.button(confirm_label, key=f"{key_pref}_confirm"):
#                 with st.spinner(f"Downloading {title}..."):
#                     file_obj = next(f for f in files if f["name"] == selected)
#                     content, _ = download_sharepoint_file(
#                         st.session_state.access_token,
#                         st.session_state.site_info["id"],
#                         file_obj["id"],
#                     )
#                     if content:
#                         st.session_state[state_key_data] = content
#                         if state_key_name:
#                             st.session_state[state_key_name] = file_obj["name"]
#                         st.success(f"Downloaded {file_obj['name']}")

# def _render_adp_picker():
#     with st.expander("ADP Files", expanded=True):
#         source = st.radio("ADP Source:", ["Desktop", "SharePoint"], key="adp_source", horizontal=True)
#         if source == "Desktop":
#             uploads = st.file_uploader(
#                 "Upload ADP CSV files",
#                 type=["csv"],
#                 accept_multiple_files=True,
#                 key="adp_upload",
#             )
#             if uploads:
#                 st.session_state.adp_files_data = [f.read() for f in uploads]
#                 st.success(f"Loaded {len(uploads)} ADP files")
#             return

#         if not st.session_state.get("site_info"):
#             st.warning("Connect to SharePoint first")
#             return

#         col1, col2 = st.columns([3, 1])
#         with col1:
#             st.caption(f"Path: {st.session_state.input_path}")
#         with col2:
#             if st.button("Up", key="adp_path_up") and st.session_state.input_path != "root":
#                 if "/" in st.session_state.input_path:
#                     st.session_state.input_path = "/".join(st.session_state.input_path.split("/")[:-1])
#                 else:
#                     st.session_state.input_path = "root"
#                 st.rerun()

#         items, _ = list_sharepoint_files(
#             st.session_state.access_token,
#             st.session_state.site_info["id"],
#             st.session_state.input_path,
#         )
#         if not items:
#             st.info("No items in current folder")
#             return

#         folders = [i for i in items if "folder" in i]
#         csv_files = [i for i in items if i["name"].lower().endswith(".csv")]

#         if folders:
#             folder = st.selectbox(
#                 "Navigate to folder:",
#                 ["-- Stay Here --"] + [f["name"] for f in folders],
#                 key="adp_folder_select",
#             )
#             if folder != "-- Stay Here --" and st.button("Enter Folder", key="adp_enter_folder"):
#                 if st.session_state.input_path == "root":
#                     st.session_state.input_path = folder
#                 else:
#                     st.session_state.input_path = f"{st.session_state.input_path}/{folder}"
#                 st.rerun()

#         if not csv_files:
#             st.info("No CSV files in current folder")
#             return

#         selected = st.multiselect(
#             "Select ADP CSV files:",
#             [f["name"] for f in csv_files],
#             key="adp_selected_files",
#         )
#         if st.button("Confirm ADP Files", key="adp_confirm_files"):
#             with st.spinner("Downloading ADP files..."):
#                 st.session_state.adp_files_data = []
#                 for filename in selected:
#                     file_obj = next(f for f in csv_files if f["name"] == filename)
#                     content, _ = download_sharepoint_file(
#                         st.session_state.access_token,
#                         st.session_state.site_info["id"],
#                         file_obj["id"],
#                     )
#                     if content:
#                         st.session_state.adp_files_data.append(content)
#                 st.success(f"Downloaded {len(st.session_state.adp_files_data)} ADP files")

# def _render_output_config():
#     st.markdown("---")
#     st.subheader("Output Configuration")
#     output_dest = st.radio("Save to:", ["Download Only", "SharePoint"], horizontal=True, key="out_dest_radio")
#     workbook_action = None

#     if output_dest == "SharePoint" and st.session_state.get("site_info"):
#         st.markdown("### Select Output Folder")
#         col1, col2 = st.columns([3, 1])
#         with col1:
#             st.caption(f"Output Path: {st.session_state.output_path}")
#         with col2:
#             if st.button("Up", key="output_up") and st.session_state.output_path != "root":
#                 if "/" in st.session_state.output_path:
#                     st.session_state.output_path = "/".join(st.session_state.output_path.split("/")[:-1])
#                 else:
#                     st.session_state.output_path = "root"
#                 st.rerun()

#         items, _ = list_sharepoint_files(
#             st.session_state.access_token,
#             st.session_state.site_info["id"],
#             st.session_state.output_path,
#         )
#         if items:
#             folders = [i for i in items if "folder" in i]
#             if folders:
#                 folder = st.selectbox(
#                     "Navigate to output folder:",
#                     ["-- Use Current --"] + [f["name"] for f in folders],
#                     key="output_folder_select",
#                 )
#                 if folder != "-- Use Current --" and st.button("Enter Folder", key="output_enter_folder"):
#                     if st.session_state.output_path == "root":
#                         st.session_state.output_path = folder
#                     else:
#                         st.session_state.output_path = f"{st.session_state.output_path}/{folder}"
#                     st.rerun()

#         workbook_action = st.radio(
#             "Workbook Action:",
#             ["Add Sheet to Existing ADP.xlsx", "Create New Workbook"],
#             horizontal=True,
#             key="wb_action_radio"
#         )

#     return output_dest, workbook_action

# def _handle_process(output_dest, workbook_action, start_date, end_date):
#     """Core processing handler."""
#     if not st.session_state.get("adp_files_data"):
#         st.error("Missing ADP files")
#         st.stop()
#     if not st.session_state.get("relay_files_data"):
#         st.error("Missing Relay files")
#         st.stop()
#     if not st.session_state.get("driverpay_data"):
#         st.error("Missing DriverPay file")
#         st.stop()

#     with st.spinner("Processing payroll data..."):
#         # 1. Process Input Files
#         adp_df, _ = process_adp(st.session_state.adp_files_data)
#         relay_df, _ = process_relay(st.session_state.relay_files_data)
        
#         dp_df = read_flexible_file(st.session_state.driverpay_data, st.session_state.dp_name)
#         if dp_df is None:
#             st.error(f"Could not read DriverPay file: {st.session_state.dp_name}")
#             st.stop()

#         # 2. Process Overrides
#         override_map = {}
#         if st.session_state.get("override_data"):
#             override_df = read_flexible_file(st.session_state.override_data, st.session_state.ov_name)
#             if override_df is not None:
#                 override_map = build_override_dict(override_df, start_date, end_date, adp_df)

#         # 3. Build Dataset (Handles dropping first chronological relay date)
#         final_df, relay_cols, adp_cols = build_final_dataset(adp_df, relay_df, dp_df)
        
#         # 4. Generate Excel
#         excel_out, _ = create_excel(final_df, relay_cols, adp_cols, override_map)

#         # 5. Dynamic naming logic
#         try:
#             all_dates = [datetime.strptime(col.replace("A_", ""), "%Y-%m-%d").date() for col in adp_cols]
#             period_str = f"{min(all_dates).strftime('%d-%b')} to {max(all_dates).strftime('%d-%b')}"
#         except:
#             period_str = f"{start_date.strftime('%d-%b')} to {end_date.strftime('%d-%b')}"
        
#         filename = f"Payroll_Report_{period_str}.xlsx"

#         # 6. UI Metrics and Preview
#         st.success(f"Excel generated for period: {period_str}")
#         col1, col2, col3, col4 = st.columns(4)
#         col1.metric("Drivers", len(final_df))
#         col2.metric("Relay Days (Processed)", len(relay_cols))
#         col3.metric("ADP Days", len(adp_cols))
#         col4.metric("Overrides", len(override_map))
        
#         st.dataframe(final_df, use_container_width=True, height=400)

#         # 7. Final Output Delivery
#         if output_dest == "Download Only":
#             st.download_button(
#                 "📥 Download Payroll Excel",
#                 excel_out.getvalue(),
#                 filename,
#                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#                 type="primary",
#                 use_container_width=True
#             )
#         else:
#             if not st.session_state.get("access_token") or not st.session_state.get("site_info"):
#                 st.error("SharePoint connection lost. Please download manually.")
#                 st.download_button("Download Excel", excel_out.getvalue(), filename)
#             else:
#                 site_id = st.session_state.site_info["id"]
#                 token = st.session_state.access_token
#                 path = st.session_state.output_path

#                 if workbook_action == "Create New Workbook":
#                     upload_to_sharepoint(token, site_id, path, filename, excel_out.getvalue())
#                     st.info(f"File uploaded to {path}")
#                 else:
#                     add_sheet_to_workbook(token, site_id, path, "ADP.xlsx", excel_out.getvalue(), period_str)
#                     st.info(f"Sheet '{period_str}' added to ADP.xlsx")

# def run_app():
#     st.set_page_config(page_title="Payroll Processor", page_icon="💸", layout="wide")
#     init_session_state()
#     creds = load_azure_credentials()

#     # OAuth Handler
#     if "code" in st.query_params and not st.session_state.get("access_token") and creds["configured"]:
#         token, _ = exchange_code_for_token(
#             st.query_params["code"],
#             creds["client_id"],
#             creds["tenant_id"],
#             creds["client_secret"],
#         )
#         if token:
#             st.session_state.access_token = token
#             st.session_state.user_info = get_user_info(token)
#         st.query_params.clear()
#         st.rerun()

#     st.title("Driver Payroll Processing System")
#     st.markdown("---")

#     # Sidebar Authentication
#     with st.sidebar:
#         st.header("Authentication")
#         if not creds["configured"]:
#             st.warning("Azure credentials missing in config.")
#         elif not st.session_state.get("access_token"):
#             if st.button("Login with Microsoft", type="primary"):
#                 st.markdown(f"[Click here to login]({get_auth_url(creds['client_id'], creds['tenant_id'])})")
#                 st.stop()
#         else:
#             name = (st.session_state.user_info or {}).get("displayName", "User")
#             st.success(f"Logged in as: {name}")
#             if st.button("Logout"):
#                 for key in list(st.session_state.keys()): del st.session_state[key]
#                 st.rerun()

#         if st.session_state.get("access_token"):
#             st.header("SharePoint Site")
#             site_url = st.text_input("Enter SharePoint Site URL")
#             if st.button("Connect Site") and site_url:
#                 site, _ = get_site_from_url(st.session_state.access_token, site_url)
#                 if site: st.session_state.site_info = site

#     # Main UI Steps
#     st.subheader("1. Select Payroll Period")
#     d_col1, d_col2 = st.columns(2)
#     with d_col1:
#         start_date = st.date_input("Start Date", datetime.now() - timedelta(days=14))
#     with d_col2:
#         end_date = st.date_input("End Date", datetime.now())

#     st.markdown("---")
#     st.subheader("2. Load Data")
#     _render_adp_picker()
#     _render_sharepoint_file_picker("Relay Files", ["csv"], "relay_files_data", select_mode="multiple")
#     _render_sharepoint_file_picker("DriverPay File", ["csv", "xlsx"], "driverpay_data", state_key_name="dp_name")
#     _render_sharepoint_file_picker("Override File", ["csv", "xlsx"], "override_data", state_key_name="ov_name")

#     output_dest, workbook_action = _render_output_config()

#     if st.button("🚀 Process and Generate Payroll", type="primary", use_container_width=True):
#         _handle_process(output_dest, workbook_action, start_date, end_date)

iimport io
from datetime import datetime, timedelta
import streamlit as st
from openpyxl import load_workbook
 
# Internal Imports
from payroll_app.config import load_azure_credentials
from payroll_app.excel_builder import create_excel
from payroll_app.processing import (
    build_final_dataset,
    build_override_dict,
    process_adp,
    process_relay,
    read_flexible_file,
)
from payroll_app.sharepoint import (
    add_sheet_to_workbook,
    check_workbook_exists,
    download_sharepoint_file,
    exchange_code_for_token,
    get_auth_url,
    get_site_from_url,
    get_user_info,
    list_sharepoint_files,
    upload_to_sharepoint,
)
from payroll_app.state import init_session_state
 
# ── Shared CSS injected once ──────────────────────────────────────────────────
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600&family=DM+Mono:wght@400;500&display=swap');
 
/* Global font */
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
 
/* File picker card */
.fp-card {
    background: #0f1117;
    border: 1px solid #2a2d3a;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 12px;
}
 
/* Breadcrumb bar */
.fp-breadcrumb {
    display: flex;
    align-items: center;
    gap: 6px;
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 8px 14px;
    margin-bottom: 12px;
    font-family: 'DM Mono', monospace;
    font-size: 12px;
    color: #8b92a5;
    flex-wrap: wrap;
}
.fp-breadcrumb .crumb { color: #c9d1e0; }
.fp-breadcrumb .sep { color: #3d4155; }
.fp-breadcrumb .crumb-root { color: #5b6bff; font-weight: 600; }
 
/* Folder grid */
.fp-folder-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 8px;
    margin-bottom: 14px;
}
.fp-folder-btn {
    display: flex;
    align-items: center;
    gap: 8px;
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 10px 12px;
    cursor: pointer;
    transition: all 0.15s ease;
    font-size: 13px;
    color: #c9d1e0;
    text-align: left;
    width: 100%;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.fp-folder-btn:hover {
    background: #21253a;
    border-color: #5b6bff;
    color: #fff;
}
.fp-folder-btn .icon { font-size: 16px; flex-shrink: 0; }
 
/* File list */
.fp-file-item {
    display: flex;
    align-items: center;
    gap: 10px;
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 6px;
    font-size: 13px;
    color: #c9d1e0;
}
.fp-file-item .file-icon { font-size: 18px; }
.fp-file-item .file-name { flex: 1; font-family: 'DM Mono', monospace; font-size: 12px; }
.fp-file-item .file-size { color: #5b6170; font-size: 11px; }
 
/* Status badge */
.fp-status-ok {
    display: inline-flex; align-items: center; gap: 6px;
    background: #0d2b1e; border: 1px solid #1a5c38;
    color: #4ade80; border-radius: 20px;
    padding: 4px 12px; font-size: 12px; font-weight: 500;
    margin-top: 8px;
}
.fp-status-pending {
    display: inline-flex; align-items: center; gap: 6px;
    background: #1a1d27; border: 1px solid #2a2d3a;
    color: #8b92a5; border-radius: 20px;
    padding: 4px 12px; font-size: 12px;
    margin-top: 8px;
}
</style>
"""
 
def _inject_css():
    if not st.session_state.get("_css_injected"):
        st.markdown(_CSS, unsafe_allow_html=True)
        st.session_state["_css_injected"] = True
 
def _breadcrumb(path):
    """Render a styled breadcrumb from a path string."""
    if path == "root":
        parts_html = '<span class="crumb-root">⌂ Root</span>'
    else:
        segments = path.split("/")
        crumbs = ['<span class="crumb-root">⌂ Root</span>']
        for seg in segments:
            crumbs.append(f'<span class="sep">›</span><span class="crumb">{seg}</span>')
        parts_html = "".join(crumbs)
    st.markdown(f'<div class="fp-breadcrumb">{parts_html}</div>', unsafe_allow_html=True)
 
def _file_icon(name):
    name = name.lower()
    if name.endswith(".csv"):  return "📄"
    if name.endswith(".xlsx"): return "📊"
    if name.endswith(".pdf"):  return "📕"
    return "📎"
 
def _render_sharepoint_file_picker(
    title,
    file_types,
    state_key_data,
    state_key_name=None,
    select_mode="single",
    confirm_label="Confirm File",
    current_path_key=None,
):
    key_pref = title.lower().replace(" ", "_")
    path_key = current_path_key or f"{key_pref}_path"
    if path_key not in st.session_state:
        st.session_state[path_key] = "root"
 
    # Check if already loaded — show persistent status badge
    already_loaded = bool(st.session_state.get(state_key_data))
    loaded_name = st.session_state.get(state_key_name, "")
 
    with st.expander(title, expanded=not already_loaded):
        # Status badge at top
        if already_loaded:
            if select_mode == "multiple":
                count = len(st.session_state[state_key_data])
                st.markdown(f'<div class="fp-status-ok">✓ {count} file{"s" if count != 1 else ""} loaded</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="fp-status-ok">✓ {loaded_name or "File loaded"}</div>', unsafe_allow_html=True)
            if st.button("↺ Change", key=f"{key_pref}_clear", help="Load a different file"):
                del st.session_state[state_key_data]
                if state_key_name and state_key_name in st.session_state:
                    del st.session_state[state_key_name]
                st.rerun()
            return
 
        source = st.radio(
            "Source:",
            ["💻 Desktop", "☁️ SharePoint"],
            key=f"{key_pref}_source",
            horizontal=True,
        )
 
        if source == "💻 Desktop":
            if select_mode == "multiple":
                uploads = st.file_uploader(
                    f"Upload {title}", type=file_types,
                    accept_multiple_files=True, key=f"{key_pref}_upload",
                )
                if uploads:
                    st.session_state[state_key_data] = [f.read() for f in uploads]
                    st.rerun()
            else:
                upload = st.file_uploader(
                    f"Upload {title}", type=file_types, key=f"{key_pref}_upload",
                )
                if upload:
                    st.session_state[state_key_data] = upload.read()
                    if state_key_name:
                        st.session_state[state_key_name] = upload.name
                    st.rerun()
            return
 
        if not st.session_state.get("access_token") or not st.session_state.get("site_info"):
            st.warning("🔗 Connect to SharePoint first (sidebar)")
            return
 
        # ── Breadcrumb navigation ──
        _breadcrumb(st.session_state[path_key])
 
        # Up button only when not at root
        if st.session_state[path_key] != "root":
            if st.button("⬆ Back", key=f"{key_pref}_up", help="Go up one level"):
                current = st.session_state[path_key]
                st.session_state[path_key] = (
                    "/".join(current.split("/")[:-1]) if "/" in current else "root"
                )
                st.rerun()
 
        # ── Fetch folder contents ──
        items, err = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state[path_key],
        )
        if err or not items:
            st.info("📂 Folder is empty")
            return
 
        folders = [i for i in items if "folder" in i]
        files   = [
            i for i in items
            if i["name"].lower().endswith(
                tuple(f".{t}" if not t.startswith(".") else t for t in file_types)
            )
        ]
 
        # ── Folder tiles ──
        if folders:
            # Render folder buttons as a grid using columns
            cols_per_row = 4
            for row_start in range(0, len(folders), cols_per_row):
                row_folders = folders[row_start:row_start + cols_per_row]
                cols = st.columns(len(row_folders))
                for col, f in zip(cols, row_folders):
                    with col:
                        if st.button(
                            f"📁  {f['name']}",
                            key=f"{key_pref}_folder_{f['id']}",
                            use_container_width=True,
                        ):
                            current = st.session_state[path_key]
                            st.session_state[path_key] = (
                                f["name"] if current == "root" else f"{current}/{f['name']}"
                            )
                            st.rerun()
 
        # ── File selection ──
        if not files:
            st.markdown(
                f'<div class="fp-status-pending">No {", ".join(file_types).upper()} files in this folder</div>',
                unsafe_allow_html=True,
            )
            return
 
        if select_mode == "multiple":
            selected = st.multiselect(
                "Select files:",
                [f["name"] for f in files],
                key=f"{key_pref}_select",
            )
        else:
            sel = st.selectbox(
                "Select file:",
                ["-- None --"] + [f["name"] for f in files],
                key=f"{key_pref}_select",
            )
            selected = [] if sel == "-- None --" else [sel]
 
        if selected and st.button(
            f"⬇ {confirm_label}",
            key=f"{key_pref}_confirm",
            type="primary",
            use_container_width=True,
        ):
            with st.spinner(f"Downloading..."):
                if select_mode == "multiple":
                    downloaded = []
                    for filename in selected:
                        file_obj = next(f for f in files if f["name"] == filename)
                        content, _ = download_sharepoint_file(
                            st.session_state.access_token,
                            st.session_state.site_info["id"],
                            file_obj["id"],
                        )
                        if content:
                            downloaded.append(content)
                    # Only persist if download succeeded — never wipe existing data
                    if downloaded:
                        st.session_state[state_key_data] = downloaded
                        st.rerun()
                else:
                    file_obj = next(f for f in files if f["name"] == selected[0])
                    content, _ = download_sharepoint_file(
                        st.session_state.access_token,
                        st.session_state.site_info["id"],
                        file_obj["id"],
                    )
                    if content:
                        st.session_state[state_key_data] = content
                        if state_key_name:
                            st.session_state[state_key_name] = file_obj["name"]
                        st.rerun()
 
def _render_adp_picker():
    if "adp_path" not in st.session_state:
        st.session_state.adp_path = "root"
 
    already_loaded = bool(st.session_state.get("adp_files_data"))
 
    with st.expander("ADP Files", expanded=not already_loaded):
        if already_loaded:
            count = len(st.session_state.adp_files_data)
            st.markdown(f'<div class="fp-status-ok">✓ {count} ADP file{"s" if count != 1 else ""} loaded</div>', unsafe_allow_html=True)
            if st.button("↺ Change", key="adp_clear"):
                del st.session_state["adp_files_data"]
                st.rerun()
            return
 
        source = st.radio("Source:", ["💻 Desktop", "☁️ SharePoint"], key="adp_source", horizontal=True)
 
        if source == "💻 Desktop":
            uploads = st.file_uploader(
                "Upload ADP CSV files", type=["csv"],
                accept_multiple_files=True, key="adp_upload",
            )
            if uploads:
                st.session_state.adp_files_data = [f.read() for f in uploads]
                st.rerun()
            return
 
        if not st.session_state.get("site_info"):
            st.warning("🔗 Connect to SharePoint first (sidebar)")
            return
 
        _breadcrumb(st.session_state.adp_path)
 
        if st.session_state.adp_path != "root":
            if st.button("⬆ Back", key="adp_path_up"):
                current = st.session_state.adp_path
                st.session_state.adp_path = (
                    "/".join(current.split("/")[:-1]) if "/" in current else "root"
                )
                st.rerun()
 
        items, _ = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state.adp_path,
        )
        if not items:
            st.info("📂 Folder is empty")
            return
 
        folders  = [i for i in items if "folder" in i]
        csv_files = [i for i in items if i["name"].lower().endswith(".csv")]
 
        # Folder tiles
        if folders:
            cols_per_row = 4
            for row_start in range(0, len(folders), cols_per_row):
                row_folders = folders[row_start:row_start + cols_per_row]
                cols = st.columns(len(row_folders))
                for col, f in zip(cols, row_folders):
                    with col:
                        if st.button(f"📁  {f['name']}", key=f"adp_folder_{f['id']}", use_container_width=True):
                            current = st.session_state.adp_path
                            st.session_state.adp_path = (
                                f["name"] if current == "root" else f"{current}/{f['name']}"
                            )
                            st.rerun()
 
        if not csv_files:
            st.info("No CSV files in this folder")
            return
 
        selected = st.multiselect("Select ADP CSV files:", [f["name"] for f in csv_files], key="adp_selected_files")
 
        if selected and st.button("⬇ Confirm ADP Files", key="adp_confirm_files", type="primary", use_container_width=True):
            with st.spinner("Downloading ADP files..."):
                downloaded = []
                for filename in selected:
                    file_obj = next(f for f in csv_files if f["name"] == filename)
                    content, _ = download_sharepoint_file(
                        st.session_state.access_token,
                        st.session_state.site_info["id"],
                        file_obj["id"],
                    )
                    if content:
                        downloaded.append(content)
                if downloaded:
                    st.session_state.adp_files_data = downloaded
                    st.rerun()
 
def _render_output_config():
    st.markdown("---")
    st.subheader("Output Configuration")
    output_dest = st.radio("Save to:", ["Download Only", "SharePoint"], horizontal=True, key="out_dest_radio")
    workbook_action = None
 
    if output_dest == "SharePoint" and st.session_state.get("site_info"):
        st.markdown("### Select Output Folder")
        _breadcrumb(st.session_state.output_path)
 
        if st.session_state.output_path != "root":
            if st.button("⬆ Back", key="output_up"):
                current = st.session_state.output_path
                st.session_state.output_path = (
                    "/".join(current.split("/")[:-1]) if "/" in current else "root"
                )
                st.rerun()
 
        items, _ = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state.output_path,
        )
        if items:
            folders = [i for i in items if "folder" in i]
            if folders:
                cols_per_row = 4
                for row_start in range(0, len(folders), cols_per_row):
                    row_folders = folders[row_start:row_start + cols_per_row]
                    cols = st.columns(len(row_folders))
                    for col, f in zip(cols, row_folders):
                        with col:
                            if st.button(f"📁  {f['name']}", key=f"out_folder_{f['id']}", use_container_width=True):
                                current = st.session_state.output_path
                                st.session_state.output_path = (
                                    f["name"] if current == "root" else f"{current}/{f['name']}"
                                )
                                st.rerun()
 
        workbook_action = st.radio(
            "Workbook Action:",
            ["Add Sheet to Existing Workbook", "Create New Workbook"],
            horizontal=True,
            key="wb_action_radio"
        )
 
        # When adding to existing — let user pick the target .xlsx from the same folder
        if workbook_action == "Add Sheet to Existing Workbook":
            items_now, _ = list_sharepoint_files(
                st.session_state.access_token,
                st.session_state.site_info["id"],
                st.session_state.output_path,
            )
            xlsx_files = [i for i in (items_now or []) if i["name"].lower().endswith(".xlsx")]
            if xlsx_files:
                chosen = st.selectbox(
                    "Select target workbook:",
                    [f["name"] for f in xlsx_files],
                    key="target_wb_select",
                )
                # Store the chosen file's id and name in session state
                chosen_obj = next(f for f in xlsx_files if f["name"] == chosen)
                st.session_state["target_wb_id"]   = chosen_obj["id"]
                st.session_state["target_wb_name"] = chosen_obj["name"]
            else:
                st.warning("No .xlsx files found in this folder — switch to 'Create New Workbook' or navigate to the right folder.")
                st.session_state.pop("target_wb_id",   None)
                st.session_state.pop("target_wb_name", None)
 
    return output_dest, workbook_action
 
def _handle_process(output_dest, workbook_action, start_date, end_date):
    """Core processing handler."""
    if not st.session_state.get("adp_files_data"):
        st.error("Missing ADP files")
        st.stop()
    if not st.session_state.get("relay_files_data"):
        st.error("Missing Relay files")
        st.stop()
    if not st.session_state.get("driverpay_data"):
        st.error("Missing DriverPay file")
        st.stop()
 
    with st.spinner("Processing payroll data..."):
        # 1. Process Input Files
        adp_df, _ = process_adp(st.session_state.adp_files_data)
        relay_df, _ = process_relay(st.session_state.relay_files_data)
        
        dp_df = read_flexible_file(st.session_state.driverpay_data, st.session_state.dp_name)
        if dp_df is None:
            st.error(f"Could not read DriverPay file: {st.session_state.dp_name}")
            st.stop()
 
        # 2. Process Overrides
        override_map = {}
        if st.session_state.get("override_data"):
            override_df = read_flexible_file(st.session_state.override_data, st.session_state.ov_name)
            if override_df is not None:
                override_map = build_override_dict(override_df, start_date, end_date, adp_df)
 
        # 3. Build Dataset (Handles dropping first chronological relay date)
        final_df, relay_cols, adp_cols = build_final_dataset(adp_df, relay_df, dp_df)
        
        # 4. Generate Excel
        excel_out, _ = create_excel(final_df, relay_cols, adp_cols, override_map)
 
        # 5. Dynamic naming logic
        try:
            all_dates = [datetime.strptime(col.replace("A_", ""), "%Y-%m-%d").date() for col in adp_cols]
            period_str = f"{min(all_dates).strftime('%d-%b')} to {max(all_dates).strftime('%d-%b')}"
        except:
            period_str = f"{start_date.strftime('%d-%b')} to {end_date.strftime('%d-%b')}"
        
        filename = f"Payroll_Report_{period_str}.xlsx"
 
        # 6. UI Metrics and Preview
        st.success(f"Excel generated for period: {period_str}")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Drivers", len(final_df))
        col2.metric("Relay Days (Processed)", len(relay_cols))
        col3.metric("ADP Days", len(adp_cols))
        col4.metric("Overrides", len(override_map))
        
        st.dataframe(final_df, use_container_width=True, height=400)
 
        # 7. Final Output Delivery
        if output_dest == "Download Only":
            st.download_button(
                "📥 Download Payroll Excel",
                excel_out.getvalue(),
                filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            if not st.session_state.get("access_token") or not st.session_state.get("site_info"):
                st.error("SharePoint connection lost. Please download manually.")
                st.download_button("Download Excel", excel_out.getvalue(), filename)
            else:
                site_id = st.session_state.site_info["id"]
                token = st.session_state.access_token
                path = st.session_state.output_path
 
                if workbook_action == "Create New Workbook":
                    ok, err = upload_to_sharepoint(token, site_id, path, filename, excel_out.getvalue())
                    if ok:
                        st.success(f"✅ Uploaded **{filename}** to `{path}`")
                    else:
                        st.error(f"Upload failed: {err}")
                        st.download_button("📥 Download instead", excel_out.getvalue(), filename)
                else:
                    # Add Sheet to Existing Workbook:
                    # 1. Download the existing workbook bytes from SharePoint
                    # 2. Merge new sheet into it using add_sheet_to_workbook()
                    # 3. Re-upload the merged workbook back to SharePoint
                    wb_id   = st.session_state.get("target_wb_id")
                    wb_name = st.session_state.get("target_wb_name", "ADP.xlsx")
 
                    if not wb_id:
                        st.error("No target workbook selected. Please choose one in Output Configuration.")
                        st.download_button("📥 Download instead", excel_out.getvalue(), filename)
                    else:
                        with st.spinner(f"Downloading {wb_name}..."):
                            existing_bytes, dl_err = download_sharepoint_file(token, site_id, wb_id)
 
                        if dl_err or not existing_bytes:
                            st.error(f"Could not download {wb_name}: {dl_err}")
                            st.download_button("📥 Download instead", excel_out.getvalue(), filename)
                        else:
                            with st.spinner("Merging sheet into workbook..."):
                                merged, merge_err = add_sheet_to_workbook(
                                    existing_bytes,       # existing workbook bytes
                                    excel_out.getvalue(), # new sheet bytes
                                    period_str,           # sheet name
                                )
 
                            if merge_err or not merged:
                                st.error(f"Merge failed: {merge_err}")
                                st.download_button("📥 Download instead", excel_out.getvalue(), filename)
                            else:
                                with st.spinner(f"Uploading merged {wb_name}..."):
                                    ok, up_err = upload_to_sharepoint(
                                        token, site_id, path, wb_name, merged.getvalue()
                                    )
                                if ok:
                                    st.success(f"✅ Sheet **'{period_str}'** added to **{wb_name}**")
                                else:
                                    st.error(f"Upload failed: {up_err}")
                                    st.download_button("📥 Download instead", excel_out.getvalue(), filename)
 
def run_app():
    st.set_page_config(page_title="Payroll Processor", page_icon="💸", layout="wide")
    init_session_state()
    _inject_css()
    creds = load_azure_credentials()
 
    # OAuth Handler
    if "code" in st.query_params and not st.session_state.get("access_token") and creds["configured"]:
        token, _ = exchange_code_for_token(
            st.query_params["code"],
            creds["client_id"],
            creds["tenant_id"],
            creds["client_secret"],
        )
        if token:
            st.session_state.access_token = token
            st.session_state.user_info = get_user_info(token)
        st.query_params.clear()
        st.rerun()
 
    st.title("Driver Payroll Processing System")
    st.markdown("---")
 
    # Sidebar Authentication
    with st.sidebar:
        st.header("Authentication")
        if not creds["configured"]:
            st.warning("Azure credentials missing in config.")
        elif not st.session_state.get("access_token"):
            if st.button("Login with Microsoft", type="primary"):
                st.markdown(f"[Click here to login]({get_auth_url(creds['client_id'], creds['tenant_id'])})")
                st.stop()
        else:
            name = (st.session_state.user_info or {}).get("displayName", "User")
            st.success(f"Logged in as: {name}")
            if st.button("Logout"):
                for key in list(st.session_state.keys()): del st.session_state[key]
                st.rerun()
 
        if st.session_state.get("access_token"):
            st.header("SharePoint Site")
            site_url = st.text_input("Enter SharePoint Site URL")
            if st.button("Connect Site") and site_url:
                site, _ = get_site_from_url(st.session_state.access_token, site_url)
                if site: st.session_state.site_info = site
 
    # Main UI Steps
    st.subheader("1. Select Payroll Period")
    d_col1, d_col2 = st.columns(2)
    with d_col1:
        start_date = st.date_input("Start Date", datetime.now() - timedelta(days=14))
    with d_col2:
        end_date = st.date_input("End Date", datetime.now())
 
    st.markdown("---")
    st.subheader("2. Load Data")
    _render_adp_picker()
    _render_sharepoint_file_picker("Relay Files", ["csv"], "relay_files_data", select_mode="multiple")
    _render_sharepoint_file_picker("DriverPay File", ["csv", "xlsx"], "driverpay_data", state_key_name="dp_name")
    _render_sharepoint_file_picker("Override File", ["csv", "xlsx"], "override_data", state_key_name="ov_name")
 
    output_dest, workbook_action = _render_output_config()
 
    if st.button("🚀 Process and Generate Payroll", type="primary", use_container_width=True):
        _handle_process(output_dest, workbook_action, start_date, end_date)
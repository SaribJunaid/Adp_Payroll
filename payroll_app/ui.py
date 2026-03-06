# import io
# from datetime import datetime, timedelta

# import streamlit as st
# from openpyxl import load_workbook

# from payroll_app.config import load_azure_credentials
# from payroll_app.excel_builder import create_excel
# from payroll_app.processing import (
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
#     # Create a unique key prefix based on the title to avoid Duplicate Widget ID errors
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
#         # 1. Basic Processing - Do ADP first to get the name list
#         adp_df, _ = process_adp(st.session_state.adp_files_data)
#         relay_df, _ = process_relay(st.session_state.relay_files_data)
#         dp_df = read_flexible_file(st.session_state.driverpay_data, st.session_state.dp_name)
        
#         if dp_df is None:
#             st.error(f"Could not read DriverPay file: {st.session_state.dp_name}")
#             st.stop()

#         # 2. FIXED: Override Logic passing adp_df for name matching
#         override_map = {}
#         if st.session_state.get("override_data"):
#             override_df = read_flexible_file(st.session_state.override_data, st.session_state.ov_name)
#             if override_df is not None:
#                 # FIX: Added adp_df as the required 4th argument
#                 override_map = build_override_dict(override_df, start_date, end_date, adp_df)

#         # 3. Build Dataset
#         final_df, relay_cols, adp_cols = build_final_dataset(adp_df, relay_df, dp_df)
        
#         # 4. Generate Excel
#         excel_out, anomaly_rows = create_excel(final_df, relay_cols, adp_cols, override_map)

#         # Dynamic Sheet Naming
#         all_dates = [datetime.strptime(col.replace("A_", ""), "%Y-%m-%d").date() for col in adp_cols]
#         sheet_name = f"{min(all_dates).strftime('%d-%b')} to {max(all_dates).strftime('%d-%b')}" if all_dates else "Payroll"
        
#         st.success(f"Excel generated. Sheet: {sheet_name}")

#         # SharePoint Upload Logic (Placeholder - use your existing logic)
#         if output_dest == "SharePoint" and st.session_state.get("site_info"):
#             # Put your specific upload/append logic here
#             pass

#         # UI Indicators
#         if anomaly_rows:
#             st.error(f"{len(anomaly_rows)} anomalies detected (Relay without ADP)")
#         if override_map:
#             st.warning(f"{len(override_map)} override payments applied for the period {start_date} to {end_date}")

#         # Preview Section
#         st.subheader("Payroll Preview")
#         col1, col2, col3, col4 = st.columns(4)
#         col1.metric("Drivers", len(final_df))
#         col2.metric("Relay Days", len(relay_cols))
#         col3.metric("ADP Days", len(adp_cols))
#         col4.metric("Overrides", len(override_map))
#         st.dataframe(final_df, use_container_width=True, height=400)

#         st.download_button(
#             "Download Excel",
#             excel_out.getvalue(),
#             f"Payroll_{sheet_name}.xlsx",
#             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#             type="primary",
#             use_container_width=True,
#         )


# def run_app():
#     st.set_page_config(page_title="Payroll Processor", page_icon="💸", layout="wide")
#     init_session_state()
#     creds = load_azure_credentials()

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

#     with st.sidebar:
#         st.header("Microsoft Login")
#         if not creds["configured"]:
#             st.warning("Azure credentials are not configured in Streamlit secrets")
#         elif not st.session_state.get("access_token"):
#             if st.button("Login with Microsoft", type="primary"):
#                 st.markdown(f"[Click here to login]({get_auth_url(creds['client_id'], creds['tenant_id'])})")
#                 st.info("Sign in with your Microsoft email and password")
#                 st.stop()
#         else:
#             name = (st.session_state.user_info or {}).get("displayName", "User")
#             st.success(name)
#             if st.button("Logout"):
#                 for key in list(st.session_state.keys()):
#                     del st.session_state[key]
#                 st.rerun()

#         st.markdown("---")
#         if st.session_state.get("access_token"):
#             st.header("SharePoint Connection")
#             site_url = st.text_input(
#                 "SharePoint Site URL",
#                 placeholder="https://company.sharepoint.com/sites/YourSite",
#             )
#             if st.button("Connect") and site_url:
#                 site, err = get_site_from_url(st.session_state.access_token, site_url)
#                 if site:
#                     st.session_state.site_info = site
#                     st.success(f"Connected: {site.get('displayName', 'Site')}")
#                 else:
#                     st.error(f"Connection failed: {err}")

#     # 1. DATE RANGE PICKER
#     st.subheader("1. Select Payroll Period")
#     st.info("This range filters the Override File to ensure only relevant extra payments are applied.")
#     d_col1, d_col2 = st.columns(2)
#     with d_col1:
#         start_date = st.date_input("Payroll Start Date", datetime.now() - timedelta(days=14))
#     with d_col2:
#         end_date = st.date_input("Payroll End Date", datetime.now())

#     st.markdown("---")
#     st.subheader("2. Input Files")
#     _render_adp_picker()
    
#     _render_sharepoint_file_picker(
#         title="Relay Files",
#         file_types=["csv"],
#         state_key_data="relay_files_data",
#         select_mode="multiple",
#         confirm_label="Confirm Relay Files",
#     )
#     _render_sharepoint_file_picker(
#         title="DriverPay File",
#         file_types=["csv", "xlsx", "xls"],
#         state_key_data="driverpay_data",
#         state_key_name="dp_name",
#         select_mode="single",
#         confirm_label="Confirm DriverPay File",
#     )
#     _render_sharepoint_file_picker(
#         title="Override File",
#         file_types=["csv", "xlsx", "xls"],
#         state_key_data="override_data",
#         state_key_name="ov_name",
#         select_mode="single",
#         confirm_label="Confirm Override File",
#     )

#     output_dest, workbook_action = _render_output_config()

#     st.markdown("---")
#     if st.button("Process Payroll", type="primary", use_container_width=True):
#         _handle_process(output_dest, workbook_action, start_date, end_date)

#     st.markdown("---")
#     st.caption("Login -> Connect SharePoint -> Select Period -> Select Files -> Process")

import io
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

def _render_sharepoint_file_picker(
    title,
    file_types,
    state_key_data,
    state_key_name=None,
    select_mode="single",
    confirm_label="Confirm File",
    current_path_key="input_path",
):
    key_pref = title.lower().replace(" ", "_")
    
    with st.expander(title, expanded=True):
        source_key = f"{key_pref}_source"
        source = st.radio(
            f"{title} Source:",
            ["Desktop", "SharePoint"],
            key=source_key,
            horizontal=True,
        )

        if source == "Desktop":
            if select_mode == "multiple":
                uploads = st.file_uploader(
                    f"Upload {title}",
                    type=file_types,
                    accept_multiple_files=True,
                    key=f"{key_pref}_upload",
                )
                if uploads:
                    st.session_state[state_key_data] = [f.read() for f in uploads]
                    st.success(f"Loaded {len(uploads)} files")
            else:
                upload = st.file_uploader(
                    f"Upload {title}",
                    type=file_types,
                    key=f"{key_pref}_upload",
                )
                if upload:
                    st.session_state[state_key_data] = upload.read()
                    if state_key_name:
                        st.session_state[state_key_name] = upload.name
                    st.success(f"Loaded {upload.name}")
            return

        if not st.session_state.get("access_token"):
            st.warning("Connect to SharePoint first")
            return

        items, _ = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state[current_path_key],
        )
        if not items:
            st.info("No files found in current folder")
            return

        files = [item for item in items if item["name"].lower().endswith(tuple(file_types))]
        if not files:
            st.info("No matching files in current folder")
            return

        if select_mode == "multiple":
            selected = st.multiselect(
                f"Select {title}:",
                [f["name"] for f in files],
                key=f"{key_pref}_select",
            )
            if st.button(confirm_label, key=f"{key_pref}_confirm"):
                with st.spinner(f"Downloading {title}..."):
                    st.session_state[state_key_data] = []
                    for filename in selected:
                        file_obj = next(f for f in files if f["name"] == filename)
                        content, _ = download_sharepoint_file(
                            st.session_state.access_token,
                            st.session_state.site_info["id"],
                            file_obj["id"],
                        )
                        if content:
                            st.session_state[state_key_data].append(content)
                    st.success(f"Downloaded {len(st.session_state[state_key_data])} files")
        else:
            selected = st.selectbox(
                f"Select {title}:",
                ["-- None --"] + [f["name"] for f in files],
                key=f"{key_pref}_select",
            )
            if selected != "-- None --" and st.button(confirm_label, key=f"{key_pref}_confirm"):
                with st.spinner(f"Downloading {title}..."):
                    file_obj = next(f for f in files if f["name"] == selected)
                    content, _ = download_sharepoint_file(
                        st.session_state.access_token,
                        st.session_state.site_info["id"],
                        file_obj["id"],
                    )
                    if content:
                        st.session_state[state_key_data] = content
                        if state_key_name:
                            st.session_state[state_key_name] = file_obj["name"]
                        st.success(f"Downloaded {file_obj['name']}")

def _render_adp_picker():
    with st.expander("ADP Files", expanded=True):
        source = st.radio("ADP Source:", ["Desktop", "SharePoint"], key="adp_source", horizontal=True)
        if source == "Desktop":
            uploads = st.file_uploader(
                "Upload ADP CSV files",
                type=["csv"],
                accept_multiple_files=True,
                key="adp_upload",
            )
            if uploads:
                st.session_state.adp_files_data = [f.read() for f in uploads]
                st.success(f"Loaded {len(uploads)} ADP files")
            return

        if not st.session_state.get("site_info"):
            st.warning("Connect to SharePoint first")
            return

        col1, col2 = st.columns([3, 1])
        with col1:
            st.caption(f"Path: {st.session_state.input_path}")
        with col2:
            if st.button("Up", key="adp_path_up") and st.session_state.input_path != "root":
                if "/" in st.session_state.input_path:
                    st.session_state.input_path = "/".join(st.session_state.input_path.split("/")[:-1])
                else:
                    st.session_state.input_path = "root"
                st.rerun()

        items, _ = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state.input_path,
        )
        if not items:
            st.info("No items in current folder")
            return

        folders = [i for i in items if "folder" in i]
        csv_files = [i for i in items if i["name"].lower().endswith(".csv")]

        if folders:
            folder = st.selectbox(
                "Navigate to folder:",
                ["-- Stay Here --"] + [f["name"] for f in folders],
                key="adp_folder_select",
            )
            if folder != "-- Stay Here --" and st.button("Enter Folder", key="adp_enter_folder"):
                if st.session_state.input_path == "root":
                    st.session_state.input_path = folder
                else:
                    st.session_state.input_path = f"{st.session_state.input_path}/{folder}"
                st.rerun()

        if not csv_files:
            st.info("No CSV files in current folder")
            return

        selected = st.multiselect(
            "Select ADP CSV files:",
            [f["name"] for f in csv_files],
            key="adp_selected_files",
        )
        if st.button("Confirm ADP Files", key="adp_confirm_files"):
            with st.spinner("Downloading ADP files..."):
                st.session_state.adp_files_data = []
                for filename in selected:
                    file_obj = next(f for f in csv_files if f["name"] == filename)
                    content, _ = download_sharepoint_file(
                        st.session_state.access_token,
                        st.session_state.site_info["id"],
                        file_obj["id"],
                    )
                    if content:
                        st.session_state.adp_files_data.append(content)
                st.success(f"Downloaded {len(st.session_state.adp_files_data)} ADP files")

def _render_output_config():
    st.markdown("---")
    st.subheader("Output Configuration")
    output_dest = st.radio("Save to:", ["Download Only", "SharePoint"], horizontal=True, key="out_dest_radio")
    workbook_action = None

    if output_dest == "SharePoint" and st.session_state.get("site_info"):
        st.markdown("### Select Output Folder")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.caption(f"Output Path: {st.session_state.output_path}")
        with col2:
            if st.button("Up", key="output_up") and st.session_state.output_path != "root":
                if "/" in st.session_state.output_path:
                    st.session_state.output_path = "/".join(st.session_state.output_path.split("/")[:-1])
                else:
                    st.session_state.output_path = "root"
                st.rerun()

        items, _ = list_sharepoint_files(
            st.session_state.access_token,
            st.session_state.site_info["id"],
            st.session_state.output_path,
        )
        if items:
            folders = [i for i in items if "folder" in i]
            if folders:
                folder = st.selectbox(
                    "Navigate to output folder:",
                    ["-- Use Current --"] + [f["name"] for f in folders],
                    key="output_folder_select",
                )
                if folder != "-- Use Current --" and st.button("Enter Folder", key="output_enter_folder"):
                    if st.session_state.output_path == "root":
                        st.session_state.output_path = folder
                    else:
                        st.session_state.output_path = f"{st.session_state.output_path}/{folder}"
                    st.rerun()

        workbook_action = st.radio(
            "Workbook Action:",
            ["Add Sheet to Existing ADP.xlsx", "Create New Workbook"],
            horizontal=True,
            key="wb_action_radio"
        )

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
                    upload_to_sharepoint(token, site_id, path, filename, excel_out.getvalue())
                    st.info(f"File uploaded to {path}")
                else:
                    add_sheet_to_workbook(token, site_id, path, "ADP.xlsx", excel_out.getvalue(), period_str)
                    st.info(f"Sheet '{period_str}' added to ADP.xlsx")

def run_app():
    st.set_page_config(page_title="Payroll Processor", page_icon="💸", layout="wide")
    init_session_state()
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


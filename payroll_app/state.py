import streamlit as st


DEFAULT_SESSION_STATE = {
    "access_token": None,
    "user_info": None,
    "site_info": None,
    "input_path": "root",
    "output_path": "root",
    "adp_files_data": [],
    "relay_files_data": [],
    "driverpay_data": None,
    "override_data": None,
    "dp_name": "",
    "ov_name": "",
}


def init_session_state():
    for key, default_value in DEFAULT_SESSION_STATE.items():
        if key not in st.session_state:
            st.session_state[key] = default_value


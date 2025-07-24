import os
import json
import streamlit as st
import streamlit_authenticator as stauth
import gspread
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
import datetime
import pytz
import re
import dropbox
from googleapiclient.errors import HttpError
import uuid
import pandas as pd
import time

jpm_logo = "https://github.com/marko-londo/coa_testing/blob/main/1752457645003.png?raw=true"
sidebar_logo = "https://github.com/marko-londo/jpm/blob/main/logo_elephant.png?raw=true"

credentials_json = st.secrets["auth_users"]["usernames"]

credentials = json.loads(credentials_json)

authenticator = stauth.Authenticate(
    credentials, 'missed_stops_app', 'some_secret_key', cookie_expiry_days=3)

app_key = st.secrets["dropbox"]["app_key"]

app_secret = st.secrets["dropbox"]["app_secret"]

refresh_token = st.secrets["dropbox"]["refresh_token"]

dbx = dropbox.Dropbox(
    oauth2_refresh_token=refresh_token,
    app_key=app_key,
    app_secret=app_secret
)

SERVICE_ACCOUNT_INFO = st.secrets["google_service_account"]

FOLDER_ID = '1iTHUFwGHpWCAIz88SPBrmjDFJdGsOBJO'

ADDRESS_LIST_SHEET_URL = "https://docs.google.com/spreadsheets/d/1JJeufDkoQ6p_LMe5F-Nrf_t0r_dHrAHu8P8WXi96V9A/edit#gid=0"

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

credentials_gs = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)

gs_client = gspread.authorize(credentials_gs)

st.set_page_config(
    page_title="JPM Ops | JP Mascaro & Sons",
    page_icon="https://raw.githubusercontent.com/marko-londo/coa_testing/refs/heads/main/favicon.ico",
    layout="centered",  # or "wide"
    initial_sidebar_state="collapsed",
    )

st.logo(image=sidebar_logo)

@st.cache_data(ttl=1800)
def load_address_df(_gs_client, sheet_url):
    ws = _gs_client.open_by_url(sheet_url).sheet1
    df = pd.DataFrame(ws.get_all_records())
    return df


with st.spinner("Loading address data..."):
    address_df = load_address_df(gs_client, ADDRESS_LIST_SHEET_URL)

def user_login(authenticator, credentials):
    name, authentication_status, username = authenticator.login('main')

    if authentication_status is False:
        st.error("Incorrect username or password. Please try again.", icon=":material/error:")
        st.stop()
    elif authentication_status is None:
        st.info("Please enter your username and password.", icon=":material/passkey:")
        st.stop()

    user_obj = credentials["usernames"].get(username, {})
    user_role = user_obj.get("role", "city")
    st.info(f"Welcome, {name}!", icon=":material/account_circle:")
    authenticator.logout("Logout", "sidebar")
    return name, username, user_role

def header():
    st.markdown(
        f"""
        <div style='display: flex; justify-content: center; align-items: center; margin-bottom: 12px;'>
            <img src='{jpm_logo}' width='320'>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("""
        <style>
        h1 {
            font-family: 'Poppins', sans-serif !important;
            font-weight: 700 !important;
            font-size: 3em !important;
            letter-spacing: 1.5px !important;
            text-shadow:
                -1px -1px 0 #181b20,
                 1px -1px 0 #181b20,
                -1px  1px 0 #181b20,
                 1px  1px 0 #181b20,
                 0  3px 12px #6CA0DC55;
        }
        </style>
        """, unsafe_allow_html=True)

    st.markdown(
        """
        <div style='text-align:center;'>
            <h1 style='color:#6CA0DC; margin-bottom:0;'>Operations Portal</h1>
            </div>
            <hr style='border:1px solid #ececec; margin-top:0;'>
        </div>
        """,
        unsafe_allow_html=True
    )

def ensure_completion_times_gsheet_exists(drive, folder_id, title):
    results = drive.files().list(
        q=f"'{folder_id}' in parents and name='{title}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    else:
        # If you want to create it automatically, implement creation logic here.
        st.error(
            f"Completion Times sheet '{title}' does not exist in the specified folder.\n"
            "Please contact your admin to create this week's completion log sheet.", icon=":material/error:"
        )
        st.stop()

def get_today_operating_zone(address_df):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    today = datetime.datetime.now().date()
    today_idx = today.weekday()
    if today_idx == 6:  # Sunday
        zone_day = "Friday"
    else:
        zone_day = days[today_idx - 1]  # minus one day
    return zone_day

def get_yw_zone_color(today=None):
    if today is None:
        today = datetime.datetime.now().date()
    year = 2025
    june_first = datetime.date(year, 6, 1)
    first_monday = june_first + datetime.timedelta(days=(0 - june_first.weekday() + 7) % 7)
    weeks_since = (today - first_monday).days // 7
    if weeks_since % 2 == 0:
        return "140", "Blue"
    else:
        return "141", "Yellow"

def dashboard():
    # 1. Operating zone
    zone_day = get_today_operating_zone(address_df)
    st.markdown(f"## Operating Zone (for today): <span style='color:#FF8C8C;'>{zone_day}</span>", unsafe_allow_html=True)

    # 2. YW (Yardwaste) zone color this week
    yw_route, yw_color = get_yw_zone_color()
    color_code = "#3980ec" if yw_color == "Blue" else "#EAC100"
    st.markdown(f"**Yardwaste Zone this week:** <span style='color:{color_code};font-weight:bold;'>{yw_route} ({yw_color})</span>", unsafe_allow_html=True)

    # 3. Count unique routes per service type for today's zone
    service_info = [
        ("MSW", "MSW Zone", "MSW Route", "#57B560"),    # Light green
        ("SS",  "SS Zone",  "SS Route", "#4FC3F7"),    # Light blue
        ("YW",  "YW Zone",  "YW Route", "#F6C244"),    # Mustard/yellow
    ]
    col1, col2, col3 = st.columns(3)
    for i, (label, zone_col, route_col, color) in enumerate(service_info):
        # Filter addresses where the zone for this service == today's operating zone
        valid = address_df[address_df[zone_col].astype(str).str.lower() == zone_day.lower()]
        # For YW, also filter by zone color
        if label == "YW":
            valid = valid[valid["YW Route"].astype(str).str.endswith(yw_route)]
        routes = valid[route_col].unique()
        count = len(routes)
        with [col1, col2, col3][i]:
            st.markdown(
                f"<div style='background-color:{color};padding:18px 0;border-radius:10px;text-align:center;'>"
                f"<span style='font-weight:bold;font-size:1.6em;'>{count}</span><br>"
                f"<span style='font-size:1.1em'>{label}</span></div>", unsafe_allow_html=True
            )
        


def hotlist():
    st.write("Hotlist")

def testing():
    st.write("Testing")

def ops(name, user_role):
    st.sidebar.subheader("Operations")
    op_select = st.sidebar.radio("Select Operation:", ["Dashboard", "Hotlist", "Testing"])
    if op_select == "Dashboard":
        dashboard()
    elif op_select == "Hotlist":
        hotlist()
    elif op_select == "Testing":
        testing()

name, username, user_role = user_login(authenticator, credentials)
header()
if user_role == "jpm":
    ops(name, user_role)



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

# --------------------------
# GLOBAL CONSTANTS & CONFIG
# --------------------------

# Logos and Branding
JPM_LOGO = "https://github.com/marko-londo/coa_testing/blob/main/1752457645003.png?raw=true"
SIDEBAR_LOGO = "https://github.com/marko-londo/jpm/blob/main/logo_elephant.png?raw=true"

# Credentials and Secrets
CREDENTIALS_JSON = st.secrets["auth_users"]["usernames"]
CREDENTIALS = json.loads(CREDENTIALS_JSON)
SERVICE_ACCOUNT_INFO = st.secrets["google_service_account"]
COOKIE_SECRET = st.secrets["auth"]["cookie_secret"] 

authenticator = stauth.Authenticate(
    CREDENTIALS, 'missed_stops_app', COOKIE_SECRET, cookie_expiry_days=3)

# Google/Dropbox Config
FOLDER_ID = '1ogx3zPeIdTKp7C5EJ5jKavFv21mDmySj'
ADDRESS_LIST_SHEET_URL = "https://docs.google.com/spreadsheets/d/1JJeufDkoQ6p_LMe5F-Nrf_t0r_dHrAHu8P8WXi96V9A/edit#gid=0"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Service Types & Statuses
SERVICE_TYPES = ["MSW", "SS", "YW"]
RESOLVED_STATUSES = {
    "PICKED UP", "REJECTED", "CONFIRMED PREMATURE",
    "ONE TIME EXCEPTION", "NOT OUT", "CREATED IN ERROR"
}
LEGITIMATE_STATUS = "PICKED UP"

# Timezone and Date
NY_TZ = pytz.timezone("America/New_York")
TODAY = datetime.datetime.now(NY_TZ).date()
THIS_MONTH = TODAY.strftime("%Y-%m")

# Google API Auth
CREDENTIALS_GS = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
GS_CLIENT = gspread.authorize(CREDENTIALS_GS)
DRIVE_SERVICE = build('drive', 'v3', credentials=CREDENTIALS_GS)

# Dropbox Auth
APP_KEY = st.secrets["dropbox"]["app_key"]
APP_SECRET = st.secrets["dropbox"]["app_secret"]
REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
DBX = dropbox.Dropbox(
    oauth2_refresh_token=REFRESH_TOKEN,
    app_key=APP_KEY,
    app_secret=APP_SECRET
)

# Streamlit Config
st.set_page_config(
    page_title="JPM Ops | JP Mascaro & Sons",
    page_icon="https://raw.githubusercontent.com/marko-londo/coa_testing/refs/heads/main/favicon.ico",
    layout="centered",  # or "wide"
    initial_sidebar_state="collapsed",
)
st.logo(image=SIDEBAR_LOGO)

# --------------------------
# UTILITY FUNCTIONS
# --------------------------

def clean_status(val):
    return str(val).strip().upper()

@st.cache_data(ttl=1800)
def load_address_df(_gs_client, sheet_url):
    ws = _gs_client.open_by_url(sheet_url).sheet1
    df = pd.DataFrame(ws.get_all_records())
    return df

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
            <img src='{JPM_LOGO}' width='320'>
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
        st.error(
            f"Completion Times sheet '{title}' does not exist in the specified folder.\n"
            "Please contact your admin to create this week's completion log sheet.", icon=":material/error:"
        )
        st.stop()

def get_today_operating_zone(address_df):
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    today_idx = TODAY.weekday()
    if today_idx == 6:  # Sunday
        zone_day = "Friday"
    else:
        zone_day = days[today_idx - 1]  # minus one day
    return zone_day

def get_yw_zone_color(today=None):
    if today is None:
        today = TODAY
    year = today.year
    june_first = datetime.date(year, 6, 1)
    first_monday = june_first + datetime.timedelta(days=(0 - june_first.weekday() + 7) % 7)
    weeks_since = (today - first_monday).days // 7
    return "140" if weeks_since % 2 == 0 else "141"

def get_tab_date(day="today"):
    # Returns correct date for today/yesterday logic
    if day == "today":
        if TODAY.weekday() == 6:  # Sunday
            return TODAY - datetime.timedelta(days=1)
        return TODAY
    elif day == "yesterday":
        if TODAY.weekday() == 0:  # Monday -> Sat
            return TODAY - datetime.timedelta(days=2)
        elif TODAY.weekday() == 6:  # Sunday -> Fri
            return TODAY - datetime.timedelta(days=2)
        return TODAY - datetime.timedelta(days=1)
    else:
        raise ValueError("day must be 'today' or 'yesterday'")

def get_sheet_title(date):
    # Replicate your week-ending logic here if needed!
    # Placeholder example:
    next_saturday = date + datetime.timedelta((5-date.weekday()) % 7)
    return f"Misses Week Ending {next_saturday.strftime('%Y-%m-%d')}"

def get_today_tab_name(date):
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    # Find this week's Monday
    next_saturday = date + datetime.timedelta((5-date.weekday()) % 7)
    monday = next_saturday - datetime.timedelta(days=5)
    idx = (date - monday).days
    label = weekdays[idx] if 0 <= idx < 6 else weekdays[0]
    return f"{label} {date.month}/{date.day}/{str(date.year)[-2:]}"

def get_tab_records(day="today"):
    date = get_tab_date(day)
    sheet_title = get_sheet_title(date)
    tab_name = get_today_tab_name(date)
    # Find the sheet ID
    results = DRIVE_SERVICE.files().list(
        q=f"'{FOLDER_ID}' in parents and name='{sheet_title}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if not files:
        return []
    sheet_id = files[0]['id']
    weekly_ss = GS_CLIENT.open_by_key(sheet_id)
    try:
        ws = weekly_ss.worksheet(tab_name)
        records = ws.get_all_records()
        return records
    except Exception:
        return []

def get_week_records():
    date = TODAY
    sheet_title = get_sheet_title(date)
    results = DRIVE_SERVICE.files().list(
        q=f"'{FOLDER_ID}' in parents and name='{sheet_title}' and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if not files:
        return []
    sheet_id = files[0]['id']
    weekly_ss = GS_CLIENT.open_by_key(sheet_id)
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    week_records = []
    next_saturday = date + datetime.timedelta((5-date.weekday()) % 7)
    monday = next_saturday - datetime.timedelta(days=5)
    for i in range(6):
        tab_date = monday + datetime.timedelta(days=i)
        tab_name = f"{weekdays[i]} {tab_date.month}/{tab_date.day}/{str(tab_date.year)[-2:]}"
        try:
            ws = weekly_ss.worksheet(tab_name)
            week_records.extend(ws.get_all_records())
        except Exception:
            continue
    return week_records

def get_month_records():
    # Master Misses Log: must be named exactly as such in folder
    results = DRIVE_SERVICE.files().list(
        q=f"'{FOLDER_ID}' in parents and name = 'Master Misses Log' and mimeType = 'application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = results.get('files', [])
    if not files:
        return []
    sheet_id = files[0]['id']
    master_ws = GS_CLIENT.open_by_key(sheet_id).sheet1
    records = master_ws.get_all_records()
    filtered = [
        row for row in records
        if str(row.get("Time Sent to JPM", "")).startswith(THIS_MONTH)
    ]
    return filtered

# --- CACHED SHEETS READS ---

@st.cache_data(ttl=300)  # 5 minutes; adjust as needed
def get_tab_records_cached(day="today"):
    return get_tab_records(day)

@st.cache_data(ttl=300)
def get_week_records_cached():
    return get_week_records()

@st.cache_data(ttl=300)
def get_month_records_cached():
    return get_month_records()


def compute_stats(records, service_types=SERVICE_TYPES):
    result = {}
    for service in service_types + ["ALL"]:
        result[service] = {
            "total_misses": 0,
            "legit_misses": 0,
            "illegit_misses": 0,
            "resolved": 0,
            "pct_resolved": 0.0,
            "pct_legit": 0.0,
        }
    for row in records:
        addr = row.get("Address", "").strip()
        if not addr:
            continue
        status = clean_status(row.get("Collection Status", ""))
        service = row.get("Service Type", "").strip().upper()
        is_resolved = status in RESOLVED_STATUSES
        is_legit = status == LEGITIMATE_STATUS
        applicable_services = [service] if service in service_types else []
        applicable_services.append("ALL")  # Always track total
        for s in applicable_services:
            result[s]["total_misses"] += 1
            if is_legit:
                result[s]["legit_misses"] += 1
            if is_resolved:
                result[s]["resolved"] += 1
    for s in result:
        result[s]["illegit_misses"] = result[s]["resolved"] - result[s]["legit_misses"]
        t = result[s]["total_misses"]
        result[s]["pct_resolved"] = (result[s]["resolved"] / t * 100) if t else 0
        result[s]["pct_legit"] = (result[s]["legit_misses"] / t * 100) if t else 0
    return result


# --------------------------
# PAGE LOGIC FUNCTIONS
# --------------------------

def dashboard():
    header()

    # Today's Zone
    st.markdown("### Today's Zone")
    st.markdown(f"<span style='color:#FF8C8C;font-size:1.3em'>{zone_day}</span>", unsafe_allow_html=True)

    # Weekly Yardwaste Zone
    st.markdown("### Weekly Yardwaste Zone")
    st.markdown(f"<span style='color:{color_code};font-weight:bold;font-size:1.3em'>{yw_route}</span>", unsafe_allow_html=True)


    # 2. Route Counts by Service Type
    st.markdown("#### Route Counts by Service")
    service_info = [
        ("MSW Routes", "MSW Zone", "MSW Route", "#57B560"),
        ("SS Routes",  "SS Zone",  "SS Route", "#4FC3F7"),
        ("YW Routes",  "YW Zone",  "YW Route", "#F6C244"),
    ]
    col1, col2, col3 = st.columns([1, 1, 1], gap="medium")
    for i, (label, zone_col, route_col, color) in enumerate(service_info):
        valid = address_df[address_df[zone_col].astype(str).str.lower() == zone_day.lower()]
        if "YW" in label:
            valid = valid[valid["YW Route"].astype(str).str.contains(yw_route, na=False)]
        routes = valid[route_col].unique()
        count = len(routes)
        label_display = label.replace("Routes", "Route" if count == 1 else "Routes")
        with [col1, col2, col3][i]:
            st.markdown(
                f"""
                <div style='background-color:{color};
                            padding:10px 0 6px 0;
                            border-radius:10px;
                            text-align:center;
                            min-width:80px;
                            min-height:60px;
                            margin:0 auto;
                            box-shadow:0 1px 6px #2222;'>
                <span style='font-weight:700;font-size:1.15em;'>{count}</span><br>
                <span style='font-size:0.95em'>{label_display}</span>
                </div>
                """, unsafe_allow_html=True
            )

    st.divider()

    # 3. Missed Stop Statistics Section
    with st.spinner("Loading missed stop stats..."):
        today_stats = compute_stats(get_tab_records_cached("today"))
        yesterday_stats = compute_stats(get_tab_records_cached("yesterday"))
        week_stats = compute_stats(get_week_records_cached())
        month_stats = compute_stats(get_month_records_cached())


    def stats_table(stats, title):
        st.markdown(f"**{title}**")
        table = []
        for key in ["ALL"] + SERVICE_TYPES:
            s = stats[key]
            label = "Total" if key == "ALL" else key
            table.append({
                "Service": label,
                "Submitted": s["total_misses"],
                "Legitimate": s["legit_misses"],
                "Illegitimate": s["illegit_misses"],
                "Resolved": s["resolved"],
                "% Resolved": f"{s['pct_resolved']:.1f}%",
                "% Legitimate": f"{s['pct_legit']:.1f}%"
            })
        st.dataframe(pd.DataFrame(table), hide_index=True, use_container_width=True)

    stats_table(today_stats, "Today's Missed Stops")
    stats_table(yesterday_stats, "Yesterday's Missed Stops")
    stats_table(week_stats, "This Week's Missed Stops")
    stats_table(month_stats, "This Month's Missed Stops")
    st.divider()

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

# --------------------------
# MAIN APP EXECUTION
# --------------------------

with st.spinner("Loading address data..."):
    address_df = load_address_df(GS_CLIENT, ADDRESS_LIST_SHEET_URL)

name, username, user_role = user_login(authenticator, CREDENTIALS)
if user_role == "jpm":
    ops(name, user_role)


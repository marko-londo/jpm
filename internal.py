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
    page_title="MPU Portal | JP Mascaro & Sons",
    page_icon="https://raw.githubusercontent.com/marko-londo/coa_testing/refs/heads/main/favicon.ico",
    layout="centered",  # or "wide"
    initial_sidebar_state="collapsed",
    )

st.logo(image=sidebar_logo)

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

def updates():
    APP_VERSION = "v2.3"
    CHANGELOG = """
    - **v2.3** (2025-07-18):  
        - Added “Submit Completion Times” section for JPM
        - Stops submitted before completion time will be flagged as potentially premature


    """

    
    st.markdown("<br>", unsafe_allow_html=True)  # One blank line
    
    # --- Centered Logo ---
    st.markdown(
        f"""
        <div style='display: flex; justify-content: center; align-items: center; margin-bottom: 12px;'>
            <img src='{jpm_logo}' width='320'>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # --- H1 Style ---
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
    
    # --- Centered Header, Subtitle, and Divider ---
    st.markdown(
        """
        <div style='text-align:center;'>
            <h1 style='color:#6CA0DC; margin-bottom:0;'>Missed Pickup Portal</h1>
            <div style='font-size:1.1em; font-style:italic; margin-bottom:12px;'>
                <span style='color:#FF8C8C;'>City of Allentown</span>
                <span style='color:#fff; padding:0 10px;'>|</span>
                <span style='color:#FF8C8C;'>JP Mascaro & Sons</span>
            </div>
            <hr style='border:1px solid #ececec; margin-top:0;'>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # --- App Version (left-aligned) ---
    st.markdown(f"<div style='color:gray;margin-bottom:8px;'>{APP_VERSION}</div>", unsafe_allow_html=True)

    with st.expander("What's New?", expanded=False):
            st.markdown(CHANGELOG)
        
    doc_col, sht_col, fold_col = st.columns(3)
    
    with doc_col:
            
        DOC_LINK = "https://docs.google.com/document/d/1UkKj56Qn-25gMWheC-G2rC6YRJzeGsfxk9k2XNLpeTw"
        st.link_button("📄 View Full Docs", DOC_LINK)

    with sht_col:
        st.link_button("Open Sheet", f"https://docs.google.com/spreadsheets/d/{weekly_id}/edit")

    with fold_col:
        st.link_button("Open Folder", f"https://drive.google.com/drive/u/0/folders/1ogx3zPeIdTKp7C5EJ5jKavFv21mDmySj")


updates()
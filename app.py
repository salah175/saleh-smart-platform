import streamlit as st
import pandas as pd
import gspread
import urllib.parse
import datetime
import hashlib
import io
import re
from google.oauth2.service_account import Credentials
try:`n    from weasyprint import HTML`nexcept Exception:`n    HTML = None
import math

# ==========================================
# âڑ™ï¸ڈ 1. ط¥ط¹ط¯ط§ط¯ط§طھ ط§ظ„ظ†ط¸ط§ظ…
# ==========================================
st.set_page_config(page_title="ظ…ظ†طµط© طµط§ظ„ط­ ط§ظ„ط°ظƒظٹط©", layout="wide", initial_sidebar_state="collapsed")

# --- ًںژ¨ طھط¹ط±ظٹظپ ط§ظ„ط£ظ„ظˆط§ظ† (ط§ظ„ط«ظٹظ… ط§ظ„ظ…ط¤ط³ط³ظٹ ط§ظ„ط­ط¯ظٹط« ط­ط³ط¨ ط§ظ„طھطµظ…ظٹظ… ط§ظ„ط¬ط¯ظٹط¯) ---
main_bg = "#F8FAFC"
card_bg = "#FFFFFF"
text_color = "#0F172A"
sub_text = "#64748B"
border_color = "#E2E8F0"
primary_color = "#2563EB"
accent_color = "#1E40AF"
success_color = "#10B981"
warning_color = "#F59E0B"
danger_color = "#EF4444"
header_grad = "linear-gradient(135deg, #1E40AF 0%, #2563EB 100%)"
shadow_val = "0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03)"

# --- [ط§ظ„ط¯ظˆط§ظ„ ط§ظ„ظ…ط³ط§ط¹ط¯ط© ظˆط§ظ„ط§طھطµط§ظ„ ط§ظ„ط°ظƒظٹ] ---

@st.cache_resource(ttl=2700) 
def get_gspread_client():
    try:
        creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds).open_by_key(st.secrets["SHEET_ID"])
    except Exception as e: 
        st.error(f"âڑ ï¸ڈ ط®ط·ط£ ط§طھطµط§ظ„ ط¨ط³ظٹط±ظپط± ط¬ظˆط¬ظ„: {e}")
        return None

sh = get_gspread_client()

def normalize_arabic(text):
    if not isinstance(text, str): text = str(text)
    text = re.sub(r'[ط£ط¥ط¢ط§]', 'ط§', text)
    text = re.sub(r'ط©', 'ظ‡', text)
    text = re.sub(r'ظ‰', 'ظٹ', text)
    return text.strip()

def clean_phone_number(phone):
    p = str(phone).strip().replace(" ", "")
    if p.startswith("0"): p = p[1:]
    if not p.startswith("966") and p != "": p = "966" + p
    return p

def get_professional_msg(name, b_type, b_desc, date):
    msg = (f"ًں”” *ط¥ط´ط¹ط§ط± ظ…ظ† ظ…ظ†طµط© ط§ظ„ط£ط³طھط§ط° طµط§ظ„ط­*\nًں‘¤ *ط§ظ„ط·ط§ظ„ط¨:* {name}\nًں“چ *ط§ظ„ظ…ظ„ط§ط­ط¸ط©:* {b_type}\nًں“‌ *ط§ظ„طھظپط§طµظٹظ„:* {b_desc if b_desc else 'ظ…طھط§ط¨ط¹ط©'}\nًں“… *ط§ظ„طھط§ط±ظٹط®:* {date}")
    return urllib.parse.quote(msg)

def show_footer():
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='text-align: center; color: {sub_text}; padding: 20px; border-top: 1px solid {border_color};'>
        <p style='margin-bottom: 10px; font-size: 0.9rem;'>ط¬ظ…ظٹط¹ ط§ظ„ط­ظ‚ظˆظ‚ ظ…ط­ظپظˆط¸ط© ظ„ظ…ظ†طµط© ط§ظ„ط£ط³طھط§ط° طµط§ظ„ط­ ط§ظ„ط°ظƒظٹط© آ© 2026</p>
    </div>
    """, unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.link_button("ًں“¢ طھظ„ظٹط¬ط±ط§ظ… ط§ظ„ط¥ط¯ط§ط±ط©", "https://t.me/@0530291027", use_container_width=True)
    c2.link_button("ًں’¬ ظˆط§طھط³ط§ط¨ ط§ظ„ظ…ط¹ظ„ظ…", "https://wa.me/966544449694", use_container_width=True)
    c3.link_button("ًں“§ ط§ظ„ط¨ط±ظٹط¯ ط§ظ„ط¥ظ„ظƒطھط±ظˆظ†ظٹ", "mailto:salah-talal1@hotmail.com", use_container_width=True)

@st.cache_data(ttl=300)
def fetch_safe(worksheet_name):
    try:
        current_sh = get_gspread_client()
        if not current_sh: return pd.DataFrame()
        ws = current_sh.worksheet(worksheet_name)
        data = ws.get_all_values()
        if not data: return pd.DataFrame()
        df = pd.DataFrame(data[1:], columns=data[0])
        if not df.empty: df.iloc[:, 0] = df.iloc[:, 0].astype(str).str.strip()
        return df
    except Exception:
        st.cache_resource.clear()
        st.toast("âڑ ï¸ڈ طھظ… طھط­ط¯ظٹط« ط§ظ„ط§طھطµط§ظ„ ظ…ط¹ ط§ظ„ط³ظٹط±ظپط±طŒ ط­ط§ظˆظ„ ظ…ط¬ط¯ط¯ط§ظ‹.", icon="ًں”„")
        return pd.DataFrame()

def safe_append_row(worksheet_name, data_dict):
    try:
        current_sh = get_gspread_client()
        ws = current_sh.worksheet(worksheet_name)
        headers = ws.row_values(1)
        row = [data_dict.get(h, "") for h in headers]
        ws.append_row(row)
        return True
    except Exception:
        st.cache_resource.clear()
        st.error("âڑ ï¸ڈ ط­ط¯ط« ط§ظ†ظ‚ط·ط§ط¹طŒ طھظ… ط§ظ„طھط­ط¯ظٹط«. ظٹط±ط¬ظ‰ ط§ظ„ط¶ط؛ط· ظ…ط±ط© ط£ط®ط±ظ‰.")
        return False

# --- طھط­ظ…ظٹظ„ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ ---
if "class_options" not in st.session_state:
    try:
        sett = sh.worksheet("settings").get_all_records()
        s_map = {row['key']: row['value'] for row in sett}
        st.session_state.max_tasks = int(s_map.get('max_tasks', 60))
        st.session_state.max_quiz = int(s_map.get('max_quiz', 40))
        st.session_state.current_year = str(s_map.get('current_year', '1447ظ‡ظ€'))
        st.session_state.class_options = [x.strip() for x in str(s_map.get('class_list', 'ط§ظ„ط£ظˆظ„')).split(',') if x.strip()]
        st.session_state.stage_options = [x.strip() for x in str(s_map.get('stage_list', 'ط§ط¨طھط¯ط§ط¦ظٹ')).split(',') if x.strip()]
    except:
        st.session_state.max_tasks, st.session_state.max_quiz = 60, 40
        st.session_state.current_year = "1447ظ‡ظ€"
        st.session_state.class_options = ["ط§ظ„ط£ظˆظ„"]; st.session_state.stage_options = ["ط§ط¨طھط¯ط§ط¦ظٹ"]

if "role" not in st.session_state: st.session_state.role = None
if "username" not in st.session_state: st.session_state.username = None

# ==========================================
# ًںژ¨ 2. ط§ظ„طھطµظ…ظٹظ… (CSS)
# ==========================================
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;900&display=swap');
    
    section[data-testid="stSidebar"] {{ display: none; }}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    
    html, body, [data-testid="stAppViewContainer"] {{ 
        font-family: 'Tajawal', sans-serif !important; 
        direction: RTL; text-align: right; 
        background-color: {main_bg} !important; color: {text_color} !important; 
    }}
    
    .block-container {{ padding-top: 1rem !important; padding-bottom: 5rem; max-width: 1000px; }}
    
    .header-container {{
        background: {header_grad};
        padding: 70px 20px 40px 20px; /* âœ³ï¸ڈ طھظ… ط²ظٹط§ط¯ط© ط§ظ„ظ…ط³ط§ط­ط© ط§ظ„ط¹ظ„ظˆظٹط© ظ„طھط¬ظ†ط¨ ظ‚طµ ط§ظ„ظ‚ط¨ط¹ط© */
        border-radius: 0 0 40px 40px;
        margin: -1rem -5rem 30px -5rem;
        box-shadow: 0 10px 30px -10px rgba(37, 99, 235, 0.4);
        color: white; 
        text-align: center;
    }}
    
    .title-wrapper {{
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 15px; 
        margin-bottom: 10px;
    }}
    
    .logo-icon {{ 
        font-size: 3.8rem; 
        filter: drop-shadow(0 4px 6px rgba(0,0,0,0.2));
        animation: float 4s ease-in-out infinite;
        line-height: 1;
        margin-top: 12px; /* âœ³ï¸ڈ طھظ†ط²ظٹظ„ ط§ظ„ظ‚ط¨ط¹ط© ظ„ظ„ط£ط³ظپظ„ ظ„ظƒظٹ طھطھظˆط§ط²ظ‰ ظ…ط¹ ط§ظ„ظ†طµ طھظ…ط§ظ…ط§ظ‹ */
    }}
    
    .main-title {{ 
        margin: 0; 
        font-size: 3rem; 
        font-weight: 900; 
        color: #ffffff !important; 
        line-height: 1;
    }}
    
    .sub-title {{ 
        margin: 0; 
        color: #DBEAFE; 
        font-size: 1.2rem; 
        font-weight: 500; 
    }}
    
    div[data-baseweb="input"], div[data-baseweb="base-input"], div[data-baseweb="select"] {{ 
        background-color: #F8FAFC !important; border: 2px solid #E2E8F0 !important; border-radius: 12px !important; height: 50px; 
    }}
    input, textarea, select {{ 
        color: #0F172A !important; -webkit-text-fill-color: #0F172A !important; caret-color: {primary_color} !important;
        background-color: transparent !important; font-weight: 700 !important; font-size: 1.05rem !important;
    }}
    ::placeholder {{ color: #94A3B8 !important; opacity: 1 !important; -webkit-text-fill-color: #94A3B8 !important; }}
    div[data-baseweb="select"] div {{ color: #0F172A !important; }}
    
    [data-testid="stFormSubmitButton"] button, 
    [data-testid="baseButton-primary"], 
    div.stButton > button {{
        background: linear-gradient(135deg, #1E40AF 0%, #2563EB 100%) !important;
        background-color: #2563EB !important;
        color: white !important; border: none !important; font-weight: 800 !important;
        font-size: 1.1rem !important; border-radius: 12px !important; padding: 12px 20px !important;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2) !important; transition: all 0.2s !important; width: 100%; height: 50px;
    }}
    [data-testid="stFormSubmitButton"] button:hover, 
    [data-testid="baseButton-primary"]:hover, 
    div.stButton > button:hover {{ 
        background: linear-gradient(135deg, #1e3a8a 0%, #1d4ed8 100%) !important; 
        box-shadow: 0 6px 12px rgba(30, 64, 175, 0.3) !important; 
    }}
    
    .app-card {{ background: {card_bg}; padding: 20px; border-radius: 16px; box-shadow: {shadow_val}; border: 1px solid {border_color}; margin-bottom: 15px; }}
    
    .stTabs [data-baseweb="tab-list"] {{ gap: 10px; background-color: transparent; border: none; }}
    .stTabs [data-baseweb="tab"] {{ height: 50px; background-color: #FFFFFF; border-radius: 12px; border: 1px solid #E2E8F0; color: #64748B; font-weight: bold; flex: 1; justify-content: center; transition: 0.3s; }}
    .stTabs [aria-selected="true"] {{ background-color: {primary_color} !important; color: white !important; border: none !important; box-shadow: 0 4px 6px rgba(37, 99, 235, 0.2); }}

    .mobile-list-item {{ background: white; border-radius: 12px; padding: 16px; margin-bottom: 12px; border: 1px solid #E2E8F0; box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; align-items: center; justify-content: space-between; }}
    
    .medal-flex {{ display: flex; gap: 10px; margin: 20px 0; direction: rtl; }}
    .m-card {{ flex: 1; background: white; padding: 15px 5px; border-radius: 16px; text-align: center; border: 1px solid #E2E8F0; box-shadow: {shadow_val}; }}
    .m-active {{ border: 2px solid {warning_color} !important; background: #FFFBEB !important; }}
    
    .points-banner {{ background: {warning_color}; color: white; padding: 25px; border-radius: 16px; text-align: center; margin-bottom: 25px; box-shadow: 0 4px 10px rgba(245, 158, 11, 0.3); }}
    .welcome-card {{ background: {header_grad}; color: white; padding: 20px; border-radius: 16px; margin-bottom: 15px; box-shadow: 0 4px 10px rgba(37, 99, 235, 0.3); }}

    @keyframes float {{ 0%, 100% {{ transform: translateY(0); }} 50% {{ transform: translateY(-10px); }} }}
    @keyframes pulse-red {{ 0% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.5); }} 70% {{ box-shadow: 0 0 0 10px rgba(239, 68, 68, 0); }} 100% {{ box-shadow: 0 0 0 0 rgba(239, 68, 68, 0); }} }}
    .urgent-box {{ background-color: #FEF2F2; border: 2px solid {danger_color}; color: #991B1B; padding: 15px; border-radius: 12px; text-align: center; animation: pulse-red 2s infinite; font-weight: bold; margin-bottom: 25px; }}

    @media (max-width: 768px) {{
        .header-container {{ padding: 60px 20px 30px 20px; }}
        .main-title {{ font-size: 2rem; }}
        .logo-icon {{ font-size: 2.5rem; }}
    }}
    </style>

    <div class="header-container">
        <div class="title-wrapper">
            <div class="logo-icon">ًںژ“</div>
            <h1 class="main-title">ظ…ظ†طµط© طµط§ظ„ط­ ط§ظ„ط°ظƒظٹط©</h1>
        </div>
        <p class="sub-title">ظ…ظ†طµظ‡ طµط§ظ„ط­ ط§ظ„ط°ظƒظٹط© 2026</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# ًں”گ 3. ظ†ط¸ط§ظ… ط§ظ„ط¯ط®ظˆظ„
# ==========================================
if st.session_state.role is None:
    c1, c2 = st.columns([1, 10]) 
    t1, t2, t3 = st.tabs(["ًںژ“ ط¨ظˆط§ط¨ط© ط§ظ„ط·ظ„ط§ط¨", "ًں‘¨â€چًں’¼ ط¨ظˆط§ط¨ط© ط§ظ„ظ…ط¹ظ„ظ…", "ًںڈ« ط¨ظˆط§ط¨ط© ط§ظ„ط¥ط¯ط§ط±ط© (ظ…ط´ط§ظ‡ط¯)"])
    
    with t1:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("st_login"):
            st.markdown("<h4 style='text-align:center; margin-bottom:20px;'>طھط³ط¬ظٹظ„ ط¯ط®ظˆظ„ ط§ظ„ط·ط§ظ„ط¨</h4>", unsafe_allow_html=True)
            sid = st.text_input("ط±ظ‚ظ… ط§ظ„ظ‡ظˆظٹط© / ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ", placeholder="ط£ط¯ط®ظ„ ط§ظ„ط±ظ‚ظ… ظ‡ظ†ط§...").strip()
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("ًںڑ€ ط¯ط®ظˆظ„ ظ„ظ„ظ…ظ†طµط©", type="primary", use_container_width=True):
                df = fetch_safe("students")
                if not df.empty:
                    df['clean_id'] = df.iloc[:,0].astype(str).str.split('.').str[0].str.strip()
                    if sid.split('.')[0] in df['clean_id'].values:
                        st.session_state.username = sid.split('.')[0]
                        st.session_state.role = "student"
                        st.rerun()
                    else: st.error("âڑ ï¸ڈ ط§ظ„ط±ظ‚ظ… ط؛ظٹط± ظ…ط³ط¬ظ„ ظپظٹ ط§ظ„ظ†ط¸ط§ظ…")
                    
    with t2:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("tr_login"):
            st.markdown("<h4 style='text-align:center; margin-bottom:20px;'>طھط³ط¬ظٹظ„ ط¯ط®ظˆظ„ ط§ظ„ظ…ط¹ظ„ظ…</h4>", unsafe_allow_html=True)
            u = st.text_input("ط§ط³ظ… ط§ظ„ظ…ط³طھط®ط¯ظ…", placeholder="User")
            p = st.text_input("ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±", type="password", placeholder="******")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("ًں› ï¸ڈ ط¯ط®ظˆظ„ ظ„ظˆط­ط© ط§ظ„طھط­ظƒظ…", type="primary", use_container_width=True):
                df = fetch_safe("users")
                if not df.empty and u in df['username'].values:
                    ud = df[df['username']==u].iloc[0]
                    if hashlib.sha256(p.encode()).hexdigest() == ud['password_hash']:
                        if ud.get('role', 'teacher') in ['teacher', '']:
                            st.session_state.username = u
                            st.session_state.role = "teacher"
                            st.rerun()
                        else:
                            st.error("â‌Œ ظ‡ط°ط§ ط§ظ„ط­ط³ط§ط¨ ظ„ط§ ظٹظ…ظ„ظƒ طµظ„ط§ط­ظٹط© ط§ظ„ظ…ط¹ظ„ظ….")
                    else:
                        st.error("â‌Œ ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط± ط؛ظٹط± طµط­ظٹط­ط©.")
                else:
                    st.error("â‌Œ ط§ط³ظ… ط§ظ„ظ…ط³طھط®ط¯ظ… ط؛ظٹط± ظ…ظˆط¬ظˆط¯.")
                    
    with t3:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("admin_login"):
            st.markdown("<h4 style='text-align:center; margin-bottom:20px;'>ط¯ط®ظˆظ„ ط§ظ„ط¥ط¯ط§ط±ط© (ظ‚ط±ط§ط،ط© ظˆط§ط³طھط®ط±ط§ط¬ طھظ‚ط§ط±ظٹط±)</h4>", unsafe_allow_html=True)
            u_admin = st.text_input("ط§ط³ظ… ط§ظ„ظ…ط³طھط®ط¯ظ…", placeholder="ط£ط¯ط®ظ„ ط§ط³ظ… ط§ظ„ظ…ط³طھط®ط¯ظ… ظ„ظ„ط¥ط¯ط§ط±ط©...")
            p_admin = st.text_input("ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±", type="password", placeholder="ط£ط¯ط®ظ„ ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±...")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.form_submit_button("ًں‘پï¸ڈ ط¯ط®ظˆظ„ ط§ظ„ط¥ط´ط±ط§ظپ", type="primary", use_container_width=True):
                df_u = fetch_safe("users")
                if not df_u.empty and u_admin in df_u['username'].values:
                    ud = df_u[df_u['username']==u_admin].iloc[0]
                    if hashlib.sha256(p_admin.encode()).hexdigest() == ud['password_hash']:
                        if ud.get('role', '') == 'viewer':
                            st.session_state.username = u_admin
                            st.session_state.role = "viewer"
                            st.rerun()
                        else:
                            st.error("â‌Œ ظ‡ط°ط§ ط§ظ„ط­ط³ط§ط¨ ظ„ط§ ظٹظ…ظ„ظƒ طµظ„ط§ط­ظٹط© ط§ظ„ط¥ط¯ط§ط±ط© (ظ‚ط±ط§ط،ط© ظپظ‚ط·).")
                    else:
                        st.error("â‌Œ ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط± ط؛ظٹط± طµط­ظٹط­ط©.")
                else:
                    st.error("â‌Œ ط¨ظٹط§ظ†ط§طھ ط§ظ„ط¯ط®ظˆظ„ ط؛ظٹط± طµط­ظٹط­ط©.")
                    
    show_footer()

# ==========================================
# ًں‘¨â€چًںڈ« 4. ظˆط§ط¬ظ‡ط© ط§ظ„ظ…ط¹ظ„ظ… / ط§ظ„ط¥ط¯ط§ط±ط© (ظ…ط´ط§ظ‡ط¯)
# ==========================================
else:
    if 'db_loaded' not in st.session_state:
        with st.spinner("âڈ³ ط¬ط§ط±ظٹ ط§ظ„ط§طھطµط§ظ„ ط¨ط§ظ„ظ…ظ†طµط© ظˆطھط­ط¯ظٹط« ط§ظ„ط¨ظٹط§ظ†ط§طھ ط§ظ„ظ…ط±ظƒط²ظٹط©..."):
            try:
                st.session_state.df_students = fetch_safe("students")
                st.session_state.df_grades = fetch_safe("grades")
                st.session_state.df_behavior = fetch_safe("behavior")
                st.session_state.db_loaded = True
            except Exception as e:
                st.error(f"â‌Œ ط­ط¯ط« ط®ط·ط£ ط£ط«ظ†ط§ط، ط§ظ„ط§طھطµط§ظ„: {e}")
                st.stop()
                
    if st.session_state.get('show_refresh_success'):
        st.toast("âœ… طھظ… طھط­ط¯ظٹط« ط§ظ„ط¨ظٹط§ظ†ط§طھ ظˆظ…ط²ط§ظ…ظ†طھظ‡ط§ ط¨ظ†ط¬ط§ط­!", icon="ًں”„")
        st.session_state['show_refresh_success'] = False 

    if st.session_state.role in ["teacher", "viewer"]:
        
        if st.session_state.role == "teacher":
            menu = st.tabs(["ًں‘¥ ط§ظ„ط·ظ„ط§ط¨", "ًں“ٹ ط§ظ„طھظ‚ظٹظٹظ…", "ًں“¢ ط§ظ„طھظ†ط¨ظٹظ‡ط§طھ", "âڑ™ï¸ڈ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ", "ًں›‘ ط®ط±ظˆط¬"])
            tab_students, tab_eval, tab_alerts, tab_settings, tab_logout = menu[0], menu[1], menu[2], menu[3], menu[4]
        else:
            menu = st.tabs(["ًں‘¥ ط§ظ„ط·ظ„ط§ط¨", "ًں“ٹ ط§ظ„طھظ‚ظٹظٹظ…", "ًں“¢ ط§ظ„طھظ†ط¨ظٹظ‡ط§طھ", "ًں›‘ ط®ط±ظˆط¬"])
            tab_students, tab_eval, tab_alerts, tab_logout = menu[0], menu[1], menu[2], menu[3]
            
        # --- ًں‘¥ ط§ظ„ط·ظ„ط§ط¨ ---
        with tab_students:
            st.subheader("ًں‘¥ ط¥ط¯ط§ط±ط© ط§ظ„ط·ظ„ط§ط¨ ظˆط§ظ„طھظ‚ط§ط±ظٹط±")
            df_st = st.session_state.df_students
            
            if not df_st.empty:
                df_st['clean_id'] = df_st.iloc[:,0].astype(str).str.split('.').str[0].str.strip()
                df_st['ط§ظ„ظ†ظ‚ط§ط·'] = pd.to_numeric(df_st['ط§ظ„ظ†ظ‚ط§ط·'], errors='coerce').fillna(0)
                
                sub_tabs = st.tabs(["ًں“‹ ظ‚ط§ط¦ظ…ط© ط§ظ„ط·ظ„ط§ط¨", "ًںڈ† ظ„ظˆط­ط© ط§ظ„ط´ط±ظپ (ظ†ظ‚ط§ط·)", "ًںŒں ط§ظ„ظ…طھظپظˆظ‚ظٹظ† (90%+)", "ًں“‘ طھظ‚ط±ظٹط± ط§ظ„ط·ط§ظ„ط¨ ط§ظ„ط´ط§ظ…ظ„"])
                
                # --- 1. ظ‚ط§ط¦ظ…ط© ط§ظ„ط·ظ„ط§ط¨ ---
                with sub_tabs[0]:
                    if 'toast_msg' in st.session_state:
                        st.toast(st.session_state.toast_msg, icon="ًں””")
                        del st.session_state['toast_msg']
        
                    total_students = len(df_st)
                    total_classes = len(df_st['class'].unique()) if 'class' in df_st.columns else 0
                    avg_points = round(df_st['ط§ظ„ظ†ظ‚ط§ط·'].mean(), 1) if 'ط§ظ„ظ†ظ‚ط§ط·' in df_st.columns else 0
                    
                    cards_css = f"""
                    <style>
                    .metric-container {{ display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px; direction: rtl; }}
                    .metric-card {{
                        background-color: #ffffff; border: 1px solid {border_color}; border-radius: 12px;
                        padding: 20px; flex: 1; display: flex; justify-content: space-between; align-items: center;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
                    }}
                    .metric-info {{ text-align: right; }}
                    .metric-title {{ color: {sub_text}; font-size: 14px; font-weight: bold; margin-bottom: 5px; }}
                    .metric-val {{ color: {text_color}; font-size: 28px; font-weight: 900; }}
                    .metric-sub {{ color: #94A3B8; font-size: 13px; }}
                    .metric-icon {{ width: 55px; height: 55px; border-radius: 12px; display: flex; justify-content: center; align-items: center; font-size: 26px; }}
                    .ic-green {{ background-color: #D1FAE5; color: {success_color}; }}
                    .ic-blue {{ background-color: #DBEAFE; color: {primary_color}; }}
                    .ic-red {{ background-color: #FEE2E2; color: {danger_color}; }}
                    
                    [data-testid="stDataFrame"] {{ direction: rtl; }}
                    </style>
                    """
                    
                    cards_html = f"""
                    {cards_css}
                    <div class="metric-container">
                        <div class="metric-card">
                            <div class="metric-info">
                                <div class="metric-title">ط§ظ„ط¹ط¯ط¯ ط§ظ„ط¥ط¬ظ…ط§ظ„ظٹ</div>
                                <div class="metric-val">{total_students}</div>
                                <div class="metric-sub">ط·ط§ظ„ط¨ ظ…ط³ط¬ظ„</div>
                            </div>
                            <div class="metric-icon ic-green">ًں‘¥</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-info">
                                <div class="metric-title">ط§ظ„ظپطµظˆظ„</div>
                                <div class="metric-val">{total_classes}</div>
                                <div class="metric-sub">ظپطµظˆظ„ ط¯ط±ط§ط³ظٹط©</div>
                            </div>
                            <div class="metric-icon ic-blue">ًںڈ«</div>
                        </div>
                        <div class="metric-card">
                            <div class="metric-info">
                                <div class="metric-title">ظ…طھظˆط³ط· ط§ظ„ظ†ظ‚ط§ط·</div>
                                <div class="metric-val" dir="ltr">{avg_points}</div>
                                <div class="metric-sub">ظ†ظ‚ط·ط© طھظ…ظٹط²</div>
                            </div>
                            <div class="metric-icon ic-red">ًں“ˆ</div>
                        </div>
                    </div>
                    """
                    st.markdown(cards_html, unsafe_allow_html=True)
        
                    action_tabs = st.tabs(["ًں”چ ط¹ط±ط¶ ط§ظ„ط·ظ„ط§ط¨", "â‍• ط¥ط¶ط§ظپط© ط·ط§ظ„ط¨", "âœڈï¸ڈ طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھ ط·ط§ظ„ط¨", "ًں—‘ï¸ڈ ط­ط°ظپ ط·ط§ظ„ط¨"])
        
                    # -------------------------------------
                    # ًں“„ ط¹ط±ط¶ ط§ظ„ط·ظ„ط§ط¨
                    # -------------------------------------
                    with action_tabs[0]:
                        
                        if 'current_page' not in st.session_state:
                            st.session_state.current_page = 1
                        
                        col_search, col_rows = st.columns([3, 1])
                        with col_search:
                            sq = st.text_input("ًں”چ ط¨ط­ط«:", placeholder="ط£ط¯ط®ظ„ ط§ط³ظ… ط§ظ„ط·ط§ظ„ط¨ ط£ظˆ ط±ظ‚ظ…ظ‡...", label_visibility="collapsed")
                        with col_rows:
                            rows_per_page = st.selectbox("ط§ظ„طµظپظˆظپ:", [10, 20, 50, 100], index=0, label_visibility="collapsed")
                        
                        display_df = df_st.copy()
                        if 'clean_id' in display_df.columns:
                            display_df = display_df.drop(columns=['clean_id'])
                        
                        rename_dict = {
                            'id': 'ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ', 'name': 'ط§ظ„ط§ط³ظ…', 'class': 'ط§ظ„طµظپ',
                            'year': 'ط§ظ„ط¹ط§ظ…', 'sem': 'ط§ظ„ظ…ط±ط­ظ„ط©', 'subject': 'ط§ظ„ظ…ط§ط¯ط©'
                        }
                        display_df = display_df.rename(columns=rename_dict)
                        
                        # ظ‚ظ„ط¨ ط§ظ„طھط±طھظٹط¨
                        display_df = display_df[display_df.columns[::-1]]
                        
                        if sq: 
                            norm_sq = normalize_arabic(sq)
                            mask = display_df['ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ'].astype(str).str.contains(norm_sq) | display_df['ط§ظ„ط§ط³ظ…'].astype(str).apply(normalize_arabic).str.contains(norm_sq)
                            display_df = display_df[mask]
                            st.session_state.current_page = 1 
                        
                        total_rows = len(display_df)
                        total_pages = max(1, math.ceil(total_rows / rows_per_page))
                        
                        if st.session_state.current_page > total_pages:
                            st.session_state.current_page = total_pages
                            
                        start_idx = (st.session_state.current_page - 1) * rows_per_page
                        end_idx = start_idx + rows_per_page
                        
                        st.dataframe(display_df.iloc[start_idx:end_idx], use_container_width=True, hide_index=True)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        pc1, pc2, pc3 = st.columns([1, 2, 1])
                        
                        with pc1:
                            if st.button("â–¶ ط§ظ„طھط§ظ„ظٹ", disabled=(st.session_state.current_page >= total_pages), use_container_width=True):
                                st.session_state.current_page += 1
                                st.rerun()
                                
                        with pc2:
                            st.markdown(f"<div style='text-align: center; font-weight: bold; padding-top: 5px; color: #64748b;'>طµظپط­ط© {st.session_state.current_page} ظ…ظ† {total_pages} <br><small>(ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ط·ظ„ط§ط¨: {total_rows})</small></div>", unsafe_allow_html=True)
                            
                        with pc3:
                            if st.button("ط§ظ„ط³ط§ط¨ظ‚ â—€", disabled=(st.session_state.current_page <= 1), use_container_width=True):
                                st.session_state.current_page -= 1
                                st.rerun()
        
                    # -------------------------------------
                    # 2. ط¥ط¶ط§ظپط© ط·ط§ظ„ط¨
                    # -------------------------------------
                    with action_tabs[1]:
                        if st.session_state.role == "teacher":
                            with st.form("add_st_v26", clear_on_submit=True):
                                st.markdown("#### ًں‘¤ ط¨ظٹط§ظ†ط§طھ ط§ظ„ط·ط§ظ„ط¨ ط§ظ„ط£ط³ط§ط³ظٹط©")
                                c1, c2 = st.columns(2)
                                f_name = c1.text_input("ط§ظ„ط¥ط³ظ… ط§ظ„ظƒط§ظ…ظ„", placeholder="ط£ط¯ط®ظ„ ط§ظ„ط§ط³ظ… ط§ظ„ظƒط§ظ…ظ„")
                                f_id = c2.text_input("ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ", placeholder="ط£ط¯ط®ظ„ ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ")
                                
                                c3, c4, c5 = st.columns(3)
                                f_stage = c3.selectbox("ط§ظ„ظ…ط±ط­ظ„ط©", st.session_state.stage_options)
                                f_class = c4.selectbox("ط§ظ„طµظپ", st.session_state.class_options)
                                f_year = c5.text_input("ط§ظ„ط¹ط§ظ…", st.session_state.current_year)
                                
                                st.markdown("#### ًں“‍ ظ…ط¹ظ„ظˆظ…ط§طھ ط§ظ„طھظˆط§طµظ„")
                                c6, c7 = st.columns(2)
                                f_mail = c6.text_input("ط§ظ„ط¥ظٹظ…ظٹظ„", placeholder="ط£ط¯ط®ظ„ ط§ظ„ط¨ط±ظٹط¯ ط§ظ„ط¥ظ„ظƒطھط±ظˆظ†ظٹ")
                                f_phone = c7.text_input("ط§ظ„ط¬ظˆط§ظ„", placeholder="ط£ط¯ط®ظ„ ط±ظ‚ظ… ط§ظ„ط¬ظˆط§ظ„")
                                
                                submit_btn = st.form_submit_button("âœ… ط­ظپط¸ ط§ظ„ط·ط§ظ„ط¨", type="primary", use_container_width=True)
                                
                                if submit_btn:
                                    if f_id and f_name:
                                        if f_id.strip() in df_st['clean_id'].values:
                                            st.error(f"âڑ ï¸ڈ ط§ظ„ط±ظ‚ظ… {f_id} ظ…ط³ط¬ظ„ ظ…ط³ط¨ظ‚ط§ظ‹!")
                                        else:
                                            cl_p = clean_phone_number(f_phone) if f_phone else ""
                                            st_map = {"id": f_id.strip(), "name": f_name.strip(), "class": f_class, "year": f_year, "sem": f_stage, "ط§ظ„ط¬ظˆط§ظ„": cl_p, "ط§ظ„ط¥ظٹظ…ظٹظ„": f_mail.strip(), "ط§ظ„ظ†ظ‚ط§ط·": "0"}
                                            if safe_append_row("students", st_map):
                                                st.session_state.toast_msg = f"âœ… طھظ… ط¥ط¶ط§ظپط© ط§ظ„ط·ط§ظ„ط¨ '{f_name}' ط¨ظ†ط¬ط§ط­!"
                                                if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                                st.cache_data.clear()
                                                st.rerun()
                                    else: st.warning("ط§ظ„ط±ط¬ط§ط، ط¥ظƒظ…ط§ظ„ ط§ظ„ط¨ظٹط§ظ†ط§طھ ط§ظ„ط£ط³ط§ط³ظٹط© (ط§ظ„ط§ط³ظ… ظˆط§ظ„ط±ظ‚ظ…)!")
                        else:
                            st.info("ظ„ظٹط³ ظ„ط¯ظٹظƒ طµظ„ط§ط­ظٹط© ظ„ط¥ط¶ط§ظپط© ط·ظ„ط§ط¨.")
        
                    # -------------------------------------
                    # 3. طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھ ط·ط§ظ„ط¨
                    # -------------------------------------
                    with action_tabs[2]:
                        if st.session_state.role == "teacher":
                            if not df_st.empty:
                                edit_options = df_st.index.tolist()
                                selected_idx = st.selectbox(
                                    "âœڈï¸ڈ ط§ط®طھط± ط§ظ„ط·ط§ظ„ط¨ ط§ظ„ظ…ط·ظ„ظˆط¨ طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھظ‡ (ظٹظ…ظƒظ†ظƒ ط§ظ„ط¨ط­ط« ط¨ط§ظ„ظƒطھط§ط¨ط© ظ‡ظ†ط§):", 
                                    edit_options, 
                                    format_func=lambda x: f"{df_st.loc[x, 'name']} - (ط§ظ„ط±ظ‚ظ…: {df_st.loc[x, 'id']})",
                                    key="edit_select"
                                )
                                
                                with st.form("edit_form_single"):
                                    st.markdown(f"**ًں“‌ طھط¹ط¯ظٹظ„ ط¨ظٹط§ظ†ط§طھ: <span style='color:{primary_color};'>{df_st.loc[selected_idx, 'name']}</span>**", unsafe_allow_html=True)
                                    cols = st.columns(3)
                                    new_vals = []
                                    
                                    valid_columns = [c for c in df_st.columns if c not in ['clean_id'] and not str(c).startswith('Unnamed') and str(c).strip() != ""]
                                    
                                    for col_idx, col_name in enumerate(valid_columns):
                                        with cols[col_idx % 3]:
                                            current_val = "" if pd.isna(df_st.loc[selected_idx, col_name]) else str(df_st.loc[selected_idx, col_name])
                                            val = st.text_input(col_name, current_val, key=f"inp_edit_{selected_idx}_{col_idx}")
                                            new_vals.append(val)
                                    
                                    if st.form_submit_button("ًں’¾ ط­ظپط¸ ط§ظ„طھط¹ط¯ظٹظ„ط§طھ", type="primary", use_container_width=True):
                                        row_index = int(selected_idx) + 2
                                        try:
                                            sh.worksheet("students").update(f"A{row_index}", [new_vals])
                                            st.session_state.toast_msg = f"ًں”„ طھظ… طھط­ط¯ظٹط« ط¨ظٹط§ظ†ط§طھ '{new_vals[1]}' ط¨ظ†ط¬ط§ط­!"
                                            if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                            st.cache_data.clear()
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"â‌Œ ط­ط¯ط« ط®ط·ط£ ط£ط«ظ†ط§ط، ط§ظ„طھط¹ط¯ظٹظ„: {e}")
                            else:
                                st.info("ظ„ط§ طھظˆط¬ط¯ ط¨ظٹط§ظ†ط§طھ ظ„ظ„ط·ظ„ط§ط¨ ط¨ط¹ط¯.")
        
                    # -------------------------------------
                    # 4. ط­ط°ظپ ط·ط§ظ„ط¨
                    # -------------------------------------
                    with action_tabs[3]:
                        if st.session_state.role == "teacher":
                            if not df_st.empty:
                                del_options = df_st.index.tolist()
                                del_idx = st.selectbox(
                                    "ًں—‘ï¸ڈ ط§ط®طھط± ط§ظ„ط·ط§ظ„ط¨ ط§ظ„ظ…ط·ظ„ظˆط¨ ط­ط°ظپظ‡ ظ†ظ‡ط§ط¦ظٹط§ظ‹ (ظٹظ…ظƒظ†ظƒ ط§ظ„ط¨ط­ط« ط¨ط§ظ„ظƒطھط§ط¨ط© ظ‡ظ†ط§):", 
                                    del_options, 
                                    format_func=lambda x: f"{df_st.loc[x, 'name']} - (ط§ظ„ط±ظ‚ظ…: {df_st.loc[x, 'id']})",
                                    key="del_select"
                                )
                                
                                student_to_delete = df_st.loc[del_idx, 'name']
                                st.error(f"âڑ ï¸ڈ ط³ظٹطھظ… ط­ط°ظپ ط¨ظٹط§ظ†ط§طھ ( {student_to_delete} ) ظ†ظ‡ط§ط¦ظٹط§ظ‹ ظ…ظ† ظ‚ط§ط¹ط¯ط© ط§ظ„ط¨ظٹط§ظ†ط§طھ. ظ‡ظ„ ط£ظ†طھ ظ…طھط£ظƒط¯طں")
                                
                                if st.button(f"ًںڑ¨ ظ†ط¹ظ…طŒ ط§ط­ط°ظپ {student_to_delete}", type="primary"):
                                    sh.worksheet("students").delete_rows(int(del_idx)+2)
                                    st.session_state.toast_msg = f"ًں—‘ï¸ڈ طھظ… ط­ط°ظپ '{student_to_delete}' ط¨ط´ظƒظ„ ظ†ظ‡ط§ط¦ظٹ!"
                                    if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                    st.cache_data.clear()
                                    st.rerun()
                            else:
                                st.info("ظ„ط§ طھظˆط¬ط¯ ط¨ظٹط§ظ†ط§طھ ظ„ظ„ط·ظ„ط§ط¨ ط¨ط¹ط¯.")
                
                # --- 2. ظ„ظˆط­ط© ط§ظ„ط´ط±ظپ (ط§ظ„ظ†ظ‚ط§ط· ظˆط§ظ„ط³ظ„ظˆظƒ) ---
                with sub_tabs[1]:
                    st.markdown("#### ًںŒں ط£ظپط¶ظ„ 10 ط·ظ„ط§ط¨ (ط­ط³ط¨ ظ†ظ‚ط§ط· ط§ظ„طھظ…ظٹط²)")
                    
                    lux_css = """
                        * { box-sizing: border-box; } 
                        body { margin: 0; padding: 0; background: #F8FAFC; font-family: 'Cairo', sans-serif; text-align: center; direction: rtl; }
                        .page { width: 210mm; padding: 10mm; display: flex; flex-wrap: wrap; justify-content: center; gap: 4%; margin: 0 auto; }
                        
                        .card { 
                            width: 46%; height: 135mm; 
                            border-radius: 15px; position: relative; overflow: hidden;
                            background: #fff;
                            box-shadow: 0 15px 35px rgba(0,0,0,0.15); 
                            page-break-inside: avoid;
                            margin-bottom: 20px;
                        }
                        
                        .ribbon {
                            position: absolute; top: 20px; right: -35px;
                            padding: 5px 40px; font-weight: 900; font-size: 15px;
                            transform: rotate(45deg); z-index: 10;
                            box-shadow: 0 4px 6px rgba(0,0,0,0.2);
                            font-family: 'Cairo', sans-serif; letter-spacing: 1px;
                        }
                        .ribbon.gold { background: linear-gradient(45deg, #FFD700, #FFA500); color: #8B4513; }
                        .ribbon.silver { background: linear-gradient(45deg, #E2E8F0, #94A3B8); color: #1E293B; }
                        .ribbon.bronze { background: linear-gradient(45deg, #FDBA74, #D97706); color: #78350F; }
                        
                        .card-inner {
                            position: absolute; top: 12px; bottom: 12px; left: 12px; right: 12px;
                            border: 2px dashed #b68a36; border-radius: 10px; padding: 20px 10px;
                            display: flex; flex-direction: column; justify-content: space-between; align-items: center;
                            background-color: rgba(255,255,255,0.95);
                            background-image: url("data:image/svg+xml,%3Csvg width='20' height='20' viewBox='0 0 20 20' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='%23b68a36' fill-opacity='0.05' fill-rule='evenodd'%3E%3Ccircle cx='3' cy='3' r='3'/%3E%3Ccircle cx='13' cy='13' r='3'/%3E%3C/g%3E%3C/svg%3E");
                        }
                        
                        .c-icon { font-size: 45px; line-height: 1; margin-bottom: 5px; width: 100%; text-align: center; filter: drop-shadow(0 4px 4px rgba(0,0,0,0.1)); }
                        .c-header { font-family: 'Aref Ruqaa', serif; font-size: 30px; font-weight: bold; width: 100%; text-align: center; margin-bottom: 5px; color: #b68a36; }
                        .c-teacher { font-size: 14px; color: #475569; font-weight: bold; width: 100%; text-align: center; margin-bottom: 10px; }
                        
                        .c-name { 
                            font-size: 26px; font-weight: 900; line-height: 1.4; 
                            width: 100%; text-align: center; 
                            display: flex; align-items: center; justify-content: center; flex-grow: 1;
                            padding: 0 10px; word-wrap: break-word;
                        }
                        
                        .c-badge { 
                            width: 80%; margin: 10px auto; padding: 15px 5px; border-radius: 12px;
                            background: #fff; text-align: center;
                            box-shadow: 0 4px 10px rgba(0,0,0,0.08); border: 1px solid #E2E8F0;
                        }
                        
                        .c-badge.rank-1 { background: linear-gradient(145deg, #FFF8DC, #FFD700); border-color: #DAA520; }
                        .c-badge.rank-2 { background: linear-gradient(145deg, #F8F9FA, #E2E8F0); border-color: #94A3B8; }
                        .c-badge.rank-3 { background: linear-gradient(145deg, #FFF1F2, #FECDD3); border-color: #CD7F32; }
                        
                        .b-val { display: block; font-size: 30px; font-weight: 900; line-height: 1; margin-bottom: 5px; }
                        .b-lbl { display: block; font-size: 14px; font-weight: bold; color: #64748B; }
                        
                        .rank-1 .b-val { color: #8B4513; font-size: 36px; } .rank-1 .b-lbl { color: #A0522D; }
                        .rank-2 .b-val { color: #1E293B; font-size: 34px; } .rank-2 .b-lbl { color: #334155; }
                        .rank-3 .b-val { color: #78350F; font-size: 34px; } .rank-3 .b-lbl { color: #92400E; }
                        
                        .c-footer { font-family: 'Amiri', serif; font-size: 18px; font-weight: bold; width: 100%; text-align: center; padding-top: 10px; border-top: 1px solid rgba(0,0,0,0.1); }
                        
                        .theme-honor { border: 14px solid #1E40AF; } 
                        .theme-honor .c-name, .theme-honor .b-val:not(.rank-1 .b-val):not(.rank-2 .b-val):not(.rank-3 .b-val), .theme-honor .c-footer { color: #1E40AF; }
                        
                        .theme-academic { border: 14px solid #881337; } 
                        .theme-academic .c-name, .theme-academic .b-val, .theme-academic .c-footer { color: #881337; }
                        
                        @media print { 
                            @page { size: A4 portrait; margin: 5mm; } 
                            body { background: white; -webkit-print-color-adjust: exact; print-color-adjust: exact; } 
                            .card { box-shadow: none; margin-bottom: 10mm; border-width: 10px; }
                        }
                    """
        
                    if not df_st.empty:
                        top_10 = df_st.sort_values('ط§ظ„ظ†ظ‚ط§ط·', ascending=False).head(10)
                        
                        for i, (_, r) in enumerate(top_10.iterrows(), 1):
                            ic = "ًں¥‡" if i==1 else "ًں¥ˆ" if i==2 else "ًں¥‰" if i==3 else f"#{i}"
                            brd_col = "#F59E0B" if i<=3 else "#E2E8F0"
                            st.markdown(f"""
                                <div style='background:#ffffff; border:1px solid #E2E8F0; border-right:5px solid {brd_col}; padding:15px; border-radius:12px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;'>
                                    <div style='display:flex; align-items:center; gap:15px;'>
                                        <span style='font-size:1.5rem; font-weight:bold; width:30px; text-align:center;'>{ic}</span>
                                        <div>
                                            <b style='font-size:1.1rem; color:{accent_color};'>{r.get('name', '')}</b><br>
                                            <small style='color:#64748B;'>ًںڈ« ط§ظ„طµظپ: {r.get('class', '')} | ًں†” ID: {r.get('clean_id', '')}</small>
                                        </div>
                                    </div>
                                    <div style='background:#FEF3C7; padding:5px 15px; border-radius:8px; color:#B45309; font-weight:900; font-size:1.2rem;'>
                                        {int(r.get('ط§ظ„ظ†ظ‚ط§ط·', 0))} ظ†ظ‚ط·ط©
                                    </div>
                                </div>
                            """, unsafe_allow_html=True)
                        
                        st.markdown("---")
                        st.subheader("ًں–¨ï¸ڈ ط·ط¨ط§ط¹ط© ط¨ط·ط§ظ‚ط§طھ ظ„ظˆط­ط© ط§ظ„ط´ط±ظپ")
                        
                        honor_cards_content = ""
                        for rank, (_, row) in enumerate(top_10.iterrows(), 1):
                            student_name = row.get('name', 'ط§ط³ظ… ط؛ظٹط± ظ…طھظˆظپط±')
                            score = int(row.get('ط§ظ„ظ†ظ‚ط§ط·', 0))
                            
                            if rank == 1:
                                rank_text = "ط§ظ„ظ…ط±ظƒط² ط§ظ„ط£ظˆظ„"
                                icon = "ًںڈ†"
                                ribbon_html = '<div class="ribbon gold">ط§ظ„ط£ظˆظ„</div>'
                                badge_class = "rank-1"
                            elif rank == 2:
                                rank_text = "ط§ظ„ظ…ط±ظƒط² ط§ظ„ط«ط§ظ†ظٹ"
                                icon = "ًں¥ˆ"
                                ribbon_html = '<div class="ribbon silver">ط§ظ„ط«ط§ظ†ظٹ</div>'
                                badge_class = "rank-2"
                            elif rank == 3:
                                rank_text = "ط§ظ„ظ…ط±ظƒط² ط§ظ„ط«ط§ظ„ط«"
                                icon = "ًں¥‰"
                                ribbon_html = '<div class="ribbon bronze">ط§ظ„ط«ط§ظ„ط«</div>'
                                badge_class = "rank-3"
                            else:
                                rank_text = f"ط§ظ„ظ…ط±ظƒط² {rank}"
                                icon = "ًںŒں"
                                ribbon_html = ""
                                badge_class = ""
                            
                            honor_cards_content += f"""
                            <div class="card theme-honor">
                                {ribbon_html}
                                <div class="card-inner">
                                    <div class="c-icon">{icon}</div>
                                    <div class="c-header">ط¨ط·ط§ظ‚ط© طھظ…ظٹط² ط·ط§ظ„ط¨</div>
                                    <div class="c-teacher">ط¥ط´ط±ط§ظپ ط§ظ„ط£ط³طھط§ط°/ طµط§ظ„ط­ ط§ظ„ط±ظˆظٹط«ظٹ</div>
                                    <div class="c-name">{student_name}</div>
                                    <div class="c-badge {badge_class}">
                                        <span class="b-val">{score}</span>
                                        <span class="b-lbl">ظ†ظ‚ط·ط© طھظ…ظٹط²</span>
                                    </div>
                                    <div class="c-footer">{rank_text}</div>
                                </div>
                            </div>
                            """
                        
                        honor_full_html = f"""<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Aref+Ruqaa:wght@400;700&family=Amiri:wght@400;700&family=Cairo:wght@400;700;900&display=swap" rel="stylesheet"><style>{lux_css}</style></head><body><div class="page">{honor_cards_content}</div><script>window.onload = function() {{ window.print(); }}</script></body></html>"""
                        
                        st.download_button(
                            label="ًںŒگ طھط­ظ…ظٹظ„ ط¨ط·ط§ظ‚ط§طھ ط§ظ„ط´ط±ظپ (طھطµظ…ظٹظ… ظˆط§ظ‚ط¹ظٹ ظ„ظ„ط·ط¨ط§ط¹ط©)", 
                            data=honor_full_html, 
                            file_name=f"Honor_Cards_{datetime.date.today()}.html", 
                            mime="text/html", 
                            use_container_width=True,
                            type="primary"
                        )
        
                # --- 3. ط§ظ„ظ…طھظپظˆظ‚ظٹظ† (ط£ظƒط§ط¯ظٹظ…ظٹط§ظ‹ 90% ظپظ…ط§ ظپظˆظ‚) ---
                with sub_tabs[2]:
                    st.markdown("#### ًںژ“ ظ„ظˆط­ط© ط§ظ„ظ…طھظپظˆظ‚ظٹظ† ط£ظƒط§ط¯ظٹظ…ظٹط§ظ‹")
                    if 'df_grades' in st.session_state and not st.session_state.df_grades.empty and not df_st.empty:
                        df_g = st.session_state.df_grades.copy()
                        df_g['clean_id'] = df_g.iloc[:,0].astype(str).str.split('.').str[0]
                        merged_df = pd.merge(df_g, df_st[['clean_id', 'name', 'class']], on='clean_id', how='inner')
                        
                        if not merged_df.empty:
                            max_total = st.session_state.get('max_tasks', 0) + st.session_state.get('max_quiz', 0)
                            if max_total > 0:
                                merged_df['perf_num'] = pd.to_numeric(merged_df['perf'], errors='coerce').fillna(0)
                                merged_df['percentage'] = (merged_df['perf_num'] / max_total) * 100
                                top_academic = merged_df[merged_df['percentage'] >= 90].sort_values('percentage', ascending=False)
                                
                                if not top_academic.empty:
                                    for i, (_, r) in enumerate(top_academic.iterrows(), 1):
                                        st.markdown(f"<div style='background:#ffffff; border:1px solid #E2E8F0; border-right:5px solid {success_color}; padding:15px; border-radius:12px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center;'><div style='display:flex; align-items:center; gap:15px;'><span style='font-size:1.5rem;'>ًںژ“</span><div><b style='font-size:1.1rem; color:{success_color};'>{r.get('name', '')}</b><br><small>ًںڈ« {r.get('class', '')}</small></div></div><div style='background:#D1FAE5; padding:5px 15px; border-radius:8px; color:{success_color}; font-weight:900;'>ظ…ظ…طھط§ط²</div></div>", unsafe_allow_html=True)
                                        
                                    st.markdown("---")
                                    st.subheader("ًں–¨ï¸ڈ ط·ط¨ط§ط¹ط© ط¨ط·ط§ظ‚ط§طھ ط§ظ„طھظپظˆظ‚")
                                    
                                    academic_cards_content = ""
                                    for _, row in top_academic.iterrows():
                                        academic_cards_content += f"""
                                        <div class="card theme-academic">
                                            <div class="ribbon gold">ظ…طھظپظˆظ‚</div>
                                            <div class="card-inner">
                                                <div class="c-icon">ًںژ–ï¸ڈ</div>
                                                <div class="c-header">ظˆط³ط§ظ… ط§ظ„طھظ…ظٹط² ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ</div>
                                                <div class="c-teacher">ط¥ط´ط±ط§ظپ ط§ظ„ط£ط³طھط§ط°/ طµط§ظ„ط­ ط§ظ„ط±ظˆظٹط«ظٹ</div>
                                                <div class="c-name">{row.get('name', 'ط·ط§ظ„ط¨')}</div>
                                                <div class="c-badge rank-1">
                                                    <span class="b-val">ظ…ظ…طھط§ط²</span>
                                                    <span class="b-lbl">ظپظٹ ظ…ط§ط¯ط© ط§ظ„ظ„ط؛ط© ط§ظ„ط¥ظ†ط¬ظ„ظٹط²ظٹط©</span>
                                                </div>
                                                <div class="c-footer">ظ…ط¹ طھظ…ظ†ظٹط§طھظ†ط§ ط¨ط¯ظˆط§ظ… ط§ظ„طھط£ظ„ظ‚ ظˆط§ظ„ظ†ط¬ط§ط­</div>
                                            </div>
                                        </div>
                                        """
        
                                    academic_full_html = f"""<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Aref+Ruqaa:wght@400;700&family=Amiri:wght@400;700&family=Cairo:wght@400;700;900&display=swap" rel="stylesheet"><style>{lux_css}</style></head><body><div class="page">{academic_cards_content}</div><script>window.onload = function() {{ window.print(); }}</script></body></html>"""
                                    
                                    st.download_button(
                                        label="ًںŒگ طھط­ظ…ظٹظ„ ط¨ط·ط§ظ‚ط§طھ ط§ظ„ظ…طھظپظˆظ‚ظٹظ† (طھطµظ…ظٹظ… ظˆط§ظ‚ط¹ظٹ ظ„ظ„ط·ط¨ط§ط¹ط©)", 
                                        data=academic_full_html, 
                                        file_name=f"Excellence_Cards_{datetime.date.today()}.html", 
                                        mime="text/html", 
                                        use_container_width=True,
                                        type="primary"
                                    )
                                else:
                                    st.info("ظ„ظ… ظٹطµظ„ ط£ط­ط¯ ظ„ظ†ط³ط¨ط© 90% ط­طھظ‰ ط§ظ„ط¢ظ†.")
                        else:
                            st.info("ظ„ط§ طھظˆط¬ط¯ ط¯ط±ط¬ط§طھ ظ…ط·ط§ط¨ظ‚ط© ظ„ظ„ط·ظ„ط§ط¨.")
                    else:
                        st.info("ظ„ظ… ظٹطھظ… ط±طµط¯ ط¯ط±ط¬ط§طھ ط¨ط¹ط¯.")
                        
                # --- 4. طھظ‚ط±ظٹط± ط§ظ„ط·ط§ظ„ط¨ ط§ظ„ط´ط§ظ…ظ„ ---
                with sub_tabs[3]:
                    st.markdown("#### ًں“‘ ط§ظ„طھظ‚ط±ظٹط± ط§ظ„ط´ط§ظ…ظ„ ط§ظ„ظ…ظپطµظ„")
                    st_dict = {f"{r['name']} ({r['clean_id']})": r['clean_id'] for _, r in df_st.iterrows()}
                    sel_rep = st.selectbox("ًں”چ ط§ط¨ط­ط« ط¹ظ† ط§ظ„ط·ط§ظ„ط¨ ظ„ط§ط³طھط®ط±ط§ط¬ ط§ظ„طھظ‚ط±ظٹط±:", [""] + list(st_dict.keys()), key="rep_sel")
                    
                    if sel_rep:
                        sid = st_dict[sel_rep]
                        s_inf = df_st[df_st['clean_id'] == sid].iloc[0]
                        st.markdown("---")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.info(f"ًں‘¤ ط§ظ„ط§ط³ظ…:\n\n**{s_inf['name']}**")
                        c2.success(f"ًں†” ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ:\n\n**{sid}**")
                        c3.warning(f"ًںڈ« ط§ظ„طµظپ:\n\n**{s_inf.get('class', 'ط؛ظٹط± ظ…ط­ط¯ط¯')}**")
                        c4.error(f"ًںŒں ط¥ط¬ظ…ط§ظ„ظٹ ط§ظ„ظ†ظ‚ط§ط·:\n\n**{int(s_inf['ط§ظ„ظ†ظ‚ط§ط·'])}**")
                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        grades_html_table = "<div style='text-align:center; padding:20px; color:#64748B;'>ظ„ط§ طھظˆط¬ط¯ ط¯ط±ط¬ط§طھ ظ…ط±طµظˆط¯ط© ظ„ظ‡ط°ط§ ط§ظ„ط·ط§ظ„ط¨.</div>"
                        behavior_html_table = "<div style='text-align:center; padding:20px; color:#64748B;'>âœ¨ ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ظ†ط¸ظٹظپ.</div>"

                        st.markdown("##### ًں“ٹ ط§ظ„ط¯ط±ط¬ط§طھ ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹط©")
                        df_g = st.session_state.df_grades
                        if not df_g.empty:
                            df_g['clean_id'] = df_g.iloc[:,0].astype(str).str.split('.').str[0]
                            my_g = df_g[df_g['clean_id'] == sid]
                            if not my_g.empty:
                                g_inf = my_g.iloc[0]
                                k1, k2, k3 = st.columns(3)
                                k1.metric("ًں“‌ ط§ظ„ظ…ط´ط§ط±ظƒط© ظˆط§ظ„ظˆط§ط¬ط¨ط§طھ", g_inf.get('p1', 0))
                                k2.metric("âœچï¸ڈ ط§ظ„ط§ط®طھط¨ط§ط±ط§طھ", g_inf.get('p2', 0))
                                k3.metric("ًںڈ† ط§ظ„ظ…ط¬ظ…ظˆط¹ ط§ظ„ظƒظ„ظٹ", g_inf.get('perf', 0))
                                
                                grades_html_table = f"""
                                <table>
                                    <tr><th>ط§ظ„ظ…ط´ط§ط±ظƒط© ظˆط§ظ„ظˆط§ط¬ط¨ط§طھ</th><th>ط§ظ„ط§ط®طھط¨ط§ط±ط§طھ</th><th>ط§ظ„ظ…ط¬ظ…ظˆط¹ ط§ظ„ظƒظ„ظٹ</th></tr>
                                    <tr>
                                        <td style="text-align: center;">{g_inf.get('p1', 0)}</td>
                                        <td style="text-align: center;">{g_inf.get('p2', 0)}</td>
                                        <td style="text-align: center; font-weight:bold; color:{primary_color};">{g_inf.get('perf', 0)}</td>
                                    </tr>
                                </table>
                                """
                            else: st.info("ظ„ظ… ظٹطھظ… ط±طµط¯ ط¯ط±ط¬ط§طھ ط£ظƒط§ط¯ظٹظ…ظٹط© ظ„ظ‡ط°ط§ ط§ظ„ط·ط§ظ„ط¨ ط¨ط¹ط¯.")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        st.markdown("##### ًں“œ ط³ط¬ظ„ ط§ظ„ظ…ظ„ط§ط­ط¸ط§طھ ظˆط§ظ„ط³ظ„ظˆظƒ ط§ظ„طھظپطµظٹظ„ظٹ")
                        df_b = st.session_state.df_behavior
                        if not df_b.empty:
                            df_b['clean_id'] = df_b.iloc[:,0].astype(str).str.split('.').str[0]
                            my_b = df_b[df_b['clean_id'] == sid]
                            if not my_b.empty:
                                display_df = my_b[['date', 'type', 'note']].rename(columns={'date':'ًں“… ط§ظ„طھط§ط±ظٹط®', 'type':'ًںژ¯ ظ†ظˆط¹ ط§ظ„ط³ظ„ظˆظƒ', 'note':'ًں“‌ ط§ظ„طھظپط§طµظٹظ„'})
                                st.dataframe(display_df, use_container_width=True, hide_index=True)
                                rows_html = ""
                                for _, r_b in display_df.iterrows():
                                    rows_html += f"<tr><td>{r_b['ًں“… ط§ظ„طھط§ط±ظٹط®']}</td><td>{r_b['ًںژ¯ ظ†ظˆط¹ ط§ظ„ط³ظ„ظˆظƒ']}</td><td>{r_b['ًں“‌ ط§ظ„طھظپط§طµظٹظ„']}</td></tr>"
                                behavior_html_table = f"<table><tr><th>ط§ظ„طھط§ط±ظٹط®</th><th>ظ†ظˆط¹ ط§ظ„ط³ظ„ظˆظƒ</th><th>ط§ظ„طھظپط§طµظٹظ„</th></tr>{rows_html}</table>"
                            else: st.success("âœ¨ ط³ط¬ظ„ظ‡ ظ†ط¸ظٹظپطŒ ظ„ط§ طھظˆط¬ط¯ ظ…ظ„ط§ط­ط¸ط§طھ ظ…ط³ط¬ظ„ط© ظپظٹ ط§ظ„ط³ط¬ظ„.")
                        else: st.info("ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ظپط§ط±ط؛ طھظ…ط§ظ…ط§ظ‹ ظپظٹ ط§ظ„ظ…ظ†طµط©.")

                        st.divider()
                        final_report = f"""
                        <!DOCTYPE html>
                        <html dir="rtl" lang="ar">
                        <head>
                            <meta charset="UTF-8">
                            <title>طھظ‚ط±ظٹط± ط§ظ„ط·ط§ظ„ط¨: {s_inf['name']}</title>
                            <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;800&display=swap" rel="stylesheet">
                            <style>
                                body {{ font-family: 'Cairo', sans-serif; background-color: #F8FAFC; padding: 20px; color: #0F172A; line-height: 1.6; }}
                                .container {{ max-width: 800px; margin: 0 auto; background: #FFFFFF; padding: 40px; border-radius: 16px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); }}
                                .banner {{ background: {header_grad}; color: white; text-align: center; padding: 15px; border-radius: 12px; margin-bottom: 30px; font-weight: 800; font-size: 24px; letter-spacing: 1px; box-shadow: 0 4px 15px rgba(37, 99, 235, 0.3); }}
                                .header {{ text-align: center; margin-bottom: 30px; }}
                                .header h1 {{ color: #0F172A; margin-bottom: 5px; font-weight: 800; font-size: 28px; }}
                                .header p {{ color: #64748B; font-size: 14px; margin-top: 0; }}
                                .student-card {{ background: #F8FAFC; border-right: 5px solid {primary_color}; padding: 25px; border-radius: 12px; margin-bottom: 40px; display: grid; grid-template-columns: 1fr 1fr; gap: 15px; box-shadow: 0 2px 10px rgba(0,0,0,0.03); }}
                                .student-card h2 {{ grid-column: 1 / -1; margin-top: 0; color: {accent_color}; border-bottom: 1px dashed {border_color}; padding-bottom: 15px; margin-bottom: 10px; }}
                                .student-card .info-item {{ font-size: 16px; }}
                                .student-card .info-item span {{ font-weight: 800; color: #475569; margin-left: 5px; }}
                                h3 {{ color: {accent_color}; display: flex; align-items: center; gap: 10px; margin-top: 40px; border-bottom: 2px solid {border_color}; padding-bottom: 10px; }}
                                .table-container {{ overflow: hidden; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); margin-bottom: 20px; border: 1px solid {border_color}; }}
                                table {{ width: 100%; border-collapse: collapse; background: #fff; text-align: right; }}
                                th {{ background-color: #F8FAFC; color: #0F172A; font-weight: 800; padding: 15px; border-bottom: 2px solid {border_color}; }}
                                td {{ padding: 15px; border-bottom: 1px solid #F1F5F9; color: #475569; font-weight: 600; }}
                                tr:last-child td {{ border-bottom: none; }}
                                tr:nth-child(even) {{ background-color: #F8FAFC; }}
                                .footer-sigs {{ margin-top: 60px; display: flex; justify-content: space-between; align-items: center; padding-top: 30px; border-top: 2px dashed #CBD5E1; color: #0F172A; font-weight: 800; }}
                                .footer-sigs > div {{ text-align: center; flex: 1; }}
                                .sig-line {{ margin-top: 30px; color: #94A3B8; font-weight: normal; }}
                                @media print {{ body {{ background: white; padding: 0; }} .container {{ box-shadow: none; padding: 0; max-width: 100%; border: none; }} .banner, th, .student-card {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }} }}
                            </style>
                        </head>
                        <body>
                            <div class="container">
                                <div class="banner">âœ¨ ظ…ظ†طµط© طµط§ظ„ط­ ط§ظ„ط°ظƒظٹط© âœ¨</div>
                                <div class="header"><h1>طھظ‚ط±ظٹط± ظ…طھط§ط¨ط¹ط© ط·ط§ظ„ط¨</h1><p>طھط§ط±ظٹط® ط§ط³طھط®ط±ط§ط¬ ط§ظ„طھظ‚ط±ظٹط±: {pd.Timestamp.now().strftime('%Y-%m-%d')}</p></div>
                                <div class="student-card">
                                    <h2>ًں‘¤ {s_inf['name']}</h2>
                                    <div class="info-item"><span>ًں†” ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ:</span> {sid}</div>
                                    <div class="info-item"><span>ًںڈ« ط§ظ„طµظپ:</span> {s_inf.get('class', 'ط؛ظٹط± ظ…ط­ط¯ط¯')}</div>
                                    <div class="info-item"><span>â­گ ظ†ظ‚ط§ط· ط§ظ„طھظ…ظٹط²:</span> <span style="color:#d97706; font-size:1.2em;">{int(s_inf['ط§ظ„ظ†ظ‚ط§ط·'])}</span></div>
                                </div>
                                <h3>ًں“ٹ ط§ظ„ط£ط¯ط§ط، ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ</h3>
                                <div class="table-container">{grades_html_table}</div>
                                <h3>ًں“œ ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ظˆط§ظ„ظ…ظ„ط§ط­ط¸ط§طھ</h3>
                                <div class="table-container">{behavior_html_table}</div>
                                <div class="footer-sigs">
                                    <div>ظˆظƒظٹظ„ ط´ط¤ظˆظ† ط§ظ„ط·ظ„ط§ط¨<div class="sig-line">.................................</div></div>
                                    <div>ط§ظ„ظ…ط¹ظ„ظ…<div style="margin-top: 20px; color: {accent_color}; font-size: 18px;">طµط§ظ„ط­ ط§ظ„ط±ظˆظٹط«ظٹ</div></div>
                                    <div>ط§ظ„ظ…ظˆط¬ظ‡ ط§ظ„ط·ظ„ط§ط¨ظٹ<div class="sig-line">.................................</div></div>
                                </div>
                            </div>
                        </body>
                        </html>
                        """
                        col_btn1, col_btn2 = st.columns([1, 4])
                        with col_btn1:
                            st.download_button(label="ًں–¨ï¸ڈ طھط­ظ…ظٹظ„ ط§ظ„طھظ‚ط±ظٹط±", data=final_report, file_name=f"Report_{sid}_{s_inf['name']}.html", mime="text/html", type="primary")
                        with col_btn2: st.caption("ًں‘ˆ ط§ظ„طھطµظ…ظٹظ… ط¬ط§ظ‡ط²! ط­ظ…ظ„ ط§ظ„ظ…ظ„ظپ ظˆط§ط¶ط؛ط· Ctrl+P ظ„ظ„ط·ط¨ط§ط¹ط©.")
            
        # ًں“ٹ ط§ظ„طھظ‚ظٹظٹظ… ظˆط§ظ„ظ…طھط§ط¨ط¹ط© (ظپط±ط¯ظٹ ظˆط¬ظ…ط§ط¹ظٹ)
        with tab_eval:
            st.markdown("### ًں“ٹ ط§ظ„طھظ‚ظٹظٹظ… ظˆط§ظ„ظ…طھط§ط¨ط¹ط©")
            eval_tabs = st.tabs(["ًں‘¤ ط§ظ„طھظ‚ظٹظٹظ… ط§ظ„ظپط±ط¯ظٹ", "ًں‘¥ ط§ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ ط§ظ„ط³ط±ظٹط¹"])
            
            # --- 1. ط§ظ„طھظ‚ظٹظٹظ… ط§ظ„ظپط±ط¯ظٹ ---
            with eval_tabs[0]:
                df_ev = st.session_state.df_students
                
                if not df_ev.empty:
                    st_dict = {f"{r.iloc[1]} ({r.iloc[0]})": r.iloc[0] for _, r in df_ev.iterrows()}
                    sel = st.selectbox("ًںژ¯ ط§ط®طھط± ط§ظ„ط·ط§ظ„ط¨ ظ…ظ† ط§ظ„ظ‚ط§ط¦ظ…ط©:", [""] + list(st_dict.keys()), key="single_eval_sel")
                    
                    if sel:
                        sid = str(st_dict[sel]).strip().split('.')[0]
                        student_idx = df_ev[df_ev.iloc[:,0].astype(str).str.split('.').str[0] == sid].index[0]
                        s_inf = df_ev.loc[student_idx]
                        
                        s_nm = s_inf['name']
                        clp = clean_phone_number(s_inf.get('ط§ظ„ط¬ظˆط§ظ„',''))
                        s_eml = s_inf.get('ط§ظ„ط¥ظٹظ…ظٹظ„', '')
                        current_points = int(pd.to_numeric(s_inf.get('ط§ظ„ظ†ظ‚ط§ط·', 0), errors='coerce') or 0)
                        
                        c1, c2 = st.columns(2)
                        
                        # --- ظ‚ط³ظ… ط§ظ„ط³ظ„ظˆظƒ ط§ظ„ظپط±ط¯ظٹ ---
                        with c2:
                            st.container(border=True)
                            st.markdown(f"##### ًںژ­ ط§ظ„ط³ظ„ظˆظƒ (ط§ظ„ط±طµظٹط¯: {current_points} ظ†ظ‚ط·ط©)")
                            if st.session_state.role == "teacher":
                                with st.form("beh_add", clear_on_submit=True):
                                    bt = st.selectbox("ظ†ظˆط¹ ط§ظ„ط³ظ„ظˆظƒ", [
                                        "ًںŒں ظ…طھظ…ظٹط² (+10)", "âœ… ط¥ظٹط¬ط§ط¨ظٹ (+5)", "ًں“‌ ط­ظ„ ط§ظ„ظˆط§ط¬ط¨ (+5)", "ًںژ¯ ط£ط¯ط§ط، ط§ظ„ظ…ظ‡ظ…ط© (+10)", "ًں“‚ ظ…ظ„ظپ ط§ظ„ط¥ظ†ط¬ط§ط² (+10)", 
                                        "âڑ ï¸ڈ طھظ†ط¨ظٹظ‡ (0)", "ًں“ڑ ظ†ظ‚طµ ظƒطھط§ط¨ (-5)", "âœچï¸ڈ ظ†ظ‚طµ ظˆط§ط¬ط¨ (-5)", "ًں–ٹï¸ڈ ظ†ظ‚طµ ط£ط¯ظˆط§طھ ط§ظ„ظƒطھط§ط¨ط© (-5)", "ًں’¤ ط§ظ„ظ†ظˆظ… ط¯ط§ط®ظ„ ط§ظ„ظپطµظ„ (-3)", 
                                        "ًںڈƒ طھط£ط®ط± ط¹ظ† ط§ظ„ط­طµط© (-5)", "â‌Œ ط¹ط¯ظ… ط¥ط­ط¶ط§ط± ظ…ظ„ظپ ط§ظ„ط¥ظ†ط¬ط§ط² (-10)", "ًںڑ« ط³ظ„ط¨ظٹ (-10)"
                                    ])
                                    bn = st.text_area("طھظپط§طµظٹظ„ ط§ظ„ظ…ظ„ط§ط­ط¸ط©")
                                    
                                    if st.form_submit_button("ًں’¾ طھط³ط¬ظٹظ„ ط§ظ„ط³ظ„ظˆظƒ", type="primary"):
                                        new_b_row = {"student_id": sid, "date": str(datetime.date.today()), "type": bt, "note": bn}
                                        
                                        safe_append_row("behavior", new_b_row)
                                        
                                        new_b_df = pd.DataFrame([new_b_row])
                                        st.session_state.df_behavior = pd.concat([st.session_state.df_behavior, new_b_df], ignore_index=True)
                                        
                                        match = re.search(r'\(([\+\-]?\d+)\)', bt)
                                        chg = int(match.group(1)) if match else 0
                                        if chg != 0:
                                            try:
                                                ws = sh.worksheet("students"); c = ws.find(sid)
                                                if c:
                                                    h = ws.row_values(1)
                                                    if 'ط§ظ„ظ†ظ‚ط§ط·' in h:
                                                        idx = h.index('ط§ظ„ظ†ظ‚ط§ط·') + 1
                                                        new_val = current_points + chg
                                                        ws.update_cell(c.row, idx, new_val)
                                                        st.session_state.df_students.loc[student_idx, 'ط§ظ„ظ†ظ‚ط§ط·'] = int(new_val)
                                            except Exception as e: st.error(f"ط®ط·ط£: {e}")
                                        
                                        st.toast(f"âœ… طھظ… ط¥ط¶ط§ظپط© ط§ظ„ظ…ظ„ط§ط­ط¸ط© ظ„ظ„ط·ط§ظ„ط¨ {s_nm} ظˆطھط­ط¯ظٹط« ط±طµظٹط¯ظ‡!", icon="ًںژ‰")
                            else: st.info("ًں’، ظˆط¶ط¹ ط§ظ„ظ‚ط±ط§ط،ط© ظپظ‚ط·.")

                        # --- ظ‚ط³ظ… ط§ظ„ط¯ط±ط¬ط§طھ ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹط© ---
                        with c1:
                            st.container(border=True)
                            st.markdown("##### ًں“‌ ط±طµط¯ ط§ظ„ط¯ط±ط¬ط§طھ ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹط©")
                            df_g = st.session_state.df_grades
                            cur_p1 = 0; cur_p2 = 0
                            grade_idx = None
                            
                            if not df_g.empty:
                                df_g['clean_id'] = df_g.iloc[:,0].astype(str).str.split('.').str[0]
                                gr_match = df_g[df_g['clean_id'] == sid]
                                if not gr_match.empty:
                                    grade_idx = gr_match.index[0]
                                    cur_p1 = int(pd.to_numeric(gr_match.iloc[0]['p1'], errors='coerce') or 0)
                                    cur_p2 = int(pd.to_numeric(gr_match.iloc[0]['p2'], errors='coerce') or 0)
                            
                            if st.session_state.role == "teacher":
                                with st.form("gr_upd", clear_on_submit=False):
                                    v1 = st.number_input("ط¯ط±ط¬ط© ط§ظ„ظ…ط´ط§ط±ظƒط©", 0, st.session_state.max_tasks, cur_p1)
                                    v2 = st.number_input("ط¯ط±ط¬ط© ط§ظ„ط§ط®طھط¨ط§ط±", 0, st.session_state.max_quiz, cur_p2)
                                    
                                    if st.form_submit_button("ًں’¾ ط­ظپط¸ ط§ظ„ط¯ط±ط¬ط§طھ", type="primary"):
                                        tot = v1 + v2
                                        ws_g = sh.worksheet("grades")
                                        cell = ws_g.find(sid)
                                        if cell:
                                            ws_g.update_cell(cell.row, 2, v1); ws_g.update_cell(cell.row, 3, v2)
                                            ws_g.update_cell(cell.row, 4, tot); ws_g.update_cell(cell.row, 5, str(datetime.date.today()))
                                        else: 
                                            ws_g.append_row([sid, v1, v2, tot, str(datetime.date.today())])
                                            
                                        if grade_idx is not None:
                                            st.session_state.df_grades.loc[grade_idx, 'p1'] = str(v1)
                                            st.session_state.df_grades.loc[grade_idx, 'p2'] = str(v2)
                                            st.session_state.df_grades.loc[grade_idx, 'perf'] = str(tot)
                                        else:
                                            new_row = pd.DataFrame([[sid, v1, v2, tot, str(datetime.date.today()), sid]], columns=df_g.columns)
                                            st.session_state.df_grades = pd.concat([st.session_state.df_grades, new_row], ignore_index=True)

                                        st.toast("âœ… طھظ… ط§ط¹طھظ…ط§ط¯ ط§ظ„ط¯ط±ط¬ط§طھ ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹط© ط¨ظ†ط¬ط§ط­!", icon="ًںژ“")
                            else: st.info("ًں’، ظˆط¶ط¹ ط§ظ„ظ‚ط±ط§ط،ط© ظپظ‚ط·.")
                            st.caption(f"ًں“ٹ ط§ظ„ظ…ط¬ظ…ظˆط¹ ط§ظ„ط­ط§ظ„ظٹ ظ„ظ„ط¯ط±ط¬ط§طھ: {cur_p1 + cur_p2}")

                        # --- ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ط§ظ„ط³ظپظ„ظٹ ---
                        st.markdown("#### ًں“œ ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ط§ظ„ط£ط®ظٹط±")
                        df_b = st.session_state.df_behavior
                        if not df_b.empty:
                            cid = 'student_id' if 'student_id' in df_b.columns else df_b.columns[0]
                            my_b = df_b[df_b[cid].astype(str) == str(sid)]
                            
                            def delete_behavior(row_idx, global_idx):
                                try: 
                                    sh.worksheet("behavior").delete_rows(int(row_idx) + 2)
                                    st.session_state.df_behavior = st.session_state.df_behavior.drop(global_idx).reset_index(drop=True)
                                except: pass

                            for global_idx, r in my_b.iloc[::-1].iterrows():
                                with st.container():
                                    color = danger_color if "ط³ظ„ط¨ظٹ" in str(r.get('type')) or "-" in str(r.get('type')) else success_color
                                    st.markdown(f"""
                                    <div class="mobile-list-item" style="border-right: 4px solid {color}">
                                        <div><b>{r.get('type')}</b> | <small>{r.get('date')}</small><br><span style="color:#64748B">{r.get('note')}</span></div>
                                    </div>
                                    """, unsafe_allow_html=True)
                                    
                                    c_del, c_wa, c_em = st.columns([0.5, 1, 1])
                                    lnk = get_professional_msg(s_nm, r.get('type'), r.get('note'), r.get('date'))
                                    c_wa.link_button("ظˆط§طھط³ط§ط¨", f"https://api.whatsapp.com/send?phone={clp}&text={lnk}", use_container_width=True)
                                    c_em.link_button("ط¥ظٹظ…ظٹظ„", f"mailto:{s_eml}?subject=ظ…ظ„ط§ط­ط¸ط©: {s_nm}&body={lnk}", use_container_width=True)
                                    
                                    if st.session_state.role == "teacher": 
                                        c_del.button("â‌Œ", key=f"dl_beh_{global_idx}", on_click=delete_behavior, args=(global_idx, global_idx))
            
            # --- 2. ط§ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ ط§ظ„ط³ط±ظٹط¹ ---
            with eval_tabs[1]:
                if st.session_state.role == "teacher":
                    st.markdown("#### ًںڑ€ ط§ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ ظ„ظ„ظ…ظ„ط§ط­ط¸ط§طھ ظˆط§ظ„ظˆط§ط¬ط¨ط§طھ")
                    st.info("ًں’، ط§ط®طھط± ط§ظ„طµظپطŒ ظˆط­ط¯ط¯ ط§ظ„ظ…ظ„ط§ط­ط¸ط§طھ ظ„ظ„ط·ظ„ط§ط¨ ط§ظ„ظ…ط¹ظ†ظٹظٹظ† ظپظ‚ط·طŒ ط«ظ… ط§ط¶ط؛ط· ط­ظپط¸ ط¨ط§ظ„ط£ط³ظپظ„ ظ„طھط±طµط¯ ظ„ظ„ط¬ظ…ظٹط¹ ط¯ظپط¹ط© ظˆط§ط­ط¯ط©.")
                    
                    bulk_class = st.selectbox("ًںژ¯ ط§ط®طھط± ط§ظ„طµظپ ظ„ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ:", st.session_state.class_options, key="bulk_class_sel")
                    df_st_bulk = fetch_safe("students")
                    
                    if not df_st_bulk.empty:
                        df_st_bulk['clean_class'] = df_st_bulk.iloc[:, 2].astype(str).str.strip()
                        class_students = df_st_bulk[df_st_bulk['clean_class'] == bulk_class.strip()]
                        
                        if not class_students.empty:
                            with st.form("bulk_behavior_form", clear_on_submit=True):
                                beh_options = [
                                    "--- ط¨ط¯ظˆظ† ظ…ظ„ط§ط­ط¸ط© ---",
                                    "ًںŒں ظ…طھظ…ظٹط² (+10)", "âœ… ط¥ظٹط¬ط§ط¨ظٹ (+5)", "ًں“‌ ط­ظ„ ط§ظ„ظˆط§ط¬ط¨ (+5)", "ًںژ¯ ط£ط¯ط§ط، ط§ظ„ظ…ظ‡ظ…ط© (+10)", "ًں“‚ ظ…ظ„ظپ ط§ظ„ط¥ظ†ط¬ط§ط² (+10)", 
                                    "âڑ ï¸ڈ طھظ†ط¨ظٹظ‡ (0)", "ًں“ڑ ظ†ظ‚طµ ظƒطھط§ط¨ (-5)", "âœچï¸ڈ ظ†ظ‚طµ ظˆط§ط¬ط¨ (-5)", "ًں–ٹï¸ڈ ظ†ظ‚طµ ط£ط¯ظˆط§طھ ط§ظ„ظƒطھط§ط¨ط© (-5)", "ًں’¤ ط§ظ„ظ†ظˆظ… ط¯ط§ط®ظ„ ط§ظ„ظپطµظ„ (-3)", 
                                    "ًںڈƒ طھط£ط®ط± ط¹ظ† ط§ظ„ط­طµط© (-5)", "â‌Œ ط¹ط¯ظ… ط¥ط­ط¶ط§ط± ظ…ظ„ظپ ط§ظ„ط¥ظ†ط¬ط§ط² (-10)", "ًںڑ« ط³ظ„ط¨ظٹ (-10)"
                                ]
                                
                                bulk_data = {}
                                
                                st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)
                                col_n, col_b, col_t = st.columns([1.5, 2, 2])
                                col_n.markdown("**ًں‘¤ ط§ط³ظ… ط§ظ„ط·ط§ظ„ط¨**"); col_b.markdown("**ًںژ­ ط§ظ„ط³ظ„ظˆظƒ**"); col_t.markdown("**ًں“‌ ظ…ظ„ط§ط­ط¸ط© (ط§ط®طھظٹط§ط±ظٹ)**")
                                st.markdown("<hr style='margin:10px 0'>", unsafe_allow_html=True)

                                for _, row in class_students.iterrows():
                                    sid_b = str(row.iloc[0]).split('.')[0].strip()
                                    sname_b = row.iloc[1]
                                    
                                    c1, c2, c3 = st.columns([1.5, 2, 2])
                                    c1.markdown(f"<div style='padding-top:15px;'>{sname_b}</div>", unsafe_allow_html=True)
                                    b_type = c2.selectbox("ط§ظ„ط³ظ„ظˆظƒ", beh_options, key=f"b_type_{sid_b}", label_visibility="collapsed")
                                    b_note = c3.text_input("طھظپط§طµظٹظ„", key=f"b_note_{sid_b}", label_visibility="collapsed", placeholder="ط£ط¶ظپ طھظپط§طµظٹظ„...")
                                    
                                    bulk_data[sid_b] = {"type": b_type, "note": b_note}
                                    st.markdown("<div style='border-bottom: 1px dashed #E2E8F0; margin: 5px 0;'></div>", unsafe_allow_html=True)

                                if st.form_submit_button("ًںڑ€ ط­ظپط¸ ط§ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ ظ„ظ„ط¬ظ…ظٹط¹", type="primary"):
                                    behavior_rows_to_add = []
                                    point_updates = {}
                                    
                                    for sid_key, data in bulk_data.items():
                                        if data["type"] != "--- ط¨ط¯ظˆظ† ظ…ظ„ط§ط­ط¸ط© ---":
                                            behavior_rows_to_add.append([sid_key, str(datetime.date.today()), data["type"], data["note"]])
                                            match = re.search(r'\(([\+\-]?\d+)\)', data["type"])
                                            if match:
                                                point_updates[sid_key] = int(match.group(1))

                                    if behavior_rows_to_add:
                                        try:
                                            with st.spinner("ط¬ط§ط±ظٹ ط­ظپط¸ ط§ظ„ط±طµط¯ ط§ظ„ط¬ظ…ط§ط¹ظٹ ظˆطھط­ط¯ظٹط« ط§ظ„ظ†ظ‚ط§ط·..."):
                                                sh.worksheet("behavior").append_rows(behavior_rows_to_add)
                                                
                                                ws_st = sh.worksheet("students")
                                                all_st = ws_st.get_all_records()
                                                headers = ws_st.row_values(1)
                                                
                                                if 'ط§ظ„ظ†ظ‚ط§ط·' in headers:
                                                    p_idx = headers.index('ط§ظ„ظ†ظ‚ط§ط·') + 1
                                                    from gspread import Cell
                                                    cells_to_update = []
                                                    
                                                    for i, r in enumerate(all_st):
                                                        st_id = str(r.get('id', '')).split('.')[0].strip()
                                                        if st_id in point_updates:
                                                            cur_p = int(pd.to_numeric(r.get('ط§ظ„ظ†ظ‚ط§ط·', 0), errors='coerce') or 0)
                                                            new_p = cur_p + point_updates[st_id]
                                                            cells_to_update.append(Cell(row=i+2, col=p_idx, value=new_p))
                                                    
                                                    if cells_to_update:
                                                        ws_st.update_cells(cells_to_update)
                                                
                                                st.success(f"âœ… طھظ…طھ ط§ظ„ظ…ظ‡ظ…ط© ط¨ظ†ط¬ط§ط­! طھظ… ط±طµط¯ ({len(behavior_rows_to_add)}) ظ…ظ„ط§ط­ط¸ط©.")
                                                if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                                st.cache_data.clear()
                                                st.rerun()
                                        except Exception as e:
                                            st.error(f"â‌Œ ط­ط¯ط« ط®ط·ط£ ط£ط«ظ†ط§ط، ط§ظ„ط­ظپط¸: {e}")
                                    else:
                                        st.warning("âڑ ï¸ڈ ظ„ظ… طھظ‚ظ… ط¨ط§ط®طھظٹط§ط± ط£ظٹ ط³ظ„ظˆظƒ ظ„ط£ظٹ ط·ط§ظ„ط¨ ظ„ظٹطھظ… ط­ظپط¸ظ‡.")
                        else:
                            st.info("ظ„ط§ ظٹظˆط¬ط¯ ط·ظ„ط§ط¨ ظ…ط³ط¬ظ„ظٹظ† ظپظٹ ظ‡ط°ط§ ط§ظ„طµظپ.")
                else:
                    st.info("ًں’، ظˆط¶ط¹ ط§ظ„ظ‚ط±ط§ط،ط© ظپظ‚ط·.")
                    
        # ًں“¢ ط§ظ„طھظ†ط¨ظٹظ‡ط§طھ
        with tab_alerts:
            st.markdown("### ًں“¢ ظ„ظˆط­ط© ط§ظ„ط¥ط¹ظ„ط§ظ†ط§طھ ظˆط§ظ„طھط¹ط§ظ…ظٹظ…")
            def perform_delete(row_index):
                try: 
                    sh.worksheet("exams").delete_rows(int(row_index) + 2)
                    st.cache_data.clear(); st.toast("âœ… طھظ… ط­ط°ظپ ط§ظ„طھظ†ط¨ظٹظ‡ ط¨ظ†ط¬ط§ط­")
                except Exception as e: st.toast(f"â‌Œ ط­ط¯ط« ط®ط·ط£: {e}")

            if st.session_state.role == "teacher":
                with st.form("ann_add"):
                    c1, c2 = st.columns([3, 1])
                    at = c1.text_input("ط¹ظ†ظˆط§ظ† ط§ظ„ط¥ط¹ظ„ط§ظ†")
                    atg = c2.selectbox("ط§ظ„ظپط¦ط© ط§ظ„ظ…ط³طھظ‡ط¯ظپط©", ["ط§ظ„ظƒظ„"] + st.session_state.class_options)
                    ad = st.text_area("ظ†طµ ط§ظ„ط¥ط¹ظ„ط§ظ† ط£ظˆ ط§ظ„ط±ط§ط¨ط·")
                    au = c1.checkbox("ًں”¥ طھط¹ظ…ظٹظ… ط¹ط§ط¬ظ„ (ظٹط¸ظ‡ط± ط¨ظˆظ…ظٹط¶)")
                    if st.form_submit_button("ًں“£ ظ†ط´ط± ط§ظ„طھط¹ظ…ظٹظ…", type="primary"):
                        safe_append_row("exams", {"ط§ظ„طµظپ": atg, "ط¹ط§ط¬ظ„": "ظ†ط¹ظ…" if au else "ظ„ط§", "ط§ظ„ط¹ظ†ظˆط§ظ†": at, "ط§ظ„طھط§ط±ظٹط®": str(datetime.date.today()), "ط§ظ„ط±ط§ط¨ط·": ad})
                        st.success("âœ… طھظ… ط§ظ„ظ†ط´ط±"); st.cache_data.clear(); st.rerun()
            
            st.divider()
            df_a = fetch_safe("exams")
            if not df_a.empty:
                for i, r in df_a.iloc[::-1].iterrows():
                    with st.container():
                        is_urgent = str(r.get('ط¹ط§ط¬ظ„')).strip() == 'ظ†ط¹ظ…'
                        anim_class = "urgent-box" if is_urgent else ""
                        border_style = f"2px solid {danger_color}" if is_urgent else f"1px solid {border_color}"
                        bg_style = "#FEF2F2" if is_urgent else "#FFFFFF"
                        st.markdown(f"""
                        <div class="{anim_class}" style="background:{bg_style}; border:{border_style}; border-radius:12px; padding:15px; margin-bottom:10px;">
                            <div style="display:flex; justify-content:space-between;"><h4 style="margin:0; color:#0F172A;">{r.get('ط§ظ„ط¹ظ†ظˆط§ظ†')}</h4><span style="background:white; padding:2px 8px; border-radius:8px; font-size:0.8rem; color:#64748B;">{r.get('ط§ظ„طھط§ط±ظٹط®')}</span></div>
                            <p style="margin:5px 0 0 0; color:#475569">{r.get('ط§ظ„ط±ط§ط¨ط·')}</p><small style="color:{accent_color}; font-weight:bold;">ًںژ¯ ط§ظ„ظپط¦ط©: {r.get('ط§ظ„طµظپ')}</small>
                        </div>
                        """, unsafe_allow_html=True)
                        kc1, kc2 = st.columns([1, 4])
                        msg_text = f"ًں“¢ *طھط¹ظ…ظٹظ… ظ‡ط§ظ… ظ…ظ† ظ…ظ†طµط© ط§ظ„ط£ط³طھط§ط° طµط§ظ„ط­*\nâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پ\nًں“Œ *ط§ظ„ط¹ظ†ظˆط§ظ†:* {r.get('ط§ظ„ط¹ظ†ظˆط§ظ†')}\nًں“„ *ط§ظ„طھظپط§طµظٹظ„:* {r.get('ط§ظ„ط±ط§ط¨ط·')}\nًں“… *ط§ظ„طھط§ط±ظٹط®:* {r.get('ط§ظ„طھط§ط±ظٹط®')}\nâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پâ”پ"
                        grp_msg = urllib.parse.quote(msg_text)
                        kc2.link_button("ًں“² ظ…ط´ط§ط±ظƒط© ط¹ط¨ط± ظˆط§طھط³ط§ط¨", f"https://api.whatsapp.com/send?text={grp_msg}", use_container_width=True)
                        if st.session_state.role == "teacher":
                            kc1.button("ًں—‘ï¸ڈ ط­ط°ظپ", key=f"del_btn_unique_{i}", type="secondary", on_click=perform_delete, args=(i,), use_container_width=True)
            else: st.info("ًں’، ظ„ط§ طھظˆط¬ط¯ طھظ†ط¨ظٹظ‡ط§طھ ظ…ظ†ط´ظˆط±ط© ط­ط§ظ„ظٹط§ظ‹.")

        # --- âڑ™ï¸ڈ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ (ط§ظ„ظ…ط¹ظ„ظ… ظپظ‚ط·) ---
        if st.session_state.role == "teacher":
            with tab_settings:
                st.subheader("âڑ™ï¸ڈ ط¥ط¹ط¯ط§ط¯ط§طھ ط§ظ„ظ†ط¸ط§ظ…")
                with st.expander("ًں› ï¸ڈ ط£ط¯ظˆط§طھ ط§ظ„طµظٹط§ظ†ط© ظˆط§ظ„ظ†ط³ط® ط§ظ„ط§ط­طھظٹط§ط·ظٹ", expanded=True):
                    c1, c2 = st.columns(2)
                    
                    if c1.button("ًں”„ طھط­ط¯ظٹط« ط§ظ„ط¨ظٹط§ظ†ط§طھ (Refresh)", use_container_width=True):
                        st.cache_data.clear()
                        if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                        st.session_state['show_refresh_success'] = True 
                        st.rerun()
                        
                    if c2.button("ًں§¹ طھطµظپظٹط± ط¬ظ…ظٹط¹ ط§ظ„ظ†ظ‚ط§ط·", use_container_width=True):
                        try:
                            ws = sh.worksheet("students"); d = ws.get_all_values()
                            if len(d) > 1: 
                                ws.update(range_name=f"I2:I{len(d)}", values=[[0]]*(len(d)-1))
                                st.success("âœ… طھظ… طھطµظپظٹط± ط¬ظ…ظٹط¹ ط§ظ„ظ†ظ‚ط§ط·")
                                if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"ط®ط·ط£: {e}")

                    if st.button("ًں§® ط¥ط¹ط§ط¯ط© ط§ط­طھط³ط§ط¨ ط§ظ„ظ†ظ‚ط§ط· ظ…ظ† ط§ظ„ط³ط¬ظ„ (طھطµط­ظٹط­ ط´ط§ظ…ظ„)", type="primary", use_container_width=True):
                        try:
                            with st.spinner("ط¬ط§ط±ظٹ ظ…ط±ط§ط¬ط¹ط© ط§ظ„ط³ط¬ظ„ط§طھ ظˆطھطµط­ظٹط­ ط£ط±طµط¯ط© ط§ظ„ط·ظ„ط§ط¨..."):
                                df_beh = fetch_safe("behavior"); ws_st = sh.worksheet("students"); students_data = ws_st.get_all_records()
                                true_scores = {}
                                if not df_beh.empty:
                                    for _, row in df_beh.iterrows():
                                        raw_id = str(row.get('student_id', row.get('id', ''))).strip().split('.')[0]
                                        if not raw_id: continue
                                        match = re.search(r'\(([\+\-]?\d+)\)', str(row.get('type', '')))
                                        if match: true_scores[raw_id] = true_scores.get(raw_id, 0) + int(match.group(1))
                                headers = ws_st.row_values(1)
                                if 'ط§ظ„ظ†ظ‚ط§ط·' in headers:
                                    col_idx = headers.index('ط§ظ„ظ†ظ‚ط§ط·') + 1; new_values = []
                                    for st_row in students_data:
                                        sid_v = str(st_row.get('id', '')).strip().split('.')[0]
                                        new_values.append([true_scores.get(sid_v, 0)])
                                    from gspread.utils import rowcol_to_a1
                                    ws_st.update(f"{rowcol_to_a1(2, col_idx)}:{rowcol_to_a1(len(new_values) + 1, col_idx)}", new_values)
                                    st.success("âœ… طھظ… طھطµط­ظٹط­ ط¬ظ…ظٹط¹ ط§ظ„ط£ط±طµط¯ط© ط¨ظ†ط¬ط§ط­!")
                                    if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                    st.cache_data.clear(); st.rerun()
                                else: st.error("ظ„ظ… ظٹطھظ… ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ط¹ظ…ظˆط¯ 'ط§ظ„ظ†ظ‚ط§ط·'")
                        except Exception as e: st.error(f"ط­ط¯ط« ط®ط·ط£: {e}")

                st.divider()
                st.markdown("##### ًں“¥ طھظ†ط²ظٹظ„ ظ†ط³ط®ط© ظƒط§ظ…ظ„ط© ظ…ظ† ط§ظ„ط¨ظٹط§ظ†ط§طھ (Backup)")
                df_st_full = fetch_safe("students")
                if not df_st_full.empty:
                    b_st = io.BytesIO()
                    with pd.ExcelWriter(b_st, engine='xlsxwriter') as writer: df_st_full.to_excel(writer, index=False, sheet_name='Students')
                    st.download_button(label="ًں“‚ طھظ†ط²ظٹظ„ ط¨ظٹط§ظ†ط§طھ ط§ظ„ط·ظ„ط§ط¨ (Excel)", data=b_st.getvalue(), file_name=f"students_backup_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                
                df_gr_full = fetch_safe("grades")
                if not df_gr_full.empty:
                    b_gr = io.BytesIO()
                    with pd.ExcelWriter(b_gr, engine='xlsxwriter') as writer: df_gr_full.to_excel(writer, index=False, sheet_name='Grades')
                    st.download_button(label="ًں“ٹ طھظ†ط²ظٹظ„ ط³ط¬ظ„ ط§ظ„ط¯ط±ط¬ط§طھ (Excel)", data=b_gr.getvalue(), file_name=f"grades_backup_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

                with st.expander("ًں“‌ طھظ‡ظٹط¦ط© ط§ظ„طµظپظˆظپ ظˆط§ظ„ط¯ط±ط¬ط§طھ"):
                    cy = st.text_input("ط§ظ„ط¹ط§ظ… ط§ظ„ط¯ط±ط§ط³ظٹ", st.session_state.current_year)
                    cls = st.text_area("ظ‚ط§ط¦ظ…ط© ط§ظ„طµظپظˆظپ (ط§ظپطµظ„ ط¨ظپط§طµظ„ط©)", ",".join(st.session_state.class_options))
                    stg = st.text_area("ظ‚ط§ط¦ظ…ط© ط§ظ„ظ…ط±ط§ط­ظ„", ",".join(st.session_state.stage_options))
                    c1, c2 = st.columns(2)
                    mt = c1.number_input("ط§ظ„ط¯ط±ط¬ط© ط§ظ„ط¹ط¸ظ…ظ‰ (ظ…ط´ط§ط±ظƒط©)", 0, 100, st.session_state.max_tasks)
                    mq = c2.number_input("ط§ظ„ط¯ط±ط¬ط© ط§ظ„ط¹ط¸ظ…ظ‰ (ط§ط®طھط¨ط§ط±)", 0, 100, st.session_state.max_quiz)
                    if st.button("ًں’¾ ط­ظپط¸ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ", type="primary"):
                        sh.worksheet("settings").batch_update([
                            {'range': 'A2:B2', 'values': [['max_tasks', mt]]}, {'range': 'A3:B3', 'values': [['max_quiz', mq]]},
                            {'range': 'A4:B4', 'values': [['current_year', cy]]}, {'range': 'A5:B5', 'values': [['class_list', cls]]},
                            {'range': 'A6:B6', 'values': [['stage_list', stg]]} 
                        ])
                        st.session_state.max_tasks = mt; st.session_state.max_quiz = mq; st.session_state.current_year = cy
                        st.session_state.class_options = [x.strip() for x in cls.split(',') if x.strip()]
                        st.session_state.stage_options = [x.strip() for x in stg.split(',') if x.strip()]
                        st.success("âœ… طھظ… ط§ظ„ط­ظپط¸ ط¨ظ†ط¬ط§ط­")
                        if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                        st.cache_data.clear(); st.rerun()

                with st.expander("ًں“¤ ط§ط³طھظٹط±ط§ط¯ ط§ظ„ط¨ظٹط§ظ†ط§طھ (Excel) - ط³ط±ظٹط¹"):
                    up = st.file_uploader("ط±ظپط¹ ظ…ظ„ظپ Excel", type=['xlsx'])
                    ts = st.radio("ظ†ظˆط¹ ط§ظ„ط¨ظٹط§ظ†ط§طھ", ["students", "grades"], horizontal=True, format_func=lambda x: "ط¨ظٹط§ظ†ط§طھ ط§ظ„ط·ظ„ط§ط¨" if x == "students" else "ط§ظ„ط¯ط±ط¬ط§طھ")
                    if st.button("ًںڑ€ ط¨ط¯ط، ط§ظ„ظ…ط²ط§ظ…ظ†ط© ط§ظ„ط³ط±ظٹط¹ط©", type="primary") and up:
                        try:
                            with st.spinner('ط¬ط§ط±ظٹ ظ…ط¹ط§ظ„ط¬ط© ظˆط±ظپط¹ ط§ظ„ط¨ظٹط§ظ†ط§طھ ط¯ظپط¹ط© ظˆط§ط­ط¯ط©...'):
                                df = pd.read_excel(up).fillna("").dropna(how='all'); ws = sh.worksheet(ts)
                                existing_ids = set(str(r.get('id', r.get('student_id', ''))).strip().split('.')[0] for r in ws.get_all_records())
                                hd = ws.row_values(1); new_rows_to_append = []; progress_bar = st.progress(0)
                                
                                for idx, row in df.iterrows():
                                    d = row.to_dict()
                                    raw_id = str(d.get('student_id', d.get('id', ''))).strip().split('.')[0]
                                    if not raw_id or raw_id == '0' or raw_id.lower() == 'nan': continue
                                    if ts == "grades":
                                        d.update({"student_id": raw_id, "p1": int(d.get('p1',0)), "p2": int(d.get('p2',0)), "perf": int(d.get('p1',0))+int(d.get('p2',0)), "date": str(datetime.date.today())})
                                        if 'id' in d: del d['id']
                                    else:
                                        d['id'] = raw_id; d['ط§ظ„ط¬ظˆط§ظ„'] = clean_phone_number(d.get('ط§ظ„ط¬ظˆط§ظ„',''))
                                        if 'ط§ظ„ظ†ظ‚ط§ط·' not in d or str(d.get('ط§ظ„ظ†ظ‚ط§ط·', '')).strip() == "": d['ط§ظ„ظ†ظ‚ط§ط·'] = 0
                                    if raw_id not in existing_ids: new_rows_to_append.append([str(d.get(k, "")) for k in hd])
                                    progress_bar.progress(min((idx + 1) / len(df), 1.0))

                                if new_rows_to_append: 
                                    ws.append_rows(new_rows_to_append)
                                    st.success(f"âœ… طھظ… ط¥ط¶ط§ظپط© {len(new_rows_to_append)} ط³ط¬ظ„ ط¬ط¯ظٹط¯ ط¨ظ†ط¬ط§ط­!")
                                else: 
                                    st.info("ًں’، ط¬ظ…ظٹط¹ ط§ظ„ط¨ظٹط§ظ†ط§طھ ظ…ظˆط¬ظˆط¯ط© ظ…ط³ط¨ظ‚ط§ظ‹طŒ ظ„ظ… ظٹطھظ… ط¥ط¶ط§ظپط© ط¬ط¯ظٹط¯.")
                                if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                                st.cache_data.clear(); st.rerun()
                        except Exception as e: st.error(f"ط­ط¯ط« ط®ط·ط£ ط£ط«ظ†ط§ط، ط§ظ„ظ…ط²ط§ظ…ظ†ط©: {e}")
                
                    st.divider(); c1, c2 = st.columns(2)
                    b1 = io.BytesIO(); pd.DataFrame(columns=["id", "name", "class", "year", "sem", "ط§ظ„ط¬ظˆط§ظ„", "ط§ظ„ط¥ظٹظ…ظٹظ„", "ط§ظ„ظ†ظ‚ط§ط·"]).to_excel(b1, index=False)
                    c1.download_button("ًں“¥ ظ‚ط§ظ„ط¨ ظپط§ط±ط؛ (ط·ظ„ط§ط¨)", b1.getvalue(), "students_template.xlsx", use_container_width=True)
                    b2 = io.BytesIO(); pd.DataFrame(columns=["student_id", "p1", "p2"]).to_excel(b2, index=False)
                    c2.download_button("ًں“¥ ظ‚ط§ظ„ط¨ ظپط§ط±ط؛ (ط¯ط±ط¬ط§طھ)", b2.getvalue(), "grades_template.xlsx", use_container_width=True)

                with st.expander("ًں”چ ظ…ط¯ظ‚ظ‚ ط±طµط¯ ط§ظ„ط¯ط±ط¬ط§طھ (ظƒط´ظپ ط§ظ„ظ†ظˆط§ظ‚طµ)", expanded=False):
                    st.markdown("##### ظ‚ط§ط¦ظ…ط© ط§ظ„ط·ظ„ط§ط¨ ط§ظ„ط°ظٹظ† ظ„ظ… ظٹطھظ… ط±طµط¯ ط¯ط±ط¬ط§طھظ‡ظ… ط¨ط¹ط¯")
                    
                    df_st_audit = fetch_safe("students")
                    df_gr_audit = fetch_safe("grades")
                    
                    if not df_st_audit.empty:
                        df_st_audit['clean_id'] = df_st_audit.iloc[:, 0].astype(str).str.strip().str.split('.').str[0]
                        
                        if not df_gr_audit.empty:
                            df_gr_audit['clean_id'] = df_gr_audit.iloc[:, 0].astype(str).str.strip().str.split('.').str[0]
                            audit_merge = pd.merge(
                                df_st_audit[['clean_id', 'name', 'class', 'sem']], 
                                df_gr_audit[['clean_id', 'perf']], 
                                on='clean_id', 
                                how='left'
                            )
                            missing_grades = audit_merge[audit_merge['perf'].isna()]
                        else:
                            missing_grades = df_st_audit[['clean_id', 'name', 'class', 'sem']]

                        if not missing_grades.empty:
                            st.warning(f"âڑ ï¸ڈ ظٹظˆط¬ط¯ {len(missing_grades)} ط·ط§ظ„ط¨ ظ„ظ… طھط±طµط¯ ظ„ظ‡ظ… ط¯ط±ط¬ط§طھ.")
                            display_missing = missing_grades[['clean_id', 'name', 'class', 'sem']].rename(columns={
                                'clean_id': 'ط§ظ„ط±ظ‚ظ… ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ',
                                'name': 'ط§ط³ظ… ط§ظ„ط·ط§ظ„ط¨',
                                'class': 'ط§ظ„طµظپ',
                                'sem': 'ط§ظ„ظ…ط±ط­ظ„ط©'
                            })
                            
                            st.dataframe(display_missing, use_container_width=True, hide_index=True)
                            
                            b_audit = io.BytesIO()
                            with pd.ExcelWriter(b_audit, engine='xlsxwriter') as writer:
                                display_missing.to_excel(writer, index=False, sheet_name='Missing_Grades')
                            
                            st.download_button(
                                label="ًں“¥ طھط­ظ…ظٹظ„ ظ‚ط§ط¦ظ…ط© ط§ظ„ظ†ظˆط§ظ‚طµ ظ„ظ„ط·ط¨ط§ط¹ط©",
                                data=b_audit.getvalue(),
                                file_name=f"missing_grades_{datetime.date.today()}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        else:
                            st.success("âœ… ط¬ظ…ظٹط¹ ط§ظ„ط·ظ„ط§ط¨ ط§ظ„ظ…ط³ط¬ظ„ظٹظ† طھظ… ط±طµط¯ ط¯ط±ط¬ط§طھظ‡ظ… ط¨ظ†ط¬ط§ط­!")
                    else:
                        st.info("ظ„ط§ طھظˆط¬ط¯ ط¨ظٹط§ظ†ط§طھ ط·ظ„ط§ط¨ ظ„ظ„طھط¯ظ‚ظٹظ‚.")

                with st.expander("ًں”گ ط¥ط¯ط§ط±ط© ط§ظ„ظ…ط³طھط®ط¯ظ…ظٹظ† (ظ…ط¹ظ„ظ…ظٹظ† / ط¥ط¯ط§ط±ط©)"):
                    t1, t2, t3 = st.tabs(["â‍• ط¥ط¶ط§ظپط© ظ…ط³طھط®ط¯ظ…", "ًں”‘ طھط¹ط¯ظٹظ„ ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±", "ًں—‘ï¸ڈ ط­ط°ظپ ظ…ط³طھط®ط¯ظ…"])
                    with t1:
                        with st.form("add_u"):
                            nu = st.text_input("ط§ط³ظ… ط§ظ„ظ…ط³طھط®ط¯ظ… ط§ظ„ط¬ط¯ظٹط¯"); np = st.text_input("ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±", type="password")
                            nrole_label = st.selectbox("ظ†ظˆط¹ ط§ظ„ط­ط³ط§ط¨ ظˆط§ظ„طµظ„ط§ط­ظٹط©", ["ًں‘¨â€چًںڈ« ظ…ط¹ظ„ظ… (طµظ„ط§ط­ظٹط§طھ ظƒط§ظ…ظ„ط©)", "ًں‘پï¸ڈ ط¥ط¯ط§ط±ط© (ظ…ط´ط§ظ‡ط¯ ظˆظ‚ط±ط§ط،ط© ظپظ‚ط·)"])
                            if st.form_submit_button("ط¥ط¶ط§ظپط© ط§ظ„ظ…ط³طھط®ط¯ظ…", type="primary"):
                                if nu and np:
                                    role_val = "teacher" if "ظ…ط¹ظ„ظ…" in nrole_label else "viewer"
                                    safe_append_row("users", {"username": nu, "password_hash": hashlib.sha256(np.encode()).hexdigest(), "role": role_val})
                                    st.success(f"âœ… طھظ…طھ ط¥ط¶ط§ظپط© ط§ظ„ط­ط³ط§ط¨ ({nu}) ط¨ظ†ط¬ط§ط­ ظƒظ€ {role_val}.")
                                    st.cache_data.clear()
                                else: st.warning("ط§ظ„ط±ط¬ط§ط، ط¥ظƒظ…ط§ظ„ ط§ظ„ط¨ظٹط§ظ†ط§طھ.")
                    with t2:
                        with st.form("chg_pass"):
                            df_u_edit = fetch_safe("users")
                            if not df_u_edit.empty:
                                target_u = st.selectbox("ط§ط®طھط± ط§ظ„ظ…ط³طھط®ط¯ظ… ظ„طھط¹ط¯ظٹظ„ ط±ظ‚ظ…ظ‡ ط§ظ„ط³ط±ظٹ:", df_u_edit['username'].tolist())
                                npwd = st.text_input("ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط± ط§ظ„ط¬ط¯ظٹط¯ط©", type="password")
                                if st.form_submit_button("طھط­ط¯ظٹط« ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط±", type="primary"):
                                    if npwd:
                                        idx = df_u_edit[df_u_edit['username']==target_u].index[0] + 2
                                        sh.worksheet("users").update_cell(idx, 2, hashlib.sha256(npwd.encode()).hexdigest())
                                        st.success(f"âœ… طھظ… طھط­ط¯ظٹط« ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط± ظ„ظ€ ({target_u}).")
                                        st.cache_data.clear()
                                    else: st.warning("ط§ظ„ط±ط¬ط§ط، ط¥ط¯ط®ط§ظ„ ظƒظ„ظ…ط© ط§ظ„ظ…ط±ظˆط± ط§ظ„ط¬ط¯ظٹط¯ط©.")
                            else: st.info("ظ„ط§ ظٹظˆط¬ط¯ ظ…ط³طھط®ط¯ظ…ظٹظ† ط¨ط¹ط¯.")
                    with t3:
                        df_u_del = fetch_safe("users")
                        if not df_u_del.empty:
                            del_u = st.selectbox("ط§ط®طھط± ط§ظ„ظ…ط³طھط®ط¯ظ… ط§ظ„ظ…ط±ط§ط¯ ط­ط°ظپظ‡:", [""] + df_u_del['username'].tolist())
                            if st.button("ًں—‘ï¸ڈ ط­ط°ظپ ط§ظ„ظ…ط³طھط®ط¯ظ… ظ†ظ‡ط§ط¦ظٹط§ظ‹", type="primary"):
                                if del_u:
                                    if del_u == st.session_state.username: st.error("âڑ ï¸ڈ ظ„ط§ ظٹظ…ظƒظ†ظƒ ط­ط°ظپ ط­ط³ط§ط¨ظƒ ط§ظ„ط­ط§ظ„ظٹ!")
                                    else:
                                        idx = df_u_del[df_u_del['username']==del_u].index[0] + 2
                                        sh.worksheet("users").delete_rows(int(idx)); st.success(f"âœ… طھظ… ط­ط°ظپ ط§ظ„ظ…ط³طھط®ط¯ظ… ({del_u}).")
                                        st.cache_data.clear(); st.rerun()
                                else: st.warning("ط§ظ„ط±ط¬ط§ط، ط§ط®طھظٹط§ط± ظ…ط³طھط®ط¯ظ… ظ„ظ„ط­ط°ظپ.")

        with tab_logout:
            st.markdown("<br><br>", unsafe_allow_html=True)
            if st.button("طھط³ط¬ظٹظ„ ط§ظ„ط®ط±ظˆط¬ ظ…ظ† ظ„ظˆط­ط© ط§ظ„طھط­ظƒظ…", type="secondary"): 
                st.session_state.role = None
                if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                st.rerun()
                
        show_footer()

    # ==========================================
    # ًں‘¨â€چًںژ“ 5. ظˆط§ط¬ظ‡ط© ط§ظ„ط·ط§ظ„ط¨ (ظ…ط¹ ط§ظ„ط£ظ„ظ‚ط§ط¨ ط§ظ„طھط­ظپظٹط²ظٹط©)
    # ==========================================
    elif st.session_state.role == "student":
        sid = str(st.session_state.get('username', '')).strip()
        df_st = fetch_safe("students"); df_gr = fetch_safe("grades"); df_beh = fetch_safe("behavior"); df_ann = fetch_safe("exams")
        
        if not df_st.empty:
            df_st['clean_id'] = df_st.iloc[:,0].astype(str).str.split('.').str[0].str.strip()
            info = df_st[df_st['clean_id'] == sid]
        else: info = pd.DataFrame()

        if not info.empty:
            s_dat = info.iloc[0]
            s_nm = s_dat.get('name', 'ط·ط§ظ„ط¨'); s_cls = str(s_dat.get('class', '')).strip()
            pts = int(pd.to_numeric(s_dat.get('ط§ظ„ظ†ظ‚ط§ط·', 0), errors='coerce') or 0)

            if not df_ann.empty:
                df_ann['ط¹ط§ط¬ظ„'] = df_ann['ط¹ط§ط¬ظ„'].astype(str).str.strip(); df_ann['ط§ظ„طµظپ'] = df_ann['ط§ظ„طµظپ'].astype(str).str.strip()
                urg = df_ann[(df_ann['ط¹ط§ط¬ظ„']=='ظ†ط¹ظ…') & (df_ann['ط§ظ„طµظپ'].isin(['ط§ظ„ظƒظ„', s_cls]))]
                if not urg.empty:
                    u = urg.tail(1).iloc[0]
                    link_text = str(u.get('ط§ظ„ط±ط§ط¨ط·', ''))
                    link_display = f"<a href='{link_text}' target='_blank' style='color:{danger_color}; text-decoration:underline;'>ط§ط¶ط؛ط· ظ‡ظ†ط§</a>" if link_text.startswith('http') else link_text if link_text.lower() != 'none' else ""
                    st.markdown(f"<div class='urgent-box'>ًںڑ¨ {u.get('ط§ظ„ط¹ظ†ظˆط§ظ†')}<br><small style='color:{danger_color}'>{link_display}</small></div>", unsafe_allow_html=True)

            st.markdown(f"""
                <div class="welcome-card">
                    <div style="display:flex; justify-content:space-between; align-items:center;">
                        <div><h2 style="color:white; margin:0; font-size:1.5rem;">ًں‘‹ ط£ظ‡ظ„ط§ظ‹ ط¨ظƒطŒ {s_nm}</h2><p style="color:#DBEAFE; margin:5px 0 0 0;">{s_cls}</p></div>
                        <div style="background:rgba(255,255,255,0.2); padding:5px 15px; border-radius:12px;"><span style="font-weight:bold; font-size:0.9rem; color:#FFFFFF;">ID: {sid}</span></div>
                    </div>
                </div>
                <div class="points-banner">
                    <p style="margin:0; opacity:0.9; font-size:0.9rem;">ط±طµظٹط¯ ط§ظ„ظ†ظ‚ط§ط· ط§ظ„ط­ط§ظ„ظٹ</p>
                    <h1 style="margin:0; font-size:3.5rem; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">{pts}</h1>
                    <p style="margin:0; font-size:0.8rem;">ط§ط³طھظ…ط± ظپظٹ ط§ظ„طھظپظˆظ‚!</p>
                </div>
                <div class="medal-flex">
                    <div class="m-card {'m-active' if pts>=100 else ''}" style="color: {warning_color};">ًں¥‡<br><b>ط°ظ‡ط¨ظٹ</b></div>
                    <div class="m-card {'m-active' if pts>=50 else ''}" style="color: {sub_text};">ًں¥ˆ<br><b>ظپط¶ظٹ</b></div>
                    <div class="m-card m-active" style="color: #B45309;">ًں¥‰<br><b>ط¨ط±ظˆظ†ط²ظٹ</b></div>
                </div>
            """, unsafe_allow_html=True)

            tabs = st.tabs(["ًں“¢ ط§ظ„طھظ†ط¨ظٹظ‡ط§طھ", "ًں“‌ ط§ظ„ط³ظ„ظˆظƒ", "ًں“ٹ ط§ظ„ط¯ط±ط¬ط§طھ", "ًںڈ† ط§ظ„ط´ط±ظپ", "âڑ™ï¸ڈ ط§ظ„ط¥ط¹ط¯ط§ط¯ط§طھ"])

            with tabs[0]: 
                st.caption("ط§ظ„طھط¹ط§ظ…ظٹظ… ظˆط§ظ„طھظ†ط¨ظٹظ‡ط§طھ")
                if not df_ann.empty:
                    anns = df_ann[df_ann['ط§ظ„طµظپ'].astype(str).str.strip().isin(['ط§ظ„ظƒظ„', s_cls])]
                    for _, r in anns.iloc[::-1].iterrows():
                        row_link = str(r.get('ط§ظ„ط±ط§ط¨ط·', ''))
                        row_link_display = f"<a href='{row_link}' target='_blank' style='color:{primary_color}; text-decoration:underline;'>ط§ط¶ط؛ط· ظ‡ظ†ط§ ظ„ظ„ظپطھط­</a>" if row_link.startswith('http') else row_link if row_link.lower() != 'none' else ""
                        st.markdown(f"""
                        <div class='mobile-list-item'>
                            <div style="width:100%">
                                <div style="display:flex; justify-content:space-between; margin-bottom:5px;"><b>ًں“¢ {r.get('ط§ظ„ط¹ظ†ظˆط§ظ†')}</b><small style="background:#EFF6FF; color:{primary_color}; padding:2px 6px; border-radius:4px;">{r.get('ط§ظ„طھط§ط±ظٹط®')}</small></div>
                                <span style="color:#475569; font-size:0.9rem;">{row_link_display}</span>
                            </div>
                        </div>""", unsafe_allow_html=True)
                else: st.info("ظ„ط§ ظٹظˆط¬ط¯ طھظ†ط¨ظٹظ‡ط§طھ ط­ط§ظ„ظٹط§ظ‹")

            with tabs[1]: 
                st.caption("ط³ط¬ظ„ ط§ظ„ط³ظ„ظˆظƒ ظˆط§ظ„ظ…ظ„ط§ط­ط¸ط§طھ")
                if not df_beh.empty:
                    df_beh['clean_id'] = df_beh.iloc[:,0].astype(str).str.split('.').str[0]
                    nts = df_beh[df_beh['clean_id']==sid]
                    if not nts.empty:
                        for _, n in nts.iloc[::-1].iterrows():
                            color = danger_color if "ط³ظ„ط¨ظٹ" in str(n.get('type')) else primary_color
                            st.markdown(f"<div class='mobile-list-item' style='border-right: 4px solid {color};'><div><b style='color:{color}'>{n.get('type')}</b><p style='margin:0; font-size:0.9rem; color:#334155;'>{n.get('note')}</p><small style='color:#94A3B8;'>{n.get('date')}</small></div></div>", unsafe_allow_html=True)
                    else: st.success("ًںŒں ط³ط¬ظ„ظƒ ظ†ط¸ظٹظپ طھظ…ط§ظ…ط§ظ‹!")

            with tabs[2]: 
                st.caption("ط¯ط±ط¬ط§طھظٹ")
                if not df_gr.empty:
                    df_gr['clean_id'] = df_gr.iloc[:,0].astype(str).str.strip().str.split('.').str[0]
                    grs = df_gr[df_gr['clean_id']==sid]
                    if not grs.empty:
                        g = grs.iloc[0]
                        max_total = st.session_state.max_tasks + st.session_state.max_quiz
                        perf_score = int(pd.to_numeric(g.get('perf', 0), errors='coerce') or 0)
                        percentage = (perf_score / max_total) * 100 if max_total > 0 else 0
                        
                        if percentage >= 90: title, title_color = "ًںŒں ط£ط³ط·ظˆط±ط© ط§ظ„ظ…ظ†طµط©", warning_color
                        elif percentage >= 80: title, title_color = "ًںڑ€ ط¨ط·ظ„ ظ…ط¨ط¯ط¹", accent_color
                        elif percentage >= 70: title, title_color = "ًں‘چ ظ…طھط£ظ„ظ‚ ظˆظ…ط¬طھظ‡ط¯", success_color
                        elif percentage >= 60: title, title_color = "ًں’ھ ظˆط§طµظ„ طھظ‚ط¯ظ…ظƒ", primary_color
                        else: title, title_color = "ًںŒ± ط£ظ†طھ ظ‚ط§ط¯ط± ط¹ظ„ظ‰ ط§ظ„ط£ظپط¶ظ„", sub_text
    
                        st.markdown(f"""
                        <div class='mobile-list-item'><span>ًں“‌ ط§ظ„ظ…ط´ط§ط±ظƒط© ظˆط§ظ„ظˆط§ط¬ط¨ط§طھ</span><b>{g.get('p1')} / {st.session_state.max_tasks}</b></div>
                        <div class='mobile-list-item'><span>âœچï¸ڈ ط§ظ„ط§ط®طھط¨ط§ط±ط§طھ ط§ظ„ظ‚طµظٹط±ط©</span><b>{g.get('p2')} / {st.session_state.max_quiz}</b></div>
                        <div class='mobile-list-item' style='background:#EFF6FF; border-color:{accent_color}; display:flex; flex-direction:column; align-items:flex-start;'>
                            <div style="width:100%; display:flex; justify-content:space-between;">
                                <span style="color:{accent_color}; font-weight:bold;">ًںڈ† ط§ظ„ظ…ط¬ظ…ظˆط¹ ط§ظ„ظ†ظ‡ط§ط¦ظٹ</span><b style="color:{accent_color}; font-size:1.2rem;">{perf_score} / {max_total}</b>
                            </div>
                            <div style="margin-top:8px; width:100%; text-align:center; padding:5px; background:white; border-radius:8px; color:{title_color}; font-weight:bold; font-size:1.1rem; border:1px solid {title_color}33;">
                                {title}
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
    
                        if percentage >= 90:
                            st.divider()
                            st.success("ًںژ‰ ظ…ط¨ط±ظˆظƒ! ظ„طھظپظˆظ‚ظƒ ظˆط­طµظˆظ„ظƒ ط¹ظ„ظ‰ ط¯ط±ط¬ط© ط§ظ„ط§ظ…طھظٹط§ط²طŒ طھظ… طھظپط¹ظٹظ„ ظ…ظٹط²ط© ط§ط³طھط®ط±ط§ط¬ 'ط´ظ‡ط§ط¯ط© ط§ظ„طھظپظˆظ‚'.")
                            certificate_html = f"""
                        <!DOCTYPE html>
                        <html dir="rtl" lang="ar">
                        <head>
                            <meta charset="UTF-8">
                            <title>ط´ظ‡ط§ط¯ط© طھظپظˆظ‚ - {s_nm}</title>
                            <link href="https://fonts.googleapis.com/css2?family=Aref+Ruqaa:wght@400;700&family=Amiri:wght@400;700&family=Cairo:wght@400;700;900&display=swap" rel="stylesheet">
                            <style>
                                * {{ box-sizing: border-box; }}
                                body {{
                                    margin: 0; padding: 0;
                                    background-color: #f0f2f5;
                                    display: flex; justify-content: center; align-items: center;
                                    min-height: 100vh;
                                }}
                                .cert-page {{
                                    width: 297mm; height: 210mm;
                                    padding: 10mm;
                                    background: #ffffff;
                                    position: relative;
                                    box-shadow: 0 0 20px rgba(0,0,0,0.1);
                                    overflow: hidden;
                                }}
                                .border-outer {{
                                    border: 14px solid #193b68;
                                    height: 100%; width: 100%;
                                    padding: 6px;
                                    position: relative;
                                }}
                                .border-inner {{
                                    border: 3px solid #b68a36;
                                    height: 100%; width: 100%;
                                    position: relative;
                                    padding: 30px 40px;
                                    text-align: center;
                                    background-image: radial-gradient(#e2e8f0 1px, transparent 1px);
                                    background-size: 25px 25px;
                                    -webkit-print-color-adjust: exact !important;
                                    print-color-adjust: exact !important;
                                }}
                                
                                .corner {{ position: absolute; width: 30px; height: 30px; border: 4px solid #b68a36; }}
                                .tl {{ top: -8px; left: -8px; border-right: none; border-bottom: none; }}
                                .tr {{ top: -8px; right: -8px; border-left: none; border-bottom: none; }}
                                .bl {{ bottom: -8px; left: -8px; border-right: none; border-top: none; }}
                                .br {{ bottom: -8px; right: -8px; border-left: none; border-top: none; }}

                                .top-badge {{ position: absolute; top: 25px; left: 35px; width: 90px; height: 90px; }}

                                h1 {{ 
                                    font-family: 'Aref Ruqaa', serif; font-size: 70px; color: #b68a36; 
                                    margin: 0; font-weight: normal; text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
                                }}
                                
                                h2 {{ 
                                    font-family: 'Cairo', sans-serif; font-size: 34px; color: #193b68; 
                                    margin: 5px 0 25px 0; font-weight: 900;
                                    display: flex; justify-content: center; align-items: center; gap: 15px;
                                }}
                                
                                p.intro {{ font-family: 'Cairo', sans-serif; font-size: 22px; color: #193b68; margin-bottom: 10px; }}
                                
                                .student-name {{ 
                                    font-family: 'Cairo', sans-serif; font-size: 55px; font-weight: 900; color: #d32f2f; 
                                    margin: 15px auto; display: inline-block;
                                    border-bottom: 3px solid #b68a36; padding-bottom: 5px; padding-left: 30px; padding-right: 30px;
                                    letter-spacing: 1px;
                                }}
                                
                                .reason-container {{ margin-top: 15px; margin-bottom: 25px; }}
                                
                                p.reason {{ 
                                    font-family: 'Cairo', sans-serif; font-size: 24px; font-weight: 600; color: #444; 
                                    line-height: 1.8; margin: 0;
                                }}
                                
                                .footer-section {{ 
                                    display: flex; justify-content: space-between; align-items: flex-end; 
                                    margin-top: 30px; padding: 0 30px; font-family: 'Cairo', sans-serif; 
                                }}
                                
                                .sig-box, .date-box {{ text-align: center; color: #193b68; width: 220px; }}
                                .sig-title {{ font-size: 22px; font-weight: bold; }}
                                .sig-line {{ border-bottom: 1px solid #b68a36; width: 150px; margin: 8px auto; }}
                                .sig-name {{ font-family: 'Aref Ruqaa', serif; font-size: 34px; color: #193b68; line-height: 1; }}
                                .date-val {{ font-size: 22px; font-weight: bold; color: #333; margin-top: 8px; }}

                                .stamp-wrapper {{
                                    display: flex; justify-content: center; align-items: center;
                                    margin-bottom: 10px; width: 150px;
                                }}

                                @media print {{
                                    @page {{ size: A4 landscape; margin: 0mm; }}
                                    body {{ min-height: auto; align-items: flex-start; justify-content: flex-start; background: white; }}
                                    .cert-page {{ padding: 0; box-shadow: none; width: 297mm; height: 210mm; overflow: hidden; page-break-after: avoid; page-break-before: avoid; }}
                                }}
                            </style>
                        </head>
                        <body>
                            <div class="cert-page">
                                <div class="border-outer">
                                    <div class="border-inner">
                                        <div class="corner tl"></div><div class="corner tr"></div>
                                        <div class="corner bl"></div><div class="corner br"></div>
                                        
                                        <img src="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMDAgMTAwIj48Y2lyY2xlIGN4PSI1MCIgY3k9IjUwIiByPSI0MCIgZmlsbD0iIzE5M2I2OCIgc3Ryb2tlPSIjYjY4YTM2IiBzdHJva2Utd2lkdGg9IjUiLz48Y2lyY2xlIGN4PSI1MCIgY3k9IjUwIiByPSIzMCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjYjY4YTM2IiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1kYXNoYXJyYXk9IjQgNCIvPjxwYXRoIGQ9Ik0zNSA2NUw1MCA0MEw2NSA2NVoiIGZpbGw9IiNiNjhhMzYiLz48cGF0aCBkPSJNMzUgMzVMNjUgMzVMMTUgNjVIMzVaIiBmaWxsPSIjYjY4YTM2IiBvcGFjaXR5PSIwLjciLz48L3N2Zz4=" class="top-badge">

                                        <h1>ط´ظ‡ط§ط¯ط© ط´ظƒط± ظˆطھظ‚ط¯ظٹط±</h1>
                                        
                                        <h2>
                                            <img src="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjYjY4YTM2IiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlnb24gcG9pbnRzPSIxMiAyIDE1LjA5IDguMjYgMjIgOS4yNyAxNyAxNC4xNCAxOC4xOCAyMS4wMiAxMiAxNy43NyA1LjgyIDIxLjAyIDcgMTQuMTQgMiA5LjI3IDguOTEgOC4yNiAxMiAyIiBmaWxsPSIjYjY4YTM2IiAvPjwvc3ZnPg==" style="width:30px; height:30px;">
                                            ظˆط³ط§ظ… ط§ظ„طھظ…ظٹط² ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ
                                            <img src="data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgc3Ryb2tlPSIjYjY4YTM2IiBzdHJva2Utd2lkdGg9IjIiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIgc3Ryb2tlLWxpbmVqb2luPSJyb3VuZCI+PHBvbHlnb24gcG9pbnRzPSIxMiAyIDE1LjA5IDguMjYgMjIgOS4yNyAxNyAxNC4xNCAxOC4xOCAyMS4wMiAxMiAxNy43NyA1LjgyIDIxLjAyIDcgMTQuMTQgMiA5LjI3IDguOTEgOC4yNiAxMiAyIiBmaWxsPSIjYjY4YTM2IiAvPjwvc3ZnPg==" style="width:30px; height:30px;">
                                        </h2>

                                        <p class="intro">ظٹطھظ‚ط¯ظ… ط§ظ„ط£ط³طھط§ط°/ <strong>طµط§ظ„ط­ ط§ظ„ط±ظˆظٹط«ظٹ</strong> ط¨ظˆط§ظپط± ط§ظ„ط´ظƒط± ظˆط§ظ„طھظ‚ط¯ظٹط± ظ„ظ„ط·ط§ظ„ط¨ ط§ظ„ظ…ط¨ط¯ط¹ ظˆط§ظ„ظ…طھط£ظ„ظ‚:</p>
                                        
                                        <div class="student-name">{s_nm}</div>
                                        
                                        <div class="reason-container">
                                            <p class="reason">ظˆط°ظ„ظƒ ظ†ط¸ظٹط± طھظپظˆظ‚ظ‡ ط§ظ„ط¹ظ„ظ…ظٹ ظˆط­طµظˆظ„ظ‡ ط¹ظ„ظ‰ طھظ‚ط¯ظٹط± <b style="color:#d32f2f;">ظ…ظ…طھط§ط²</b> ظپظٹ ظ…ط§ط¯ط© ط§ظ„ظ„ط؛ط© ط§ظ„ط¥ظ†ط¬ظ„ظٹط²ظٹط©.</p>
                                            <p class="reason">ظ…طھظ…ظ†ظٹظ† ظ„ظ‡ ط¯ظˆط§ظ… ط§ظ„طھظˆظپظٹظ‚ ظˆظ…ط²ظٹط¯ط§ظ‹ ظ…ظ† ط§ظ„طھط£ظ„ظ‚ ظˆط§ظ„ظ†ط¬ط§ط­.</p>
                                        </div>

                                        <div class="footer-section">
                                            <div class="date-box">
                                                <div class="sig-title">طھط§ط±ظٹط® ط§ظ„ط¥طµط¯ط§ط±</div>
                                                <div class="sig-line"></div>
                                                <div class="date-val">{datetime.date.today().strftime('%Y-%m-%d')}</div>
                                            </div>
                                            
                                            <div class="stamp-wrapper">
                                                <div style="width: 140px; height: 140px; border: 3px dashed #d32f2f; border-radius: 50%; transform: rotate(-15deg); color: #d32f2f; text-align: center; opacity: 0.85; position: relative; padding-top: 25px; margin: 0 auto; background-color: rgba(255, 255, 255, 0.7);">
                                                    <div style="position: absolute; top: 4px; left: 4px; right: 4px; bottom: 4px; border: 1px solid #d32f2f; border-radius: 50%;"></div>
                                                    <div style="font-family: 'Cairo', sans-serif; font-size: 16px; font-weight: 900; line-height: 1; margin-top: 5px;">ظˆط³ط§ظ…</div>
                                                    <div style="font-family: 'Aref Ruqaa', serif; font-size: 34px; font-weight: bold; line-height: 1.2; margin: 2px 0;">ط®طھظ… ط§ظ„طھظ…ظٹط²</div>
                                                    <div style="font-family: 'Cairo', sans-serif; font-size: 12px; font-weight: bold; line-height: 1;">ط§ظ„ط£ظƒط§ط¯ظٹظ…ظٹ</div>
                                                </div>
                                            </div>

                                            <div class="sig-box">
                                                <div class="sig-title">طھظˆظ‚ظٹط¹ ط§ظ„ظ…ط¹ظ„ظ…</div>
                                                <div class="sig-line"></div>
                                                <div class="sig-name">طµط§ظ„ط­ ط§ظ„ط±ظˆظٹط«ظٹ</div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </body>
                        </html>
                        """
                            
                            try:
                                try:`n    from weasyprint import HTML`nexcept Exception:`n    HTML = None
                                with st.spinner("âڈ³ ط¬ط§ط±ظٹ ط¥ط¹ط¯ط§ط¯ ط´ظ‡ط§ط¯ط© ط§ظ„طھظپظˆظ‚ ط¨طµظٹط؛ط© PDF..."):
                                    pdf_bytes = HTML(string=certificate_html).write_pdf()
                                    
                                    st.download_button(
                                        label="ًں“¥ طھط­ظ…ظٹظ„ ط´ظ‡ط§ط¯ط© ط§ظ„طھظپظˆظ‚ (PDF ظ…ط¨ط§ط´ط±)",
                                        data=pdf_bytes,
                                        file_name=f"Certificate_{sid}_{s_nm}.pdf",
                                        mime="application/pdf",
                                        type="primary",
                                        use_container_width=True
                                    )
                            except Exception as e:
                                st.error(f"âڑ ï¸ڈ ظپط´ظ„ طھظˆظ„ظٹط¯ ط§ظ„ظ€ PDF ط¨ط³ط¨ط¨: {e}")
                                st.info("ًں’، ظ…ظ„ط§ط­ط¸ط©: طھظ… طھظپط¹ظٹظ„ طھط­ظ…ظٹظ„ ظ†ط³ط®ط© ط§ظ„ظˆظٹط¨ ظ„ط³ط±ط¹ط© ط§ظ„ظˆطµظˆظ„.")
                                st.download_button(
                                    label="ًں“œ ط§ط³طھط®ط±ط§ط¬ ط´ظ‡ط§ط¯ط© ط§ظ„طھظپظˆظ‚ (ظ†ط³ط®ط© ظˆظٹط¨)",
                                    data=certificate_html,
                                    file_name=f"Certificate_{sid}.html",
                                    mime="text/html",
                                    type="primary",
                                    use_container_width=True
                                )
    
                    else: st.info("ظ„ظ… ظٹطھظ… ط±طµط¯ ط¯ط±ط¬ط§طھ ط¨ط¹ط¯")
            with tabs[3]: 
                st.caption("ظ„ظˆط­ط© ط§ظ„ط´ط±ظپ (ط£ظپط¶ظ„ 10 ط·ظ„ط§ط¨)")
                df_st['p_num'] = pd.to_numeric(df_st['ط§ظ„ظ†ظ‚ط§ط·'], errors='coerce').fillna(0)
                for i, (_, r) in enumerate(df_st.sort_values('p_num', ascending=False).head(10).iterrows(), 1):
                    ic = "ًں¥‡" if i==1 else "ًں¥ˆ" if i==2 else "ًں¥‰" if i==3 else f"#{i}"
                    sty = f"border:2px solid {primary_color}; background:#EFF6FF;" if str(r['clean_id']) == sid else ""
                    st.markdown(f"<div class='mobile-list-item' style='{sty}'><div style='display:flex; align-items:center; gap:10px;'><span style='font-weight:900; font-size:1.2rem; width:30px;'>{ic}</span><span>{r['name']}</span></div><span style='color:{warning_color}; font-weight:900;'>{int(r['p_num'])}</span></div>", unsafe_allow_html=True)

            with tabs[4]:
                st.caption("ط¥ط¯ط§ط±ط© ط§ظ„ظ…ظ„ظپ ط§ظ„ط´ط®طµظٹ")
                with st.form("my_profile"):
                    nm = st.text_input("ًں“§ ط§ظ„ط¨ط±ظٹط¯ ط§ظ„ط¥ظ„ظƒطھط±ظˆظ†ظٹ", s_dat.get('ط§ظ„ط¥ظٹظ…ظٹظ„',''))
                    np = st.text_input("ًں“± ط±ظ‚ظ… ط§ظ„ط¬ظˆط§ظ„", s_dat.get('ط§ظ„ط¬ظˆط§ظ„',''))
                    if st.form_submit_button("ًں’¾ طھط­ط¯ظٹط« ط¨ظٹط§ظ†ط§طھظٹ", type="primary", use_container_width=True):
                        try:
                            fp = clean_phone_number(np) if np else ""
                            ws = sh.worksheet("students"); c = ws.find(sid)
                            if c:
                                h = ws.row_values(1)
                                if 'ط§ظ„ط¥ظٹظ…ظٹظ„' in h and 'ط§ظ„ط¬ظˆط§ظ„' in h:
                                    ws.update_cell(c.row, h.index('ط§ظ„ط¥ظٹظ…ظٹظ„')+1, nm); ws.update_cell(c.row, h.index('ط§ظ„ط¬ظˆط§ظ„')+1, fp); st.success("âœ… طھظ… ط§ظ„طھط­ط¯ظٹط«")
                                else: st.error("ط®ط·ط£ ظ‡ظٹظƒظ„ظٹ")
                        except Exception as e: st.error(f"ط®ط·ط£: {e}")
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("ًںڑھ طھط³ط¬ظٹظ„ ط§ظ„ط®ط±ظˆط¬", type="secondary", use_container_width=True): 
                    st.session_state.role = None
                    if 'db_loaded' in st.session_state: del st.session_state['db_loaded']
                    st.rerun()

        else: st.error("ط¹ط°ط±ط§ظ‹طŒ ظ„ظ… ظٹطھظ… ط§ظ„ط¹ط«ظˆط± ط¹ظ„ظ‰ ط¨ظٹط§ظ†ط§طھظƒ"); st.button("ط§ظ„ط¹ظˆط¯ط© ظ„ظ„ظ‚ط§ط¦ظ…ط© ط§ظ„ط±ط¦ظٹط³ظٹط©", on_click=st.rerun)
        
        show_footer()

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
KST = ZoneInfo('Asia/Seoul')
from streamlit_cookies_controller import CookieController

# 1. 페이지 기본 설정
st.set_page_config(
    page_title="레이븐 리더 길드 아지트",
    page_icon=None,
    layout="wide"
)

controller = CookieController()

SHEET_ID = "1rQmw3ryJyQZw_ou9sQ1NuewMSg_wcH5fkxfDg7uy55k"
SCOPES = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def get_gspread_client():
    try:
        # Streamlit Cloud 환경 (Secrets 사용)
        secret_dict = {k: (dict(v) if hasattr(v, '_asdict') else v) for k, v in st.secrets["gcp_service_account"].items()}
        creds = Credentials.from_service_account_info(secret_dict, scopes=SCOPES)
    except (KeyError, FileNotFoundError):
        # 로컬 환경 (credentials.json 사용)
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    return gspread.authorize(creds)

def load_sheet_data(sheet_name):
    try:
        client = get_gspread_client()
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet(sheet_name)
        all_values = ws.get_all_values()
        if not all_values or len(all_values) < 2:
            return pd.DataFrame()
        headers = [str(h).strip().lower() for h in all_values[0]]
        rows = all_values[1:]
        df = pd.DataFrame(rows, columns=headers)
        df = df.replace('', pd.NA).dropna(how='all', axis=0).fillna('')
        return df
    except Exception as e:
        st.error(f"시트 로드 실패 ({sheet_name}): {e}")
        return pd.DataFrame()

def save_member_to_sheet(name, member_data):
    try:
        client = get_gspread_client()
        sh = client.open_by_key(SHEET_ID)
        ws = sh.worksheet("guildmembers")
        all_values = ws.get_all_values()
        headers = [h.strip().lower() for h in all_values[0]]
        # job 컬럼 없으면 자동 추가
        if 'job' not in headers:
            next_col = len(headers) + 1
            ws.update_cell(1, next_col, 'job')
            headers.append('job')
        new_row_data = {
            'name': name,
            'password': member_data.get('password', ''),
            'gold': member_data.get('gold', 0),
            'atk': member_data.get('atk', 0),
            'def': member_data.get('def', 0),
            'hit': member_data.get('hit', 0),
            'power': member_data.get('power', 0),
            'job': member_data.get('job', '-'),
            'updated_at': member_data.get('updated_at', ''),
            'boss_le': member_data['attendance'].get('레기카', False),
            'boss_si': member_data['attendance'].get('시온', False),
            'boss_fl': member_data['attendance'].get('플라우드', False),
        }
        new_row = [str(new_row_data.get(h, '')) for h in headers]
        cell = ws.find(name)
        if cell:
            ws.update(f"A{cell.row}", [new_row])
        else:
            ws.append_row(new_row)
        return True
    except Exception as e:
        st.error(f"시트 저장 실패: {e}")
        return False

def save_auction_to_sheet(auction_items):
    try:
        client = get_gspread_client()
        sh = client.open_by_key(SHEET_ID)
        try:
            ws = sh.worksheet("auction_items")
            ws.clear()
        except:
            ws = sh.add_worksheet(title="auction_items", rows="100", cols="10")
        headers = ["name", "boss", "price", "status", "bidder", "bidders", "registered_at", "deadline"]
        ws.update("A1", [headers])
        rows = []
        for item in auction_items:
            rows.append([
                item.get("name", ""),
                item.get("boss", ""),
                item.get("price", 0),
                item.get("status", ""),
                item.get("bidder", ""),
                ",".join(item.get("bidders", [])),
                item.get("registered_at", ""),
                item.get("deadline", ""),
            ])
        if rows:
            ws.update("A2", rows)
        return True
    except Exception as e:
        st.error(f"입찰 시트 저장 실패: {e}")
        return False

def save_finance_to_sheet(finance):
    try:
        client = get_gspread_client()
        sh = client.open_by_key(SHEET_ID)
        try:
            ws = sh.worksheet("guild_finance")
            ws.clear()
        except:
            ws = sh.add_worksheet(title="guild_finance", rows="20", cols="2")
        ws.update("A1", [["category", "amount"]])
        rows = [[k, v] for k, v in finance.items()]
        if rows:
            ws.update("A2", rows)
        return True
    except Exception as e:
        st.error(f"자산 시트 저장 실패: {e}")
        return False

def load_auction_from_sheet():
    try:
        df = load_sheet_data("auction_items")
        if df.empty:
            return []
        items = []
        for _, row in df.iterrows():
            items.append({
                "name": str(row.get("name", "")),
                "boss": str(row.get("boss", "")),
                "price": int(pd.to_numeric(row.get("price", 0), errors="coerce") or 0),
                "status": str(row.get("status", "경매중")),
                "bidder": str(row.get("bidder", "-")),
                "bidders": [b for b in str(row.get("bidders", "")).split(",") if b and b != "nan"],
                "registered_at": str(row.get("registered_at", "")),
                "deadline": str(row.get("deadline", "")),
            })
        return items
    except:
        return []

def convert_sheets_to_dict(members_df, finance_df):
    db_data = {"guildmembers": {}, "guild_finance": {}}
    if not finance_df.empty:
        for _, row in finance_df.iterrows():
            cat = str(row.get('category', '')).strip()
            amt = pd.to_numeric(row.get('amount', 0), errors='coerce')
            if pd.notna(amt):
                db_data["guild_finance"][cat] = int(amt)
    else:
        db_data["guild_finance"] = {"guild_money": 250000, "power_dist": 120000, "attend_dist": 130000}
    if not members_df.empty:
        for _, row in members_df.iterrows():
            name_key = str(row.get('name', '')).strip()
            if not name_key or name_key == 'nan':
                continue
            db_data["guildmembers"][name_key] = {
                "password": str(row.get('password', '')),
                "gold": int(pd.to_numeric(row.get('gold', 0), errors='coerce') or 0),
                "atk": int(pd.to_numeric(row.get('atk', 0), errors='coerce') or 0),
                "def": int(pd.to_numeric(row.get('def', 0), errors='coerce') or 0),
                "hit": int(pd.to_numeric(row.get('hit', 0), errors='coerce') or 0),
                "power": int(pd.to_numeric(row.get('power', 0), errors='coerce') or 0) if pd.notna(pd.to_numeric(row.get('power', 0), errors='coerce')) else 0,
                "job": str(row.get('job', '-')) if str(row.get('job', '-')) not in ['nan', ''] else '-',
                "updated_at": str(row.get('updated_at', '-')),
                "attendance": {
                    "레기카": str(row.get('boss_le', 'FALSE')).upper() == 'TRUE',
                    "시온": str(row.get('boss_si', 'FALSE')).upper() == 'TRUE',
                    "플라우드": str(row.get('boss_fl', 'FALSE')).upper() == 'TRUE'
                }
            }
    return db_data

# 낙찰 처리 함수
def finalize_auction(item, db_data):
    """마감 시간이 지난 아이템 낙찰 처리"""
    if item["status"] != "경매중":
        return item
    if not item.get("deadline") or str(item["deadline"]).strip() in ["", "nan"]:
        return item
    try:
        deadline = datetime.strptime(item["deadline"], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return item
    if datetime.now(KST) < deadline.replace(tzinfo=KST):
        return item
    bidders = item.get("bidders", [])
    if not bidders:
        item["status"] = "유찰"
        item["bidder"] = "-"
        return item
    # 전투력 1위 낙찰
    best = max(bidders, key=lambda n: db_data["guildmembers"].get(n, {}).get("power", 0))
    item["status"] = "낙찰"
    item["bidder"] = best
    return item

if "db_data" not in st.session_state:
    m_df = load_sheet_data("guildmembers")
    f_df = load_sheet_data("guild_finance")
    st.session_state.db_data = convert_sheets_to_dict(m_df, f_df)

if "auction_items" not in st.session_state:
    st.session_state.auction_items = load_auction_from_sheet()

# 마감된 아이템 자동 낙찰 처리
changed = False
for i, item in enumerate(st.session_state.auction_items):
    updated = finalize_auction(item, st.session_state.db_data)
    if updated["status"] != item["status"]:
        st.session_state.auction_items[i] = updated
        changed = True
if changed:
    save_auction_to_sheet(st.session_state.auction_items)

cookie_user = controller.get('saved_user_id')
if "logged_in" not in st.session_state:
    if cookie_user and cookie_user in st.session_state.db_data["guildmembers"]:
        st.session_state.logged_in = True
        st.session_state.login_user = cookie_user
    else:
        st.session_state.logged_in = False
if "login_user" not in st.session_state:
    st.session_state.login_user = cookie_user if cookie_user else ""

st.markdown("""
<style>
    .stApp { background-color: #0b0e14; color: #ffffff; font-family: 'Inter', sans-serif; }
    .neon-title {
        text-align: center; font-size: 2.2rem; font-weight: 900; color: #ffffff;
        text-shadow: 0 0 10px rgba(0, 149, 255, 0.6), 0 0 20px rgba(0, 149, 255, 0.3);
        margin-bottom: 35px; letter-spacing: 2px;
    }
    .stElementContainer div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #111622 !important; border: 1px solid #1e293b !important;
        border-radius: 12px !important; padding: 20px !important;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25) !important;
    }
    .section-title {
        font-size: 1.1rem; font-weight: 700; color: #ffffff; margin-bottom: 15px;
        letter-spacing: 0.5px; border-bottom: 1px solid #1e293b; padding-bottom: 10px;
    }
    p, span, label, div { color: #ffffff !important; }
    .stat-val { font-size: 1.8rem; font-weight: 700; margin-top: 5px; margin-bottom: 15px; }
    .blue-txt { color: #38bdf8 !important; text-shadow: 0 0 5px rgba(56,189,248,0.4); }
    .purple-txt { color: #c084fc !important; text-shadow: 0 0 5px rgba(192,132,252,0.4); }
    .green-txt { color: #4ade80 !important; text-shadow: 0 0 5px rgba(74,222,128,0.4); }
    .guild-roster-table { width: 100% !important; max-width: 500px; border-collapse: collapse !important; margin: 5px 0 !important; }
    .guild-roster-table th { background-color: #1a2333 !important; color: #ffffff !important; text-align: center !important; padding: 6px 2px !important; font-size: 0.75rem !important; font-weight: 700 !important; border-bottom: 2px solid #2e3d56 !important; }
    .guild-roster-table td { padding: 8px 2px !important; font-size: 0.75rem !important; color: #ffffff !important; border-bottom: 1px solid #1e293b !important; text-align: center !important; }
    .member-name-tag { font-weight: 700 !important; color: #ffffff !important; display: block; text-align: left !important; padding-left: 4px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .member-power-tag { font-weight: 800 !important; color: #4ade80 !important; }
    .member-time-tag { font-size: 0.65rem !important; color: #cbd5e1 !important; text-align: center !important; display: block; }
    .spec-atk { color: #f43f5e !important; font-weight: 700; }
    .spec-def { color: #3b82f6 !important; font-weight: 700; }
    .spec-hit { color: #eab308 !important; font-weight: 700; }
    div[data-testid="stNumberInput"] input { color: #111622 !important; font-weight: 700 !important; font-size: 0.95rem !important; text-align: center !important; }
    div[data-testid="stNumberInput"] label { color: #ffffff !important; font-weight: 600; }
    div.stButton > button:first-child { background: rgba(0, 149, 255, 0.1) !important; color: #0095ff !important; border: 1px solid #0095ff !important; }
    .stButton-withdraw > div > button { background: rgba(0, 229, 255, 0.1) !important; color: #00e5ff !important; border: 1px solid #00e5ff !important; white-space: nowrap !important; font-size: 0.85rem !important; }
    .stButton-attend > div > button { background: rgba(168, 85, 247, 0.1) !important; color: #a855f7 !important; border: 1px solid #a855f7 !important; white-space: nowrap !important; font-size: 0.85rem !important; }
    .stButton-refresh > div > button { background: rgba(16, 185, 129, 0.1) !important; color: #10b981 !important; border: 1px solid #10b981 !important; font-weight: 700 !important; white-space: nowrap !important; font-size: 0.85rem !important; }
    .item-card { background: #141b29; border: 1px solid #1e293b; border-radius: 10px; padding: 14px; margin-bottom: 10px; }
    .item-card-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
    .item-name { font-weight: 700; font-size: 0.95rem; color: #ffffff !important; }
    .item-boss { font-size: 0.75rem; color: #94a3b8 !important; margin-top: 2px; }
    .item-price { font-size: 1rem; font-weight: 800; color: #f59e0b !important; }
    .item-footer { display: flex; justify-content: space-between; align-items: center; margin-top: 8px; font-size: 0.78rem; }
    .item-bidder { color: #4ade80 !important; font-weight: 700; }
    .item-deadline { color: #94a3b8 !important; }
    .item-bidders { color: #cbd5e1 !important; font-size: 0.72rem; margin-top: 4px; }
    .status-badge { font-size: 0.65rem; padding: 2px 8px; border-radius: 20px; font-weight: 700; }
    .status-ongoing { background: rgba(239,68,68,0.2); color: #ef4444 !important; border: 1px solid #ef4444; }
    .status-done { background: rgba(74,222,128,0.2); color: #4ade80 !important; border: 1px solid #4ade80; }
    .status-fail { background: rgba(100,116,139,0.2); color: #94a3b8 !important; border: 1px solid #475569; }
    div[data-testid="stSelectbox"] > div > div { background-color: #1a2333 !important; color: #ffffff !important; border: 1px solid #2e3d56 !important; }
    div[data-testid="stSelectbox"] span { color: #ffffff !important; }
    div[data-testid="stSelectbox"] svg { fill: #ffffff !important; }
    [data-baseweb="select"] { background-color: #1a2333 !important; }
    [data-baseweb="select"] * { color: #ffffff !important; }
    [data-baseweb="popover"] { background-color: #1a2333 !important; }
    [data-baseweb="menu"] { background-color: #1a2333 !important; border: 1px solid #2e3d56 !important; }
    [data-baseweb="menu"] li { background-color: #1a2333 !important; color: #ffffff !important; }
    [data-baseweb="menu"] li:hover { background-color: #2e3d56 !important; }
    /* 모든 selectbox 어두운 배경 + 흰 글씨 통일 */
    div[data-testid="stSelectbox"] > div > div { background-color: #1a2333 !important; border: 1px solid #2e3d56 !important; }
    div[data-testid="stSelectbox"] > div > div > div { color: #ffffff !important; }
    div[data-testid="stSelectbox"] > div > div * { color: #ffffff !important; }
    div[data-testid="stSelectbox"] svg { fill: #ffffff !important; }
    [data-baseweb="select"] { background-color: #1a2333 !important; }
    [data-baseweb="select"] > div { background-color: #1a2333 !important; }
    [data-baseweb="select"] * { color: #ffffff !important; }
    [data-baseweb="popover"] { background-color: #1a2333 !important; }
    [data-baseweb="menu"] { background-color: #1a2333 !important; border: 1px solid #2e3d56 !important; }
    [data-baseweb="menu"] li { background-color: #1a2333 !important; color: #ffffff !important; }
    [data-baseweb="menu"] li * { color: #ffffff !important; }
    [data-baseweb="menu"] li:hover { background-color: #2e3d56 !important; }
    .stButton-bid > div > button { background: rgba(251,191,36,0.1) !important; color: #fbbf24 !important; border: 1px solid #fbbf24 !important; font-size: 0.85rem !important; }
    .stButton-bid-diamond > div > button { background: rgba(56,189,248,0.1) !important; color: #38bdf8 !important; border: 1px solid #38bdf8 !important; font-size: 0.85rem !important; }
    .stButton-cancel > div > button { background: rgba(239,68,68,0.1) !important; color: #ef4444 !important; border: 1px solid #ef4444 !important; font-size: 0.85rem !important; margin-top: 0px !important; }
    .flag-left {
        position: fixed; top: 50%; left: 0px;
        transform: translateY(-50%);
        width: 130px; opacity: 0.12; z-index: 0;
        pointer-events: none;
    }
    .flag-right {
        position: fixed; top: 50%; right: 0px;
        transform: translateY(-50%) scaleX(-1);
        width: 130px; opacity: 0.12; z-index: 0;
        pointer-events: none;
    }

    div[data-testid="stVerticalBlock"] > div { gap: 0.3rem !important; }
    .item-card + div { margin-top: 0 !important; }
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] + div { margin-top: 0 !important; }
    div.stButton { margin-top: 0 !important; margin-bottom: 0 !important; }
    div.stButton > button { margin-top: 0 !important; margin-bottom: 0 !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='neon-title'>레이븐 리더 길드 아지트</div>", unsafe_allow_html=True)
st.markdown("""<style>
div[data-testid="stTextInput"]:has(input[placeholder="마스터"]) {
    position: fixed; top: 12px; right: 16px; width: 120px; z-index: 9999;
}
div[data-testid="stTextInput"]:has(input[placeholder="마스터"]) input {
    font-size: 0.75rem !important; padding: 4px 8px !important; height: 32px !important;
}
</style>""", unsafe_allow_html=True)
quick_pw_top = st.text_input("", placeholder="마스터", type="password", key="quick_pw_top", label_visibility="collapsed")
if quick_pw_top == "1234":
    if "마스터" not in st.session_state.db_data["guildmembers"]:
        st.session_state.db_data["guildmembers"]["마스터"] = {
            "password":"1234","gold":0,"atk":0,"def":0,"hit":0,"power":0,
            "updated_at":datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
            "attendance":{"레기카":False,"시온":False,"플라우드":False}
        }
    st.session_state.logged_in = True
    st.session_state.login_user = "마스터"
    st.rerun()

if "auth_mode" not in st.session_state:
    st.session_state.auth_mode = "login"

if True:
    col_left, col_center, col_right = st.columns([1.0, 1.0, 0.9])
    current_user = st.session_state.get("login_user", "")
    is_master = current_user == "마스터"
    member_info = st.session_state.db_data["guildmembers"].get(
        current_user,
        {"password":"","gold":0,"atk":0,"def":0,"hit":0,"power":0,
         "updated_at":"-","attendance":{"레기카":False,"시온":False,"플라우드":False}}
    )

    # ── 왼쪽 컬럼 ──
    with col_left:
        with st.container(border=True):
            if not st.session_state.logged_in:
                # 로그인/회원가입 폼
                input_id = st.text_input("ID", placeholder="길드원 계정명 입력", key="login_id")
                input_pw = st.text_input("PASSWORD", type="password", placeholder="비밀번호 입력", key="login_pw")
                btn_l, btn_r = st.columns(2)
                with btn_l:
                    if st.button("로그인", use_container_width=True, key="do_login"):
                        if input_id in st.session_state.db_data["guildmembers"]:
                            if st.session_state.db_data["guildmembers"][input_id]["password"] == input_pw:
                                st.session_state.logged_in = True
                                st.session_state.login_user = input_id
                                controller.set('saved_user_id', input_id)
                                st.rerun()
                            else: st.error("❌ 비밀번호 오류")
                        else: st.error("❌ 미등록 계정")
                with btn_r:
                    if st.button("📝 회원가입", use_container_width=True, key="mode_register"):
                        st.session_state.auth_mode = "register"
                        st.rerun()

                if st.session_state.auth_mode == "register":
                    st.markdown("<div style='border-top:1px solid #1e293b;margin:10px 0;'></div>", unsafe_allow_html=True)
                    reg_id = st.text_input("인게임 닉네임", key="reg_id")
                    reg_pw = st.text_input("새 비밀번호", type="password", key="reg_pw")
                    reg_pw_confirm = st.text_input("비밀번호 확인", type="password", key="reg_pw_confirm")
                    rc1, rc2 = st.columns(2)
                    with rc1:
                        if st.button("가입하기", use_container_width=True):
                            if reg_pw != reg_pw_confirm:
                                st.error("비밀번호가 일치하지 않아요!")
                            elif reg_id in st.session_state.db_data["guildmembers"]:
                                st.error("이미 존재하는 아이디예요!")
                            else:
                                new_member = {
                                    "password":reg_pw,"gold":0,"atk":0,"def":0,"hit":0,"power":0,
                                    "updated_at":datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S"),
                                    "attendance":{"레기카":False,"시온":False,"플라우드":False},
                                    "job":"-"
                                }
                                st.session_state.db_data["guildmembers"][reg_id] = new_member
                                if save_member_to_sheet(reg_id, new_member):
                                    st.success("가입 완료! ✅")
                    with rc2:
                        if st.button("↩️ 취소", use_container_width=True):
                            st.session_state.auth_mode = "login"
                            st.rerun()
            else:
                # 로그인 후 계정 정보
                st.markdown(f"<div class='section-title'>👤 현재 <span style='color:#38bdf8;'>{current_user}</span> 님의 정보</div>", unsafe_allow_html=True)
                btn1, btn2, btn3 = st.columns(3)
                with btn1:
                    if st.button("✏️ 닉네임", use_container_width=True):
                        st.session_state.show_nick_editor = not st.session_state.get("show_nick_editor", False)
                with btn2:
                    if st.button("🚪 로그아웃", use_container_width=True):
                        st.session_state.logged_in = False
                        st.session_state.login_user = ""
                        try:
                            controller.remove('saved_user_id')
                        except:
                            pass
                        st.rerun()
                with btn3:
                    if st.button("🔄 새로고침", use_container_width=True):
                        m_df = load_sheet_data("guildmembers")
                        f_df = load_sheet_data("guild_finance")
                        st.session_state.db_data = convert_sheets_to_dict(m_df, f_df)
                        st.session_state.auction_items = load_auction_from_sheet()
                        st.rerun()
            if st.session_state.get("show_nick_editor", False):
                st.markdown("<div style='background:#141b29;border:1px solid #2e3d56;border-radius:8px;padding:12px;margin-top:4px;'>", unsafe_allow_html=True)
                new_nick = st.text_input("새 닉네임 입력", key="new_nick_input")
                if st.button("✅ 변경 완료", use_container_width=True):
                    if new_nick.strip() == "":
                        st.error("닉네임을 입력해주세요!")
                    elif new_nick.strip() in st.session_state.db_data["guildmembers"]:
                        st.error("이미 사용중인 닉네임이에요!")
                    else:
                        old_data = st.session_state.db_data["guildmembers"][current_user]
                        st.session_state.db_data["guildmembers"][new_nick.strip()] = old_data
                        del st.session_state.db_data["guildmembers"][current_user]
                        save_member_to_sheet(new_nick.strip(), old_data)
                        try:
                            client = get_gspread_client()
                            sh = client.open_by_key(SHEET_ID)
                            ws = sh.worksheet("guildmembers")
                            cell = ws.find(current_user)
                            if cell:
                                ws.delete_rows(cell.row)
                        except Exception as e:
                            st.error(f"기존 닉네임 삭제 실패: {e}")
                        st.session_state.login_user = new_nick.strip()
                        controller.set('saved_user_id', new_nick.strip())
                        st.session_state.show_nick_editor = False
                        st.success(f"✅ 닉네임이 {new_nick.strip()} 으로 변경됐어요!")
                        st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)


            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown("<span style='font-size:13px;'>예상 분배금</span>", unsafe_allow_html=True)
                st.markdown(f"<div class='stat-val blue-txt'>{member_info['gold']:,} 💎</div>", unsafe_allow_html=True)
                st.markdown("<div class='stButton-withdraw'>", unsafe_allow_html=True)
                if st.button("💸 출금 신청", use_container_width=True):
                    st.session_state.db_data["guildmembers"][current_user]["gold"] = 0
                    save_member_to_sheet(current_user, st.session_state.db_data["guildmembers"][current_user])
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)
            with c2:
                st.markdown("<span style='font-size:13px;'>현재 참여도</span>", unsafe_allow_html=True)
                st.markdown("<div class='stat-val purple-txt'>94.5 %</div>", unsafe_allow_html=True)
                st.markdown("<div class='stButton-attend'>", unsafe_allow_html=True)
                if st.button("📅 출석 체크", use_container_width=True): st.success("확인 완료")
                st.markdown("</div>", unsafe_allow_html=True)
            with c3:
                st.markdown("<span style='font-size:13px;'>현재 전투력</span>", unsafe_allow_html=True)
                st.markdown(f"<div class='stat-val green-txt'>{member_info['power']:,}</div>", unsafe_allow_html=True)
                st.markdown("<div class='stButton-refresh'>", unsafe_allow_html=True)
                if st.button("⚡ 투력 최신화", use_container_width=True):
                    st.session_state.show_power_editor = not st.session_state.get("show_power_editor", False)
                st.markdown("</div>", unsafe_allow_html=True)

            if st.session_state.get("show_power_editor", False):
                st.write("")
                st.markdown("<div style='border:1px dashed #2e3d56;padding:15px;border-radius:8px;background-color:#141b29;'>", unsafe_allow_html=True)
                st.caption("⚔️ **세부 능력치 입력 콘솔**")
                job_list = ["뱅가드","버서커","디스트","레인저","엘리","디바인","어쌔신","데브","건슬","워로드"]
                cur_job = member_info.get('job', '-')
                if "selected_job" not in st.session_state:
                    st.session_state.selected_job = cur_job if cur_job in job_list else job_list[0]
                st.markdown("""
                <style>
                div[data-testid="stRadio"] > div {
                    display: flex; flex-wrap: wrap; gap: 6px;
                }
                div[data-testid="stRadio"] > div > label {
                    display: flex !important;
                    background: #2e3d56 !important;
                    color: #ffffff !important;
                    border: 1px solid #3e4d66 !important;
                    border-radius: 6px !important;
                    padding: 5px 12px !important;
                    font-size: 0.8rem !important;
                    font-weight: 700 !important;
                    cursor: pointer !important;
                    margin: 0 !important;
                }
                div[data-testid="stRadio"] > div > label:has(input:checked) {
                    background: #0095ff !important;
                    border-color: #0095ff !important;
                    color: #ffffff !important;
                }
                div[data-testid="stRadio"] > div > label > div:first-child { display: none !important; }
                </style>
                """, unsafe_allow_html=True)
                job_idx = job_list.index(st.session_state.selected_job) if st.session_state.selected_job in job_list else 0
                edit_job = st.radio("⚔️ 직업", job_list, index=job_idx, key="job_radio", horizontal=True)
                st.session_state.selected_job = edit_job
                edit_atk = st.number_input("💥 공격력", value=member_info.get('atk', 0), step=1000, key="edit_atk")
                edit_def = st.number_input("🛡️ 방어력", value=member_info.get('def', 0), step=1000, key="edit_def")
                edit_hit = st.number_input("🎯 명중률", value=member_info.get('hit', 0), step=500, key="edit_hit")
                calc_total = edit_atk + edit_def + edit_hit
                st.markdown(
                    f"<div style='margin-top:12px;background:#1c2536;padding:8px;border-radius:4px;text-align:center;'>"
                    f"<span style='font-size:12px;'>공/방/명 합산 전투력</span><br/>"
                    f"<strong style='color:#4ade80 !important;font-size:1.15rem;'>{calc_total:,}</strong></div>",
                    unsafe_allow_html=True
                )
                st.write("")
                if st.button("적용", key="apply_power"):
                    st.session_state.db_data["guildmembers"][current_user].update({
                        "atk":edit_atk,"def":edit_def,"hit":edit_hit,"power":calc_total,
                        "job":edit_job,
                        "updated_at":datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S")
                    })
                    with st.spinner("저장 중..."):
                        ok = save_member_to_sheet(current_user, st.session_state.db_data["guildmembers"][current_user])
                    if ok: st.success("✅ 저장 완료!")
                    st.session_state.show_power_editor = False
                    st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("<div class='section-title'>📋 길드원 리스트 스펙 현황</div>", unsafe_allow_html=True)
            table_rows = []
            # 정렬 상태 초기화
            if "sort_col" not in st.session_state:
                st.session_state.sort_col = "power"
                st.session_state.sort_asc = False

            def sort_members(col):
                if st.session_state.sort_col == col:
                    st.session_state.sort_asc = not st.session_state.sort_asc
                else:
                    st.session_state.sort_col = col
                    st.session_state.sort_asc = False

            def get_sort_arrow(col):
                if st.session_state.sort_col == col:
                    return " ▲" if st.session_state.sort_asc else " ▼"
                return " ▼"

            # 정렬 버튼 한 줄
            sc1,sc2,sc3,sc4,sc5,sc6,sc7 = st.columns(7)
            with sc1:
                if st.button(f"직업 🔍", key="sort_job", use_container_width=True):
                    st.session_state.show_job_filter = not st.session_state.get("show_job_filter", False)
                    st.rerun()
            with sc2:
                if st.button(f"공격{get_sort_arrow('atk')}", key="sort_atk", use_container_width=True):
                    sort_members("atk"); st.rerun()
            with sc3:
                if st.button(f"방어{get_sort_arrow('def')}", key="sort_def", use_container_width=True):
                    sort_members("def"); st.rerun()
            with sc4:
                if st.button(f"명중{get_sort_arrow('hit')}", key="sort_hit", use_container_width=True):
                    sort_members("hit"); st.rerun()
            with sc5:
                if st.button(f"총합{get_sort_arrow('power')}", key="sort_power", use_container_width=True):
                    sort_members("power"); st.rerun()
            with sc6:
                if st.button(f"갱신{get_sort_arrow('updated_at')}", key="sort_time", use_container_width=True):
                    sort_members("updated_at"); st.rerun()
            with sc7:
                if st.button("🔄", key="sort_reset", use_container_width=True):
                    st.session_state.sort_col = "power"
                    st.session_state.sort_asc = False
                    st.rerun()

            # 직업 필터 창
            if st.session_state.get("show_job_filter", False):
                job_list_all = ["전체", "뱅가드","버서커","디스트","레인저","엘리","디바인","어쌔신","데브","건슬","워로드"]
                selected_filter = st.session_state.get("job_filter", "전체")
                st.markdown("<div style='background:#141b29;border:1px solid #2e3d56;border-radius:8px;padding:10px;margin-top:4px;'>", unsafe_allow_html=True)
                st.markdown("<span style='font-size:0.8rem;color:#ffffff;'>직업 선택</span>", unsafe_allow_html=True)
                jf_cols = st.columns(6)
                for ji, jname in enumerate(job_list_all):
                    with jf_cols[ji % 6]:
                        if st.button(jname, key=f"jf_{ji}", use_container_width=True):
                            st.session_state.job_filter = jname
                            st.session_state.show_job_filter = False
                            st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

            # 직업 필터 적용
            job_filter = st.session_state.get("job_filter", "전체")

            # 정렬 적용
            def get_sort_val(item):
                if st.session_state.sort_col == 'job':
                    return (str(item[1].get('job', '')), -int(item[1].get('power', 0)))
                val = item[1].get(st.session_state.sort_col, '')
                if st.session_state.sort_col in ['atk','def','hit','power','gold']:
                    try: return int(val)
                    except: return 0
                return str(val)

            sorted_members = sorted(
                st.session_state.db_data["guildmembers"].items(),
                key=get_sort_val,
                reverse=not st.session_state.sort_asc
            )
            if job_filter != "전체":
                sorted_members = [(n, d) for n, d in sorted_members if d.get('job', '-') == job_filter]
            crown_icons = {0: "👑", 1: "🥈", 2: "🥉"}
            for rank, (name, m_data) in enumerate(sorted_members):
                t_val = str(m_data.get('updated_at', '-'))
                if t_val not in ["-", "nan"]:
                    try: t_val = datetime.strptime(t_val, "%Y-%m-%d %H:%M:%S").strftime("%m-%d %H:%M")
                    except: pass
                else: t_val = "-"
                crown = crown_icons.get(rank, "")
                if crown:
                    crown_html = f"<span style='font-size:0.75rem;'>{crown}</span> "
                else:
                    crown_html = f"<span style='font-size:0.72rem;color:#64748b;font-weight:700;'>{rank+1}. </span>"
                job_val = m_data.get('job', '-')
                if job_val in ['nan', '', None]: job_val = '-'
                table_rows.append(
                    f"<tr>"
                    f"<td style='text-align:left;width:22%;'><span class='member-name-tag'>{crown_html}{name}</span></td>"
                    f"<td style='width:10%;'><span style='font-size:0.72rem;color:#c084fc;font-weight:700;'>{job_val}</span></td>"
                    f"<td style='width:12%;'><span class='spec-atk'>{int(m_data.get('atk',0)):,}</span></td>"
                    f"<td style='width:12%;'><span class='spec-def'>{int(m_data.get('def',0)):,}</span></td>"
                    f"<td style='width:10%;'><span class='spec-hit'>{int(m_data.get('hit',0)):,}</span></td>"
                    f"<td style='width:12%;'><span class='member-power-tag'>{int(m_data.get('power',0)):,}</span></td>"
                    f"<td style='width:22%;'><span class='member-time-tag'>{t_val}</span></td>"
                    f"</tr>"
                )
            st.html(
                f"<table class='guild-roster-table'><thead><tr>"
                f"<th style='width:22%;'>이름</th><th style='width:10%;'>직업</th><th style='width:12%;'>공격</th>"
                f"<th style='width:12%;'>방어</th><th style='width:10%;'>명중</th>"
                f"<th style='width:12%;'>총합</th><th style='width:22%;'>갱신</th>"
                f"</tr></thead><tbody>{''.join(table_rows)}</tbody></table>"
            )

    # ── 가운데 컬럼 ──
    with col_center:
        # ── 아이템 입찰 현황 ──
        with st.container(border=True):
            st.markdown("<div class='section-title'>🏆 아이템 입찰 현황</div>", unsafe_allow_html=True)

            # 마스터 전용: 아이템 등록
            if is_master:
                if st.button("➕ 아이템 등록", key="toggle_item_reg", use_container_width=True):
                    st.session_state.show_item_reg = not st.session_state.get("show_item_reg", False)
                if st.session_state.get("show_item_reg", False):
                    st.markdown("<div style='background:#141b29;border:1px solid #2e3d56;border-radius:8px;padding:12px;margin-top:4px;'>", unsafe_allow_html=True)
                    new_name = st.text_input("아이템 이름", key="new_item_name")
                    new_boss = st.selectbox("드롭 보스", ["레기카", "시온", "플라우드"], key="new_item_boss")
                    new_price = st.number_input("판매 가격 (💎)", min_value=0, step=1000, key="new_item_price")
                    if st.button("📌 아이템 등록", use_container_width=True):
                        if new_name.strip():
                            now = datetime.now()
                            deadline = now + timedelta(hours=12)
                            new_item = {
                                "name": new_name.strip(),
                                "boss": new_boss,
                                "price": new_price,
                                "status": "경매중",
                                "bidder": "-",
                                "bidders": [],
                                "registered_at": now.strftime("%Y-%m-%d %H:%M:%S"),
                                "deadline": deadline.strftime("%Y-%m-%d %H:%M:%S"),
                            }
                            st.session_state.auction_items.append(new_item)
                            with st.spinner("시트에 저장 중..."):
                                save_auction_to_sheet(st.session_state.auction_items)
                            st.success(f"✅ '{new_name}' 등록 완료! 마감: {deadline.strftime('%m/%d %H:%M')}")
                            st.session_state.show_item_reg = False
                            st.rerun()
                        else:
                            st.error("아이템 이름을 입력해주세요!")
                    st.markdown("</div>", unsafe_allow_html=True)

            # 아이템 목록 표시
            if not st.session_state.auction_items:
                st.markdown("<div style='text-align:center;color:#475569;padding:20px;'>등록된 아이템이 없어요</div>", unsafe_allow_html=True)
            else:
                # 우선순위 정렬
                def sort_priority(item):
                    if item.get("status") != "경매중":
                        return (99, 0)  # 경매중 아닌 건 맨 아래
                    try:
                        deadline_dt = datetime.strptime(item["deadline"], "%Y-%m-%d %H:%M:%S")
                        remaining_sec = (deadline_dt.replace(tzinfo=KST) - datetime.now(KST)).total_seconds()
                    except:
                        remaining_sec = 99999
                    is_urgent = remaining_sec <= 3600  # 1시간 이내
                    has_bidders = len(item.get("bidders", [])) > 0
                    price = item.get("price", 0)

                    if is_urgent and not has_bidders:
                        return (0, remaining_sec)   # 1순위: 임박 + 입찰자 없음
                    elif is_urgent and has_bidders:
                        return (1, remaining_sec)   # 2순위: 임박 + 입찰자 있음
                    elif price >= 1000:
                        return (2, -price)           # 3순위: 1000다이아 이상
                    else:
                        return (3, remaining_sec)   # 나머지

                sorted_items = sorted(enumerate(st.session_state.auction_items), key=lambda x: sort_priority(x[1]))

                for i, item in sorted_items:
                    # 마감 시간 계산
                    try:
                        deadline_dt = datetime.strptime(item["deadline"], "%Y-%m-%d %H:%M:%S")
                        remaining = deadline_dt.replace(tzinfo=KST) - datetime.now(KST)
                        if remaining.total_seconds() > 0:
                            hours, rem = divmod(int(remaining.total_seconds()), 3600)
                            mins = rem // 60
                            time_str = f"⏱ {hours}시간 {mins}분 남음"
                        else:
                            time_str = "⏰ 마감"
                    except:
                        time_str = ""

                    status = item["status"]
                    if status == "경매중":
                        badge = "<span class='status-badge status-ongoing'>경매중</span>"
                    elif status == "낙찰":
                        badge = "<span class='status-badge status-done'>낙찰</span>"
                    else:
                        badge = "<span class='status-badge status-fail'>유찰</span>"

                    bidders_list = item.get("bidders", [])
                    bidders_str = ", ".join(bidders_list) if bidders_list else "없음"

                    # 현재 전투력 1위 입찰자
                    if bidders_list:
                        top_bidder = max(bidders_list, key=lambda n: st.session_state.db_data["guildmembers"].get(n, {}).get("power", 0))
                        top_power = st.session_state.db_data["guildmembers"].get(top_bidder, {}).get("power", 0)
                        bidder_display = f"🏅 {top_bidder} ({top_power:,})"
                    else:
                        bidder_display = "-"

                    st.markdown(
                        f"<div class='item-card'>"
                        f"<div class='item-card-header'>"
                        f"<span class='item-name'>{item['name']} {badge}</span>"
                        f"<span class='item-price'>{item['price']:,} 💎</span>"
                        f"</div>"
                        f"<div class='item-boss'>👹 {item['boss']} 드롭</div>"
                        f"<div class='item-footer'>"
                        f"<span class='item-bidder'>현재 1위: {bidder_display}</span>"
                        f"<span class='item-deadline'>{time_str}</span>"
                        f"</div>"
                        f"<div class='item-bidders'>입찰자: {bidders_str}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

                    # 입찰/취소 + 강제종료 버튼 (경매중일 때만)
                    if status == "경매중":
                        already_bid = current_user in bidders_list
                        if not already_bid:
                            # 마스터: 3열 / 일반: 2열
                            if is_master:
                                btn_col1, btn_col2, btn_col3 = st.columns(3)
                            else:
                                btn_col1, btn_col2 = st.columns(2)
                            with btn_col1:
                                st.markdown("<div class='stButton-bid'>", unsafe_allow_html=True)
                                if st.button(f"💰 분배금 차감", key=f"bid_gold_{i}", use_container_width=True):
                                    price = item.get("price", 0)
                                    cur_gold = st.session_state.db_data["guildmembers"].get(current_user, {}).get("gold", 0)
                                    if cur_gold < price:
                                        st.error(f"분배금 부족! (보유: {cur_gold:,} / 필요: {price:,} 💎)")
                                    else:
                                        st.session_state.db_data["guildmembers"][current_user]["gold"] = cur_gold - price
                                        st.session_state.auction_items[i]["bidders"].append(current_user)
                                        save_member_to_sheet(current_user, st.session_state.db_data["guildmembers"][current_user])
                                        with st.spinner("저장 중..."):
                                            save_auction_to_sheet(st.session_state.auction_items)
                                        st.success(f"✅ 분배금 {price:,}💎 차감 후 입찰 완료!")
                                        st.rerun()
                                st.markdown("</div>", unsafe_allow_html=True)
                            with btn_col2:
                                st.markdown("<div class='stButton-bid-diamond'>", unsafe_allow_html=True)
                                if st.button(f"💎 다이아로 입찰", key=f"bid_dia_{i}", use_container_width=True):
                                    st.session_state.auction_items[i]["bidders"].append(current_user)
                                    with st.spinner("저장 중..."):
                                        save_auction_to_sheet(st.session_state.auction_items)
                                    st.success("✅ 입찰 완료! (다이아 직접 입금해주세요)")
                                    st.rerun()
                                st.markdown("</div>", unsafe_allow_html=True)
                            if is_master:
                                with btn_col3:
                                    st.markdown("<div class='stButton-cancel'>", unsafe_allow_html=True)
                                    if st.button(f"🔴 강제종료", key=f"force_{i}", use_container_width=True):
                                        st.session_state[f"force_confirm_{i}"] = True
                                    st.markdown("</div>", unsafe_allow_html=True)
                        else:
                            # 입찰 중 상태
                            if is_master:
                                bid_col1, bid_col2, bid_col3 = st.columns(3)
                            else:
                                bid_col1, bid_col2 = st.columns(2)
                            with bid_col1:
                                st.markdown("<div style='color:#4ade80;font-size:0.8rem;text-align:center;padding:6px;border:1px solid #4ade80;border-radius:6px;'>✅ 입찰 중</div>", unsafe_allow_html=True)
                            with bid_col2:
                                st.markdown("<div class='stButton-cancel'>", unsafe_allow_html=True)
                                if st.button(f"❌ 취소", key=f"cancel_{i}", use_container_width=True):
                                    st.session_state.auction_items[i]["bidders"].remove(current_user)
                                    with st.spinner("저장 중..."):
                                        save_auction_to_sheet(st.session_state.auction_items)
                                    st.success("입찰 취소됐어요!")
                                    st.rerun()
                                st.markdown("</div>", unsafe_allow_html=True)
                            if is_master:
                                with bid_col3:
                                    st.markdown("<div class='stButton-cancel'>", unsafe_allow_html=True)
                                    if st.button(f"🔴 강제종료", key=f"force_{i}", use_container_width=True):
                                        st.session_state[f"force_confirm_{i}"] = True
                                    st.markdown("</div>", unsafe_allow_html=True)

                        # 강제종료 선택창 (버튼 누른 후 표시)
                        if is_master and st.session_state.get(f"force_confirm_{i}", False):
                            bidders_now = st.session_state.auction_items[i].get("bidders", [])
                            if bidders_now:
                                top_now = max(bidders_now, key=lambda n: st.session_state.db_data["guildmembers"].get(n, {}).get("power", 0))
                                top_power_now = st.session_state.db_data["guildmembers"].get(top_now, {}).get("power", 0)
                                top_info = f"현재 1위: {top_now} ({top_power_now:,})"
                            else:
                                top_now = None
                                top_info = "입찰자 없음"
                            st.markdown(
                                f"<div style='background:#1a2333;border:1px solid #ef4444;border-radius:8px;padding:12px;margin-top:4px;font-size:0.85rem;'>"
                                f"⚠️ <b>강제종료 방식을 선택해주세요</b><br>"
                                f"<span style='color:#94a3b8;'>{top_info}</span>"
                                f"</div>",
                                unsafe_allow_html=True
                            )
                            fc1, fc2, fc3 = st.columns(3)
                            with fc1:
                                if st.button("🏅 낙찰로 종료", key=f"force_win_{i}", use_container_width=True):
                                    if top_now:
                                        st.session_state.auction_items[i]["status"] = "낙찰"
                                        st.session_state.auction_items[i]["bidder"] = top_now
                                    else:
                                        st.session_state.auction_items[i]["status"] = "유찰"
                                        st.session_state.auction_items[i]["bidder"] = "-"
                                    save_auction_to_sheet(st.session_state.auction_items)
                                    st.session_state[f"force_confirm_{i}"] = False
                                    st.rerun()
                            with fc2:
                                if st.button("❌ 유찰로 종료", key=f"force_fail_{i}", use_container_width=True):
                                    st.session_state.auction_items[i]["status"] = "유찰"
                                    st.session_state.auction_items[i]["bidder"] = "-"
                                    save_auction_to_sheet(st.session_state.auction_items)
                                    st.session_state[f"force_confirm_{i}"] = False
                                    st.rerun()
                            with fc3:
                                if st.button("↩️ 취소", key=f"force_back_{i}", use_container_width=True):
                                    st.session_state[f"force_confirm_{i}"] = False
                                    st.rerun()

                    # 낙찰 확정 → 자동 분배 적용 (마스터 전용)
                    if status == "낙찰" and is_master:
                        winner = item.get("bidder", "-")
                        price = item.get("price", 0)
                        admin  = int(price * 0.06)
                        guild  = int(price * 0.30)
                        attend_pool = int(price * 0.32)
                        power_pool  = int(price * 0.32)
                        st.markdown(
                            f"<div style='background:#1a2333;border-radius:8px;padding:12px;margin-top:8px;font-size:0.82rem;'>"
                            f"🏅 낙찰자: <b style='color:#4ade80;'>{winner}</b> &nbsp;|&nbsp; 낙찰가: <b style='color:#f59e0b;'>{price:,} 💎</b><br><br>"
                            f"<span style='color:#94a3b8;'>자동 분배 내역</span><br>"
                            f"🔸 총무비 (6%): <b>{admin:,} 💎</b><br>"
                            f"🔸 혈비 (30%): <b>{guild:,} 💎</b><br>"
                            f"🔸 참여분배 (32%): <b>{attend_pool:,} 💎</b><br>"
                            f"🔸 투력분배 (32%): <b>{power_pool:,} 💎</b>"
                            f"</div>",
                            unsafe_allow_html=True
                        )
                        if st.button(f"✅ 분배 확정 & 자산 반영", key=f"confirm_{i}", use_container_width=True):
                            finance = st.session_state.db_data.get("guild_finance", {})
                            finance["guild_money"]  = finance.get("guild_money", 0)  + guild
                            finance["power_dist"]   = finance.get("power_dist", 0)   + power_pool
                            finance["attend_dist"]  = finance.get("attend_dist", 0)  + attend_pool
                            st.session_state.db_data["guild_finance"] = finance
                            st.session_state.auction_items[i]["status"] = "정산완료"
                            with st.spinner("시트에 저장 중..."):
                                save_auction_to_sheet(st.session_state.auction_items)
                                save_finance_to_sheet(finance)
                            st.success(f"✅ 자산 반영 완료! 혈비 +{guild:,} / 투력분배 +{power_pool:,} / 참여분배 +{attend_pool:,} 💎")
                            st.rerun()

    # ── 오른쪽 컬럼 ──
    with col_right:
        with st.container(border=True):
            st.markdown("<div class='section-title'>📊 혈맹 자산 현황</div>", unsafe_allow_html=True)
            finance = st.session_state.db_data.get("guild_finance", {})
            st.html(
                f"<table class='guild-roster-table'><thead><tr>"
                f"<th style='text-align:left;padding-left:4px;'>항목</th>"
                f"<th style='text-align:center;'>보유 금액</th></tr></thead><tbody>"
                f"<tr><td style='text-align:left;padding-left:4px;'>💰 현재 혈비</td><td><span style='color:#f59e0b;font-weight:bold;'>{finance.get('guild_money',0):,} 💎</span></td></tr>"
                f"<tr><td style='text-align:left;padding-left:4px;'>⚔️ 투력분배금</td><td><span style='color:#4ade80;font-weight:bold;'>{finance.get('power_dist',0):,} 💎</span></td></tr>"
                f"<tr><td style='text-align:left;padding-left:4px;'>📅 참여분배금</td><td><span style='color:#c084fc;font-weight:bold;'>{finance.get('attend_dist',0):,} 💎</span></td></tr>"
                f"</tbody></table>"
            )

            # 마스터 전용: 다이아 수입 입력 및 자동 분배
            if is_master:
                st.markdown("<div style='margin-top:16px;border-top:1px solid #1e293b;padding-top:12px;'>", unsafe_allow_html=True)
                st.caption("💼 **수입 입력 및 자동 분배**")
                income = st.number_input("획득 다이아 입력 💎", min_value=0, step=1000, key="income_input")
                if income > 0:
                    admin = int(income * 0.06)
                    guild = int(income * 0.30)
                    attend_pool = int(income * 0.32)
                    power_pool = int(income * 0.32)
                    st.markdown(
                        f"<div style='background:#1c2536;border-radius:6px;padding:10px;font-size:0.8rem;margin-top:8px;'>"
                        f"🔸 총무비 (6%): <b>{admin:,} 💎</b><br>"
                        f"🔸 혈비 (30%): <b>{guild:,} 💎</b><br>"
                        f"🔸 참여 분배 (32%): <b>{attend_pool:,} 💎</b><br>"
                        f"🔸 전투력 분배 (32%): <b>{power_pool:,} 💎</b>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                    if st.button("✅ 분배 적용", use_container_width=True):
                        finance["guild_money"] = finance.get("guild_money", 0) + guild
                        finance["power_dist"] = finance.get("power_dist", 0) + power_pool
                        finance["attend_dist"] = finance.get("attend_dist", 0) + attend_pool
                        st.session_state.db_data["guild_finance"] = finance
                        with st.spinner("시트에 저장 중..."):
                            save_finance_to_sheet(finance)
                        st.success(f"✅ 분배 완료! 혈비 +{guild:,} / 투력분배 +{power_pool:,} / 참여분배 +{attend_pool:,}")
                        st.rerun()
                st.markdown("</div>", unsafe_allow_html=True)

        with st.container(border=True):
            st.markdown("<div class='section-title'>보스 출석체크</div>", unsafe_allow_html=True)
            st.caption(f"**{current_user}** 님의 오늘 레이드 참여 기록입니다.")
            st.write("")
            boss_1 = st.checkbox("👹 레기카 토벌 완료", value=member_info["attendance"].get("레기카", False))
            boss_2 = st.checkbox("⚡ 시온 토벌 완료", value=member_info["attendance"].get("시온", False))
            boss_3 = st.checkbox("🔮 플라우드 토벌 완료", value=member_info["attendance"].get("플라우드", False))
            st.write("")
            st.markdown("<div class='stButton-attend'>", unsafe_allow_html=True)
            if st.button("⚔️ 레이드 기록 저장", use_container_width=True):
                st.session_state.db_data["guildmembers"][current_user]["attendance"]["레기카"] = boss_1
                st.session_state.db_data["guildmembers"][current_user]["attendance"]["시온"] = boss_2
                st.session_state.db_data["guildmembers"][current_user]["attendance"]["플라우드"] = boss_3
                with st.spinner("저장 중..."):
                    ok = save_member_to_sheet(current_user, st.session_state.db_data["guildmembers"][current_user])
                if ok: st.success("✅ 레이드 기록 저장!")
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)

# ==========================================
# 共用工具函式與全域狀態
# ==========================================
import streamlit as st
import pandas as pd
import io
import base64
import threading
import requests
from gtts import gTTS


# ── 全域狀態 ──────────────────────────────
@st.cache_resource
def get_submit_lock():
    return threading.Lock()

@st.cache_resource
def get_global_state():
    return {"db_version": 0}

def increment_db_version():
    get_global_state()["db_version"] += 1


# ── 語音播報 ──────────────────────────────
def autoplay_audio(text: str):
    try:
        tts = gTTS(text=text, lang="zh-tw")
        fp = io.BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        b64 = base64.b64encode(fp.read()).decode()
        st.markdown(
            f"""<audio autoplay="true" style="display:none;">
            <source src="data:audio/mp3;base64,{b64}" type="audio/mp3">
            </audio>""",
            unsafe_allow_html=True,
        )
    except Exception as e:
        st.error(f"語音播報錯誤：{e}")


# ── 資料處理工具 ─────────────────────────
def format_phone(val) -> str:
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s.lower() in ("nan", "none", ""):
        return ""
    if s and not s.startswith("0"):
        return "0" + s
    return s

def convert_bool_col(df: pd.DataFrame, col: str) -> pd.DataFrame:
    mapping = {"TRUE": True, "FALSE": False, "True": True, "False": False, "1": True, "0": False}
    df[col] = df[col].replace(mapping).fillna(False).astype(bool)
    return df

def ensure_cols(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    for col in cols:
        if col not in df.columns:
            df[col] = pd.Series(dtype=object)
    return df

def get_next_station_seq(queue_df: pd.DataFrame, station: str) -> int:
    station_data = queue_df[queue_df["體驗站點"] == station].copy()
    station_data["站點序號"] = pd.to_numeric(station_data["站點序號"], errors="coerce").fillna(0)
    return int(station_data["站點序號"].max()) + 1 if not station_data.empty else 1

@st.cache_resource
def _open_spreadsheet():
    """快取 gspread Spreadsheet 物件，直接用 service account 建立連線。"""
    import gspread
    info = dict(st.secrets["connections"]["gsheets"])
    gc  = gspread.service_account_from_dict(info)
    return gc.open_by_url(info["spreadsheet"])


@st.cache_data(ttl=30, show_spinner=False)
def _fetch_from_gas(sheet_name: str, _cache_bust: int = 0) -> pd.DataFrame:
    """透過 GAS Web App 讀取工作表，所有使用者共用此快取。"""
    try:
        url = st.secrets["gas"]["read_url"]
        resp = requests.get(url, params={"sheet": sheet_name, "v": _cache_bust}, timeout=15)
        resp.raise_for_status()
        payload = resp.json()
        if "error" in payload:
            return pd.DataFrame()
        rows = payload.get("rows", [])
        headers = payload.get("headers", [])
        if not headers:
            return pd.DataFrame()
        if not rows:
            return pd.DataFrame(columns=headers)
        return pd.DataFrame(rows, columns=headers)
    except Exception:
        return pd.DataFrame()


def safe_read(conn, worksheet: str, ttl=30, default_cols: list = None) -> pd.DataFrame:
    """讀取工作表。優先使用 GAS 快取 API；寫入後 db_version 遞增可自動破快取。"""
    db_version = get_global_state()["db_version"]
    df = _fetch_from_gas(worksheet, _cache_bust=db_version)
    if df.empty and default_cols:
        df = pd.DataFrame(columns=default_cols)
    elif default_cols:
        df = ensure_cols(df, default_cols)
    return df

def fast_update_queue_status(conn, target_idx, new_status: str, full_df: pd.DataFrame) -> pd.DataFrame:
    """只更新 Queue 裡單一列的狀態欄，不重寫整張表，避免舊資料覆蓋問題。"""
    full_df.loc[target_idx, "狀態"] = new_status
    ss         = _open_spreadsheet()
    ws         = ss.worksheet("Queue")
    sheet_row  = int(target_idx) + 2          # 0-based index → 1-based + 1 header row
    status_col = list(full_df.columns).index("狀態") + 1  # 1-based column
    ws.update_cell(sheet_row, status_col, new_status)
    increment_db_version()
    _fetch_from_gas.clear()
    return full_df


# ── 欄位常數 ─────────────────────────────
REG_COLS    = ["報到序號", "姓名", "年齡", "聯繫方式", "地址", "報名項目", "有無求道", "得知管道", "報名時間", "成全進度", "gform_timestamp"]
QUEUE_COLS  = ["報到序號", "站點序號", "姓名", "體驗站點", "狀態", "報名時間"]
SET_COLS    = ["項目名稱", "老師名單", "總名額", "已報名數"]
TASK_COLS   = ["階段", "任務名稱", "負責人", "完成狀態"]
ROLE_COLS   = ["姓名", "組別", "對應儲存格"]
EQUIP_COLS  = ["器材名稱", "數量", "負責人", "取得位置", "準備狀態"]

STATUS_WAITING  = "等待中"
STATUS_SERVING  = "服務中"
STATUS_DONE     = "完成"
STATUS_MISSED   = "過號"

PROGRESS_OPTIONS = ["初次接觸", "已參加活動", "有意願", "已求道", "穩定參與", "其他"]

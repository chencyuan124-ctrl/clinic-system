# ==========================================
# 共用工具函式與全域狀態
# ==========================================
import streamlit as st
import pandas as pd
import io
import base64
import threading
import requests
import datetime
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

def append_service_log(log_rows: list):
    """將事件紀錄 append 到 ServiceLog 工作表，失敗時靜默，不阻斷主流程。"""
    if not log_rows:
        return
    try:
        ss = _open_spreadsheet()
        ws = ss.worksheet("ServiceLog")
        values = [[r.get(c, "") for c in LOG_COLS] for r in log_rows]
        ws.append_rows(values, value_input_option="USER_ENTERED")
    except Exception:
        pass


_EVENT_MAP = {
    (STATUS_WAITING, STATUS_SERVING): "CALLING",
    (STATUS_SERVING, STATUS_DONE):    "DONE",
    (STATUS_SERVING, STATUS_MISSED):  "MISSED",
    (STATUS_DONE,    STATUS_WAITING): "UNDO_DONE",
}


def fast_update_queue_status(conn, target_idx, new_status: str, full_df: pd.DataFrame) -> pd.DataFrame:
    """只更新 Queue 裡單一列的狀態欄，不重寫整張表，避免舊資料覆蓋問題。"""
    old_status = str(full_df.loc[target_idx, "狀態"])
    row_data   = full_df.loc[target_idx]

    full_df.loc[target_idx, "狀態"] = new_status
    ss         = _open_spreadsheet()
    ws         = ss.worksheet("Queue")
    sheet_row  = int(target_idx) + 2          # 0-based index → 1-based + 1 header row
    status_col = list(full_df.columns).index("狀態") + 1  # 1-based column
    ws.update_cell(sheet_row, status_col, new_status)

    now = datetime.datetime.now(_TW_TZ).strftime("%Y-%m-%d %H:%M:%S")
    append_service_log([{
        "事件時間": now,
        "報到序號": row_data.get("報到序號", ""),
        "姓名":     row_data.get("姓名", ""),
        "體驗站點": row_data.get("體驗站點", ""),
        "站點序號": row_data.get("站點序號", ""),
        "事件類型": _EVENT_MAP.get((old_status, new_status), "STATUS_CHANGE"),
        "前狀態":   old_status,
        "後狀態":   new_status,
        "備註":     "",
    }])

    increment_db_version()
    _fetch_from_gas.clear()
    return full_df


# ── 欄位常數 ─────────────────────────────
REG_COLS    = ["報到序號", "姓名", "年齡", "聯繫方式", "地址", "報名項目", "有無求道", "得知管道", "報名時間", "成全進度", "gform_timestamp"]
QUEUE_COLS  = ["報到序號", "站點序號", "姓名", "體驗站點", "狀態", "報名時間"]
SET_COLS    = ["項目名稱", "老師名單", "總名額", "已報名數", "服務人數"]
TASK_COLS   = ["階段", "任務名稱", "負責人", "完成狀態"]
ROLE_COLS   = ["姓名", "組別", "對應儲存格"]
EQUIP_COLS  = ["器材名稱", "數量", "負責人", "取得位置", "準備狀態"]

STATUS_WAITING  = "等待中"
STATUS_SERVING  = "服務中"
STATUS_DONE     = "完成"
STATUS_MISSED   = "過號"

LOG_COLS = ["事件時間", "報到序號", "姓名", "體驗站點", "站點序號", "事件類型", "前狀態", "後狀態", "備註"]
_TW_TZ   = datetime.timezone(datetime.timedelta(hours=8))

PROGRESS_OPTIONS = ["初次接觸", "已參加活動", "有意願", "已求道", "穩定參與", "其他"]


def call_next_slot(conn, station: str, teacher_count: int):
    """
    在全域 Lock 內重讀 Queue，確認服務槽有空才叫號。
    回傳 (success: bool, nxt_row | None, updated_df)
    競態保護：所有 session 共用同一個 Lock，後進者等待後重讀最新狀態。
    """
    with get_submit_lock():
        fresh = safe_read(conn, "Queue", ttl=0, default_cols=QUEUE_COLS)
        if fresh.empty:
            return False, None, fresh
        fresh["站點序號"] = pd.to_numeric(fresh["站點序號"], errors="coerce").fillna(0).astype(int)

        serving = fresh[(fresh["體驗站點"] == station) & (fresh["狀態"] == STATUS_SERVING)]
        waiting = fresh[(fresh["體驗站點"] == station) & (fresh["狀態"] == STATUS_WAITING)].sort_values("站點序號")

        if len(serving) >= teacher_count:
            return False, None, fresh
        if waiting.empty:
            return False, None, fresh

        nxt = waiting.iloc[0]
        nxt_idx = fresh[
            (fresh["站點序號"] == nxt["站點序號"]) &
            (fresh["體驗站點"] == station) &
            (fresh["狀態"] == STATUS_WAITING)
        ].index[0]
        updated = fast_update_queue_status(conn, nxt_idx, STATUS_SERVING, fresh)
        return True, nxt, updated

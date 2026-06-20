import streamlit as st
import pandas as pd
import datetime
from utils import (
    get_submit_lock, increment_db_version,
    safe_read, ensure_cols, get_next_station_seq,
    _fetch_from_gas, _open_spreadsheet,
    REG_COLS, QUEUE_COLS, SET_COLS,
    STATUS_WAITING, STATUS_SERVING, STATUS_DONE, STATUS_MISSED,
)

GFORM_SHEET_KEY = "gform_sheet_name"
GFORM_DEFAULT   = "表單回應 1"


# ─────────────────────────────────────────
# 工具函式
# ─────────────────────────────────────────
_TW_TZ = datetime.timezone(datetime.timedelta(hours=8))  # 固定 UTC+8，不依賴伺服器時區


def _today() -> str:
    """永遠回傳台灣（UTC+8）當日日期，無論伺服器時區為何。"""
    return datetime.datetime.now(_TW_TZ).strftime("%Y-%m-%d")


def _tw_date(val) -> str:
    """將 GAS 可能回傳的任何日期格式，統一轉成台灣日期 YYYY-MM-DD。
    支援：
      "2026-06-20 15:30:00"        → "2026-06-20"（台灣時間，直接截）
      "2026-06-19T23:09:13.000Z"   → "2026-06-20"（UTC→+8 轉換）
    """
    s = str(val).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return ""
    if "T" in s:
        # ISO UTC 格式："2026-06-19T23:09:13.000Z"
        try:
            clean = s[:19].replace("T", " ")          # "2026-06-19 23:09:13"
            d = datetime.datetime.strptime(clean, "%Y-%m-%d %H:%M:%S")
            return (d + datetime.timedelta(hours=8)).strftime("%Y-%m-%d")
        except Exception:
            return s[:10]
    return s[:10]  # "2026-06-20 15:30:00" → "2026-06-20"


def _norm_queue(queue_df: pd.DataFrame) -> pd.DataFrame:
    df = queue_df.copy()
    df["報到序號"] = pd.to_numeric(df["報到序號"], errors="coerce")
    return df


def get_active_count(queue_df: pd.DataFrame, serial: int, today: str) -> int:
    """今天此人的進行中（等待中／服務中）項目數"""
    if queue_df.empty:
        return 0
    q = _norm_queue(queue_df)
    person = q[q["報到序號"] == serial]
    today_rows = person[person["報名時間"].apply(_tw_date) == today]
    return len(today_rows[today_rows["狀態"].isin([STATUS_WAITING, STATUS_SERVING])])


def get_today_items(queue_df: pd.DataFrame, serial: int, today: str) -> set:
    """今天此人已有紀錄（任何狀態）的所有項目"""
    if queue_df.empty:
        return set()
    q = _norm_queue(queue_df)
    person = q[q["報到序號"] == serial]
    today_rows = person[person["報名時間"].apply(_tw_date) == today]
    return set(today_rows["體驗站點"].dropna().tolist())


def is_done(queue_df: pd.DataFrame, serial: int, item: str, today: str) -> bool:
    """此人今天該項目是否已全部完成（完成或過號）"""
    if queue_df.empty:
        return False
    q = _norm_queue(queue_df)
    rows = q[
        (q["報到序號"] == serial) &
        (q["體驗站點"] == item) &
        (q["報名時間"].apply(_tw_date) == today)
    ]
    if rows.empty:
        return False
    return all(s in (STATUS_DONE, STATUS_MISSED) for s in rows["狀態"].tolist())


# ─────────────────────────────────────────
# 自動同步 Google 表單（每 60 秒靜默執行）
# ─────────────────────────────────────────
@st.fragment(run_every=60)
def _auto_sync_gform(conn):
    sheet_name = st.session_state.get(GFORM_SHEET_KEY, GFORM_DEFAULT)
    try:
        form_df = conn.read(worksheet=sheet_name, ttl=0)
        if form_df is None or form_df.empty:
            return

        ts_col  = form_df.columns[0]   # Google 表單第一欄固定是時間戳記
        reg_df  = safe_read(conn, "Registration", ttl=0, default_cols=REG_COLS)

        existing_ts: set = set()
        if not reg_df.empty and "gform_timestamp" in reg_df.columns:
            existing_ts = set(reg_df["gform_timestamp"].dropna().astype(str).tolist())

        truly_new = form_df[~form_df[ts_col].astype(str).isin(existing_ts)]
        if truly_new.empty:
            return

        current_time = datetime.datetime.now(_TW_TZ).strftime("%Y-%m-%d %H:%M:%S")

        with get_submit_lock():
            reg_df = safe_read(conn, "Registration", ttl=0, default_cols=REG_COLS)
            if "gform_timestamp" not in reg_df.columns:
                reg_df["gform_timestamp"] = ""

            existing_ts = set(reg_df["gform_timestamp"].dropna().astype(str).tolist())
            truly_new   = form_df[~form_df[ts_col].astype(str).isin(existing_ts)]
            if truly_new.empty:
                return

            base_serial = int(pd.to_numeric(reg_df["報到序號"], errors="coerce").max()) + 1 \
                          if not reg_df.empty else 1

            add_rows = []
            for offset, (_, row) in enumerate(truly_new.iterrows()):
                name = str(row.get("姓名", "")).strip()
                if not name:
                    continue
                add_rows.append({
                    "報到序號":      base_serial + offset,
                    "姓名":          name,
                    "年齡":          row.get("年齡", ""),
                    "聯繫方式":      f"{str(row.get('聯繫方式', '')).strip()} ",
                    "地址":          str(row.get("地址", "")),
                    "報名項目":      "",
                    "有無求道":      str(row.get("有無求道", "無")),
                    "得知管道":      str(row.get("得知管道", "其他")),
                    "報名時間":      current_time,
                    "成全進度":      "初次接觸",
                    "gform_timestamp": str(row.get(ts_col, "")),
                })

            if not add_rows:
                return

            reg_df = pd.concat([reg_df, pd.DataFrame(add_rows)], ignore_index=True)
            conn.update(worksheet="Registration", data=reg_df)
            increment_db_version()
            st.toast(f"✅ 自動匯入 {len(add_rows)} 筆新表單報名！", icon="🔔")

    except Exception:
        pass   # 表單工作表尚未設定時靜默略過


# ─────────────────────────────────────────
# 主頁面：掛號分配站（工作人員）
# ─────────────────────────────────────────
def render_registration_page(conn):
    # 啟動背景自動同步（頁面開著時每 60 秒執行）
    _auto_sync_gform(conn)

    st.subheader("📋 掛號分配站（工作人員）")
    st.caption("民眾填寫 Google 表單後自動出現在下方列表，工作人員在此為他們分配體驗項目。")

    # 設定 expander
    with st.expander("⚙️ Google 表單設定", expanded=False):
        sheet_name = st.text_input(
            "表單回應工作表名稱",
            value=st.session_state.get(GFORM_SHEET_KEY, GFORM_DEFAULT),
            help="開啟 Google Sheets，表單回應的分頁名稱（預設為「表單回應 1」）",
        )
        if st.button("💾 儲存設定"):
            st.session_state[GFORM_SHEET_KEY] = sheet_name
            st.success("已儲存！")
        st.info("⚠️ 請確保該試算表的共用權限包含本系統的服務帳號可以讀取。")

    st.markdown("---")

    # 讀取資料
    today       = _today()
    reg_df      = safe_read(conn, "Registration", ttl=5,  default_cols=REG_COLS)
    queue_df    = safe_read(conn, "Queue",        ttl=5,  default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings",     ttl=10, default_cols=SET_COLS)

    if reg_df.empty:
        st.info("⏳ 尚無報名紀錄，等待民眾填寫 Google 表單…")
        return

    if settings_df.empty:
        st.warning("⚠️ 尚未設定體驗項目，請先至「體驗項目與名額設定」頁面新增。")
        return

    settings_df["總名額"]   = pd.to_numeric(settings_df["總名額"],   errors="coerce").fillna(0).astype(int)
    settings_df["已報名數"] = pd.to_numeric(settings_df["已報名數"], errors="coerce").fillna(0).astype(int)
    all_items       = settings_df["項目名稱"].tolist()
    available_items = [r["項目名稱"] for _, r in settings_df.iterrows()
                       if r["總名額"] - r["已報名數"] > 0]

    # 計算每人今日狀態
    reg_df["報到序號"] = pd.to_numeric(reg_df["報到序號"], errors="coerce")

    pending  = []   # 今天還沒有任何項目的人
    can_add  = []   # 今天有些項目完成了、還有空位的人

    for _, person in reg_df.iterrows():
        serial = person["報到序號"]
        if pd.isna(serial):
            continue
        serial       = int(serial)
        active_count = get_active_count(queue_df, serial, today)
        used_items   = get_today_items(queue_df, serial, today)
        slots        = 2 - active_count

        if slots <= 0:
            continue   # 已有 2 個進行中，跳過

        info = {
            "serial":       serial,
            "姓名":          str(person.get("姓名", "")),
            "active_count": active_count,
            "used_items":   used_items,
            "slots":        slots,
        }

        if not used_items:
            pending.append(info)
        else:
            can_add.append(info)

    col_refresh, _ = st.columns([1, 4])
    with col_refresh:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()

    tab1, tab2 = st.tabs([
        f"⏳ 待分配（{len(pending)} 人）",
        f"➕ 可加選（{len(can_add)} 人）",
    ])

    with tab1:
        if not pending:
            st.info("目前所有已報名民眾均已分配項目。")
        else:
            _render_list(conn, pending, available_items, all_items, queue_df, reg_df, settings_df, today)

    with tab2:
        if not can_add:
            st.info("目前無民眾可加選（需先完成部分項目才會出現在此）。")
        else:
            _render_list(conn, can_add, available_items, all_items, queue_df, reg_df, settings_df, today)


# ─────────────────────────────────────────
# 分配列表渲染
# ─────────────────────────────────────────
def _render_list(conn, person_list, available_items, all_items, queue_df, reg_df, settings_df, today):
    for person in person_list:
        serial   = person["serial"]
        name     = person["姓名"]
        slots    = person["slots"]
        used     = person["used_items"]

        # 此人今天已用但不能再選的項目 = 所有已出現在 queue 的項目（無論狀態）
        selectable = [i for i in available_items if i not in used]

        with st.container(border=True):
            col_info, col_form = st.columns([1, 2])

            with col_info:
                st.markdown(f"### {name}")
                st.caption(f"報到序號：{serial}")
                if used:
                    done_list   = [i for i in used if is_done(queue_df, serial, i, today)]
                    active_list = [i for i in used if not is_done(queue_df, serial, i, today)]
                    if active_list:
                        st.markdown(f"**進行中：** {'、'.join(active_list)}")
                    if done_list:
                        st.markdown(f"**已完成：** {'、'.join(done_list)}")
                st.markdown(f"**可再選：{slots} 項**")

            with col_form:
                if not selectable:
                    st.warning("所有項目今日均已體驗過，或目前全部額滿。")
                    continue

                form_key = f"form_{serial}_{today}_{len(used)}"
                with st.form(form_key):
                    chosen = st.multiselect(
                        f"為【{name}】選擇體驗項目（最多 {slots} 項）",
                        options=selectable,
                        max_selections=slots,
                    )
                    submitted = st.form_submit_button("✅ 確認分配", type="primary", use_container_width=True)

                if submitted:
                    if not chosen:
                        st.error("請至少選擇一項！")
                    else:
                        ok, msg = do_assign(conn, serial, name, chosen, today)
                        if ok:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)


# ─────────────────────────────────────────
# 執行分配（寫入 Queue / Registration / Settings）
# ─────────────────────────────────────────
def do_assign(conn, serial: int, name: str, new_items: list, today: str):
    current_time = datetime.datetime.now(_TW_TZ).strftime("%Y-%m-%d %H:%M:%S")

    with get_submit_lock():
        queue_df = safe_read(conn, "Queue",        ttl=0, default_cols=QUEUE_COLS)
        reg_df   = safe_read(conn, "Registration", ttl=0, default_cols=REG_COLS)
        live_set = safe_read(conn, "Settings",     ttl=0, default_cols=SET_COLS)
        live_set["已報名數"] = pd.to_numeric(live_set["已報名數"], errors="coerce").fillna(0).astype(int)
        live_set["總名額"]   = pd.to_numeric(live_set["總名額"],   errors="coerce").fillna(0).astype(int)

        # 再次確認：名額 & 今日同項目規則
        used_now = get_today_items(queue_df, serial, today)
        for item in new_items:
            if item in used_now:
                return False, f"⚠️【{item}】今日已報名過，無法重複！"
            row = live_set[live_set["項目名稱"] == item]
            if not row.empty:
                r = row.iloc[0]
                if int(r["總名額"]) - int(r["已報名數"]) <= 0:
                    return False, f"⚠️【{item}】剛剛已額滿，請重新選擇！"

        # 確認進行中項目數 + 新增後不超過 2
        active_now = get_active_count(queue_df, serial, today)
        if active_now + len(new_items) > 2:
            return False, f"⚠️ 此民眾目前有 {active_now} 個進行中項目，最多只能再加 {2 - active_now} 項。"

        # 取得 gspread 底層物件（避免 conn.update 的 clear+write 雙呼叫開銷）
        ss = _open_spreadsheet()

        # 寫入 Queue（append_rows：只加新列，不重寫整張表，最快）
        new_queue_rows = []
        for item in new_items:
            seq = get_next_station_seq(queue_df, item)
            new_queue_rows.append({
                "報到序號": serial, "站點序號": seq,
                "姓名": name, "體驗站點": item,
                "狀態": STATUS_WAITING, "報名時間": current_time,
            })
        append_values = [
            [r["報到序號"], r["站點序號"], r["姓名"], r["體驗站點"], r["狀態"], r["報名時間"]]
            for r in new_queue_rows
        ]
        ss.worksheet("Queue").append_rows(append_values, value_input_option="USER_ENTERED")

        # 更新 Registration 報名項目欄（單格寫入）
        reg_df["報到序號"] = pd.to_numeric(reg_df["報到序號"], errors="coerce")
        reg_df["報名項目"] = reg_df["報名項目"].fillna("").astype(str)
        t_idx = reg_df[reg_df["報到序號"] == serial].index
        if not t_idx.empty:
            old_str = reg_df.at[t_idx[0], "報名項目"].strip()
            new_str = "、".join(new_items)
            combined = (old_str + "、" + new_str) if old_str and old_str != "nan" else new_str
            sheet_row  = int(t_idx[0]) + 2  # 0-based index + 1(header) + 1(1-based)
            item_col   = list(reg_df.columns).index("報名項目") + 1
            ss.worksheet("Registration").update_cell(sheet_row, item_col, combined)

        # 更新 Settings 名額（單次 update，比 clear+write 省一個 API call）
        for item in new_items:
            idx = live_set[live_set["項目名稱"] == item].index
            if not idx.empty:
                live_set.loc[idx, "已報名數"] += 1
        settings_ws   = ss.worksheet("Settings")
        settings_data = [live_set.columns.tolist()] + live_set.fillna("").astype(str).values.tolist()
        settings_ws.update(settings_data, value_input_option="USER_ENTERED")

        increment_db_version()
        _fetch_from_gas.clear()  # 立即清除 GAS 快取，確保 rerun 後讀到最新資料

    return True, f"✅ 已為【{name}】分配：{'、'.join(new_items)}"

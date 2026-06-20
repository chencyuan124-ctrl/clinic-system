import streamlit as st
import pandas as pd
from utils import safe_read, autoplay_audio, QUEUE_COLS, SET_COLS, STATUS_SERVING, STATUS_DONE


def render_display_page(conn):
    st.markdown(
        "<h2 style='text-align:center; color:#1a1a2e; letter-spacing:2px;'>"
        "💆 身心靈保健活動 — 體驗進度總覽</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align:center; font-size:1.2em; color:#555;'>"
        "請參考下方最新進度，前往您的體驗站點</p>",
        unsafe_allow_html=True,
    )

    col_a, col_b, col_c = st.columns([3, 1, 1])
    with col_a:
        auto_refresh = st.checkbox("⚡ 啟用智慧連動（有人叫號時自動更新）", value=True)
    with col_b:
        if st.button("🔄 手動重整", use_container_width=True):
            if "cached_queue_df" in st.session_state:
                del st.session_state["cached_queue_df"]
            st.rerun()
    with col_c:
        audio_enabled = st.session_state.get("display_audio_enabled", False)
        if audio_enabled:
            if st.button("🔊 語音開啟中", use_container_width=True, type="primary"):
                st.session_state["display_audio_enabled"] = False
                st.rerun()
        else:
            if st.button("🔇 點此啟用語音", use_container_width=True):
                # 此按鈕點擊本身就是 user gesture，解鎖 Chrome 的 Autoplay Policy
                st.session_state["display_audio_enabled"] = True
                st.rerun()

    if auto_refresh:
        st.caption("✅ 自動更新中，每 5 秒從資料庫讀取最新狀態。")

    # 兩個獨立 fragment：各自擁有獨立輸出區域，互不干擾
    # audio fragment 每 10 秒刷新一次，讓 4~5 秒的語音完整播完
    _render_audio_check(conn, auto_refresh)
    _render_display_grid(conn, auto_refresh)


@st.fragment(run_every=10)
def _render_audio_check(conn, auto_refresh: bool):
    """每 10 秒偵測服務中狀態，有新叫號時直接在此 fragment 輸出 <audio>。
    display grid 每 5 秒刷新，但不會觸碰本 fragment 的輸出區，音訊不會被中斷。"""
    if not auto_refresh:
        return

    queue_df    = safe_read(conn, "Queue",    default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings", default_cols=SET_COLS)
    if queue_df.empty or settings_df.empty:
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int)
    stations = settings_df["項目名稱"].tolist()

    if not st.session_state.get("display_audio_enabled", False):
        return  # 用戶尚未點擊啟用語音，不播

    prev         = st.session_state.setdefault("prev_serving", {})
    is_first_run = "prev_serving_init" not in st.session_state

    for station in stations:
        serving = queue_df[
            (queue_df["體驗站點"] == station) & (queue_df["狀態"] == STATUS_SERVING)
        ]
        current_key = None
        if not serving.empty:
            p = serving.iloc[0]
            current_key = (int(p["站點序號"]), str(p["姓名"]), str(p.get("報名時間", "")))

        if not is_first_run and current_key and current_key != prev.get(station):
            seq, name, _ = current_key
            autoplay_audio(f"來賓 {seq} 號 {name}，{name} 請到 {station} 處報到。")
            prev[station] = current_key
            break  # 一次播一站，避免同時兩站叫號互相覆蓋

        prev[station] = current_key

    st.session_state["prev_serving_init"] = True


@st.fragment(run_every=5)
def _render_display_grid(conn, auto_refresh: bool):
    queue_df    = safe_read(conn, "Queue",    default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings", default_cols=SET_COLS)

    if settings_df.empty:
        st.info("目前無資料，請先設定體驗項目。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int)
    stations = settings_df["項目名稱"].tolist()

    cols_per_row = 3 if len(stations) <= 3 else 4
    for i in range(0, len(stations), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            if i + j >= len(stations):
                break
            station = stations[i + j]
            with cols[j]:
                _render_station_card(queue_df, station)
        if i + cols_per_row < len(stations):
            st.markdown("<div style='margin-bottom:8px'></div>", unsafe_allow_html=True)


def _render_station_card(queue_df: pd.DataFrame, station: str):
    active_q = queue_df[
        (queue_df["體驗站點"] == station) & (queue_df["狀態"] != STATUS_DONE)
    ].sort_values("站點序號")

    serving     = active_q[active_q["狀態"] == "服務中"]
    waiting_all = active_q[active_q["狀態"] == "等待中"]
    missed      = active_q[active_q["狀態"] == "過號"]

    # ── 站點標題 ──────────────────────────
    st.markdown(
        f"""<div style='background:linear-gradient(135deg,#1a1a2e,#16213e);
        color:white; border-radius:12px 12px 0 0; padding:14px 18px;
        font-size:1.3em; font-weight:bold; text-align:center; letter-spacing:1px;'>
        📍 {station}</div>""",
        unsafe_allow_html=True,
    )

    # ── 服務中卡片（支援多人同時服務）─────────
    if not serving.empty:
        serving_html = ""
        for _, p in serving.iterrows():
            serving_html += (
                f"<div style='text-align:center; padding:8px 0;'>"
                f"<div style='color:rgba(255,255,255,0.85); font-size:0.85em;'>▶ 服務中</div>"
                f"<div style='color:white; font-size:{3.0 if len(serving) == 1 else 2.2}em; font-weight:900; line-height:1.1;'>"
                f"{p['站點序號']} 號</div>"
                f"<div style='color:rgba(255,255,255,0.9); font-size:{1.6 if len(serving) == 1 else 1.2}em;'>"
                f"{p['姓名']}</div></div>"
                + ("" if _ == serving.index[-1] else "<hr style='border-color:rgba(255,255,255,0.3); margin:4px 0;'>")
            )
        st.markdown(
            f"""<div style='background:linear-gradient(135deg,#27ae60,#2ecc71);
            border-radius:0; padding:16px;'>{serving_html}</div>""",
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            """<div style='background:#f0f0f0; border-radius:0; padding:20px; text-align:center;'>
            <div style='color:#aaa; font-size:0.95em; margin-bottom:4px;'>目前空閒</div>
            <div style='color:#ccc; font-size:3.5em; font-weight:900; line-height:1.1;'>—</div>
            <div style='color:transparent; font-size:1.8em;'>-</div>
            </div>""",
            unsafe_allow_html=True,
        )

    # ── 等待名單 ──────────────────────────
    total_waiting = len(waiting_all)
    show_names    = waiting_all.head(2) if total_waiting > 3 else waiting_all
    extra_count   = total_waiting - 2 if total_waiting > 3 else 0

    wait_html = ""
    for _, w in show_names.iterrows():
        wait_html += (
            f"<div style='display:flex; justify-content:space-between; "
            f"padding:6px 0; border-bottom:1px solid #eee; font-size:1.15em;'>"
            f"<b style='color:#333;'>{w['站點序號']} 號</b>"
            f"<span style='color:#555;'>{w['姓名']}</span></div>"
        )
    if extra_count > 0:
        wait_html += (
            f"<div style='padding:6px 0; font-size:1.05em; color:#888; text-align:center;'>"
            f"⌛ 還有 {extra_count} 人等待</div>"
        )
    for _, m in missed.iterrows():
        wait_html += (
            f"<div style='display:flex; justify-content:space-between; "
            f"padding:6px 0; border-bottom:1px solid #eee; font-size:1.15em;'>"
            f"<b style='color:#e74c3c;'>{m['站點序號']} 號</b>"
            f"<span style='color:#e74c3c;'>{m['姓名']} <small>(過號)</small></span></div>"
        )
    if not wait_html:
        wait_html = "<div style='color:#aaa; padding:8px 0; font-size:1.05em;'>無等待民眾</div>"

    st.markdown(
        f"""<div style='background:white; border:1px solid #ddd; border-top:none;
        border-radius:0 0 12px 12px; padding:12px 16px;'>
        <div style='font-size:0.9em; color:#888; margin-bottom:6px; font-weight:bold;'>
        ⌛ 準備名單</div>
        {wait_html}</div>""",
        unsafe_allow_html=True,
    )
    st.markdown("<div style='margin-bottom:12px'></div>", unsafe_allow_html=True)

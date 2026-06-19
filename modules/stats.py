import streamlit as st
import pandas as pd
from utils import safe_read, QUEUE_COLS, SET_COLS, REG_COLS


def render_stats_page(conn):
    st.subheader("📊 活動統計儀表板")

    col_refresh = st.columns([5, 1])
    with col_refresh[1]:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()

    queue_df    = safe_read(conn, "Queue",        ttl=10, default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings",     ttl=10, default_cols=SET_COLS)
    reg_df      = safe_read(conn, "Registration", ttl=10, default_cols=REG_COLS)

    # ── 全站摘要 ──────────────────────────
    st.markdown("### 📌 全站摘要")
    total_reg     = len(reg_df)
    total_waiting = (queue_df["狀態"] == "等待中").sum() if not queue_df.empty else 0
    total_serving = (queue_df["狀態"] == "服務中").sum() if not queue_df.empty else 0
    total_done    = (queue_df["狀態"] == "完成").sum()   if not queue_df.empty else 0
    total_missed  = (queue_df["狀態"] == "過號").sum()   if not queue_df.empty else 0

    cols = st.columns(5)
    _card(cols[0], "👥 報名人次", total_reg,     "#4A90D9")
    _card(cols[1], "⌛ 等待中",   total_waiting, "#F0A500")
    _card(cols[2], "🟢 服務中",   total_serving, "#27AE60")
    _card(cols[3], "✅ 已完成",   total_done,    "#8E44AD")
    _card(cols[4], "⏭️ 過號",    total_missed,  "#E74C3C")

    if queue_df.empty or settings_df.empty:
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int)

    # ── 各站點分析 ─────────────────────────
    st.markdown("---")
    st.markdown("### 📍 各站點詳細統計")

    stations = settings_df["項目名稱"].tolist()
    station_data = []
    for s in stations:
        sq    = queue_df[queue_df["體驗站點"] == s]
        quota_row = settings_df[settings_df["項目名稱"] == s]
        quota = int(quota_row["總名額"].iloc[0]) if not quota_row.empty else 0
        reg_n = int(quota_row["已報名數"].iloc[0]) if not quota_row.empty else 0

        station_data.append({
            "站點":     s,
            "名額上限": quota,
            "已報名":   reg_n,
            "剩餘名額": max(quota - reg_n, 0),
            "等待中":   (sq["狀態"] == "等待中").sum(),
            "服務中":   (sq["狀態"] == "服務中").sum(),
            "已完成":   (sq["狀態"] == "完成").sum(),
            "過號":     (sq["狀態"] == "過號").sum(),
        })

    stat_df = pd.DataFrame(station_data)
    st.dataframe(stat_df, use_container_width=True, hide_index=True)

    # ── 報名來源分析 ───────────────────────
    if not reg_df.empty and "得知管道" in reg_df.columns:
        st.markdown("---")
        st.markdown("### 📣 報名來源分析")
        source_counts = reg_df["得知管道"].value_counts().reset_index()
        source_counts.columns = ["來源", "人數"]
        st.dataframe(source_counts, use_container_width=True, hide_index=True)

    # ── 成全進度分析 ───────────────────────
    if not reg_df.empty and "成全進度" in reg_df.columns:
        st.markdown("---")
        st.markdown("### 🌱 成全進度分析")
        progress_counts = reg_df["成全進度"].value_counts().reset_index()
        progress_counts.columns = ["進度", "人數"]
        st.dataframe(progress_counts, use_container_width=True, hide_index=True)

    # ── 求道狀況 ──────────────────────────
    if not reg_df.empty and "有無求道" in reg_df.columns:
        st.markdown("---")
        st.markdown("### 🙏 求道狀況")
        dao_counts = reg_df["有無求道"].value_counts().reset_index()
        dao_counts.columns = ["求道", "人數"]
        st.dataframe(dao_counts, use_container_width=True, hide_index=True)


def _card(col, label: str, value: int, color: str):
    col.markdown(
        f"""<div style='background:{color}18; border-left:5px solid {color};
        border-radius:10px; padding:16px; text-align:center;'>
        <div style='font-size:0.85em; color:#555; margin-bottom:6px;'>{label}</div>
        <div style='font-size:2.5em; font-weight:bold; color:{color};'>{value}</div>
        </div>""",
        unsafe_allow_html=True,
    )

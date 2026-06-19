import streamlit as st
import pandas as pd
from utils import safe_read, QUEUE_COLS, SET_COLS


def render_full_queue_page(conn):
    st.subheader("📋 各站點完整名單總覽")

    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.caption("顯示所有站點的完整排隊紀錄（含已完成名單）。")
    with col_b:
        if st.button("🔄 重新整理", use_container_width=True):
            st.rerun()

    queue_df    = safe_read(conn, "Queue",    ttl=0,  default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings", ttl=30, default_cols=SET_COLS)

    if queue_df.empty or settings_df.empty:
        st.info("目前無排隊資料。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int)
    stations = settings_df["項目名稱"].tolist()

    tabs = st.tabs(stations)
    for i, station in enumerate(stations):
        with tabs[i]:
            st.write(f"### 📍 {station} 完整名單")
            sq = queue_df[queue_df["體驗站點"] == station].sort_values("站點序號")
            if sq.empty:
                st.info("尚無人報名此項目。")
            else:
                st.dataframe(
                    sq[["站點序號", "報到序號", "姓名", "狀態", "報名時間"]],
                    use_container_width=True, hide_index=True,
                )

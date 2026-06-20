import streamlit as st
import pandas as pd
import datetime
import io
from utils import safe_read, format_phone, REG_COLS, PROGRESS_OPTIONS


def render_history_page(conn):
    st.subheader("🗂️ 歷史紀錄與成全進度")

    reg_df = safe_read(conn, "Registration", ttl=0, default_cols=REG_COLS)
    if reg_df.empty:
        st.info("目前尚無報名紀錄。")
        return

    display_df = reg_df.copy()
    display_df["聯繫方式"] = display_df["聯繫方式"].apply(format_phone)

    edited = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "報到序號": st.column_config.NumberColumn(disabled=True),
            "報名時間": st.column_config.TextColumn(disabled=True),
            "聯繫方式": st.column_config.TextColumn(),
            "成全進度": st.column_config.SelectboxColumn(options=PROGRESS_OPTIONS),
        },
    )

    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("💾 儲存進度", type="primary"):
            conn.update(worksheet="Registration", data=edited)
            st.success("已更新！")
    with col2:
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            edited.to_excel(w, index=False)
        out.seek(0)
        st.download_button(
            "📥 匯出完整名單",
            data=out,
            file_name=f"紀錄_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

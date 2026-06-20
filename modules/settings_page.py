import streamlit as st
import pandas as pd
from utils import safe_read, SET_COLS


_SHEET_URL  = "https://docs.google.com/spreadsheets/d/1yYWTlQ6qTOz-7R9w8tMo292oAvRP6yi8oFSugXbzXYQ"
_GAS_URL    = "https://script.google.com/macros/s/AKfycbxneDWtFzXsmGzPKVYxkRqP1bW_5-y25OWdc1VJS3eZNBHZge6NNnclvXzWgoy3diOt/exec"
_FORM_URL   = "https://forms.gle/si5W6w8dPkCLzYEb9"


def render_settings_page(conn):
    st.subheader("⚙️ 體驗項目設定概覽")

    # ── 快速連結 ──────────────────────────
    st.markdown("### 🔗 快速連結")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.link_button("📊 義診資料庫試算表", _SHEET_URL, use_container_width=True)
    with c2:
        st.link_button("⚙️ GAS 登入系統", _GAS_URL, use_container_width=True)
    with c3:
        st.link_button("📝 現場報名表單", _FORM_URL, use_container_width=True)

    st.markdown("---")
    st.info(
        "📋 **項目設定由 Google Sheets 管理**\n\n"
        "如需新增、修改或刪除體驗項目，請直接開啟 Google Sheets 的「Settings」工作表進行編輯。\n\n"
        "此頁面僅供查閱，不提供編輯功能。"
    )

    df = safe_read(conn, "Settings", ttl=30, default_cols=SET_COLS)

    if df.empty:
        st.warning("目前 Google Sheets 的「Settings」工作表尚無資料。")
        return

    df["總名額"]   = pd.to_numeric(df["總名額"],   errors="coerce").fillna(0).astype(int)
    df["已報名數"] = pd.to_numeric(df["已報名數"], errors="coerce").fillna(0).astype(int)
    df["剩餘名額"] = df["總名額"] - df["已報名數"]

    # ── 名額狀態摘要 ──────────────────────
    cols = st.columns(3)
    cols[0].metric("體驗項目數", len(df))
    cols[1].metric("總名額合計", int(df["總名額"].sum()))
    cols[2].metric("已報名合計", int(df["已報名數"].sum()))

    st.markdown("---")
    st.markdown("### 📝 項目一覽")

    display_df = df[["項目名稱", "老師名單", "總名額", "已報名數", "剩餘名額"]].copy()
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    if st.button("🔄 重新整理", use_container_width=False):
        st.rerun()

    st.markdown("---")
    st.markdown("### 🔄 新場次準備")
    st.info(
        "場次重置功能請開啟 Google Sheet 後點擊選單：\n\n"
        "**⚙️ 系統管理** → 選擇要執行的清除項目\n\n"
        "重置前請先至「**歷史紀錄**」頁面下載備份。"
    )

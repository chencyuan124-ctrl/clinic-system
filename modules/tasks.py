import streamlit as st
import pandas as pd
import datetime
import io
import openpyxl
from utils import safe_read, convert_bool_col, TASK_COLS, ROLE_COLS, EQUIP_COLS


def render_task_page(conn):
    st.subheader("📋 任務與職務管理")
    tab1, tab2, tab3 = st.tabs(["🎯 任務清單", "🧑‍🤝‍🧑 職務安排與排班表", "📦 器材清單"])

    with tab1:
        _render_tasks(conn)
    with tab2:
        _render_roles(conn)
    with tab3:
        _render_equipment(conn)


# ─────────────────────────────────────────
def _render_tasks(conn):
    st.write("### 🎯 活動任務管理")
    task_df = safe_read(conn, "Tasks", ttl=0, default_cols=TASK_COLS)
    task_df = convert_bool_col(task_df, "完成狀態")

    with st.expander("➕ 新增任務"):
        with st.form("add_task_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1: phase = st.selectbox("階段", ["活動前", "活動中", "活動後"])
            with c2: tname = st.text_input("任務名稱")
            with c3: tpic  = st.text_input("負責人")
            if st.form_submit_button("新增"):
                if tname.strip():
                    task_df = pd.concat([
                        task_df,
                        pd.DataFrame({"階段": [phase], "任務名稱": [tname], "負責人": [tpic], "完成狀態": [False]}),
                    ], ignore_index=True)
                    conn.update(worksheet="Tasks", data=task_df)
                    st.success("新增成功！")
                    st.rerun()

    edited = st.data_editor(
        task_df, num_rows="dynamic", use_container_width=True,
        column_config={
            "階段":   st.column_config.SelectboxColumn(options=["活動前", "活動中", "活動後"]),
            "完成狀態": st.column_config.CheckboxColumn("是否完成", default=False),
        },
    )
    if st.button("💾 儲存任務變更"):
        conn.update(worksheet="Tasks", data=edited)
        st.success("已儲存！")


def _render_roles(conn):
    st.write("### 🧑‍🤝‍🧑 職務安排")
    role_df = safe_read(conn, "Roles", ttl=0, default_cols=ROLE_COLS)

    with st.expander("➕ 新增人員職務"):
        with st.form("add_role_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            with c1: rname  = st.text_input("姓名")
            with c2: rgroup = st.selectbox("組別", ["祈福組", "工作人員組", "服務老師組"])
            with c3: rcell  = st.text_input("Excel 儲存格（如 A1）")
            if st.form_submit_button("新增"):
                if rname.strip():
                    role_df = pd.concat([
                        role_df,
                        pd.DataFrame({"姓名": [rname], "組別": [rgroup], "對應儲存格": [rcell]}),
                    ], ignore_index=True)
                    conn.update(worksheet="Roles", data=role_df)
                    st.success("新增成功！")
                    st.rerun()

    edited_roles = st.data_editor(
        role_df, num_rows="dynamic", use_container_width=True,
        column_config={"組別": st.column_config.SelectboxColumn(options=["祈福組", "工作人員組", "服務老師組"])},
    )
    if st.button("💾 儲存職務變更"):
        conn.update(worksheet="Roles", data=edited_roles)
        st.success("已儲存！")

    st.markdown("---")
    st.write("### 📤 一鍵套用排班表模板")
    uploaded = st.file_uploader("上傳 Excel 格式模板（.xlsx）", type=["xlsx"])
    if uploaded and st.button("✨ 產生排班表並下載", type="primary"):
        try:
            wb = openpyxl.load_workbook(uploaded)
            ws = wb.active
            for _, row in edited_roles.iterrows():
                cell = row["對應儲存格"]
                if pd.notna(cell) and str(cell).strip():
                    ws[cell] = row["姓名"]
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            st.download_button(
                "📥 下載排班表", data=out,
                file_name=f"{datetime.datetime.now().strftime('%Y%m%d')}_排班表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"匯出失敗：{e}")


def _render_equipment(conn):
    st.write("### 📦 活動器材清單")
    eq_df = safe_read(conn, "Equipment", ttl=0, default_cols=EQUIP_COLS)
    eq_df = convert_bool_col(eq_df, "準備狀態")

    with st.expander("➕ 新增器材"):
        with st.form("add_eq_form", clear_on_submit=True):
            c1, c2, c3, c4 = st.columns([2, 1, 2, 2])
            with c1: ename = st.text_input("器材名稱 *")
            with c2: eqty  = st.number_input("數量", min_value=1, value=1)
            with c3: epic  = st.text_input("負責人")
            with c4: eloc  = st.text_input("取得位置")
            if st.form_submit_button("新增"):
                if ename.strip():
                    eq_df = pd.concat([
                        eq_df,
                        pd.DataFrame({"器材名稱": [ename], "數量": [eqty], "負責人": [epic], "取得位置": [eloc], "準備狀態": [False]}),
                    ], ignore_index=True)
                    conn.update(worksheet="Equipment", data=eq_df)
                    st.success("新增成功！")
                    st.rerun()

    edited_eq = st.data_editor(
        eq_df, num_rows="dynamic", use_container_width=True,
        column_config={
            "數量":   st.column_config.NumberColumn(min_value=1, step=1),
            "準備狀態": st.column_config.CheckboxColumn("是否已準備", default=False),
        },
    )
    if st.button("💾 儲存器材變更"):
        conn.update(worksheet="Equipment", data=edited_eq)
        st.success("已儲存！")

# ==========================================
# 身心靈保健活動系統 — 主程式入口
# ==========================================
import streamlit as st
from streamlit_gsheets import GSheetsConnection

from modules.registration import render_registration_page
from modules.calling       import render_calling_page
from modules.display       import render_display_page
from modules.settings_page import render_settings_page
from modules.tasks         import render_task_page
from modules.history       import render_history_page
from modules.full_queue    import render_full_queue_page
from modules.stats         import render_stats_page
from modules.public_form   import render_public_form
from modules.self_select   import render_self_select_page

st.set_page_config(page_title="身心靈保健活動系統", page_icon="💆", layout="wide")


def _check_password(pwd: str) -> bool:
    try:
        return pwd == st.secrets["admin"]["password"]
    except Exception:
        return pwd == "10151015"  # fallback when secrets not configured


def main():
    st.markdown(
        "<h3 style='color:#888; margin-top:-20px;'>💆 身心靈保健活動系統</h3>",
        unsafe_allow_html=True,
    )

    conn     = st.connection("gsheets", type=GSheetsConnection)
    is_admin = st.session_state.get("is_admin", False)

    # ── 頁面定義 ──────────────────────────
    def page_display():       render_display_page(conn)
    def page_public_form():   render_public_form(conn)
    def page_self_select():   render_self_select_page(conn)
    def page_register():      render_registration_page(conn)
    def page_calling():      render_calling_page(conn)
    def page_full_queue():   render_full_queue_page(conn)
    def page_history():      render_history_page(conn)
    def page_settings():     render_settings_page(conn)
    def page_task():         render_task_page(conn)
    def page_stats():        render_stats_page(conn)

    # ── 導覽 ──────────────────────────────
    if hasattr(st, "navigation"):
        pages = {
            "📺 顯示專區": [st.Page(page_display,     title="民眾體驗顯示螢幕（大螢幕）")],
            "✋ 自助選項": [st.Page(page_self_select, title="自助選取體驗項目（民眾）")],
        }
        if is_admin:
            pages["📋 掛號分配（後台）"] = [
                st.Page(page_public_form, title="現場掃碼報名"),
                st.Page(page_register,    title="掛號分配站"),
            ]
            pages["📢 叫號與紀錄（後台）"] = [
                st.Page(page_calling,    title="排隊清單與叫號操作"),
                st.Page(page_full_queue, title="各站點完整名單總覽"),
                st.Page(page_history,    title="歷史紀錄與成全進度"),
            ]
            pages["📊 統計與管理（後台）"] = [
                st.Page(page_stats,    title="活動統計儀表板"),
                st.Page(page_settings, title="體驗項目與名額設定"),
                st.Page(page_task,     title="任務與職務管理"),
            ]
        pg = st.navigation(pages)
        pg.run()

    else:
        # Streamlit < 1.30 fallback sidebar navigation
        st.sidebar.markdown("### 🗂️ 系統導覽")
        menu = ["　└ 民眾體驗顯示螢幕", "　└ 自助選取體驗項目"]
        if is_admin:
            menu = [
                "　└ 民眾體驗顯示螢幕", "　└ 自助選取體驗項目",
                "　└ 現場掃碼報名", "　└ 掛號分配站",
                "　└ 排隊清單與叫號操作", "　└ 各站點完整名單總覽", "　└ 歷史紀錄與成全進度",
                "　└ 活動統計儀表板", "　└ 體驗項目與名額設定", "　└ 任務與職務管理",
            ]
        choice = st.sidebar.radio("請選擇頁面：", menu, label_visibility="collapsed")
        dispatch = {
            "　└ 民眾體驗顯示螢幕":   page_display,
            "　└ 現場掃碼報名":       page_public_form,
            "　└ 自助選取體驗項目":   page_self_select,
            "　└ 掛號分配站":         page_register,
            "　└ 排隊清單與叫號操作": page_calling,
            "　└ 各站點完整名單總覽": page_full_queue,
            "　└ 歷史紀錄與成全進度": page_history,
            "　└ 活動統計儀表板":     page_stats,
            "　└ 體驗項目與名額設定": page_settings,
            "　└ 任務與職務管理":     page_task,
        }
        dispatch.get(choice, page_display)()

    # ── 側邊欄登入 / 登出 ─────────────────
    st.sidebar.markdown("---")
    if not is_admin:
        with st.sidebar.expander("🔐 工作人員入口", expanded=False):
            with st.form("login_form"):
                pwd          = st.text_input("請輸入密碼", type="password")
                submit_login = st.form_submit_button("確認登入", use_container_width=True)
                if submit_login:
                    if _check_password(pwd):
                        st.session_state["is_admin"] = True
                        st.rerun()
                    else:
                        st.error("密碼錯誤")
    else:
        st.sidebar.success("✅ 管理員已登入")
        if st.sidebar.button("🚪 登出", use_container_width=True):
            st.session_state["is_admin"] = False
            st.rerun()


if __name__ == "__main__":
    main()

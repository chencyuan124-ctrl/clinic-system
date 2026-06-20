import streamlit as st
import requests
import datetime


def _get_gas_url() -> str:
    """從 secrets.toml 讀取 GAS Web App 網址"""
    try:
        return st.secrets["gas"]["registration_url"]
    except Exception:
        return ""


def _submit_to_gas(gas_url: str, payload: dict) -> tuple[bool, str]:
    """將報名資料 POST 至 GAS Web App，回傳 (成功與否, 訊息)"""
    try:
        resp = requests.post(
            gas_url,
            json=payload,
            timeout=15,
            headers={"Content-Type": "application/json"},
        )
        result = resp.json()
        if result.get("status") == "ok":
            return True, ""
        return False, result.get("message", "未知錯誤")
    except requests.exceptions.Timeout:
        return False, "連線逾時，請稍後再試。"
    except Exception as e:
        return False, str(e)


def render_public_form(conn=None):   # conn 保留參數相容性，此頁不使用
    # ── 頂部品牌區 ──────────────────────────────────────
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #4CAF50 0%, #2196F3 100%);
            border-radius: 16px;
            padding: 28px 24px 20px 24px;
            text-align: center;
            margin-bottom: 24px;
        ">
            <div style="font-size: 2.8em;">💆</div>
            <h2 style="color:white; margin:8px 0 4px 0; font-size:1.6em;">身心靈保健活動</h2>
            <p style="color:rgba(255,255,255,0.88); margin:0; font-size:1.05em;">現場掃碼報名</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    gas_url = _get_gas_url()

    # ── 未設定 GAS 時的提示（僅管理員看得到）──────────
    if not gas_url:
        st.warning(
            "⚠️ 尚未設定 Google Apps Script 網址。\n\n"
            "請依照 `gas_registration_handler.js` 的說明部署 GAS Web App，"
            "並將網址填入 `.streamlit/secrets.toml`：\n"
            "```toml\n[gas]\nregistration_url = \"https://script.google.com/...\"\n```"
        )
        return

    # ── 送出成功畫面 ─────────────────────────────────
    if "form_success" in st.session_state:
        st.success(
            "🎉 **報名成功！**\n\n"
            "請至工作人員處，告知您的姓名，工作人員將為您安排體驗項目。"
        )
        st.balloons()
        st.markdown("---")
        st.info("您可以關閉此頁面，或繼續為其他人報名。")
        if st.button("繼續為其他人報名", use_container_width=True):
            del st.session_state["form_success"]
            st.rerun()
        return

    # ── 報名表單 ─────────────────────────────────────
    with st.form("public_registration_form", clear_on_submit=False):
        st.markdown("#### 📝 請填寫基本資料")

        col1, col2 = st.columns(2)
        with col1:
            name  = st.text_input("姓名 ＊", placeholder="請輸入全名")
            phone = st.text_input("聯繫方式 ＊", placeholder="手機或市話")
        with col2:
            age     = st.number_input("年齡", min_value=1, max_value=120, value=30)
            address = st.text_input("居住地址", placeholder="選填")

        dao_status = st.radio(
            "是否有參加過求道開智慧的活動？",
            ["無", "有"],
            horizontal=True,
        )
        source = st.selectbox(
            "從哪裡得知此活動？",
            ["親友介紹", "FB 社團宣傳", "DM 海報", "其他"],
        )

        st.markdown(
            "<p style='color:#888; font-size:0.9em; margin-top:8px;'>"
            "＊ 為必填欄位。送出後請至工作人員處確認體驗項目。</p>",
            unsafe_allow_html=True,
        )

        submitted = st.form_submit_button(
            "✅ 確認送出報名",
            type="primary",
            use_container_width=True,
        )

    # ── 送出處理 ─────────────────────────────────────
    if submitted:
        if not name.strip() or not phone.strip():
            st.error("⚠️ 姓名與聯繫方式為必填欄位！")
            return

        with st.spinner("正在送出報名資料，請稍候…"):
            payload = {
                "姓名":    name.strip(),
                "聯繫方式": phone.strip(),
                "年齡":    int(age),
                "地址":    address.strip(),
                "有無求道": dao_status,
                "得知管道": source,
            }
            ok, err_msg = _submit_to_gas(gas_url, payload)

        if ok:
            st.session_state["form_success"] = True
            st.rerun()
        else:
            st.error(f"送出失敗，請重試或請工作人員協助。\n\n（錯誤：{err_msg}）")

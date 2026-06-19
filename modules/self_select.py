import streamlit as st
import pandas as pd
import datetime
from utils import safe_read, format_phone, REG_COLS, QUEUE_COLS, SET_COLS
from modules.registration import get_active_count, get_today_items, do_assign


def _today() -> str:
    return datetime.date.today().strftime("%Y-%m-%d")


def _normalize_phone(val: str) -> str:
    """統一正規化手機號碼：去空白、去-.、去.0後綴、補開頭0"""
    s = str(val).strip().replace("-", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    if s and not s.startswith("0") and len(s) == 9:
        s = "0" + s
    return s


def _verify(reg_df: pd.DataFrame, name: str, phone: str):
    """回傳符合姓名+手機的第一筆 Registration 紀錄，找不到回傳 None"""
    if reg_df.empty:
        return None
    name  = name.strip()
    phone = _normalize_phone(phone)
    mask = (
        reg_df["姓名"].astype(str).str.strip() == name
    ) & (
        reg_df["聯繫方式"].apply(_normalize_phone) == phone
    )
    matched = reg_df[mask]
    return matched.iloc[0] if not matched.empty else None


def render_self_select_page(conn):
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            border-radius: 16px;
            padding: 28px 24px 20px 24px;
            text-align: center;
            margin-bottom: 24px;
        ">
            <div style="font-size: 2.8em;">✋</div>
            <h2 style="color:white; margin:8px 0 4px 0; font-size:1.6em;">自助選取體驗項目</h2>
            <p style="color:rgba(255,255,255,0.75); margin:0; font-size:1.05em;">
                請輸入報名時填寫的姓名與手機號碼
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    step = st.session_state.get("self_select_step", "verify")

    if step == "success":
        _render_success()
        return

    if step == "select":
        _render_select(conn)
        return

    _render_verify(conn)


# ─────────────────────────────────────────
# Step 1：身分驗證
# ─────────────────────────────────────────
def _render_verify(conn):
    reg_df = safe_read(conn, "Registration", ttl=5, default_cols=REG_COLS)
    names  = sorted(reg_df["姓名"].astype(str).str.strip().unique().tolist()) if not reg_df.empty else []

    with st.form("self_verify_form"):
        st.markdown("#### 📋 身分確認")
        if names:
            name = st.selectbox("姓名 ＊", options=["（請選擇）"] + names)
        else:
            name = st.text_input("姓名 ＊", placeholder="請輸入您的全名")
        phone = st.text_input("手機號碼 ＊", placeholder="報名時填寫的手機號碼")
        submitted = st.form_submit_button("🔍 查詢並繼續", type="primary", use_container_width=True)

    if submitted:
        if not name or name == "（請選擇）" or not phone.strip():
            st.error("⚠️ 請選擇姓名並填寫手機號碼。")
            return

        person = _verify(reg_df, name.strip(), phone.strip())
        if person is None:
            st.error("❌ 查無符合紀錄，請確認姓名與手機號碼是否與報名時填寫的一致。")
            return

        serial = int(pd.to_numeric(person["報到序號"], errors="coerce"))
        st.session_state["self_select_serial"] = serial
        st.session_state["self_select_name"]   = str(person["姓名"]).strip()
        st.session_state["self_select_step"]   = "select"
        st.rerun()


# ─────────────────────────────────────────
# Step 2：選取項目
# ─────────────────────────────────────────
def _render_select(conn):
    serial = st.session_state.get("self_select_serial")
    name   = st.session_state.get("self_select_name", "")
    today  = _today()

    st.success(f"✅ 驗證成功！歡迎 **{name}**")
    st.markdown("---")

    queue_df    = safe_read(conn, "Queue",    ttl=5,  default_cols=QUEUE_COLS)
    settings_df = safe_read(conn, "Settings", ttl=60, default_cols=SET_COLS)

    if settings_df.empty:
        st.warning("⚠️ 目前尚未設定體驗項目，請稍後再試。")
        _back_btn()
        return

    settings_df["總名額"]   = pd.to_numeric(settings_df["總名額"],   errors="coerce").fillna(0).astype(int)
    settings_df["已報名數"] = pd.to_numeric(settings_df["已報名數"], errors="coerce").fillna(0).astype(int)

    active_count = get_active_count(queue_df, serial, today)
    used_items   = get_today_items(queue_df, serial, today)
    slots        = 2 - active_count

    # 顯示目前狀態
    if used_items:
        done_items   = [i for i in used_items if _item_is_done(queue_df, serial, i, today)]
        active_items = [i for i in used_items if i not in done_items]
        if active_items:
            st.info(f"**進行中項目：** {'、'.join(active_items)}")
        if done_items:
            st.info(f"**已完成項目：** {'、'.join(done_items)}")

    if slots <= 0:
        st.warning("⚠️ 您目前已有 2 個進行中的項目，請完成後再回來加選。")
        _back_btn()
        return

    st.markdown(f"#### 🎯 請選擇體驗項目（最多可選 **{slots}** 項）")
    st.caption("已額滿的項目無法選取；今日已體驗過的項目不重複顯示。")

    selectable = []
    full_items = []
    for _, row in settings_df.iterrows():
        item = str(row["項目名稱"])
        if item in used_items:
            continue  # 今日已選過
        if int(row["總名額"]) - int(row["已報名數"]) > 0:
            selectable.append(item)
        else:
            full_items.append(item)

    if not selectable and not full_items:
        st.info("目前所有項目均已體驗過，感謝您的參與！")
        _back_btn()
        return

    # 額滿提示
    if full_items:
        st.markdown(
            "<div style='color:#888; font-size:0.9em; margin-bottom:8px;'>"
            f"以下項目目前額滿（無法選取）：{'、'.join(full_items)}</div>",
            unsafe_allow_html=True,
        )

    if not selectable:
        st.warning("目前所有可用項目均已額滿，請詢問工作人員。")
        _back_btn()
        return

    chosen = st.multiselect(
        "選擇項目",
        options=selectable,
        max_selections=slots,
        placeholder="請點選您想體驗的項目…",
    )

    if st.button("✅ 確認送出", type="primary", use_container_width=True, disabled=not chosen):
        ok, msg = do_assign(conn, serial, name, chosen, today)
        if ok:
            st.session_state["self_select_chosen"] = chosen
            st.session_state["self_select_step"]   = "success"
            st.rerun()
        else:
            st.error(msg)

    _back_btn()


# ─────────────────────────────────────────
# Step 3：成功畫面
# ─────────────────────────────────────────
def _render_success():
    name   = st.session_state.get("self_select_name", "")
    chosen = st.session_state.get("self_select_chosen", [])
    st.success(f"🎉 **已成功排入等候！**")
    st.balloons()
    st.markdown(
        f"**{name}** 已報名以下體驗項目：\n\n"
        + "\n".join(f"- 📍 {item}" for item in chosen)
        + "\n\n請至各站點等候叫號，或查看顯示螢幕確認號碼。"
    )
    st.markdown("---")
    if st.button("🔄 換人操作", use_container_width=True):
        for k in ["self_select_step", "self_select_serial", "self_select_name", "self_select_chosen"]:
            st.session_state.pop(k, None)
        st.rerun()


def _back_btn():
    if st.button("← 重新輸入", use_container_width=False):
        for k in ["self_select_step", "self_select_serial", "self_select_name"]:
            st.session_state.pop(k, None)
        st.rerun()


def _item_is_done(queue_df, serial, item, today):
    if queue_df.empty:
        return False
    q = queue_df.copy()
    q["報到序號"] = pd.to_numeric(q["報到序號"], errors="coerce")
    rows = q[
        (q["報到序號"] == serial) &
        (q["體驗站點"] == item) &
        (q["報名時間"].astype(str).str.startswith(today, na=False))
    ]
    if rows.empty:
        return False
    return all(s in ("完成", "過號") for s in rows["狀態"].tolist())

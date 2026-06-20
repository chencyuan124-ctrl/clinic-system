import streamlit as st
import pandas as pd
from utils import (
    fast_update_queue_status, safe_read, autoplay_audio, call_next_slot,
    QUEUE_COLS, SET_COLS,
    STATUS_WAITING, STATUS_SERVING, STATUS_DONE, STATUS_MISSED,
)

_STATION_PWD = "123"


def _get_teacher_count(settings_df: pd.DataFrame, station: str) -> int:
    row = settings_df[settings_df["項目名稱"] == station]
    if row.empty:
        return 1
    val = pd.to_numeric(row.iloc[0].get("服務人數", 1), errors="coerce")
    return max(1, int(val) if not pd.isna(val) else 1)


def render_station_page(conn):
    st.subheader("🏥 站點服務台（老師操作）")

    # ── 語音在 fragment 外播放 ──
    if "station_audio" in st.session_state:
        autoplay_audio(st.session_state.pop("station_audio"))

    # ── 密碼解鎖 ──────────────────────────
    if not st.session_state.get("station_unlocked"):
        st.info("請輸入站點密碼以繼續操作。")
        with st.form("station_login"):
            pwd = st.text_input("站點密碼", type="password")
            if st.form_submit_button("確認", use_container_width=True, type="primary"):
                if pwd == _STATION_PWD:
                    st.session_state["station_unlocked"] = True
                    st.rerun()
                else:
                    st.error("❌ 密碼錯誤")
        return

    # ── 站點選擇 ──────────────────────────
    settings_df = safe_read(conn, "Settings", ttl=30, default_cols=SET_COLS)
    if settings_df.empty:
        st.info("請先至「體驗項目設定」頁面新增站點。")
        return

    station_options = settings_df["項目名稱"].tolist()
    last = st.session_state.get("station_last", station_options[0])
    default_idx = station_options.index(last) if last in station_options else 0
    current_station = st.selectbox("選擇您的服務站點：", station_options, index=default_idx)
    st.session_state["station_last"] = current_station

    teacher_count = _get_teacher_count(settings_df, current_station)
    if teacher_count > 1:
        st.caption(f"此站點有 {teacher_count} 個服務槽，可同時服務 {teacher_count} 位民眾。")

    _render_station_fragment(conn, current_station, teacher_count)

    st.markdown("---")
    if st.button("🔒 登出站點服務台", use_container_width=False):
        st.session_state["station_unlocked"] = False
        st.rerun()


@st.fragment
def _render_station_fragment(conn, current_station: str, teacher_count: int):
    queue_df = safe_read(conn, "Queue", ttl=0, default_cols=QUEUE_COLS)

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int) if not queue_df.empty else queue_df["站點序號"]
    mask       = (queue_df["體驗站點"] == current_station) & (queue_df["狀態"] != STATUS_DONE) if not queue_df.empty else pd.Series(dtype=bool)
    station_q  = queue_df[mask].sort_values("站點序號") if not queue_df.empty else pd.DataFrame(columns=QUEUE_COLS)
    serving_df = station_q[station_q["狀態"] == STATUS_SERVING]
    waiting_df = station_q[station_q["狀態"] == STATUS_WAITING]

    slots_used = len(serving_df)
    slots_free = max(0, teacher_count - slots_used)

    st.markdown(f"### 📍 【{current_station}】目前狀態")
    if teacher_count > 1:
        st.markdown(
            f"<div style='background:#f0f7ff; border-radius:8px; padding:8px 14px; margin-bottom:12px; font-size:0.95em;'>"
            f"服務槽：<b>{slots_used} / {teacher_count}</b> 使用中　空閒槽：<b>{slots_free}</b></div>",
            unsafe_allow_html=True,
        )

    # ── 所有服務中的人 ─────────────────────
    if not serving_df.empty:
        cols = st.columns(min(len(serving_df), 3))
        for i, (_, p) in enumerate(serving_df.iterrows()):
            with cols[i % len(cols)]:
                st.markdown(
                    f"""<div style='background:linear-gradient(135deg,#27ae60,#2ecc71);
                    border-radius:12px; padding:18px; text-align:center; margin-bottom:12px;'>
                    <div style='color:rgba(255,255,255,0.85); font-size:0.9em;'>▶ 服務中</div>
                    <div style='color:white; font-size:3em; font-weight:900; line-height:1.1;'>{int(p['站點序號'])} 號</div>
                    <div style='color:rgba(255,255,255,0.95); font-size:1.5em;'>{p['姓名']}</div>
                    </div>""",
                    unsafe_allow_html=True,
                )
    else:
        st.info("💡 目前無人體驗中。")

    # ── 下一位提示 ────────────────────────
    if not waiting_df.empty:
        nxt = waiting_df.iloc[0]
        st.caption(f"⌛ 下一位等待：第 {int(nxt['站點序號'])} 號 — {nxt['姓名']}")
    else:
        st.caption("目前無候補名單。")

    st.markdown("---")

    # ── 操作按鈕區 ───────────────────────
    # 「完成並叫下一位」或「叫下一位」
    if not serving_df.empty:
        # 多服務槽：用 selectbox 選擇要完成哪一位
        if len(serving_df) > 1:
            options = [f"第{int(r['站點序號'])}號 {r['姓名']}" for _, r in serving_df.iterrows()]
            sel = st.selectbox("選擇要標記完成的人：", options, key=f"done_sel_{current_station}")
            sel_seq = int(sel.split("號")[0].replace("第", ""))
        else:
            sel_seq = int(serving_df.iloc[0]["站點序號"])

        col1, col2 = st.columns(2)

        with col1:
            if st.button("✅ 完成，自動叫下一位", type="primary", use_container_width=True):
                done_idx = queue_df[
                    (queue_df["站點序號"] == sel_seq) &
                    (queue_df["體驗站點"] == current_station) &
                    (queue_df["狀態"] == STATUS_SERVING)
                ].index[0]
                fast_update_queue_status(conn, done_idx, STATUS_DONE, queue_df)

                # 完成後重新確認槽數，若有空則叫下一位（Lock 保護）
                ok, nxt, _ = call_next_slot(conn, current_station, teacher_count)
                if ok and nxt is not None:
                    st.session_state["station_audio"] = (
                        f"來賓 {int(nxt['站點序號'])} 號 {nxt['姓名']}，"
                        f"{nxt['姓名']} 請到 {current_station} 處報到。"
                    )
                st.rerun()

        with col2:
            confirm_key = f"station_confirm_missed_{current_station}"
            if st.session_state.get(confirm_key):
                if st.button("❗ 確認過號", use_container_width=True, type="primary"):
                    missed_idx = queue_df[
                        (queue_df["站點序號"] == sel_seq) &
                        (queue_df["體驗站點"] == current_station) &
                        (queue_df["狀態"] == STATUS_SERVING)
                    ].index[0]
                    fast_update_queue_status(conn, missed_idx, STATUS_MISSED, queue_df)
                    st.session_state[confirm_key] = False

                    ok, nxt, _ = call_next_slot(conn, current_station, teacher_count)
                    if ok and nxt is not None:
                        st.session_state["station_audio"] = (
                            f"來賓 {int(nxt['站點序號'])} 號 {nxt['姓名']}，"
                            f"{nxt['姓名']} 請到 {current_station} 處報到。"
                        )
                    st.rerun()
            else:
                if st.button("⏭️ 過號（未到站）", use_container_width=True):
                    st.session_state[confirm_key] = True
                    st.rerun(scope="fragment")

    else:
        # 無人服務中：只有「叫下一位」按鈕
        if not waiting_df.empty:
            if st.button("🔊 呼叫下一位", type="primary", use_container_width=True):
                ok, nxt, _ = call_next_slot(conn, current_station, teacher_count)
                if ok and nxt is not None:
                    st.session_state["station_audio"] = (
                        f"來賓 {int(nxt['站點序號'])} 號 {nxt['姓名']}，"
                        f"{nxt['姓名']} 請到 {current_station} 處報到。"
                    )
                    st.rerun()
                else:
                    st.warning("⚠️ 所有服務槽目前已滿，請稍候。")
        else:
            st.button("目前無候補，等待下一位報到", use_container_width=True, disabled=True)

    # 若還有空槽且有人等待，提示可以再叫
    if slots_free > 0 and not waiting_df.empty and not serving_df.empty:
        if st.button(f"➕ 呼叫下一位（還有 {slots_free} 個空槽）", use_container_width=True):
            ok, nxt, _ = call_next_slot(conn, current_station, teacher_count)
            if ok and nxt is not None:
                st.session_state["station_audio"] = (
                    f"來賓 {int(nxt['站點序號'])} 號 {nxt['姓名']}，"
                    f"{nxt['姓名']} 請到 {current_station} 處報到。"
                )
                st.rerun()
            else:
                st.warning("⚠️ 叫號失敗，可能剛被其他老師搶先呼叫。")

    if st.button("🔄 手動刷新", use_container_width=False):
        st.rerun(scope="fragment")

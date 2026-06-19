import streamlit as st
import pandas as pd
import datetime
from utils import (
    fast_update_queue_status, safe_read, autoplay_audio,
    _open_spreadsheet, _fetch_from_gas, increment_db_version,
    QUEUE_COLS, SET_COLS,
    STATUS_WAITING, STATUS_SERVING, STATUS_DONE, STATUS_MISSED,
)


def render_calling_page(conn):
    st.subheader("📢 叫號操作台")

    # 語音在 fragment 外播放：fragment rerun 不會移除外層頁面的 <audio> 元素
    if "calling_audio" in st.session_state:
        autoplay_audio(st.session_state.pop("calling_audio"))

    settings_df = safe_read(conn, "Settings", ttl=30, default_cols=SET_COLS)
    if settings_df.empty:
        st.info("請先至「體驗項目設定」頁面新增項目。")
        return

    station_options = settings_df["項目名稱"].tolist()
    last_station = st.session_state.get("last_station", station_options[0])
    default_idx  = station_options.index(last_station) if last_station in station_options else 0
    current_station = st.selectbox("請選擇您負責的服務站點：", station_options, index=default_idx)
    st.session_state["last_station"] = current_station

    _render_calling_fragment(conn, current_station)
    st.markdown("---")
    _render_stats_bar(conn, station_options)
    _render_overview(conn, station_options)


@st.fragment
def _render_stats_bar(conn, station_options: list):
    """頁面頂部統計摘要列（每 10 秒自動刷新）"""
    queue_df = safe_read(conn, "Queue", ttl=10, default_cols=QUEUE_COLS)
    if queue_df.empty:
        return

    total_waiting  = (queue_df["狀態"] == STATUS_WAITING).sum()
    total_serving  = (queue_df["狀態"] == STATUS_SERVING).sum()
    total_done     = (queue_df["狀態"] == STATUS_DONE).sum()
    total_missed   = (queue_df["狀態"] == STATUS_MISSED).sum()
    total_reg      = queue_df["報到序號"].nunique()

    cols = st.columns(5)
    _stat_card(cols[0], "👥 報名人次", total_reg,     "#4A90D9")
    _stat_card(cols[1], "⌛ 等待中",   total_waiting, "#F0A500")
    _stat_card(cols[2], "🟢 服務中",   total_serving, "#27AE60")
    _stat_card(cols[3], "✅ 已完成",   total_done,    "#8E44AD")
    _stat_card(cols[4], "⏭️ 過號",    total_missed,  "#E74C3C")


def _stat_card(col, label: str, value: int, color: str):
    col.markdown(
        f"""<div style='background:{color}15; border-left:4px solid {color};
        border-radius:8px; padding:12px 16px; text-align:center;'>
        <div style='font-size:0.85em; color:#555; margin-bottom:4px;'>{label}</div>
        <div style='font-size:2.2em; font-weight:bold; color:{color};'>{value}</div>
        </div>""",
        unsafe_allow_html=True,
    )


@st.fragment
def _render_calling_fragment(conn, current_station: str):

    queue_df = safe_read(conn, "Queue", ttl=0, default_cols=QUEUE_COLS)
    if queue_df.empty:
        st.info("目前無人排隊。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors="coerce").fillna(0).astype(int)
    mask         = (queue_df["體驗站點"] == current_station) & (queue_df["狀態"] != STATUS_DONE)
    display_q    = queue_df[mask].sort_values("站點序號").copy()
    serving_df   = display_q[display_q["狀態"] == STATUS_SERVING]
    waiting_df   = display_q[display_q["狀態"] == STATUS_WAITING]
    missed_df    = display_q[display_q["狀態"] == STATUS_MISSED]

    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.write(f"### 📍 【{current_station}】待處理名單")
    with col_h2:
        if st.button("🔄 手動刷新", use_container_width=True):
            st.rerun(scope="fragment")

    # 服務中橫幅
    if not serving_df.empty:
        p = serving_df.iloc[0]
        st.markdown(
            f"""<div style='background:#d4edda; border:2px solid #28a745; border-radius:10px;
            padding:12px 20px; margin-bottom:12px;'>
            <span style='color:#155724; font-size:1.1em; font-weight:bold;'>
            👨‍⚕️ 服務中：第 {p['站點序號']} 號 — {p['姓名']}</span></div>""",
            unsafe_allow_html=True,
        )
    else:
        st.info("💡 目前無人體驗，請點擊「呼叫下一位」。")

    if not display_q.empty:
        st.dataframe(
            display_q[["站點序號", "報到序號", "姓名", "狀態", "報名時間"]],
            use_container_width=True, hide_index=True,
        )

    # ── 主要叫號按鈕 ──────────────────────────
    st.markdown("---")
    st.write("### 🎛️ 叫號操作區")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        if st.button("🔊 呼叫下一位", type="primary", use_container_width=True):
            if not serving_df.empty:
                st.warning("⚠️ 請先將目前服務中的民眾標記完成或過號。")
            elif not waiting_df.empty:
                nxt = waiting_df.iloc[0]
                idx = queue_df[
                    (queue_df["站點序號"] == nxt["站點序號"]) &
                    (queue_df["體驗站點"] == current_station) &
                    (queue_df["狀態"] == STATUS_WAITING)
                ].index[0]
                queue_df = fast_update_queue_status(conn, idx, STATUS_SERVING, queue_df)
                st.session_state["calling_audio"] = f"來賓 {int(nxt['站點序號'])} 號 {nxt['姓名']}，{nxt['姓名']} 請到 {current_station} 處報到。"
                st.rerun()  # full rerun → 外層頁面播音，fragment 更新資料
            else:
                st.info("目前沒有等待中的民眾。")

    with c2:
        if st.button("📢 再次呼叫", use_container_width=True):
            if not serving_df.empty:
                p = serving_df.iloc[0]
                # 更新 Queue 的報名時間 → 觸發顯示螢幕偵測到變化後重新播音
                s_idx = queue_df[
                    (queue_df["站點序號"] == p["站點序號"]) &
                    (queue_df["體驗站點"] == current_station) &
                    (queue_df["狀態"] == STATUS_SERVING)
                ].index[0]
                new_ts   = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ss       = _open_spreadsheet()
                ws       = ss.worksheet("Queue")
                time_col = list(queue_df.columns).index("報名時間") + 1
                ws.update_cell(int(s_idx) + 2, time_col, new_ts)
                increment_db_version()
                _fetch_from_gas.clear()
                st.session_state["calling_audio"] = f"來賓 {int(p['站點序號'])} 號 {p['姓名']}，{p['姓名']} 請到 {current_station} 處報到。"
                st.rerun()  # full rerun → 外層頁面播音
            else:
                st.warning("目前無服務中民眾。")

    with c3:
        if not serving_df.empty:
            p = serving_df.iloc[0]
            confirm_key = f"confirm_missed_{current_station}"
            if st.session_state.get(confirm_key):
                if st.button("❗ 確認過號", use_container_width=True, type="primary"):
                    idx = queue_df[
                        (queue_df["站點序號"] == p["站點序號"]) &
                        (queue_df["體驗站點"] == current_station) &
                        (queue_df["狀態"] == STATUS_SERVING)
                    ].index[0]
                    fast_update_queue_status(conn, idx, STATUS_MISSED, queue_df)
                    st.session_state[confirm_key] = False
                    st.rerun(scope="fragment")
            else:
                if st.button("⏭️ 標記過號", use_container_width=True):
                    st.session_state[confirm_key] = True
                    st.rerun(scope="fragment")
        else:
            st.button("⏭️ 標記過號", use_container_width=True, disabled=True)

    with c4:
        if not serving_df.empty:
            p = serving_df.iloc[0]
            confirm_key = f"confirm_done_{current_station}"
            if st.session_state.get(confirm_key):
                if st.button("❗ 確認完成", use_container_width=True, type="primary"):
                    idx = queue_df[
                        (queue_df["站點序號"] == p["站點序號"]) &
                        (queue_df["體驗站點"] == current_station) &
                        (queue_df["狀態"] == STATUS_SERVING)
                    ].index[0]
                    fast_update_queue_status(conn, idx, STATUS_DONE, queue_df)
                    st.session_state[confirm_key] = False
                    st.rerun(scope="fragment")
            else:
                if st.button("✅ 標記完成", use_container_width=True):
                    st.session_state[confirm_key] = True
                    st.rerun(scope="fragment")
        else:
            st.button("✅ 標記完成", use_container_width=True, disabled=True)

    # ── 過號重叫 ──────────────────────────────
    st.markdown("---")
    st.write("### 🔁 過號重叫區")
    if not missed_df.empty:
        waiting_sorted = waiting_df.sort_values("站點序號")
        max_pos        = len(waiting_sorted) + 1   # 最多插到末尾

        with st.form("recall_form"):
            options = [f"第{int(r['站點序號'])}號 — {r['姓名']}" for _, r in missed_df.iterrows()]
            sel = st.selectbox("選擇要重叫的過號民眾：", options)
            pos = st.number_input(
                f"插入順位（1 = 下一位，最大 {max_pos}）",
                min_value=1, max_value=max_pos, value=1, step=1,
            )
            recall_submitted = st.form_submit_button("🔊 重新呼叫", use_container_width=True)

        if recall_submitted:
            seq  = int(sel.split("號")[0].replace("第", ""))
            name = sel.split("— ")[-1]
            t_idx = queue_df[
                (queue_df["站點序號"] == seq) &
                (queue_df["體驗站點"] == current_station)
            ].index[0]
            pos = int(pos)

            if serving_df.empty and max_pos == 1:
                # 完全沒人等待也沒人服務：直接呼叫
                fast_update_queue_status(conn, t_idx, STATUS_SERVING, queue_df)
            else:
                # 按指定順位插入等待佇列，重排此站序號
                base_seq = int(serving_df.iloc[0]["站點序號"]) if not serving_df.empty else 0

                # 目前等待清單（已排除過號者），按順序排列
                wait_indices = waiting_sorted.index.tolist()

                # 把過號者插入指定位置（1-based → 0-based）
                wait_indices.insert(pos - 1, t_idx)

                # 重新指派此站等待者的序號
                for rank, idx in enumerate(wait_indices):
                    queue_df.at[idx, "站點序號"] = base_seq + rank + 1
                    queue_df.at[idx, "狀態"]     = STATUS_WAITING

                ss   = _open_spreadsheet()
                ws   = ss.worksheet("Queue")
                data = [queue_df.columns.tolist()] + queue_df.fillna("").astype(str).values.tolist()
                ws.update(data, value_input_option="USER_ENTERED")
                increment_db_version()
                _fetch_from_gas.clear()
                st.toast(f"✅ {name} 已排入第 {pos} 順位", icon="🔁")

            st.rerun(scope="fragment")
    else:
        st.info("目前無過號名單。")

    # ── 誤觸還原 ──────────────────────────────
    st.markdown("---")
    with st.expander("🔙 誤觸完成還原區"):
        done_q = queue_df[
            (queue_df["體驗站點"] == current_station) & (queue_df["狀態"] == STATUS_DONE)
        ]
        if not done_q.empty:
            sel_undo = st.selectbox(
                "選擇要還原的人員",
                [f"{r['站點序號']}號 {r['姓名']}" for _, r in done_q.iterrows()],
            )
            if st.button("還原為等待中"):
                seq   = int(sel_undo.split("號")[0])
                u_idx = queue_df[
                    (queue_df["站點序號"] == seq) &
                    (queue_df["體驗站點"] == current_station) &
                    (queue_df["狀態"] == STATUS_DONE)
                ].index[0]
                fast_update_queue_status(conn, u_idx, STATUS_WAITING, queue_df)
                st.rerun(scope="fragment")
        else:
            st.caption("目前無已完成名單。")


@st.fragment(run_every=10)
def _render_overview(conn, station_options: list):
    with st.expander("📊 全站即時概覽", expanded=False):
        queue_df = safe_read(conn, "Queue", ttl=0, default_cols=QUEUE_COLS)
        if queue_df.empty:
            st.caption("目前無排隊資料。")
            return
        rows = []
        for station in station_options:
            sq = queue_df[queue_df["體驗站點"] == station]
            rows.append({
                "站點":    station,
                "等待中":  int((sq["狀態"] == STATUS_WAITING).sum()),
                "服務中":  int((sq["狀態"] == STATUS_SERVING).sum()),
                "已完成":  int((sq["狀態"] == STATUS_DONE).sum()),
                "過號":    int((sq["狀態"] == STATUS_MISSED).sum()),
            })
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

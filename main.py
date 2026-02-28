# ==========================================
# 載入套件與基本設定
# ==========================================
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
from gtts import gTTS
import io
import base64
import openpyxl
import time
import hashlib
import threading

st.set_page_config(page_title="身心靈保健活動系統", page_icon="💆", layout="wide")

# ==========================================
# 共用輔助函式 (鎖定機制、語音播放、狀態更新)
# ==========================================
# 【新增】：報名專用的互斥鎖，確保同一時間只有一人能寫入資料
@st.cache_resource
def get_submit_lock():
    return threading.Lock()

def autoplay_audio(text):
    try:
        tts = gTTS(text=text, lang='zh-tw')
        fp = io.BytesIO()
        tts.write_to_fp(fp)
        fp.seek(0)
        b64 = base64.b64encode(fp.read()).decode()
        md = f"""
            <audio autoplay="true" style="display:none;">
            <source src="data:audio/mp3;base64,{b64}" type="audio/mp3">
            </audio>
            """
        st.markdown(md, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"語音播報發生錯誤: {e}")

def fast_update_queue_status(conn, target_idx, new_status, full_df):
    full_df.loc[target_idx, "狀態"] = new_status
    conn.update(worksheet="Queue", data=full_df)
    return full_df

# ==========================================
# 模組 1：民眾報名專區 (前台 - 加入執行緒鎖防干涉)
# ==========================================
def render_registration_page(conn):
    st.subheader("📝 民眾報名專區")
    
    if "reg_form_key" not in st.session_state: st.session_state["reg_form_key"] = 0
    if "add_form_key" not in st.session_state: st.session_state["add_form_key"] = 0
    
    if "reg_success_msg" in st.session_state:
        st.success(st.session_state["reg_success_msg"])
        st.balloons()
        del st.session_state["reg_success_msg"]

    try:
        settings_df = conn.read(worksheet="Settings", ttl=5)
    except Exception:
        st.warning("⚠️ 目前尚未設定任何體驗項目。")
        return

    if settings_df.empty:
        st.warning("⚠️ 目前尚未設定任何體驗項目。")
        return

    available_options = [row["項目名稱"] for idx, row in settings_df.iterrows() if int(row["總名額"]) - int(row["已報名數"]) > 0]
    full_options = [row["項目名稱"] for idx, row in settings_df.iterrows() if int(row["總名額"]) - int(row["已報名數"]) <= 0]

    if full_options: st.info(f"💡 溫馨提示：以下項目已額滿 - {', '.join(full_options)}")

    mode = st.radio("請選擇您的報名身份：", ["🆕 報名服務項目", "🔄 已做完兩項，加選服務項目"], horizontal=True)

    if mode == "🆕 報名服務項目":
        with st.form(f"registration_form_{st.session_state['reg_form_key']}", clear_on_submit=False):
            col1, col2 = st.columns(2)
            with col1:
                name = st.text_input("姓名 *", placeholder="請輸入全名")
                age = st.number_input("年齡 *", min_value=1, max_value=120, value=30)
                phone = st.text_input("聯繫方式 *", placeholder="手機號碼或室內電話")
            with col2:
                address = st.text_input("地址", placeholder="請輸入居住區域")
                dao_status = st.radio("請問您是否有求過道？", ["無", "有"], horizontal=True)
                source = st.selectbox("從哪裡得知活動訊息？", ["親友介紹", "網路宣傳", "DM海報", "佛堂公告", "其他"])
            
            selected_items = st.multiselect("請選擇想體驗的項目 (最多選擇 2 項) *", options=available_options, max_selections=2)
            submit_button = st.form_submit_button("確認送出報名", type="primary")

            if submit_button:
                if not name.strip() or not phone.strip() or not selected_items:
                    st.error("⚠️ 姓名、聯繫方式、以及體驗項目為必填欄位！")
                else:
                    st.toast("🔄 正在為您處理，請稍候...")
                    
                    # 【核心防護】：取得鎖定，確保排隊寫入
                    with get_submit_lock():
                        try:
                            # 鎖定後才讀取最新資料，確保拿到的序號是絕對準確的
                            reg_df = conn.read(worksheet="Registration", ttl=0)
                            queue_df = conn.read(worksheet="Queue", ttl=0)
                            latest_settings_df = conn.read(worksheet="Settings", ttl=0)
                        except Exception:
                            reg_df, queue_df, latest_settings_df = pd.DataFrame(), pd.DataFrame(), settings_df

                        for col in ["報到序號", "姓名", "年齡", "聯繫方式", "地址", "報名項目", "有無求道", "得知管道", "報名時間", "成全進度"]:
                            if col not in reg_df.columns: reg_df[col] = pd.Series(dtype=object)
                        for col in ["報到序號", "站點序號", "姓名", "體驗站點", "狀態", "報名時間"]:
                            if col not in queue_df.columns: queue_df[col] = pd.Series(dtype=object)

                        new_serial = len(reg_df) + 1
                        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        safe_phone = f"{phone} "

                        new_reg = pd.DataFrame({
                            "報到序號": [new_serial], "姓名": [name], "年齡": [age], "聯繫方式": [safe_phone],
                            "地址": [address], "報名項目": ["、".join(selected_items)], "有無求道": [dao_status],
                            "得知管道": [source], "報名時間": [current_time], "成全進度": ["初次接觸"]
                        })
                        reg_df = pd.concat([reg_df, new_reg], ignore_index=True)
                        conn.update(worksheet="Registration", data=reg_df)

                        new_queue_rows = []
                        for item in selected_items:
                            station_data = queue_df[queue_df["體驗站點"] == item]
                            station_data["站點序號"] = pd.to_numeric(station_data["站點序號"], errors='coerce').fillna(0)
                            max_seq = station_data["站點序號"].max() if not station_data.empty else 0
                            current_station_seq = int(max_seq) + 1
                            
                            new_queue_rows.append({
                                "報到序號": new_serial, "站點序號": current_station_seq,
                                "姓名": name, "體驗站點": item, "狀態": "等待中", "報名時間": current_time
                            })
                        
                        queue_df = pd.concat([queue_df, pd.DataFrame(new_queue_rows)], ignore_index=True)
                        conn.update(worksheet="Queue", data=queue_df)

                        for item in selected_items:
                            item_idx = latest_settings_df[latest_settings_df["項目名稱"] == item].index
                            if not item_idx.empty: latest_settings_df.loc[item_idx, "已報名數"] += 1
                        conn.update(worksheet="Settings", data=latest_settings_df)

                    st.session_state["reg_success_msg"] = f"🎉 報名成功！您的總報到序號為：【 {new_serial} 】號"
                    st.session_state["reg_form_key"] += 1
                    st.rerun()

    elif mode == "🔄 已做完兩項，加選服務項目":
        try:
            reg_df = conn.read(worksheet="Registration", ttl=0)
            queue_df = conn.read(worksheet="Queue", ttl=0)
        except Exception:
            st.warning("目前尚無名單。")
            return
            
        if reg_df.empty:
            st.warning("目前尚無人報名，無法加選。")
            return

        name_list = reg_df["姓名"].dropna().unique().tolist()
        old_name = st.selectbox("請選擇您的姓名", [""] + name_list)
        
        if old_name:
            user_queues = queue_df[queue_df["姓名"] == old_name]
            unfinished = user_queues[user_queues["狀態"] != "完成"]
            
            if not unfinished.empty:
                unfinished_items = unfinished["體驗站點"].tolist()
                st.error(f"⚠️ 系統檢查到您還有尚未完成的項目：【{', '.join(unfinished_items)}】\n\n請先將目前的項目體驗「完成」後再來加選喔！")
            else:
                done_items = user_queues["體驗站點"].dropna().tolist()
                st.info(f"✅ 您已經完成的項目：{', '.join(done_items)}")
                new_available = [x for x in available_options if x not in done_items]
                
                with st.form(f"add_more_form_{st.session_state['add_form_key']}", clear_on_submit=False):
                    new_items = st.multiselect("請選擇想加選的體驗項目 (最多 2 項)", new_available, max_selections=2)
                    submit_add = st.form_submit_button("確認加選", type="primary")
                    
                    if submit_add:
                        if not new_items:
                            st.error("請至少選擇一項！")
                        else:
                            st.toast("🔄 正在為您處理，請稍候...")
                            
                            # 【核心防護】：加選區也同樣加上鎖定機制
                            with get_submit_lock():
                                # 在鎖定區內重新讀取，確保不蓋掉別人的資料
                                reg_df = conn.read(worksheet="Registration", ttl=0)
                                queue_df = conn.read(worksheet="Queue", ttl=0)
                                latest_settings_df = conn.read(worksheet="Settings", ttl=0)
                                
                                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                orig_serial = reg_df[reg_df["姓名"] == old_name].iloc[0]["報到序號"]
                                
                                new_queue_rows = []
                                for item in new_items:
                                    station_data = queue_df[queue_df["體驗站點"] == item]
                                    station_data["站點序號"] = pd.to_numeric(station_data["站點序號"], errors='coerce').fillna(0)
                                    max_seq = station_data["站點序號"].max() if not station_data.empty else 0
                                    current_station_seq = int(max_seq) + 1
                                    
                                    new_queue_rows.append({
                                        "報到序號": orig_serial, "站點序號": current_station_seq,
                                        "姓名": old_name, "體驗站點": item, "狀態": "等待中", "報名時間": current_time
                                    })
                                queue_df = pd.concat([queue_df, pd.DataFrame(new_queue_rows)], ignore_index=True)
                                conn.update(worksheet="Queue", data=queue_df)
                                
                                target_idx = reg_df[reg_df["姓名"] == old_name].index[0]
                                old_str = str(reg_df.loc[target_idx, "報名項目"])
                                reg_df.loc[target_idx, "報名項目"] = old_str + "、" + "、".join(new_items)
                                conn.update(worksheet="Registration", data=reg_df)
                                
                                for item in new_items:
                                    item_idx = latest_settings_df[latest_settings_df["項目名稱"] == item].index
                                    if not item_idx.empty: latest_settings_df.loc[item_idx, "已報名數"] += 1
                                conn.update(worksheet="Settings", data=latest_settings_df)
                                
                            st.session_state["reg_success_msg"] = "🎉 加選成功！請留意叫號廣播。"
                            st.session_state["add_form_key"] += 1
                            st.rerun()

    elif mode == "🔄 已做完兩項，加選服務項目":
        try:
            reg_df = conn.read(worksheet="Registration", ttl=0)
            queue_df = conn.read(worksheet="Queue", ttl=0)
        except Exception:
            st.warning("目前尚無名單。")
            return
            
        if reg_df.empty:
            st.warning("目前尚無人報名，無法加選。")
            return

        name_list = reg_df["姓名"].dropna().unique().tolist()
        old_name = st.selectbox("請選擇您的姓名", [""] + name_list)
        
        if old_name:
            user_queues = queue_df[queue_df["姓名"] == old_name]
            unfinished = user_queues[user_queues["狀態"] != "完成"]
            
            if not unfinished.empty:
                unfinished_items = unfinished["體驗站點"].tolist()
                st.error(f"⚠️ 系統檢查到您還有尚未完成的項目：【{', '.join(unfinished_items)}】\n\n請先將目前的項目體驗「完成」後再來加選喔！")
            else:
                done_items = user_queues["體驗站點"].dropna().tolist()
                st.info(f"✅ 您已經完成的項目：{', '.join(done_items)}")
                new_available = [x for x in available_options if x not in done_items]
                
                with st.form(f"add_more_form_{st.session_state['add_form_key']}", clear_on_submit=False):
                    new_items = st.multiselect("請選擇想加選的體驗項目 (最多 2 項)", new_available, max_selections=2)
                    submit_add = st.form_submit_button("確認加選", type="primary")
                    
                    if submit_add:
                        if not new_items:
                            st.error("請至少選擇一項！")
                        else:
                            st.toast("🔄 正在為您處理，請稍候...")
                            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            orig_serial = reg_df[reg_df["姓名"] == old_name].iloc[0]["報到序號"]
                            
                            new_queue_rows = []
                            for item in new_items:
                                station_data = queue_df[queue_df["體驗站點"] == item]
                                station_data["站點序號"] = pd.to_numeric(station_data["站點序號"], errors='coerce').fillna(0)
                                max_seq = station_data["站點序號"].max() if not station_data.empty else 0
                                current_station_seq = int(max_seq) + 1
                                
                                new_queue_rows.append({
                                    "報到序號": orig_serial, "站點序號": current_station_seq,
                                    "姓名": old_name, "體驗站點": item, "狀態": "等待中", "報名時間": current_time
                                })
                            queue_df = pd.concat([queue_df, pd.DataFrame(new_queue_rows)], ignore_index=True)
                            conn.update(worksheet="Queue", data=queue_df)
                            
                            target_idx = reg_df[reg_df["姓名"] == old_name].index[0]
                            old_str = str(reg_df.loc[target_idx, "報名項目"])
                            reg_df.loc[target_idx, "報名項目"] = old_str + "、" + "、".join(new_items)
                            conn.update(worksheet="Registration", data=reg_df)
                            
                            latest_settings_df = conn.read(worksheet="Settings", ttl=0)
                            for item in new_items:
                                item_idx = latest_settings_df[latest_settings_df["項目名稱"] == item].index
                                if not item_idx.empty: latest_settings_df.loc[item_idx, "已報名數"] += 1
                            conn.update(worksheet="Settings", data=latest_settings_df)
                            
                            st.session_state["reg_success_msg"] = "🎉 加選成功！請留意叫號廣播。"
                            st.session_state["add_form_key"] += 1
                            st.rerun()

# ==========================================
# 模組 3：排隊清單與叫號操作 (後台)
# ==========================================
@st.fragment
def render_calling_station_fragment(conn, current_station):
    if "pending_audio" in st.session_state:
        autoplay_audio(st.session_state["pending_audio"])
        st.success(f"📢 正在播報：{st.session_state['pending_audio']}")
        del st.session_state["pending_audio"]

    try:
        # 讀取全表以避免覆蓋遺失已完成資料
        queue_df = conn.read(worksheet="Queue", ttl=0)
    except Exception:
        st.warning("無法讀取排隊資料。")
        return

    if queue_df.empty:
        st.info("目前無人排隊。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors='coerce').fillna(0).astype(int)
    
    # 僅顯示未完成名單
    mask_active = (queue_df["體驗站點"] == current_station) & (queue_df["狀態"] != "完成")
    display_queue = queue_df[mask_active].sort_values(by="站點序號").copy()

    col_header1, col_header2 = st.columns([3, 1])
    with col_header1:
        st.write(f"### 📍 【{current_station}】待處理名單")
    with col_header2:
        if st.button("🔄 手動刷新名單", use_container_width=True):
            st.rerun(scope="fragment")

    serving_df = display_queue[display_queue["狀態"] == "服務中"]
    
    if display_queue.empty:
        st.info("目前尚無排隊名單。")
    else:
        if not serving_df.empty:
            serving_person = serving_df.iloc[0]
            st.success(f"👨‍⚕️ **目前服務中：** 第 {serving_person['站點序號']} 號 - {serving_person['姓名']}")
        else:
            st.info("💡 目前無人體驗，請點擊「呼叫下一位」。")
            
        st.dataframe(display_queue[["站點序號", "報到序號", "姓名", "狀態", "報名時間"]], use_container_width=True)

    st.markdown("---")
    st.write("### 🎛️ 叫號操作區")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("🔊 呼叫下一位", type="primary", use_container_width=True):
            if not serving_df.empty:
                st.warning("⚠️ 請先將目前服務中的民眾標記為完成或過號。")
            else:
                waiting_df = display_queue[display_queue["狀態"] == "等待中"]
                if not waiting_df.empty:
                    next_p = waiting_df.iloc[0]
                    idx = queue_df[(queue_df["站點序號"] == next_p["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                    queue_df.loc[idx, "狀態"] = "服務中"
                    conn.update(worksheet="Queue", data=queue_df)
                    
                    st.session_state["pending_audio"] = f"來賓 {next_p['站點序號']} 號 {next_p['姓名']}，{next_p['姓名']} 請到 {current_station} 處報到。"
                    st.rerun(scope="fragment")
                else:
                    st.info("沒有等待中的民眾了！")
                    
    with col2:
        if st.button("📢 再次呼叫當前", use_container_width=True):
            if not serving_df.empty:
                p = serving_df.iloc[0]
                autoplay_audio(f"來賓 {p['站點序號']} 號 {p['姓名']}，{p['姓名']} 請到 {current_station} 處報到。")
            else:
                st.warning("⚠️ 目前無服務中民眾。")

    with col3:
        if st.button("⏭️ 標記為「過號」", use_container_width=True):
            if not serving_df.empty:
                p = serving_df.iloc[0]
                idx = queue_df[(queue_df["站點序號"] == p["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                queue_df.loc[idx, "狀態"] = "過號"
                conn.update(worksheet="Queue", data=queue_df)
                st.rerun(scope="fragment")

    with col4:
        if st.button("✅ 標記為「完成」", use_container_width=True):
            if not serving_df.empty:
                p = serving_df.iloc[0]
                idx = queue_df[(queue_df["站點序號"] == p["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                queue_df.loc[idx, "狀態"] = "完成" 
                conn.update(worksheet="Queue", data=queue_df) 
                st.success(f"{p['姓名']} 已完成體驗。")
                st.rerun(scope="fragment")

    with st.expander("🔙 誤觸完成還原區"):
        done_df = queue_df[(queue_df["體驗站點"] == current_station) & (queue_df["狀態"] == "完成")]
        if not done_df.empty:
            selected_undo = st.selectbox("選擇要還原的人員", [f"{r['站點序號']}號 {r['姓名']}" for i, r in done_df.iterrows()])
            if st.button("還原為等待中"):
                u_seq = int(selected_undo.split("號")[0])
                u_idx = queue_df[(queue_df["站點序號"] == u_seq) & (queue_df["體驗站點"] == current_station)].index[0]
                queue_df.loc[u_idx, "狀態"] = "等待中"
                conn.update(worksheet="Queue", data=queue_df)
                st.rerun(scope="fragment")
        else:
            st.caption("目前無已完成名單。")

def render_calling_page(conn):
    st.subheader("📢 叫號操作台")
    try:
        settings_df = conn.read(worksheet="Settings", ttl=30)
    except Exception:
        st.warning("目前尚無設定資料。")
        return
        
    if settings_df.empty: 
        st.info("請先至設定頁面新增項目。")
        return
    station_options = settings_df["項目名稱"].tolist()
    current_station = st.selectbox("請選擇您負責的服務站點：", station_options)
    
    render_calling_station_fragment(conn, current_station)

# ==========================================
# 模組 5：民眾體驗顯示螢幕 (大螢幕 - 純手動版)
# ==========================================
def render_display_grid(conn):
    try:
        try:
            queue_df = conn.query('SELECT * FROM "Queue" WHERE "狀態" != \'完成\'', ttl=0)
        except Exception:
            queue_df = conn.read(worksheet="Queue", ttl=0)
        settings_df = conn.read(worksheet="Settings", ttl=60)
    except Exception:
        st.warning("無法讀取資料庫。")
        return
        
    if settings_df.empty or queue_df.empty:
        st.info("目前無資料。")
        return

    current_hash = hashlib.md5(pd.util.hash_pandas_object(queue_df).values).hexdigest()
    if st.session_state.get("last_display_hash") == current_hash:
        pass 
    st.session_state["last_display_hash"] = current_hash

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors='coerce').fillna(0).astype(int)
    stations = settings_df["項目名稱"].tolist()
    
    cols_per_row = 4
    for i in range(0, len(stations), cols_per_row):
        cols = st.columns(cols_per_row)
        for j in range(cols_per_row):
            if i + j < len(stations):
                station = stations[i + j]
                with cols[j]:
                    st.markdown(f"<h2 style='text-align:center; color:#1f77b4; background-color:#e9ecef; border-radius:10px; padding:10px;'>📍 {station}</h2>", unsafe_allow_html=True)
                    station_q = queue_df[(queue_df["體驗站點"] == station) & (queue_df["狀態"] != "完成")].sort_values(by="站點序號")
                    
                    serving = station_q[station_q["狀態"] == "服務中"]
                    if not serving.empty:
                        s_person = serving.iloc[0]
                        st.markdown(f"""
                        <div style='background-color:#d4edda; padding:20px; border-radius:15px; border: 3px solid #28a745; margin-bottom: 15px;'>
                            <h3 style='color:#155724; text-align:center; margin:0;'>服務中</h3>
                            <div style='text-align:center; font-size: 4em; font-weight:bold; color:#155724; line-height:1.2;'>{s_person['站點序號']} 號</div>
                            <div style='text-align:center; font-size: 2em; color:#155724;'>{s_person['姓名']}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div style='background-color:#f8f9fa; padding:20px; border-radius:15px; border: 3px dashed #ced4da; margin-bottom: 15px;'>
                            <h3 style='color:#6c757d; text-align:center; margin:0;'>目前空閒中</h3>
                            <div style='text-align:center; font-size: 4em; font-weight:bold; color:transparent; line-height:1.2;'>-</div>
                            <div style='text-align:center; font-size: 2em; color:transparent;'>-</div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                    st.markdown("<p style='font-size:1.4em; font-weight:bold; margin-bottom:5px; border-bottom: 2px solid #ddd;'>⌛ 準備名單：</p>", unsafe_allow_html=True)
                    
                    waiting = station_q[station_q["狀態"] == "等待中"].head(4) 
                    missed = station_q[station_q["狀態"] == "過號"]
                    
                    if not waiting.empty or not missed.empty:
                        for idx, w_person in waiting.iterrows():
                            st.markdown(f"<div style='font-size:1.4em; padding:4px 0; color:#333;'><b>{w_person['站點序號']}</b>號 {w_person['姓名']}</div>", unsafe_allow_html=True)
                        for idx, m_person in missed.iterrows():
                            st.markdown(f"<div style='font-size:1.4em; padding:4px 0; color:#dc3545;'><b>{m_person['站點序號']}</b>號 {m_person['姓名']} <span style='font-size:0.8em; font-weight:bold;'>(過號)</span></div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div style='font-size:1.3em; color:#6c757d; padding:4px 0;'>無等待民眾</div>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 2em 0; border: 1px solid #eee;'>", unsafe_allow_html=True)

def render_display_page(conn):
    st.markdown("<h3 style='text-align: center; color: #1f77b4; margin-bottom: 0;'>📺 體驗進度總覽</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 1.2em; color: #555;'>請參考下方最新進度前往體驗</p>", unsafe_allow_html=True)
    
    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.caption("⏸️ 目前設定為：純手動更新")
    with col_b:
        if st.button("🔄 手動重新整理", use_container_width=True):
            st.rerun()
    
    render_display_grid(conn)

# ==========================================
# 模組 2：體驗項目與名額設定 (後台)
# ==========================================
def render_settings_page(conn):
    st.subheader("⚙️ 體驗項目與名額設定 (後台)")
    try:
        df = conn.read(worksheet="Settings", ttl=0)
    except Exception:
        df = pd.DataFrame(columns=["項目名稱", "老師名單", "總名額", "已報名數"])

    for col in ["項目名稱", "老師名單", "總名額", "已報名數"]:
        if col not in df.columns: df[col] = pd.Series(dtype=object)

    try:
        roles_df = conn.read(worksheet="Roles", ttl=0)
        teacher_list = roles_df[roles_df["組別"] == "服務老師組"]["姓名"].dropna().unique().tolist()
    except Exception:
        teacher_list = []

    if not teacher_list: st.warning("⚠️ 系統中目前沒有「服務老師組」的名單，請先至職務管理新增。")

    col1, col2 = st.columns(2)
    with col1:
        with st.expander("➕ 新增體驗項目", expanded=False):
            with st.form("add_item_form", clear_on_submit=True):
                new_item = st.text_input("項目名稱")
                new_teachers = st.multiselect("老師名單", options=teacher_list)
                new_quota = st.number_input("總名額", min_value=1, value=20)
                if st.form_submit_button("確認新增"):
                    if not new_item.strip(): st.error("請輸入項目名稱！")
                    elif new_item in df["項目名稱"].values: st.error("此項目已存在！")
                    else:
                        teachers_str = "、".join(new_teachers)
                        new_row = pd.DataFrame({"項目名稱": [new_item], "老師名單": [teachers_str], "總名額": [new_quota], "已報名數": [0]})
                        df = pd.concat([df, new_row], ignore_index=True)
                        conn.update(worksheet="Settings", data=df)
                        st.success(f"已成功新增【{new_item}】！")
                        st.rerun()

    with col2:
        if not df.empty:
            with st.expander("✏️ 編輯現有項目", expanded=False):
                edit_target = st.selectbox("請選擇要修改的項目", df["項目名稱"].tolist())
                target_idx = df[df["項目名稱"] == edit_target].index[0]
                current_row = df.loc[target_idx]
                
                with st.form("edit_item_form"):
                    edit_teachers = st.multiselect("重新勾選老師名單", options=teacher_list, default=[t for t in str(current_row["老師名單"]).split("、") if t in teacher_list])
                    edit_quota = st.number_input("修改總名額", min_value=1, value=int(current_row["總名額"]))
                    if st.form_submit_button("儲存修改"):
                        df.loc[target_idx, "老師名單"] = "、".join(edit_teachers)
                        df.loc[target_idx, "總名額"] = edit_quota
                        conn.update(worksheet="Settings", data=df)
                        st.success(f"【{edit_target}】已更新！")
                        st.rerun()

    st.write("### 📝 項目總覽表")
    edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)
    if st.button("💾 儲存表格變更 (含刪除項目)", key="save_settings_table"):
        conn.update(worksheet="Settings", data=edited_df)
        st.success("總覽表已更新！")
        st.rerun()

    st.markdown("---")
    st.write("### 🚨 危險區域：名額重整")
    with st.container(border=True):
        st.warning("⚠️ 此操作會將「所有項目」的已報名人數重設為 0，請謹慎使用。")
        confirm_reset = st.checkbox("我確定要將所有項目的報名人數歸零")
        if confirm_reset:
            if st.button("🔥 立即將所有報名數歸零", type="primary", use_container_width=True):
                df["已報名數"] = 0
                conn.update(worksheet="Settings", data=df)
                st.success("✅ 所有項目的報名人數已成功歸零！")
                time.sleep(1)
                st.rerun()

# ==========================================
# 模組 4：任務與職務管理 (後台)
# ==========================================
def render_task_page(conn):
    st.subheader("📋 任務與職務管理")
    tab1, tab2, tab3 = st.tabs(["🎯 任務清單管理", "🧑‍🤝‍🧑 職務安排與模板設定", "📦 器材清單管理"])
    
    with tab1:
        st.write("### 🎯 前/中/後任務管理")
        try: task_df = conn.read(worksheet="Tasks", ttl=0)
        except Exception: task_df = pd.DataFrame(columns=["階段", "任務名稱", "負責人", "完成狀態"])
        for col in ["階段", "任務名稱", "負責人", "完成狀態"]:
            if col not in task_df.columns: task_df[col] = pd.Series(dtype=object)
        task_df["完成狀態"] = task_df["完成狀態"].replace({'TRUE': True, 'FALSE': False, 'True': True, 'False': False, '1': True, '0': False}).fillna(False).astype(bool)

        with st.expander("➕ 新增任務", expanded=False):
            with st.form("add_task_form", clear_on_submit=True):
                col_a, col_b, col_c = st.columns(3)
                with col_a: t_phase = st.selectbox("執行階段", ["活動前", "活動中", "活動後"])
                with col_b: t_name = st.text_input("任務名稱")
                with col_c: t_pic = st.text_input("負責人")
                if st.form_submit_button("新增任務"):
                    if t_name.strip():
                        task_df = pd.concat([task_df, pd.DataFrame({"階段": [t_phase], "任務名稱": [t_name], "負責人": [t_pic], "完成狀態": [False]})], ignore_index=True)
                        conn.update(worksheet="Tasks", data=task_df)
                        st.success("任務新增成功！")
                        st.rerun()

        edited_tasks = st.data_editor(task_df, num_rows="dynamic", use_container_width=True, column_config={"階段": st.column_config.SelectboxColumn(options=["活動前", "活動中", "活動後"]), "完成狀態": st.column_config.CheckboxColumn("是否完成", default=False)})
        if st.button("💾 儲存任務變更", key="save_tasks"):
            conn.update(worksheet="Tasks", data=edited_tasks)
            st.success("任務清單已儲存！")
            
    with tab2:
        st.write("### 🧑‍🤝‍🧑 職務安排與 Excel 模板對應")
        try: role_df = conn.read(worksheet="Roles", ttl=0)
        except Exception: role_df = pd.DataFrame(columns=["姓名", "組別", "對應儲存格"])
        for col in ["姓名", "組別", "對應儲存格"]:
            if col not in role_df.columns: role_df[col] = pd.Series(dtype=object)

        with st.expander("➕ 新增人員職務", expanded=False):
            with st.form("add_role_form", clear_on_submit=True):
                col_x, col_y, col_z = st.columns(3)
                with col_x: r_name = st.text_input("姓名")
                with col_y: r_group = st.selectbox("組別", ["祈福組", "工作人員組", "服務老師組"])
                with col_z: r_cell = st.text_input("Excel 儲存格 (例如: A1)")
                if st.form_submit_button("新增人員"):
                    if r_name.strip():
                        role_df = pd.concat([role_df, pd.DataFrame({"姓名": [r_name], "組別": [r_group], "對應儲存格": [r_cell]})], ignore_index=True)
                        conn.update(worksheet="Roles", data=role_df)
                        st.success("人員新增成功！")
                        st.rerun()

        edited_roles = st.data_editor(role_df, num_rows="dynamic", use_container_width=True, column_config={"組別": st.column_config.SelectboxColumn(options=["祈福組", "工作人員組", "服務老師組"])})
        if st.button("💾 儲存職務變更", key="save_roles"):
            conn.update(worksheet="Roles", data=edited_roles)
            st.success("職務安排已儲存！")
            
        st.write("### 📤 一鍵套用模板匯出")
        uploaded_file = st.file_uploader("請上傳 Excel 格式模板 (.xlsx)", type=["xlsx"])
        if uploaded_file and st.button("✨ 產生專屬排班表並下載", type="primary"):
            try:
                wb = openpyxl.load_workbook(uploaded_file)
                ws = wb.active
                for idx, row in edited_roles.iterrows():
                    cell = row["對應儲存格"]
                    if pd.notna(cell) and str(cell).strip() != "": ws[cell] = row["姓名"]
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                st.download_button(label="📥 下載排班表", data=output, file_name=f"{datetime.datetime.now().strftime('%Y%m%d')}_排班表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e: st.error(f"匯出失敗：{e}")

    with tab3:
        st.write("### 📦 活動器材清單管理")
        try: eq_df = conn.read(worksheet="Equipment", ttl=0)
        except Exception: eq_df = pd.DataFrame(columns=["器材名稱", "數量", "負責人", "取得位置", "準備狀態"])
        for col in ["器材名稱", "數量", "負責人", "取得位置", "準備狀態"]:
            if col not in eq_df.columns: eq_df[col] = pd.Series(dtype=object)
        eq_df["準備狀態"] = eq_df["準備狀態"].replace({'TRUE': True, 'FALSE': False, 'True': True, 'False': False, '1': True, '0': False}).fillna(False).astype(bool)

        with st.expander("➕ 新增器材", expanded=False):
            with st.form("add_eq_form", clear_on_submit=True):
                col_e1, col_e2, col_e3, col_e4 = st.columns([2, 1, 2, 2])
                with col_e1: e_name = st.text_input("器材名稱 *")
                with col_e2: e_qty = st.number_input("數量", min_value=1, value=1)
                with col_e3: e_pic = st.text_input("負責準備人員")
                with col_e4: e_loc = st.text_input("取得位置 (放哪/去哪買)")
                if st.form_submit_button("新增器材"):
                    if e_name.strip():
                        eq_df = pd.concat([eq_df, pd.DataFrame({"器材名稱": [e_name], "數量": [e_qty], "負責人": [e_pic], "取得位置": [e_loc], "準備狀態": [False]})], ignore_index=True)
                        conn.update(worksheet="Equipment", data=eq_df)
                        st.success("器材新增成功！")
                        st.rerun()

        edited_eq = st.data_editor(eq_df, num_rows="dynamic", use_container_width=True, column_config={"數量": st.column_config.NumberColumn("數量", min_value=1, step=1), "準備狀態": st.column_config.CheckboxColumn("是否已準備好", default=False)})
        if st.button("💾 儲存器材變更", key="save_eq"):
            conn.update(worksheet="Equipment", data=edited_eq)
            st.success("器材清單已儲存！")

# ==========================================
# 模組 6：歷史紀錄與進度 (後台)
# ==========================================
def render_history_page(conn):
    st.subheader("🗂️ 歷史紀錄與成全進度")
    try: reg_df = conn.read(worksheet="Registration", ttl=0)
    except Exception: st.warning("尚無紀錄。"); return
    if reg_df.empty: return

    display_df = reg_df.copy()
    def format_phone(val):
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2]
        if s.lower() in ['nan', 'none', '']: return ""
        if s and not s.startswith('0'): return '0' + s
        return s
    display_df['聯繫方式'] = display_df['聯繫方式'].apply(format_phone)

    progress_options = ["初次接觸", "已參加活動", "有意願", "已求道", "穩定參與", "其他"]
    edited_history = st.data_editor(display_df, use_container_width=True, column_config={"聯繫方式": st.column_config.TextColumn("聯繫方式"), "成全進度": st.column_config.SelectboxColumn(options=progress_options), "報到序號": st.column_config.NumberColumn(disabled=True), "報名時間": st.column_config.TextColumn(disabled=True)})
    
    col1, col2 = st.columns([1, 5])
    with col1:
        if st.button("💾 儲存進度", type="primary"):
            conn.update(worksheet="Registration", data=edited_history)
            st.success("已更新！")
    with col2:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer: edited_history.to_excel(writer, index=False)
        output.seek(0)
        st.download_button("📥 匯出完整名單", data=output, file_name=f"紀錄_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==========================================
# 模組 7：各站點完整名單總覽 (後台)
# ==========================================
def render_full_queue_page(conn):
    st.subheader("📋 各站點完整名單總覽")
    
    col_a, col_b = st.columns([3, 1])
    with col_a:
        st.caption("✅ 此頁面顯示所有站點的完整排隊紀錄（包含已體驗「完成」的名單）。")
    with col_b:
        if st.button("🔄 重新整理資料", use_container_width=True):
            st.rerun()

    try:
        queue_df = conn.read(worksheet="Queue", ttl=0)
        settings_df = conn.read(worksheet="Settings", ttl=30)
    except Exception:
        st.warning("目前尚無資料。")
        return

    if queue_df.empty or settings_df.empty:
        st.info("目前無排隊資料。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors='coerce').fillna(0).astype(int)
    stations = settings_df["項目名稱"].tolist()

    if stations:
        tabs = st.tabs(stations)
        for i, station in enumerate(stations):
            with tabs[i]:
                st.write(f"### 📍 {station} 完整名單")
                station_queue = queue_df[queue_df["體驗站點"] == station].sort_values(by="站點序號")
                
                if station_queue.empty:
                    st.info("尚無人報名此項目。")
                else:
                    st.dataframe(
                        station_queue[["站點序號", "報到序號", "姓名", "狀態", "報名時間"]], 
                        use_container_width=True, 
                        hide_index=True
                    )

# ==========================================
# 主程式路由
# ==========================================
def main():
    st.markdown("<h3 style='color: #888; margin-top:-20px;'>💆 身心靈保健活動系統</h3>", unsafe_allow_html=True)
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    is_admin = st.session_state.get("is_admin", False)
    
    def page_display(): render_display_page(conn)
    def page_registration(): render_registration_page(conn)
    def page_calling(): render_calling_page(conn)
    def page_full_queue(): render_full_queue_page(conn)
    def page_history(): render_history_page(conn)
    def page_settings(): render_settings_page(conn)
    def page_task(): render_task_page(conn)

    if hasattr(st, "navigation"):
        pages = {"📺 顯示專區": [st.Page(page_display, title="民眾體驗顯示螢幕 (大螢幕)")], "📝 報名專區": [st.Page(page_registration, title="民眾報名專區 (前台)")]}
        if is_admin:
            pages["📝 叫號與紀錄 (後台)"] = [
                st.Page(page_calling, title="排隊清單與叫號操作 (後台)"), 
                st.Page(page_full_queue, title="各站點完整名單總覽 (後台)"), 
                st.Page(page_history, title="歷史紀錄與進度 (後台)")
            ]
            pages["⚙️ 系統與後台管理"] = [
                st.Page(page_settings, title="體驗項目與名額設定 (後台)"), 
                st.Page(page_task, title="任務與職務管理 (後台)")
            ]
        pg = st.navigation(pages)
        pg.run()
    else:
        st.sidebar.markdown("### 🗂️ 系統導覽選單")
        tree_menu = ["📺 顯示專區", "　└ 民眾體驗顯示螢幕 (大螢幕)", "📝 報名專區", "　└ 民眾報名專區 (前台)"]
        if is_admin:
            tree_menu = [
                "📺 顯示專區", "　├ 民眾體驗顯示螢幕 (大螢幕)", 
                "📝 報名與叫號專區", "　├ 民眾報名專區 (前台)", "　├ 排隊清單與叫號操作 (後台)", "　├ 各站點完整名單總覽 (後台)", "　└ 歷史紀錄與進度 (後台)", 
                "⚙️ 系統與後台管理", "　├ 體驗項目與名額設定 (後台)", "　└ 任務與職務管理 (後台)"
            ]
            
        choice = st.sidebar.radio("請選擇頁面：", tree_menu, label_visibility="collapsed")
        
        if choice in ["　├ 民眾體驗顯示螢幕 (大螢幕)", "　└ 民眾體驗顯示螢幕 (大螢幕)"]: render_display_page(conn)
        elif choice in ["　├ 民眾報名專區 (前台)", "　└ 民眾報名專區 (前台)"]: render_registration_page(conn)
        elif choice == "　├ 排隊清單與叫號操作 (後台)": render_calling_page(conn)
        elif choice == "　├ 各站點完整名單總覽 (後台)": render_full_queue_page(conn)
        elif choice == "　└ 歷史紀錄與進度 (後台)": render_history_page(conn)
        elif choice == "　├ 體驗項目與名額設定 (後台)": render_settings_page(conn)
        elif choice == "　└ 任務與職務管理 (後台)": render_task_page(conn)
        else: st.info("👈 這裡是分類標題，請點擊下方的子項目進入對應頁面。")

    st.sidebar.markdown("---")
    if not is_admin:
        with st.sidebar.expander("🔐 工作人員入口", expanded=False):
            with st.form("login_form"):
                pwd = st.text_input("請輸入密碼解鎖後台", type="password")
                submit_login = st.form_submit_button("確認登入", use_container_width=True)
                if submit_login:
                    if pwd == "1234":
                        st.session_state["is_admin"] = True
                        st.rerun() 
                    else: st.error("密碼錯誤")
    else:
        st.sidebar.success("✅ 管理員已登入")
        if st.sidebar.button("🚪 登出並隱藏後台", use_container_width=True):
            st.session_state["is_admin"] = False
            st.rerun()

if __name__ == '__main__':
    main()

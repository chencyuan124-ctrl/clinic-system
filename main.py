# --完整main.py程式碼開頭--
import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import datetime
from gtts import gTTS
import io
import base64
import openpyxl
import time

# --網頁基本設定開頭--
st.set_page_config(page_title="身心靈保健活動系統", page_icon="💆", layout="wide")
# --網頁基本設定結尾--

# ==========================================
# 輔助函式：文字轉語音並自動播放
# ==========================================
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

# ==========================================
# 模組 1：前台報名 (漏填防呆、成功才清空)
# ==========================================
# --前台報名模組開頭--
def render_registration_page(conn):
    st.subheader("📝 民眾報名專區")
    
    # 【升級：初始化表單金鑰與成功提示狀態】
    if "reg_form_key" not in st.session_state: st.session_state["reg_form_key"] = 0
    if "add_form_key" not in st.session_state: st.session_state["add_form_key"] = 0
    
    # 檢查是否有成功註冊的暫存訊息，有就顯示並放氣球，然後刪除暫存
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

    if full_options:
        st.info(f"💡 溫馨提示：以下項目已額滿 - {', '.join(full_options)}")

    mode = st.radio("請選擇您的報名身份：", ["🆕 報名服務項目", "🔄 已做完兩項，加選服務項目"], horizontal=True)

    if mode == "🆕 報名服務項目":
        # 【升級：clear_on_submit 改為 False，用動態 key 來控制成功後才清空】
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
                    try:
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

                    # 【升級：存入成功訊息，更改表單金鑰強制清空，然後刷新網頁】
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
                st.error(f"⚠️ 系統檢查到您還有尚未完成的項目：【{', '.join(unfinished_items)}】\n\n請您先將目前的項目體驗「完成」後，再來進行加選喔！")
            else:
                done_items = user_queues["體驗站點"].dropna().tolist()
                st.info(f"✅ 您已經完成的項目：{', '.join(done_items)}")
                
                new_available = [x for x in available_options if x not in done_items]
                
                # 同理：設定動態 key 並改為 False
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
# --前台報名模組結尾--

# ==========================================
# 模組 2：後台設定 (老師名單連動職務表 + 專屬編輯表單)
# ==========================================
# --後台設定模組開頭--
def render_settings_page(conn):
    st.subheader("⚙️ 體驗項目與名額設定 (後台)")
    
    # 1. 讀取現有項目
    try:
        df = conn.read(worksheet="Settings", ttl=0)
    except Exception:
        df = pd.DataFrame(columns=["項目名稱", "老師名單", "總名額", "已報名數"])

    for col in ["項目名稱", "老師名單", "總名額", "已報名數"]:
        if col not in df.columns: df[col] = pd.Series(dtype=object)

    # 2. 讀取職務安排中的「服務老師組」名單
    try:
        roles_df = conn.read(worksheet="Roles", ttl=0)
        teacher_list = roles_df[roles_df["組別"] == "服務老師組"]["姓名"].dropna().unique().tolist()
    except Exception:
        teacher_list = []

    if not teacher_list:
        st.warning("⚠️ 系統中目前沒有「服務老師組」的名單。請先至「任務與職務管理」中新增人員，才能在這裡使用下拉選單綁定老師。")

    col1, col2 = st.columns(2)
    
    # 3. 新增體驗項目表單
    with col1:
        with st.expander("➕ 新增體驗項目", expanded=False):
            with st.form("add_item_form", clear_on_submit=True):
                new_item = st.text_input("項目名稱")
                new_teachers = st.multiselect("老師名單 (資料來源：任務與職務管理)", options=teacher_list)
                new_quota = st.number_input("總名額", min_value=1, value=20)
                
                if st.form_submit_button("確認新增"):
                    if not new_item.strip():
                        st.error("請輸入項目名稱！")
                    elif new_item in df["項目名稱"].values:
                        st.error("此項目名稱已存在，請勿重複新增！")
                    else:
                        teachers_str = "、".join(new_teachers)
                        new_row = pd.DataFrame({
                            "項目名稱": [new_item], 
                            "老師名單": [teachers_str], 
                            "總名額": [new_quota], 
                            "已報名數": [0]
                        })
                        df = pd.concat([df, new_row], ignore_index=True)
                        conn.update(worksheet="Settings", data=df)
                        st.success(f"已成功新增【{new_item}】！")
                        st.rerun()

    # 4. 編輯現有項目表單 (解決表格無法多選的限制)
    with col2:
        if not df.empty:
            with st.expander("✏️ 編輯現有項目 (修改老師或名額)", expanded=False):
                edit_target = st.selectbox("請選擇要修改的項目", df["項目名稱"].tolist())
                
                # 取得該項目的現有資料
                current_row = df[df["項目名稱"] == edit_target].iloc[0]
                current_quota = int(current_row["總名額"])
                
                # 解析現有的老師名單，並確保他們還在 teacher_list 裡面 (避免報錯)
                current_teachers_str = str(current_row["老師名單"])
                current_teachers_list = current_teachers_str.split("、") if current_teachers_str != "nan" and current_teachers_str else []
                valid_default_teachers = [t for t in current_teachers_list if t in teacher_list]

                with st.form("edit_item_form"):
                    edit_teachers = st.multiselect("重新勾選老師名單", options=teacher_list, default=valid_default_teachers)
                    edit_quota = st.number_input("修改總名額", min_value=1, value=current_quota)
                    
                    if st.form_submit_button("儲存修改"):
                        target_idx = df[df["項目名稱"] == edit_target].index[0]
                        df.loc[target_idx, "老師名單"] = "、".join(edit_teachers)
                        df.loc[target_idx, "總名額"] = edit_quota
                        
                        conn.update(worksheet="Settings", data=df)
                        st.success(f"【{edit_target}】已成功更新！")
                        st.rerun()

    st.write("### 📝 項目總覽表")
    st.caption("提示：下方表格僅供檢視與刪除。若要修改老師名單或總名額，請使用上方的「✏️ 編輯現有項目」。若要刪除項目，請選取左側核取方塊後按 Delete 鍵。")
    
    # 將表格設為唯讀 (除了刪除功能以外)，強迫使用者透過上方表單修改，確保資料一致性
    edited_df = st.data_editor(
        df, 
        num_rows="dynamic", 
        use_container_width=True, 
        disabled=["項目名稱", "老師名單", "總名額", "已報名數"]
    )
    
    # 由於表格只開放刪除，所以按下儲存時就是同步刪除狀態
    if st.button("💾 確認刪除並同步至雲端", type="primary"):
        conn.update(worksheet="Settings", data=edited_df)
        st.success("資料已同步！")
        st.rerun()
# --後台設定模組結尾--

# ==========================================
# 模組 3：叫號系統 (修復語音中斷問題)
# ==========================================
# --叫號系統模組開頭--
def render_calling_page(conn):
    st.subheader("📢 叫號操作台")
    
    # 【關鍵修復】網頁重新載入後，檢查是否有待播報的語音，有的話立刻播放！
    if "pending_audio" in st.session_state:
        autoplay_audio(st.session_state["pending_audio"])
        st.success(f"📢 正在播報：{st.session_state['pending_audio']}")
        del st.session_state["pending_audio"] # 播完就刪除，避免下次整理網頁又重播

    try:
        queue_df = conn.read(worksheet="Queue", ttl=0)
        settings_df = conn.read(worksheet="Settings", ttl=5)
    except Exception:
        st.warning("目前尚無排隊資料。")
        return
        
    if settings_df.empty: return

    station_options = settings_df["項目名稱"].tolist()
    current_station = st.selectbox("請選擇您負責的服務站點：", station_options)

    if queue_df.empty:
        st.info("目前無人排隊。")
        return

    queue_df["站點序號"] = pd.to_numeric(queue_df["站點序號"], errors='coerce').fillna(0).astype(int)

    mask = (queue_df["體驗站點"] == current_station) & (queue_df["狀態"] != "完成")
    station_queue = queue_df[mask].sort_values(by="站點序號").copy()

    st.write(f"### 📍 【{current_station}】排隊清單")
    if station_queue.empty:
        st.info("目前尚無排隊名單。")
    else:
        serving_df = station_queue[station_queue["狀態"] == "服務中"]
        if not serving_df.empty:
            serving_person = serving_df.iloc[0]
            st.success(f"👨‍⚕️ **目前服務中：** 第 {serving_person['站點序號']} 號 - {serving_person['姓名']} (總號: {serving_person['報到序號']})")
        else:
            st.info("💡 目前無人體驗，請點擊「呼叫下一位」。")
            
        st.dataframe(station_queue[["站點序號", "報到序號", "姓名", "狀態", "報名時間"]], use_container_width=True)

    st.markdown("---")
    st.write("### 🎛️ 叫號操作區")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        if st.button("🔊 呼叫下一位", type="primary", use_container_width=True):
            if not serving_df.empty:
                st.warning("⚠️ 目前已有「服務中」的民眾，請先標記為完成或過號！")
            else:
                waiting_df = station_queue[station_queue["狀態"] == "等待中"]
                if not waiting_df.empty:
                    next_person = waiting_df.iloc[0]
                    target_idx = queue_df[(queue_df["站點序號"] == next_person["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                    
                    queue_df.loc[target_idx, "狀態"] = "服務中"
                    conn.update(worksheet="Queue", data=queue_df)
                    
                    name = next_person['姓名']
                    seq = next_person['站點序號']
                    announce_text = f"來賓 {seq} 號 {name}，{name} 請到 {current_station} 處報到。"
                    
                    # 【關鍵修復】不直接播音，而是存入暫存，讓網頁先更新！
                    st.session_state["pending_audio"] = announce_text
                    st.rerun()
                else:
                    st.info("沒有等待中的民眾了！")
                    
    with col2:
        if st.button("📢 再次呼叫當前", use_container_width=True):
            if not serving_df.empty:
                serving_person = serving_df.iloc[0]
                name = serving_person['姓名']
                seq = serving_person['站點序號']
                announce_text = f"來賓 {seq} 號 {name}，{name} 請到 {current_station} 處報到。"
                
                # 這個按鈕沒有觸發 st.rerun()，所以原本就直接播放不會中斷
                autoplay_audio(announce_text)
                st.success(f"📢 正在重新播報：{announce_text}")
            else:
                st.warning("⚠️ 目前沒有正在「服務中」的民眾可以重叫！")

    with col3:
        if st.button("⏭️ 標記為「過號」", use_container_width=True):
            if not serving_df.empty:
                serving_person = serving_df.iloc[0]
                target_idx = queue_df[(queue_df["站點序號"] == serving_person["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                queue_df.loc[target_idx, "狀態"] = "過號"
                conn.update(worksheet="Queue", data=queue_df)
                st.warning("已標記為過號。")
                st.rerun()

    with col4:
        if st.button("✅ 標記為「完成」", use_container_width=True):
            if not serving_df.empty:
                serving_person = serving_df.iloc[0]
                target_idx = queue_df[(queue_df["站點序號"] == serving_person["站點序號"]) & (queue_df["體驗站點"] == current_station)].index[0]
                queue_df.loc[target_idx, "狀態"] = "完成"
                conn.update(worksheet="Queue", data=queue_df)
                st.success("已標記為完成！")
                st.rerun()

    st.markdown("---")
    st.write("### 🔁 過號重叫區")
    missed_df = station_queue[station_queue["狀態"] == "過號"]
    if not missed_df.empty:
        missed_options = [f"第{int(row['站點序號'])}號 - {row['姓名']}" for idx, row in missed_df.iterrows()]
        selected_missed = st.selectbox("請選擇要重叫的過號民眾：", missed_options)
        
        if st.button("🔊 過號重叫", use_container_width=True):
            if not serving_df.empty:
                st.warning("⚠️ 目前已有服務中名單，請先完成。")
            else:
                missed_seq = int(selected_missed.split("號")[0].replace("第", ""))
                target_idx = queue_df[(queue_df["站點序號"] == missed_seq) & (queue_df["體驗站點"] == current_station)].index[0]
                
                queue_df.loc[target_idx, "狀態"] = "服務中"
                conn.update(worksheet="Queue", data=queue_df)
                
                name = queue_df.loc[target_idx, "姓名"]
                announce_text = f"來賓 {missed_seq} 號 {name}，{name} 請到 {current_station} 處報到。"
                
                # 【關鍵修復】存入暫存，讓網頁先更新再播放
                st.session_state["pending_audio"] = announce_text
                st.rerun()
# --叫號系統模組結尾--

# ==========================================
# 模組 4：任務與職務管理 (修正 Checkbox 型態)
# ==========================================
# --任務與職務管理模組開頭--
def render_task_page(conn):
    st.subheader("📋 任務與職務管理")
    tab1, tab2 = st.tabs(["任務清單管理", "職務安排與模板設定"])
    
    with tab1:
        st.write("### 🎯 前/中/後任務管理")
        try:
            task_df = conn.read(worksheet="Tasks", ttl=0)
        except Exception:
            task_df = pd.DataFrame(columns=["階段", "任務名稱", "負責人", "完成狀態"])

        for col in ["階段", "任務名稱", "負責人", "完成狀態"]:
            if col not in task_df.columns: task_df[col] = pd.Series(dtype=object)

        # 強制修正完成狀態為 Boolean，避免 Streamlit 報 Float 錯誤
        task_df["完成狀態"] = task_df["完成狀態"].replace({'TRUE': True, 'FALSE': False, 'True': True, 'False': False, '1': True, '0': False})
        task_df["完成狀態"] = task_df["完成狀態"].fillna(False).astype(bool)

        with st.expander("➕ 新增任務", expanded=False):
            with st.form("add_task_form", clear_on_submit=True):
                col_a, col_b, col_c = st.columns(3)
                with col_a: t_phase = st.selectbox("執行階段", ["活動前", "活動中", "活動後"])
                with col_b: t_name = st.text_input("任務名稱")
                with col_c: t_pic = st.text_input("負責人")
                if st.form_submit_button("新增任務"):
                    if t_name.strip():
                        new_row = pd.DataFrame({"階段": [t_phase], "任務名稱": [t_name], "負責人": [t_pic], "完成狀態": [False]})
                        task_df = pd.concat([task_df, new_row], ignore_index=True)
                        conn.update(worksheet="Tasks", data=task_df)
                        st.success("任務新增成功！")
                        st.rerun()

        st.write("**編輯現有任務**")
        edited_tasks = st.data_editor(task_df, num_rows="dynamic", use_container_width=True,
            column_config={"階段": st.column_config.SelectboxColumn(options=["活動前", "活動中", "活動後"]), "完成狀態": st.column_config.CheckboxColumn("是否完成", default=False)})
        if st.button("💾 儲存任務變更", key="save_tasks"):
            conn.update(worksheet="Tasks", data=edited_tasks)
            st.success("任務清單已儲存！")
            
    with tab2:
        st.write("### 🧑‍🤝‍🧑 職務安排與 Excel 模板對應")
        try:
            role_df = conn.read(worksheet="Roles", ttl=0)
        except Exception:
            role_df = pd.DataFrame(columns=["姓名", "組別", "對應儲存格"])
            
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
                        new_row = pd.DataFrame({"姓名": [r_name], "組別": [r_group], "對應儲存格": [r_cell]})
                        role_df = pd.concat([role_df, new_row], ignore_index=True)
                        conn.update(worksheet="Roles", data=role_df)
                        st.success("人員新增成功！")
                        st.rerun()

        st.write("**編輯現有人員**")
        edited_roles = st.data_editor(role_df, num_rows="dynamic", use_container_width=True,
            column_config={"組別": st.column_config.SelectboxColumn(options=["祈福組", "工作人員組", "服務老師組"])})
        if st.button("💾 儲存職務變更", key="save_roles"):
            conn.update(worksheet="Roles", data=edited_roles)
            st.success("職務安排已儲存！")
            
        st.markdown("---")
        st.write("### 📤 一鍵套用模板匯出")
        uploaded_file = st.file_uploader("請上傳 Excel 格式模板 (.xlsx)", type=["xlsx"])
        if uploaded_file and st.button("✨ 產生專屬排班表並下載", type="primary"):
            try:
                wb = openpyxl.load_workbook(uploaded_file)
                ws = wb.active
                for idx, row in edited_roles.iterrows():
                    cell = row["對應儲存格"]
                    name = row["姓名"]
                    if pd.notna(cell) and str(cell).strip() != "": ws[cell] = name
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                filename = f"{datetime.datetime.now().strftime('%Y%m%d')}_排班表.xlsx"
                st.download_button(label="📥 下載排班表", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"匯出失敗，錯誤訊息：{e}")
# --任務與職務管理模組結尾--

# ==========================================
# 模組 5：大螢幕顯示面板 (新增過號名單顯示)
# ==========================================
# --大螢幕顯示面板開頭--
def render_display_page(conn):
    st.markdown("<h3 style='text-align: center; color: #1f77b4; margin-bottom: 0;'>📺 體驗進度總覽</h3>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 1.2em; color: #555;'>請參考下方最新進度前往體驗</p>", unsafe_allow_html=True)
    
    col_a, col_b = st.columns([3, 1])
    with col_a:
        auto_refresh = st.checkbox("🔄 啟用自動接收叫號更新 (每 3 秒刷新一次)", value=False)
    with col_b:
        if st.button("🔄 手動重新整理", use_container_width=True):
            st.rerun()
    
    try:
        queue_df = conn.read(worksheet="Queue", ttl=0)
        settings_df = conn.read(worksheet="Settings", ttl=5)
    except Exception:
        st.warning("無法讀取資料庫。")
        return
        
    if settings_df.empty or queue_df.empty:
        st.info("目前無資料。")
        return

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
                    
                    # 撈取等待與過號名單
                    waiting = station_q[station_q["狀態"] == "等待中"].head(4) 
                    missed = station_q[station_q["狀態"] == "過號"]
                    
                    if not waiting.empty or not missed.empty:
                        # 顯示正常等待名單
                        for idx, w_person in waiting.iterrows():
                            st.markdown(f"<div style='font-size:1.4em; padding:4px 0; color:#333;'><b>{w_person['站點序號']}</b>號 {w_person['姓名']}</div>", unsafe_allow_html=True)
                        # 顯示過號名單 (加上紅字標示)
                        for idx, m_person in missed.iterrows():
                            st.markdown(f"<div style='font-size:1.4em; padding:4px 0; color:#dc3545;'><b>{m_person['站點序號']}</b>號 {m_person['姓名']} <span style='font-size:0.8em; font-weight:bold;'>(過號)</span></div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div style='font-size:1.3em; color:#6c757d; padding:4px 0;'>無等待民眾</div>", unsafe_allow_html=True)
        
        st.markdown("<hr style='margin: 2em 0; border: 1px solid #eee;'>", unsafe_allow_html=True)

    if auto_refresh:
        time.sleep(3)
        st.rerun()
# --大螢幕顯示面板結尾--

# ==========================================
# 模組 6：歷史紀錄與進度 (電話補零修復)
# ==========================================
# --歷史紀錄模組開頭--
def render_history_page(conn):
    st.subheader("🗂️ 歷史紀錄與成全進度")
    try:
        reg_df = conn.read(worksheet="Registration", ttl=0)
    except Exception:
        st.warning("尚無紀錄。"); return
    if reg_df.empty: return

    display_df = reg_df.copy()
    
    # 【升級：智慧修復電話號碼格式】
    def format_phone(val):
        s = str(val).strip()
        if s.endswith('.0'): s = s[:-2] # 砍掉 .0
        if s.lower() in ['nan', 'none', '']: return ""
        if s and not s.startswith('0'): return '0' + s # 如果不是0開頭就補0
        return s

    display_df['聯繫方式'] = display_df['聯繫方式'].apply(format_phone)

    progress_options = ["初次接觸", "已參加活動", "有意願", "已求道", "穩定參與", "其他"]
    edited_history = st.data_editor(
        display_df, 
        use_container_width=True, 
        column_config={
            "聯繫方式": st.column_config.TextColumn("聯繫方式"),
            "成全進度": st.column_config.SelectboxColumn(options=progress_options), 
            "報到序號": st.column_config.NumberColumn(disabled=True), 
            "報名時間": st.column_config.TextColumn(disabled=True)
        }
    )
    
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
# --歷史紀錄模組結尾--

# ==========================================
# 主程式路由 (調整歷史紀錄分類)
# ==========================================
# --主程式架構開頭--
def main():
    st.markdown("<h3 style='color: #888; margin-top:-20px;'>💆 身心靈保健活動系統</h3>", unsafe_allow_html=True)
    conn = st.connection("gsheets", type=GSheetsConnection)
    
    # 判斷系統是否支援最新版的 st.navigation 原生樹狀選單
    if hasattr(st, "navigation"):
        # 建立無參數的包裝函式供 st.Page 呼叫
        def page_display(): render_display_page(conn)
        def page_registration(): render_registration_page(conn)
        def page_calling(): render_calling_page(conn)
        def page_settings(): render_settings_page(conn)
        def page_task(): render_task_page(conn)
        def page_history(): render_history_page(conn)

        # 定義完美的樹狀分類結構
        pages = {
            "📺 顯示專區": [
                st.Page(page_display, title="民眾體驗顯示螢幕 (大螢幕)"),
            ],
            "📝 報名與叫號專區": [
                st.Page(page_registration, title="民眾報名專區 (前台)"),
                st.Page(page_calling, title="排隊清單與叫號操作 (後台)"),
                st.Page(page_history, title="歷史紀錄與進度 (後台)"), # 已經搬移到這裡囉！
            ],
            "⚙️ 系統與後台管理": [
                st.Page(page_settings, title="體驗項目與名額設定 (後台)"),
                st.Page(page_task, title="任務與職務管理 (後台)"),
            ]
        }
        
        # 啟動樹狀導覽
        pg = st.navigation(pages)
        pg.run()
        
    else:
        # 萬一版本較舊，使用符號畫出備用的樹狀選單
        st.sidebar.markdown("### 🗂️ 系統導覽選單")
        tree_menu = [
            "📺 顯示專區",
            "　├ 民眾體驗顯示螢幕 (大螢幕)",
            "📝 報名與叫號專區",
            "　├ 民眾報名專區 (前台)",
            "　├ 排隊清單與叫號操作 (後台)", # 符號改為 ├
            "　└ 歷史紀錄與進度 (後台)",   # 搬移到這裡，符號為 └
            "⚙️ 系統與後台管理",
            "　├ 體驗項目與名額設定 (後台)",
            "　└ 任務與職務管理 (後台)"    # 符號改為 └
        ]
        
        choice = st.sidebar.radio("請選擇頁面：", tree_menu, label_visibility="collapsed")
        st.sidebar.markdown("---")
        
        # 根據選擇的節點執行對應的頁面 (注意符號也有跟著對應修改)
        if choice == "　├ 民眾體驗顯示螢幕 (大螢幕)": render_display_page(conn)
        elif choice == "　├ 民眾報名專區 (前台)": render_registration_page(conn)
        elif choice == "　├ 排隊清單與叫號操作 (後台)": render_calling_page(conn)
        elif choice == "　└ 歷史紀錄與進度 (後台)": render_history_page(conn)
        elif choice == "　├ 體驗項目與名額設定 (後台)": render_settings_page(conn)
        elif choice == "　└ 任務與職務管理 (後台)": render_task_page(conn)
        else:
            # 點到大分類標題時的防呆提示
            st.info("👈 這裡是分類標題，請點擊下方的子項目進入對應頁面。")

if __name__ == '__main__':
    main()
# --主程式架構結尾--
# --完整main.py程式碼結尾--

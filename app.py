import streamlit as st
import pandas as pd
import datetime
import re
import uuid

st.title("Department Roster to Calendar Generator")
st.write("Upload your original department Excel roster file to extract your customized calendar.")

uploaded_file = st.file_uploader("Upload Roster Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1. 讀取主分頁 "Duty List" 並強制將所有內容先轉為字串
        df_raw_duty = pd.read_excel(uploaded_file, sheet_name="Duty List", header=None)
        df_raw_duty = df_raw_duty.fillna("nan").astype(str)
        
        # 從 A1 儲存格動態提取年份與月份
        header_text = str(df_raw_duty.iloc[0, 0]).strip()
        
        # 尋找連續 4 個數字作為年份
        year_match = re.search(r'\b(\d{4})\b', header_text)
        year = int(year_match.group(1)) if year_match else datetime.date.today().year
        
        # 月份名稱與數字映射
        months_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        
        month = datetime.date.today().month  # 預設預防值
        for month_name, month_num in months_map.items():
            if month_name in header_text.lower():
                month = month_num
                break
        
        # 2. 定義寫死的巡房寫死欄位 (0-indexed 映射)
        # C 欄 = 2, D 欄 = 3, F 欄 = 5, G 欄 = 6, H 欄 = 7, I 欄 = 8, J 欄 = 9
        ward_column_map = {
            2: "Round: A11/E11",
            3: "Round: EPAC",
            5: "Round: C2",
            6: "Round: C2",
            7: "Round: E2",
            8: "Round: A2",
            9: "Round: A2"
        }
        
        # 讀取 "Ward round" 分頁並進行安全的字串與填補空白處理
        df_raw_round = pd.read_excel(uploaded_file, sheet_name="Ward round", header=None)
        df_raw_round = df_raw_round.fillna("nan").astype(str)
        
        # 從第 2 行提取所有員工簡寫名稱
        raw_names = df_raw_duty.iloc[1, 2:].tolist()
        ignored_tokens = ['nan', 'free', 'an', 'gyn', '']
        personnel_list = [str(name).strip() for name in raw_names if str(name).strip() and str(name).strip().lower() not in ignored_tokens]
        
        # 生成下拉選單
        selected_staff = st.selectbox("Select your name to extract duties:", sorted(list(set(personnel_list))))
        
        if selected_staff:
            events = []
            
            # --- 階段 A: 提取巡房表 (Ward Round) ---
            current_round_date = None
            round_records = {} # 用於標記當天是否有巡房，以便後續將 AM 動態切換為 10:30
            
            # 巡房數據行從第 5 行 (index 4) 開始
            for r_idx in range(4, len(df_raw_round)):
                col_a = str(df_raw_round.iloc[r_idx, 0]).strip()
                
                if col_a != "nan" and col_a != "":
                    try:
                        # 處理可能帶有浮點數點號的日期字串，例如 "1.0" 轉為 "1"
                        day_num = int(float(col_a))
                        current_round_date = datetime.date(year, month, day_num).strftime('%Y-%m-%d')
                    except ValueError:
                        pass
                
                if not current_round_date:
                    continue
                    
                # 掃描寫死的特定幾欄
                for c_idx, ward_name in ward_column_map.items():
                    if c_idx < df_raw_round.shape[1]:
                        cell_txt = str(df_raw_round.iloc[r_idx, c_idx]).strip()
                        if cell_txt and cell_txt != 'nan':
                            # 子字串包含比對，確保帶有標籤（如 A Yam[T1]）也能正確識別
                            if selected_staff.lower() in cell_txt.lower():
                                events.append({
                                    'date': current_round_date,
                                    'start_time': '08:30',
                                    'end_time': '10:30',
                                    'summary': ward_name,
                                    'description': f"Roster notation: {cell_txt}"
                                })
                                # 記錄該醫生當天已有巡房
                                round_records[current_round_date] = True

            # --- 階段 B: 提取主班表 (Duty List) ---
            staff_col_idx = None
            for idx, col_val in enumerate(df_raw_duty.iloc[1, :]):
                if str(col_val).strip() == selected_staff:
                    staff_col_idx = idx
                    break
            
            if staff_col_idx is not None:
                current_duty_date = None
                for r_idx in range(3, len(df_raw_duty)):
                    col_a = str(df_raw_duty.iloc[r_idx, 0]).strip()
                    col_b = str(df_raw_duty.iloc[r_idx, 1]).strip()
                    
                    # 碰到結尾統計欄位（如 free session）安全退出
                    if "free session" in col_a.lower() or "admin duty" in col_a.lower():
                        break
                        
                    if col_a != "nan" and col_a != "":
                        try:
                            day_num = int(float(col_a))
                            current_duty_date = datetime.date(year, month, day_num).strftime('%Y-%m-%d')
                        except ValueError:
                            pass
                    
                    if not current_duty_date:
                        continue
                        
                    time_slot = "AM" if "AM" in col_b else ("PM" if "PM" in col_b else None)
                    duty_val = str(df_raw_duty.iloc[r_idx, staff_col_idx]).strip()
                    
                    if duty_val and duty_val.lower() not in ['nan', '', 'x', '-']:
                        slot_info = {
                            'date': current_duty_date,
                            'summary': f"{time_slot}: {duty_val}",
                            'description': ""
                        }
                        
                        if time_slot == "AM":
                            # 判斷當天是否有巡房
                            has_round = current_duty_date in round_records
                            slot_info.update({'start_time': '10:30' if has_round else '08:30', 'end_time': '13:00'})
                            events.append(slot_info)
                        elif time_slot == "PM":
                            slot_info.update({'start_time': '14:00', 'end_time': '17:00'})
                            events.append(slot_info)

            # --- 階段 C: 編譯並生成產出檔案 ---
            if len(events) > 0:
                lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Duty//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
                for e in events:
                    date_raw = e['date'].replace('-', '')
                    st_raw = e['start_time'].replace(':', '') + "00"
                    end_raw = e['end_time'].replace(':', '') + "00"
                    uid = f"{date_raw}-{st_raw}-{uuid.uuid4().hex[:8]}@duty.local"
                    
                    lines.extend([
                        "BEGIN:VEVENT", f"UID:{uid}", 
                        f"DTSTART:{date_raw}T{st_raw}", f"DTEND:{date_raw}T{end_raw}", 
                        f"SUMMARY:{e['summary']}"
                    ])
                    if e.get('description'):
                        desc = e['description'].replace(',', '\\,').replace(';', '\\;')
                        lines.append(f"DESCRIPTION:{desc}")
                    lines.append("END:VEVENT")
                lines.append("END:VCALENDAR")
                ics_text = "\n".join(lines)
                
                st.info(f"Detected Target Period: {month}/{year}")
                st.success(f"Successfully compiled {len(events)} calendar items for {selected_staff}!")
                st.download_button(
                    label=f"Download {selected_staff}'s .ics File",
                    data=ics_text,
                    file_name=f"Roster_{selected_staff}_{month}_{year}.ics",
                    mime="text/calendar"
                )
            else:
                st.warning(f"No active schedule items mapped for {selected_staff}.")
                
    except Exception as e:
        st.error(f"Error parsing structural workbook parameters: {e}")

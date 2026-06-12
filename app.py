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
        xl = pd.ExcelFile(uploaded_file)
        
        # 1. 讀取主分頁 "Duty List"
        df_raw_duty = xl.parse("Duty List", header=None)
        df_raw_duty = df_raw_duty.fillna("nan").astype(str)
        
        # 動態提取年份與月份
        header_text = str(df_raw_duty.iloc[0, 0]).strip()
        year_match = re.search(r'\b(\d{4})\b', header_text)
        year = int(year_match.group(1)) if year_match else datetime.date.today().year
        
        months_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
        }
        
        month = datetime.date.today().month
        for month_name, month_num in months_map.items():
            if month_name in header_text.lower():
                month = month_num
                break
        
        # 2. 讀取 "Ward round" 分頁
        ward_column_map = {
            2: "Round: A11/E11",
            3: "Round: EPAC",
            5: "Round: C2",
            6: "Round: C2",
            7: "Round: E2",
            8: "Round: A2",
            9: "Round: A2"
        }
        df_raw_round = xl.parse("Ward round", header=None)
        df_raw_round = df_raw_round.fillna("nan").astype(str)
        
        # 從主表第 2 行提取所有員工簡寫名稱
        raw_names = df_raw_duty.iloc[1, 2:].tolist()
        ignored_tokens = ['nan', 'free', 'an', 'gyn', '']
        personnel_list = [str(name).strip() for name in raw_names if str(name).strip() and str(name).strip().lower() not in ignored_tokens]
        
        # UI 介面
        selected_staff = st.selectbox("Select your name to extract duties:", sorted(list(set(personnel_list))))
        include_leaves = st.checkbox("Include Leaves (e.g. AL, CO, OFF) as events", value=True)
        show_preop = st.checkbox("Show pre-op (general)", value=False)
        
        if selected_staff:
            events = []
            
            # --- 階段 A: 提取巡房表 (Ward Round) ---
            current_round_date = None
            round_records = {} 
            
            for r_idx in range(4, len(df_raw_round)):
                col_a = str(df_raw_round.iloc[r_idx, 0]).strip()
                if col_a != "nan" and col_a != "":
                    try:
                        day_num = int(float(col_a))
                        current_round_date = datetime.date(year, month, day_num).strftime('%Y-%m-%d')
                    except ValueError:
                        pass
                
                if not current_round_date:
                    continue
                    
                for c_idx, ward_name in ward_column_map.items():
                    if c_idx < df_raw_round.shape[1]:
                        cell_txt = str(df_raw_round.iloc[r_idx, c_idx]).strip()
                        if cell_txt and cell_txt != 'nan':
                            if selected_staff.lower() in cell_txt.lower():
                                events.append({
                                    'all_day': False,
                                    'date': current_round_date,
                                    'start_time': '08:30',
                                    'end_time': '10:30',
                                    'summary': ward_name,
                                    'description': f"Roster notation: {cell_txt}"
                                })
                                round_records[current_round_date] = True

            # --- 階段 B: 提取主班表 (Duty List) ---
            staff_col_idx = None
            for idx, col_val in enumerate(df_raw_duty.iloc[1, :]):
                if str(col_val).strip() == selected_staff:
                    staff_col_idx = idx
                    break
            
            if staff_col_idx is not None:
                current_duty_date_obj = None
                current_duty_date = None
                leave_tokens = ['al', 'co', 'off']
                
                for r_idx in range(3, len(df_raw_duty)):
                    col_a = str(df_raw_duty.iloc[r_idx, 0]).strip()
                    col_b = str(df_raw_duty.iloc[r_idx, 1]).strip()
                    
                    if "free session" in col_a.lower() or "admin duty" in col_a.lower():
                        break
                        
                    if col_a != "nan" and col_a != "":
                        try:
                            day_num = int(float(col_a))
                            current_duty_date_obj = datetime.date(year, month, day_num)
                            current_duty_date = current_duty_date_obj.strftime('%Y-%m-%d')
                        except ValueError:
                            pass
                    
                    if not current_duty_date or not current_duty_date_obj:
                        continue
                        
                    time_slot = "AM" if "AM" in col_b else ("PM" if "PM" in col_b else None)
                    duty_val = str(df_raw_duty.iloc[r_idx, staff_col_idx]).strip()
                    
                    if duty_val and duty_val.lower() not in ['nan', '', 'x', '-']:
                        # 處理自動推算前一週 8 天前的 pre-op 機制
                        if show_preop and duty_val.upper() == "OT" and any(day in col_b for day in ["Tue", "Wed"]):
                            preop_date_obj = current_duty_date_obj - datetime.timedelta(days=8)
                            events.append({
                                'all_day': False,
                                'date': preop_date_obj.strftime('%Y-%m-%d'),
                                'start_time': '09:30',
                                'end_time': '10:30',
                                'summary': 'pre-op',
                                'description': f"Automated pre-op session for OT duty scheduled on {current_duty_date} ({col_b})"
                            })
                        
                        # 檢查是否為請假/放假項目
                        is_leave = duty_val.lower() in leave_tokens
                        if is_leave and not include_leaves:
                            continue
                            
                        slot_info = {
                            'all_day': False,
                            'date': current_duty_date,
                            'summary': f"{time_slot}: {duty_val}",
                            'description': ""
                        }
                        
                        if time_slot == "AM":
                            has_round = current_duty_date in round_records
                            slot_info.update({'start_time': '10:30' if has_round else '08:30', 'end_time': '13:00'})
                            events.append(slot_info)
                        elif time_slot == "PM":
                            slot_info.update({'start_time': '14:00', 'end_time': '17:00'})
                            events.append(slot_info)

            # --- 階段 C: 生成 ICS 檔案 ---
            if len(events) > 0:
                lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Duty//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
                for e in events:
                    uid = f"{e['date'].replace('-', '')}-{uuid.uuid4().hex[:8]}@duty.local"
                    lines.append("BEGIN:VEVENT")
                    lines.append(f"UID:{uid}")
                    
                    date_raw = e['date'].replace('-', '')
                    st_raw = e['start_time'].replace(':', '') + "00"
                    end_raw = e['end_time'].replace(':', '') + "00"
                    
                    lines.append(f"DTSTART:{date_raw}T{st_raw}")
                    lines.append(f"DTEND:{date_raw}T{end_raw}")
                    lines.append(f"SUMMARY:{e['summary']}")
                    
                    if e.get('description'):
                        desc = e['description'].replace(',', '\\,').replace(';', '\\;')
                        lines.append(f"DESCRIPTION:{desc}")
                        
                    lines.append("END:VEVENT")
                lines.append("END:VCALENDAR")
                
                ics_text = "\r\n".join(lines) + "\r\n"
                
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

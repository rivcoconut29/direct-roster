import streamlit as st
import pandas as pd
import datetime
import re
import uuid

st.title("Department Roster to Calendar Generator")
st.write("Upload your original department Excel roster file to extract your customized calendar.")

uploaded_file = st.file_uploader("Upload Roster Excel File (.xlsx)", type=["xlsx"])

def extract_day_num(val):
    """安全提取 Call List 中的日期數字，相容 1900-01-XX 及 (SH) XX 格式"""
    val = str(val).strip()
    if not val or val.lower() == 'nan':
        return None
    
    # 處理 Excel 自動轉成的 1900 年日期字串
    date_match = re.search(r'1900-01-(\d+)', val)
    if date_match:
        return int(date_match.group(1))
        
    # 提取任何包含的數字序列 (例如 "(SH) 19" -> 19)
    num_match = re.search(r'\d+', val)
    if num_match:
        num = int(num_match.group(0))
        if num <= 31:
            return num
            
    return None

if uploaded_file is not None:
    try:
        # 1. 讀取主分頁 "Duty List" 並強制將所有內容先轉為字串
        df_raw_duty = pd.read_excel(uploaded_file, sheet_name="Duty List", header=None)
        df_raw_duty = df_raw_duty.fillna("nan").astype(str)
        
        # 從 A1 儲存格動態提取年份與月份
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
        df_raw_round = pd.read_excel(uploaded_file, sheet_name="Ward round", header=None)
        df_raw_round = df_raw_round.fillna("nan").astype(str)
        
        # 3. 讀取 "Call List" 分頁
        df_call = pd.read_excel(uploaded_file, sheet_name="Call List", header=None)
        df_call = df_call.fillna("nan").astype(str)
        
        # 從主表第 2 行提取所有員工簡寫名稱
        raw_names = df_raw_duty.iloc[1, 2:].tolist()
        ignored_tokens = ['nan', 'free', 'an', 'gyn', '']
        personnel_list = [str(name).strip() for name in raw_names if str(name).strip() and str(name).strip().lower() not in ignored_tokens]
        
        # UI 介面：選擇人員與開關 On-Call 功能
        selected_staff = st.selectbox("Select your name to extract duties:", sorted(list(set(personnel_list))))
        include_oncall = st.checkbox("Include On-Call Events (All-Day)", value=True)
        
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
                current_duty_date = None
                for r_idx in range(3, len(df_raw_duty)):
                    col_a = str(df_raw_duty.iloc[r_idx, 0]).strip()
                    col_b = str(df_raw_duty.iloc[r_idx, 1]).strip()
                    
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

            # --- 階段 C: 提取 On-Call 表 (Call List) ---
            if include_oncall:
                curr_month = month
                curr_year = year
                prev_day = None
                
                for r_idx in range(len(df_call)):
                    col_a = str(df_call.iloc[r_idx, 0]).strip()
                    
                    # 遇到下方非排班的備註欄位文字，直接安全中斷避免誤判
                    if "urgent call" in col_a.lower() or "bolded" in col_a.lower() or "from a&e" in col_a.lower():
                        break
                        
                    day_num = extract_day_num(col_a)
                    if day_num is not None:
                        # 處理跨月邏輯
                        if prev_day is None and day_num > 20:
                            curr_month = month - 1
                            if curr_month == 0:
                                curr_month = 12
                                curr_year -= 1
                        elif prev_day is not None and prev_day > 25 and day_num < 5:
                            curr_month += 1
                            if curr_month == 13:
                                curr_month = 1
                                curr_year += 1
                                
                        prev_day = day_num
                        call_date_str = datetime.date(curr_year, curr_month, day_num).strftime('%Y-%m-%d')
                        
                        # 掃描 E 到 J 欄 (index 4 to 9)
                        for c_idx in range(4, 10):
                            if c_idx < df_call.shape[1]:
                                cell_txt = str(df_call.iloc[r_idx, c_idx]).strip()
                                if cell_txt and cell_txt != 'nan':
                                    if selected_staff.lower() in cell_txt.lower():
                                        events.append({
                                            'all_day': True,
                                            'date': call_date_str,
                                            'summary': "On Call",
                                            'description': f"Call List notation: {cell_txt}"
                                        })
                                        break

            # --- 階段 D: 編譯並生成產出檔案 ---
            if len(events) > 0:
                lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Duty//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
                for e in events:
                    uid = f"{e['date'].replace('-', '')}-{uuid.uuid4().hex[:8]}@duty.local"
                    lines.append("BEGIN:VEVENT")
                    lines.append(f"UID:{uid}")
                    
                    if e.get('all_day'):
                        # 處理全天事件的格式
                        start_dt = datetime.datetime.strptime(e['date'], '%Y-%m-%d')
                        end_dt = start_dt + datetime.timedelta(days=1)
                        date_start_raw = start_dt.strftime('%Y%m%d')
                        date_end_raw = end_dt.strftime('%Y%m%d')
                        
                        lines.append(f"DTSTART;VALUE=DATE:{date_start_raw}")
                        lines.append(f"DTEND;VALUE=DATE:{date_end_raw}")
                    else:
                        # 處理特定時間事件的格式
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

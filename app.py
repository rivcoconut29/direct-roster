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
        personnel_options = sorted(list(set(personnel_list)))
        
        # --- Session State 記憶功能 ---
        if "selected_staff_name" not in st.session_state:
            st.session_state["selected_staff_name"] = personnel_options[0] if personnel_options else None
            
        default_index = 0
        if st.session_state["selected_staff_name"] in personnel_options:
            default_index = personnel_options.index(st.session_state["selected_staff_name"])
            
        selected_staff = st.selectbox(
            "Select your name to extract duties:", 
            personnel_options, 
            index=default_index
        )
        st.session_state["selected_staff_name"] = selected_staff
        
        # UI 控制元件
        include_leaves = st.checkbox("Include Leaves (e.g. AL, CO, OFF) as events", value=True)
        show_preop = st.checkbox("Show pre-op (general)", value=False)
        include_oncall = st.checkbox("Include On-Call Duties (All-Day)", value=False)
        
        if selected_staff:
            events = []
            round_records = {} 
            
            # --- 階段 A: 提取巡房表 (Ward Round) ---
            current_round_date = None
            # 固定忽略最後 7 列備忘錄
            for r_idx in range(4, len(df_raw_round) - 7):
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
                        cell_txt = str(df_raw_round.iloc[r_idx, c_idx]).strip()import streamlit as st
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
        personnel_options = sorted(list(set(personnel_list)))
        
        # --- Session State 記憶功能 ---
        if "selected_staff_name" not in st.session_state:
            st.session_state["selected_staff_name"] = personnel_options[0] if personnel_options else None
            
        default_index = 0
        if st.session_state["selected_staff_name"] in personnel_options:
            default_index = personnel_options.index(st.session_state["selected_staff_name"])
            
        selected_staff = st.selectbox(
            "Select your name to extract duties:", 
            personnel_options, 
            index=default_index
        )
        st.session_state["selected_staff_name"] = selected_staff
        
        # UI 控制元件
        include_leaves = st.checkbox("Include Leaves (e.g. AL, CO, OFF) as events", value=True)
        show_preop = st.checkbox("Show pre-op (general)", value=False)
        include_oncall = st.checkbox("Include On-Call Duties (All-Day)", value=False)
        
        if selected_staff:
            events = []
            round_records = {} 
            
            # --- 階段 A: 提取巡房表 (Ward Round) ---
            current_round_date = None
            # 固定忽略最後 7 列備忘錄
            for r_idx in range(4, len(df_raw_round) - 7):
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
                        if show_preop and duty_val.upper() == "OT" and any(day in col_b for day in ["Tue", "Wed"]):
                            preop_date_obj = current_duty_date_obj - datetime.timedelta(days=8)
                            events.append({
                                'all_day': False,
                                'date': preop_date_obj.strftime('%Y-%m-%d'),
                                'start_time': '09:30',
                                'end_time': '10:30',
                                'summary': 'Pre-op',
                                'description': f"Automated pre-op session for OT duty scheduled on {current_duty_date} ({col_b})"
                            })
                        
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

            # --- 階段 C: 提取 On-Call 表 (Call List) ---
            if include_oncall and "Call List" in xl.sheet_names:
                df_raw_call = xl.parse("Call List", header=None)
                df_raw_call = df_raw_call.fillna("nan").astype(str)
                
                curr_month = month
                curr_year = year
                prev_day = None
                
                for r_idx in range(4, len(df_raw_call)):
                    col_a = str(df_raw_call.iloc[r_idx, 0]).strip()
                    
                    if "urgent call" in col_a.lower() or "bolded" in col_a.lower() or "from a&e" in col_a.lower():
                        break
                        
                    day_num = extract_day_num(col_a)
                    if day_num is not None:
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
                        
                        try:
                            current_call_date = datetime.date(curr_year, curr_month, day_num).strftime('%Y-%m-%d')
                            
                            call_cols_range = range(4, df_raw_call.shape[1])
                            row_staff_tokens = [str(df_raw_call.iloc[r_idx, c]).strip() for c in call_cols_range]
                            row_staff_clean = [t.replace('*', '').strip().lower() for t in row_staff_tokens]
                            
                            if selected_staff.lower() in row_staff_clean:
                                other_personnel = []
                                for c_idx in call_cols_range:
                                    val = str(df_raw_call.iloc[r_idx, c_idx]).strip()
                                    if val and val.lower() not in ['nan', '']:
                                        if val.replace('*', '').strip().lower() != selected_staff.lower():
                                            other_personnel.append(val)
                                            
                                desc_str = f"Other on-call personnel: {', '.join(other_personnel)}" if other_personnel else "No other personnel listed."
                                
                                events.append({
                                    'all_day': True,
                                    'date': current_call_date,
                                    'summary': 'On Call',
                                    'description': desc_str
                                })
                        except ValueError:
                            pass

            # --- 階段 D: 生成 ICS 檔案 與 預覽 ---
            if len(events) > 0:
                sorted_events = sorted(events, key=lambda x: (x['date'], x.get('start_time', '00:00')))
                
                lines = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Duty//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH"]
                for e in sorted_events:
                    uid = f"{e['date'].replace('-', '')}-{uuid.uuid4().hex[:8]}@duty.local"
                    lines.append("BEGIN:VEVENT")
                    lines.append(f"UID:{uid}")
                    
                    date_raw = e['date'].replace('-', '')
                    if e['all_day']:
                        curr_date_obj = datetime.datetime.strptime(e['date'], '%Y-%m-%d').date()
                        next_date_obj = curr_date_obj + datetime.timedelta(days=1)
                        next_date_raw = next_date_obj.strftime('%Y%m%d')
                        lines.append(f"DTSTART;VALUE=DATE:{date_raw}")
                        lines.append(f"DTEND;VALUE=DATE:{next_date_raw}")
                    else:
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
                st.success(f"Successfully compiled {len(sorted_events)} calendar items for {selected_staff}!")
                
                st.download_button(
                    label=f"Download {selected_staff}'s .ics File",
                    data=ics_text,
                    file_name=f"Roster_{selected_staff}_{month}_{year}.ics",
                    mime="text/calendar"
                )
                
                # --- 網頁行事曆預覽 (已移除括號時間) ---
                st.write("---")
                st.subheader("Calendar Preview")
                for e in sorted_events:
                    summary_up = e['summary'].upper()
                    
                    core_duty = summary_up
                    if ": " in summary_up:
                        core_duty = summary_up.split(": ", 1)[1]
                    
                    if "ROUND:" in summary_up:
                        color_hex = "#2E7D32"  # 綠色
                    elif "ON CALL" in summary_up:
                        color_hex = "#D32F2F"  # 紅色
                    elif "PRE-OP" in summary_up:
                        color_hex = "#1976D2"  # 藍色
                    elif core_duty in ["OFF", "AL", "CO"]:
                        color_hex = "#757575"  # 灰色
                    elif "AM:" in summary_up or "PM:" in summary_up:
                        color_hex = "#B7950B"  # 深黃/金色
                    else:
                        color_hex = "#333333"
                        
                    # 直接渲染「日期 - 班表名稱」，移除了原本的 time_str
                    st.markdown(
                        f'<span style="color:{color_hex}; font-weight:bold;">{e["date"]}</span> '
                        f'<span style="color:{color_hex};"> - <strong>{e["summary"]}</strong></span>', 
                        unsafe_allow_html=True
                    )
                    if e.get('description'):
                        st.caption(f"Description: {e['description']}")
            else:
                st.warning(f"No active schedule items mapped for {selected_staff}.")
                
    except Exception as e:
        st.error(f"Error parsing structural workbook parameters: {e}")

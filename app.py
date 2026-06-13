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
        include_oncall = st.checkbox("Include On-Call Duties (All-Day)", value=False)
        
        if selected_staff:
            events = []
            round_records = {} 
            
            # --- 階段 A: 提取巡房表 (Ward Round) ---
            current_round_date = None
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
                leave_tokens =

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
            
            # --- 階段 A: 提取巡

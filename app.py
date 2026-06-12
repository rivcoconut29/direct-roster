import streamlit as st
import pandas as pd
import datetime
import uuid

st.title("Department Roster to Calendar Generator")
st.write("Upload your original messy Excel roster file to extract your customized calendar.")

uploaded_file = st.file_uploader("Upload Roster Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 1. Read Setup configuration for Year and Month parameters
        df_setup = pd.read_excel(uploaded_file, sheet_name="Setup")
        year = int(df_setup.iloc[0]['Year'])
        month = int(df_setup.iloc[0]['Month'])
        
        # 2. Define Hardcoded Ward Round Columns (0-indexed map from Excel)
        # Column C=2, D=3, F=5, G=6, H=7, I=8, J=9
        ward_column_map = {
            2: "Round: A11/E11",
            3: "Round: EPAC",
            5: "Round: C2",
            6: "Round: C2",
            7: "Round: E2",
            8: "Round: A2",
            9: "Round: A2"
        }
        
        # Read raw Ward Round data layer
        df_raw_round = pd.read_excel(uploaded_file, sheet_name="Ward round", header=None).astype(str)

        # 3. Read the Main "Duty List" sheet layout matrix
        df_raw_duty = pd.read_excel(uploaded_file, sheet_name="Duty List", header=None).astype(str)
        
        # Extract individual personnel tracking array from Row 2 (index 1), Column C onwards
        raw_names = df_raw_duty.iloc[1, 2:].tolist()
        ignored_tokens = ['nan', 'free', 'an', 'gyn']
        personnel_list = [name.strip() for name in raw_names if name.strip() and name.strip().lower() not in ignored_tokens]
        
        # Present interactive menu picker interface
        selected_staff = st.selectbox("Select your name to extract duties:", sorted(personnel_list))
        
        if selected_staff:
            events = []
            
            # --- PHASE A: Extract Ward Round Sheets via Hardcoded Columns ---
            current_round_date = None
            round_records = {} # Map dictionary log to track active dynamic AM shift overrides
            
            # Historical schedule lines start strictly at Row 5 (index 4)
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
                    
                # Iterate across specified targeted columns
                for c_idx, ward_name in ward_column_map.items():
                    if c_idx < df_raw_round.shape[1]:
                        cell_txt = str(df_raw_round.iloc[r_idx, c_idx]).strip()
                        if cell_txt and cell_txt != 'nan':
                            # Substring containment match to correctly capture cell suffixes/tags
                            if selected_staff.lower() in cell_txt.lower():
                                events.append({
                                    'date': current_round_date,
                                    'start_time': '08:30',
                                    'end_time': '10:30',
                                    'summary': ward_name,
                                    'description': f"Roster notation: {cell_txt}"
                                })
                                # Mark that this day has an active round assignment for this clinician
                                round_records[current_round_date] = True

            # --- PHASE B: Extract Main Duty List Matrix ---
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
                    
                    # Delimiter match to break execution past active calendars safely
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
                            # Condition match: Look up round tracking array record
                            has_round = current_duty_date in round_records
                            slot_info.update({'start_time': '10:30' if has_round else '08:30', 'end_time': '13:00'})
                            events.append(slot_info)
                        elif time_slot == "PM":
                            slot_info.update({'start_time': '14:00', 'end_time': '17:00'})
                            events.append(slot_info)

            # --- PHASE C: Compile and Generate Output File ---
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

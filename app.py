import io
from collections import defaultdict
from datetime import datetime, timedelta
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter

# --- ROBUST BANK HOLIDAY LOADING ---
try:
    import holidays
    uk_holidays = holidays.country_holidays('GB', subdiv='ENG')
except ImportError:
    uk_holidays = None

# --- Configuration ---
B_WORK_DAYS = [0, 1, 2, 3, 4]
DAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

DEFAULT_STAFF = [
    {"Initials": "GB", "Full Name": "Giuseppe Barone", "Active": True, "DIL Day": "Friday"},
    {"Initials": "OS", "Full Name": "Olga Slater", "Active": True, "DIL Day": "Tuesday"},
    {"Initials": "CD", "Full Name": "Cat Duncan", "Active": True, "DIL Day": "Tuesday"},
    {"Initials": "TC", "Full Name": "Tanzina Chowdhury", "Active": True, "DIL Day": "Tuesday"},
    {"Initials": "SD", "Full Name": "Sarita Depani", "Active": True, "DIL Day": "Monday"},
    {"Initials": "MJ", "Full Name": "Mette Jorgensen", "Active": True, "DIL Day": "Friday"},
    {"Initials": "KS", "Full Name": "Kavitha Srivatsa", "Active": True, "DIL Day": "Monday"},
    {"Initials": "TN", "Full Name": "Trung Nguyen", "Active": True, "DIL Day": "Wednesday"},
    {"Initials": "ES", "Full Name": "Elwira Szychot", "Active": True, "DIL Day": "Friday"},
]

def is_b_manual_period(dt):
    if isinstance(dt, pd.Timestamp): dt = dt.date()
    year = dt.year
    if dt.month == 1 and dt.day < 15: year -= 1
    
    # Monday strictly before Dec 25th
    dec_25 = datetime(year, 12, 25).date()
    days_back = dec_25.weekday() if dec_25.weekday() != 0 else 7
    m_start = dec_25 - timedelta(days=days_back)
    
    # Friday strictly after Jan 1st
    jan_1 = datetime(year + 1, 1, 1).date()
    days_forward = (4 - jan_1.weekday()) % 7
    if days_forward == 0: days_forward = 7
    m_end = jan_1 + timedelta(days=days_forward)
    
    return m_start <= dt <= m_end

def parse_previous_rota(uploaded_file):
    counts = defaultdict(lambda: {"A_days": 0, "B_days": 0, "Nights": 0})
    master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates = [], {}, [], None, None, {}
    try:
        xls = pd.ExcelFile(uploaded_file)
        if "Summary" in xls.sheet_names:
            df_sum = pd.read_excel(xls, sheet_name="Summary")
            for _, row in df_sum.iterrows():
                counts[str(row['Initials']).strip().upper()] = {"A_days": int(row.get('YTD A', 0)), "B_days": int(row.get('YTD B', 0)), "Nights": int(row.get('YTD Nights', 0))}
        if "State" in xls.sheet_names:
            state_dict = dict(zip(pd.read_excel(xls, sheet_name="State").iloc[:, 0], pd.read_excel(xls, sheet_name="State").iloc[:, 1]))
            def s_str(k): return str(state_dict.get(k)).strip() if pd.notna(state_dict.get(k)) else ""
            master_a_seq = [x.strip().upper() for x in s_str("A_MASTER_SEQ").split(",")] if s_str("A_MASTER_SEQ") else []
            if s_str("CYCLE_A_SCHEDULE"):
                for item in s_str("CYCLE_A_SCHEDULE").split(","):
                    if ":" in item:
                        k, v = item.split(":")
                        try: cycle_a_schedule[datetime.strptime(k, '%Y-%m-%d').date()] = v
                        except: pass
            b_queue = [x.strip().upper() for x in s_str("B_QUEUE").split(",")] if s_str("B_QUEUE") else []
            curr_a, curr_b = s_str("CURRENT_A"), s_str("CURRENT_B")
            if s_str("LAST_A_END_DATES"):
                for item in s_str("LAST_A_END_DATES").split(","):
                    if ":" in item:
                        k, d = item.split(":")
                        try: last_a_dates[k] = datetime.strptime(d, '%Y-%m-%d').date()
                        except: pass
    except: pass
    return counts, master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates

def get_leave_marker(df, date, ini):
    if not df.empty and date in df.index and ini in df.columns:
        val = df.loc[date, ini]
        if isinstance(val, pd.Series): val = val.iloc[0]
        return str(val).strip().upper() if pd.notna(val) else ""
    return ""

def generate_rota(leave_df, consultants, start_date, end_date, hist_counts, hist_master_a, hist_cycle_a, hist_b_q, hist_curr_a, hist_curr_b, hist_last_a_dates, locum_dict):
    active_initials = [str(c['Initials']).strip().upper() for c in consultants if c['Active']]
    name_map = {str(c['Initials']).strip().upper(): c['Full Name'] for c in consultants}
    counts = {ini: hist_counts.get(ini, {"A_days": 0, "B_days": 0, "Nights": 0}) for ini in active_initials}
    period_counts = {ini: {"A_days": 0, "B_days": 0, "Nights": 0} for ini in active_initials}
    
    def sync(h, a):
        nl = [i for i in h if i in a]
        for i in a:
            if i not in nl: nl.append(i)
        return nl

    master_a_seq = sync(hist_master_a, active_initials) if hist_master_a else active_initials.copy()
    cycle_a_schedule, b_queue = hist_cycle_a.copy(), sync(hist_b_q, active_initials)
    rota_data = []
    curr_d, fin_d = pd.to_datetime(start_date).normalize(), pd.to_datetime(end_date).normalize()
    cur_a, cur_b, last_a_end = hist_curr_a, hist_curr_b, hist_last_a_dates.copy()

    def is_clean(c, sd):
        if any(get_leave_marker(leave_df, sd + timedelta(days=o), c) in ["AL", "SL", "NO", "OFF"] for o in range(7)): return False
        if (sd.date() - last_a_end.get(c, datetime.min.date())).days < 22: return False
        return True

    while curr_d <= fin_d:
        if curr_d.month == 12 and curr_d.day == 24:
            for i in active_initials: counts[i] = {"A_days": 0, "B_days": 0, "Nights": 0}
            
        wd = curr_d.weekday()
        man_b = is_b_manual_period(curr_d)
        is_bh_today = uk_holidays is not None and curr_d.date() in uk_holidays
        
        w_start = curr_d - timedelta(days=wd)
        is_bh_week = False
        if uk_holidays is not None:
            for d in range(5):
                if (w_start + timedelta(days=d)).date() in uk_holidays:
                    is_bh_week = True
                    break

        active_a_nights = [0, 2, 4, 5, 6] if is_bh_week else [1, 3, 4, 5, 6]
        active_b_nights = [1, 3] if is_bh_week else [0, 2]
        
        # --- Consultant A Engine (V13: Continuous Dealing Shoe) ---
        if wd == 4 or cur_a is None:
            anc = curr_d - timedelta(days=(curr_d.weekday() - 4) % 7)
            if any(locum_dict.get((anc + timedelta(days=o)).date()) for o in range(7)): 
                cur_a = "LOCUM_ACADEMIC"
            else:
                if anc.date() not in cycle_a_schedule:
                    # 1. Project an infinite timeline of valid Friday slots
                    c_fs, chk = [], pd.to_datetime(start_date).normalize()
                    chk -= timedelta(days=(chk.weekday() - 4) % 7)
                    while chk <= fin_d + timedelta(days=365):
                        if not any(locum_dict.get((chk + timedelta(days=o)).date()) for o in range(7)): 
                            c_fs.append(chk)
                        chk += timedelta(days=7)
                    
                    temp = {f: {'doc': None, 'locked': False} for f in c_fs}

                    # 2. Lock Full 7-Day Heavy REQs
                    for f in c_fs:
                        for doc in active_initials:
                            req_count = sum(1 for o in range(7) if get_leave_marker(leave_df, f + timedelta(days=o), doc) == "REQ")
                            has_wknd = any(get_leave_marker(leave_df, f + timedelta(days=o), doc) == "REQ" and (f+timedelta(days=o)).weekday() in [5,6] for o in range(7))
                            
                            is_heavy = (req_count >= 4) or (req_count >= 3 and has_wknd)
                            if is_heavy and req_count == 7: # Only lock as base if full week requested
                                temp[f]['doc'] = doc
                                temp[f]['locked'] = True
                                break

                    # 3. The 4-Week Heavy REQ Forcefield Logic
                    def has_heavy_req_nearby(doc, target_f):
                        start_w = target_f - timedelta(days=4*7)
                        end_w = target_f + timedelta(days=4*7)
                        curr_w = start_w
                        while curr_w <= end_w:
                            r_count = sum(1 for o in range(7) if get_leave_marker(leave_df, curr_w + timedelta(days=o), doc) == "REQ")
                            h_wknd = any(get_leave_marker(leave_df, curr_w + timedelta(days=o), doc) == "REQ" and (curr_w+timedelta(days=o)).weekday() in [5,6] for o in range(7))
                            if (r_count >= 4) or (r_count >= 3 and h_wknd):
                                return True
                            curr_w += timedelta(days=7)
                        return False

                    # 4. Continuous Dealing Shoe Assignment
                    deck_idx = 0
                    for f in c_fs:
                        if temp[f]['locked']:
                            continue
                        
                        assigned = False
                        attempts = 0
                        # Try to find next unburned person (max attempts = 2 full deck lengths)
                        while not assigned and attempts < len(master_a_seq) * 2:
                            candidate = master_a_seq[deck_idx % len(master_a_seq)]
                            deck_idx += 1
                            attempts += 1
                            
                            # If they don't have a Heavy REQ within a 4-week radius, deal them in.
                            if not has_heavy_req_nearby(candidate, f):
                                temp[f]['doc'] = candidate
                                assigned = True
                                
                        if not assigned:
                            temp[f]['doc'] = "UNCOVERED"

                    # 5. Omnidirectional Swap Engine (Unlocked slots only)
                    def is_valid_for_slot(doc, slot_f):
                        if doc == "UNCOVERED": return False
                        if (slot_f.date() - last_a_end.get(doc, datetime.min.date())).days < 22: return False
                        if any(get_leave_marker(leave_df, slot_f + timedelta(days=o), doc) in ["AL", "SL", "NO", "OFF"] for o in range(7)): return False
                        for other_f, inf in temp.items():
                            if inf['doc'] == doc and other_f != slot_f:
                                if abs((slot_f.date() - other_f.date()).days) < 22: return False
                        window_start = slot_f - timedelta(days=22)
                        window_end = slot_f + timedelta(days=6 + 22)
                        for d_offset in range((window_end - window_start).days + 1):
                            chk_d = window_start + timedelta(days=d_offset)
                            if slot_f <= chk_d < slot_f + timedelta(days=7): continue
                            if get_leave_marker(leave_df, chk_d, doc) == "REQ": return False 
                        return True

                    for f_idx, f_dt in enumerate(c_fs):
                        if temp[f_dt]['locked']: continue
                        doc = temp[f_dt]['doc']
                        if doc and doc != "UNCOVERED" and not is_valid_for_slot(doc, f_dt):
                            swaps = sorted([(abs(idx - f_idx), f) for idx, f in enumerate(c_fs) if idx != f_idx and not temp[f]['locked'] and temp[f]['doc'] and temp[f]['doc'] != "UNCOVERED"], key=lambda x: x[0])
                            for _, of in swaps:
                                doc_1, doc_2 = temp[f_dt]['doc'], temp[of]['doc']
                                temp[f_dt]['doc'], temp[of]['doc'] = doc_2, doc_1
                                if is_valid_for_slot(doc_2, f_dt) and is_valid_for_slot(doc_1, of): break
                                else: temp[f_dt]['doc'], temp[of]['doc'] = doc_1, doc_2 
                                    
                    for f, inf in temp.items():
                        if inf['doc']: cycle_a_schedule[f.date()] = inf['doc']
                
                cur_a = cycle_a_schedule.get(anc.date(), "UNCOVERED")
                if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC"] and not is_clean(cur_a, anc): 
                    cur_a = "UNCOVERED"

        # --- Consultant B Queue Logic ---
        if wd == 0 or cur_b is None:
            if man_b:
                cur_b = "MANUAL ENTRY"
            else:
                b_anc = curr_d - timedelta(days=curr_d.weekday())
                pr_a = cycle_a_schedule.get((curr_d + timedelta(days=(4-curr_d.weekday()))).date(), cur_a)
                idl, part, tc = [], [], []
                for i, c in enumerate(b_queue):
                    if c != pr_a and get_leave_marker(leave_df, curr_d, c) != "REQ":
                        is_tc = (b_anc.date() - last_a_end.get(c, datetime.min.date())).days <= 7
                        is_cl = not any(get_leave_marker(leave_df, b_anc + timedelta(days=o), c) in ["AL", "SL", "NO", "OFF"] for o in range(5))
                        if is_tc: tc.append((i, c))
                        elif is_cl: idl.append((i, c))
                        else: part.append((i, c))
                if idl: cur_b = b_queue.pop(idl[0][0]); b_queue.append(cur_b)
                elif part: cur_b = b_queue.pop(part[0][0]); b_queue.append(cur_b)
                elif tc: cur_b = b_queue.pop(tc[0][0]); b_queue.append(cur_b)
                else: cur_b = None

        rq_d = next((i for i in active_initials if get_leave_marker(leave_df, curr_d, i) == "REQ"), None)
        lc_d = locum_dict.get(curr_d.date())
        
        if lc_d: da_n, da_i = lc_d, None
        elif rq_d: da_n, da_i = name_map.get(rq_d, rq_d), rq_d
        else: da_n, da_i = name_map.get(cur_a, cur_a), (cur_a if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC", "MANUAL ENTRY"] else None)

        row = {"Date": curr_d.strftime("%d/%m/%Y"), "Day": DAY_NAMES[wd], "Consultant A": da_n, "Consultant B": "", "Nights": ""}
        if da_i: counts[da_i]["A_days"] += 1; period_counts[da_i]["A_days"] += 1
        if wd in active_a_nights: 
            row["Nights"] = da_n
            if da_i: counts[da_i]["Nights"] += 1; period_counts[da_i]["Nights"] += 1

        if wd in B_WORK_DAYS:
            if is_bh_today:
                row["Consultant B"] = ""
                if wd in active_b_nights:
                    if man_b: row["Nights"] = "MANUAL ENTRY"
                    else:
                        b_night_name = name_map.get(cur_b, "UNCOVERED")
                        row["Nights"] = b_night_name
                        if cur_b: counts[cur_b]["Nights"] += 1; period_counts[cur_b]["Nights"] += 1
            else:
                if man_b:
                    row["Consultant B"] = "MANUAL ENTRY"
                    if wd in active_b_nights: row["Nights"] = "MANUAL ENTRY"
                else:
                    if not cur_b or get_leave_marker(leave_df, curr_d, cur_b) in ["AL", "SL", "NO", "OFF", "DIL", "REQ"] or cur_b == da_i:
                        avail = [i for i in active_initials if i not in [da_i, cur_b] and get_leave_marker(leave_df, curr_d, i) not in ["AL", "SL", "NO", "OFF", "DIL", "REQ"]]
                        if avail:
                            best = min([c for c in avail if (curr_d.date() - last_a_end.get(c, datetime.min.date())).days > 7 and c != cur_a] or avail, key=lambda x: counts[x]["B_days"])
                            row["Consultant B"] = name_map.get(best); counts[best]["B_days"] += 1; period_counts[best]["B_days"] += 1
                        else: row["Consultant B"] = "UNCOVERED"
                    else: row["Consultant B"] = name_map.get(cur_b, "UNCOVERED"); counts[cur_b]["B_days"] += 1; period_counts[cur_b]["B_days"] += 1
                    
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        b_i = next((k for k, v in name_map.items() if v == row["Consultant B"]), None)
                        if b_i: counts[b_i]["Nights"] += 1; period_counts[b_i]["Nights"] += 1

        for i in active_initials: row[i] = get_leave_marker(leave_df, curr_d, i)
        if da_i: last_a_end[da_i] = curr_d.date()
        rota_data.append(row); curr_d += timedelta(days=1)

    ex_c = fin_d + timedelta(days=1)
    d_map = {str(c['Initials']).strip().upper(): c.get('DIL Day', '') for c in consultants if c['Active']}
    for _ in range(365):
        w = ex_c.weekday()
        er = {"Date": ex_c.strftime("%d/%m/%Y"), "Day": DAY_NAMES[w], "Consultant A": "", "Consultant B": "", "Nights": ""}
        for i in active_initials: er[i] = "DIL" if d_map.get(i) == DAY_NAMES[w] else ""
        rota_data.append(er); ex_c += timedelta(days=1)

    state = {"A_MASTER_SEQ": ",".join(master_a_seq), "CYCLE_A_SCHEDULE": ",".join([f"{k.strftime('%Y-%m-%d')}:{v}" for k, v in cycle_a_schedule.items()]), "B_QUEUE": ",".join(b_queue), "CURRENT_A": cur_a or "", "CURRENT_B": cur_b or "", "LAST_A_END_DATES": ",".join([f"{k}:{v.strftime('%Y-%m-%d')}" for k, v in last_a_end.items()])}
    summ = [{"Initials": k, "Period A": period_counts[k]["A_days"], "Period B": period_counts[k]["B_days"], "Period Nights": period_counts[k]["Nights"], "Period Total (A+N)": period_counts[k]["A_days"] + period_counts[k]["Nights"], "YTD A": counts[k]["A_days"], "YTD B": counts[k]["B_days"], "YTD Nights": counts[k]["Nights"], "YTD Total (A+N)": counts[k]["A_days"] + counts[k]["Nights"]} for k in active_initials]
    return pd.DataFrame(rota_data), pd.DataFrame(summ), state

def main():
    st.set_page_config(layout="wide", page_title="Consultant Rota Tool")
    st.markdown("<style>.stApp {background-color: #fbfbfd !important;} p, h1, h2, h3, span, label, div {color: #1d1d1f; font-family: -apple-system, sans-serif;} div.stButton > button {background-color: #0071e3 !important; color: #fff !important; border-radius: 20px !important;}</style>", unsafe_allow_html=True)
    st.title("Consultant Rota Dashboard")
    
    if 'rota_result' not in st.session_state: st.session_state.rota_result = None
    if 'staff' not in st.session_state: st.session_state.staff = pd.DataFrame(DEFAULT_STAFF).reset_index(drop=True)
    if 'locums' not in st.session_state: 
        st.session_state.locums = pd.DataFrame({"Start Date": pd.Series(dtype='datetime64[ns]'), "End Date": pd.Series(dtype='datetime64[ns]'), "Locum Name": pd.Series(dtype='str')}).reset_index(drop=True)
    
    st.sidebar.header("1. Consultant Roster")
    edited_staff = st.sidebar.data_editor(st.session_state.staff, num_rows="dynamic", column_config={"DIL Day": st.column_config.SelectboxColumn("DIL Day", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"])})
    
    st.sidebar.header("2. Locum / Academic Overrides")
    st.sidebar.markdown("<p style='color:#86868b; font-size:14px;'>Leave 'End Date' blank for a 1-day assignment.</p>", unsafe_allow_html=True)
    edited_locums = st.sidebar.data_editor(st.session_state.locums, num_rows="dynamic", column_config={"Start Date": st.column_config.DateColumn("Start Date", format="DD/MM/YYYY"), "End Date": st.column_config.DateColumn("End Date", format="DD/MM/YYYY")}, key="locum_editor")
    
    if not edited_locums.equals(st.session_state.locums):
        changes = False
        for idx, row in edited_locums.iterrows():
            if pd.notna(row["Start Date"]) and pd.isna(row["End Date"]):
                edited_locums.at[idx, "End Date"] = row["Start Date"]
                changes = True
        if changes:
            st.session_state.locums = edited_locums
            st.rerun()

    c1, c2 = st.columns(2)
    with c1: leave_f = st.file_uploader("Upload Leave Requests (Excel) - Optional", type=["xlsx"])
    with c2: prev_f = st.file_uploader("Upload Previous Rota (Memory) - Optional", type=["xlsx"])

    def_s, def_e, df_l = datetime.today().date(), datetime.today().date() + timedelta(days=28), pd.DataFrame()
    if leave_f:
        try:
            df_l = pd.read_excel(leave_f); df_l.columns = [str(c).strip().upper() for c in df_l.columns]
            d_col = next((c for c in df_l.columns if "DATE" in c), df_l.columns[1])
            df_l[d_col] = pd.to_datetime(df_l[d_col], dayfirst=True, errors='coerce').dt.normalize()
            df_l = df_l.dropna(subset=[d_col]).set_index(d_col)
            if not df_l.empty: def_s, def_e = df_l.index.min().date(), df_l.index.max().date()
        except: st.error("Error loading leave file.")
            
    col_s, col_e = st.columns(2)
    with col_s: sel_s = st.date_input("Start Rota:", value=def_s, format="DD/MM/YYYY")
    with col_e: sel_e = st.date_input("End Rota:", value=max(def_e, sel_s), min_value=sel_s, format="DD/MM/YYYY")

    counts, m_seq, c_sch, bq, ca, cb, h_last = parse_previous_rota(prev_f) if prev_f else ({}, [], {}, [], None, None, {})
    
    locum_dict = {}
    for _, row in edited_locums.iterrows():
        if pd.notna(row["Start Date"]) and str(row["Locum Name"]).strip() and str(row["Locum Name"]).strip() != 'nan':
            s = pd.to_datetime(row["Start Date"]).date()
            e = pd.to_datetime(row["End Date"]).date() if pd.notna(row["End Date"]) else s
            c = s
            while c <= e:
                locum_dict[c] = str(row["Locum Name"]).strip()
                c += timedelta(days=1)

    if st.button("Generate Final Rota", type="primary"):
        rota, summ, state = generate_rota(df_l, edited_staff.to_dict('records'), sel_s, sel_e, counts, m_seq, c_sch, bq, ca, cb, h_last, locum_dict)
        st.session_state.rota_result = (rota, summ, state, sel_e)

    if st.session_state.rota_result:
        rota, summ, state, sel_e = st.session_state.rota_result
        def highlight(row):
            dt = pd.to_datetime(row['Date'], format="%d/%m/%Y")
            if dt.weekday() >= 5 or (uk_holidays is not None and dt.date() in uk_holidays): return ['background-color: #e5e5ea;'] * len(row)
            return [''] * len(row)
        
        st.subheader("Rota for the Period")
        try:
            target_str = pd.to_datetime(sel_e).strftime("%d/%m/%Y")
            l_idx = rota[rota['Date'] == target_str].index[-1]
            st.dataframe(rota.iloc[:l_idx+1].style.apply(highlight, axis=1), use_container_width=True, height=450)
        except: st.dataframe(rota.style.apply(highlight, axis=1), use_container_width=True, height=450)
        
        st.subheader("Fairness Calculator for Period and YTD")
        st.dataframe(summ, use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            rota.to_excel(writer, index=False, sheet_name='Rota')
            summ.to_excel(writer, index=False, sheet_name='Summary')
            pd.DataFrame(list(state.items()), columns=["Key", "Value"]).to_excel(writer, index=False, sheet_name='State')
            pd.DataFrame([["AL", "Annual Leave"], ["SL", "Study Leave"], ["OFF", "Off for other reasons"], ["NO", "Please do not put me on call"], ["REQ", "Please add me on this week"]], columns=["Marker", "Description"]).to_excel(writer, index=False, sheet_name='Legend')
            
            ws = writer.sheets['Rota']
            ws.freeze_panes = "C2"
            align_left = Alignment(horizontal='left', vertical='center')
            for i, col_cells in enumerate(ws.columns, 1):
                col_name = ws.cell(row=1, column=i).value
                vals = [len(str(cell.value)) for cell in col_cells if cell.value]
                max_l = max(vals + [len(str(col_name)) if col_name else 5])
                ws.column_dimensions[get_column_letter(i)].width = max_l + 2
                for cell in col_cells: cell.alignment = align_left
            
            fill = PatternFill(start_color="E5E5EA", end_color="E5E5EA", fill_type="solid")
            for i, r in enumerate(rota.itertuples(), start=2):
                dt = pd.to_datetime(r.Date, format="%d/%m/%Y")
                if dt.weekday() >= 5 or (uk_holidays is not None and dt.date() in uk_holidays):
                    for c in range(1, len(rota.columns) + 1): ws.cell(row=i, column=c).fill = fill
        
        st.download_button("Download Rota", out.getvalue(), f"Rota_{datetime.now().strftime('%d_%m')}.xlsx", type="primary")

if __name__ == "__main__": main()

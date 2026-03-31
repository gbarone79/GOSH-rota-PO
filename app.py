import io
import random
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
    {"Initials": "GB", "Full Name": "Giuseppe Barone", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "SD", "Full Name": "Sarita Depani", "Active": True, "DIL Day": "Monday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "OS", "Full Name": "Olga Slater", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "MJ", "Full Name": "Mette Jorgensen", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "CD", "Full Name": "Cat Duncan", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "KS", "Full Name": "Kavitha Srivatsa", "Active": True, "DIL Day": "Monday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "TC", "Full Name": "Tanzina Chowdhury", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "ES", "Full Name": "Elwira Szychot", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0},
    {"Initials": "TN", "Full Name": "Trung Nguyen", "Active": True, "DIL Day": "Wednesday", "FTE A": 1.0, "FTE B": 1.0},
]

def to_date(d):
    return d.date() if hasattr(d, 'date') else d

def is_b_manual_period(dt):
    dt_safe = to_date(dt)
    year = dt_safe.year
    if dt_safe.month == 1 and dt_safe.day < 15: year -= 1
    
    dec_25 = datetime(year, 12, 25).date()
    days_back = dec_25.weekday() if dec_25.weekday() != 0 else 7
    m_start = dec_25 - timedelta(days=days_back)
    
    jan_1 = datetime(year + 1, 1, 1).date()
    days_forward = (4 - jan_1.weekday()) % 7
    if days_forward == 0: days_forward = 7
    m_end = jan_1 + timedelta(days=days_forward)
    
    return m_start <= dt_safe <= m_end

def match_name(val, consultants):
    val = str(val).strip().upper()
    if not val or val in ['NAN', 'NONE', 'NAT', '']: return None
    for c in consultants:
        ini = str(c['Initials']).strip().upper()
        fname = str(c['Full Name']).strip().upper()
        if val == ini or val == fname: return ini
        if val in fname.split(): return ini
        if fname in val or val in fname: return ini
    return None

def parse_previous_rota(uploaded_file, consultants, sel_s, ytd_s, ytd_e):
    counts = defaultdict(lambda: {"A_days": 0, "B_days": 0, "Nights": 0})
    master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates = [], {}, [], None, None, {}
    hist_b_dates = {}
    prev_rota_df = None
    master_a_seq_temp = []
    memory_leave_dict = {}
    
    if not uploaded_file: 
        return counts, master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates, hist_b_dates, prev_rota_df, pd.DataFrame()

    active_initials = [str(c['Initials']).strip().upper() for c in consultants if c.get('Active', True)]
    name_map = {str(c['Initials']).strip().upper(): str(c['Full Name']).strip() for c in consultants}

    def safe_raw(val):
        if pd.isna(val): return ""
        s = str(val).strip()
        if s.upper() in ["NAN", "NONE", "NAT", ""]: return ""
        return s

    try:
        xls = pd.ExcelFile(uploaded_file)
        
        if "Rota" in xls.sheet_names:
            df_rota = pd.read_excel(xls, sheet_name="Rota")
            clean_rota_data = [] 
            
            date_col = next((c for c in df_rota.columns if str(c).strip().upper() == "DATE"), None)
            if not date_col:
                if "Day" in df_rota.columns and df_rota.columns.get_loc("Day") > 0:
                    date_col = df_rota.columns[df_rota.columns.get_loc("Day") - 1]
                elif len(df_rota.columns) > 1:
                    date_col = df_rota.columns[1]

            col_a = next((c for c in df_rota.columns if str(c).strip().upper() in ["CONSULTANT A", "ATTENDING"]), None)
            col_b = next((c for c in df_rota.columns if str(c).strip().upper() in ["CONSULTANT B", "B CONSULTANT"]), None)
            col_n = next((c for c in df_rota.columns if str(c).strip().upper() in ["NIGHTS", "NIGHT"]), None)
            
            for _, row in df_rota.iterrows():
                try:
                    d_val = row[date_col] if date_col and date_col in row else None
                    if pd.isna(d_val): continue
                    
                    if isinstance(d_val, str): 
                        try: r_date = pd.to_datetime(d_val, dayfirst=True).date()
                        except: continue
                    else: r_date = d_val.date() if hasattr(d_val, 'date') else d_val

                    a_ini = match_name(row[col_a], consultants) if col_a else None
                    b_ini = match_name(row[col_b], consultants) if col_b else None
                    n_ini = match_name(row[col_n], consultants) if col_n else None

                    if ytd_s <= r_date <= ytd_e and r_date < sel_s:
                        if a_ini: counts[a_ini]["A_days"] += 1
                        if b_ini: counts[b_ini]["B_days"] += 1
                        if n_ini: counts[n_ini]["Nights"] += 1
                        
                    if r_date < sel_s:
                        if a_ini:
                            cycle_a_schedule[r_date] = a_ini
                            last_a_dates[a_ini] = max(last_a_dates.get(a_ini, datetime.min.date()), r_date)
                            if not master_a_seq_temp or master_a_seq_temp[-1] != a_ini:
                                master_a_seq_temp.append(a_ini)
                        if b_ini:
                            hist_b_dates[b_ini] = max(hist_b_dates.get(b_ini, datetime.min.date()), r_date)
                        if n_ini:
                            hist_b_dates[n_ini] = max(hist_b_dates.get(n_ini, datetime.min.date()), r_date)

                    leave_for_date = {}
                    for ini in active_initials:
                        leave_val = ""
                        for col in df_rota.columns:
                            if str(col).strip().upper() == ini or str(col).strip().upper() == str(name_map.get(ini, "")).upper():
                                leave_val = safe_raw(row.get(col))
                                break
                        leave_for_date[ini] = leave_val
                    
                    memory_leave_dict[pd.to_datetime(r_date)] = leave_for_date

                    clean_row = {
                        "Date": r_date.strftime("%d/%m/%Y"),
                        "Day": DAY_NAMES[r_date.weekday()],
                        "Consultant A": name_map.get(a_ini, safe_raw(row.get(col_a))) if col_a else "",
                        "Consultant B": name_map.get(b_ini, safe_raw(row.get(col_b))) if col_b else "",
                        "Nights": name_map.get(n_ini, safe_raw(row.get(col_n))) if col_n else ""
                    }
                    
                    for ini in active_initials:
                        clean_row[ini] = leave_for_date[ini]
                        
                    clean_rota_data.append(clean_row)

                except Exception:
                    pass
            
            if clean_rota_data:
                df_temp = pd.DataFrame(clean_rota_data)
                mask = (df_temp['Consultant A'] != '') | (df_temp['Consultant B'] != '') | (df_temp['Nights'] != '')
                if mask.any():
                    last_idx = mask[::-1].idxmax()
                    prev_rota_df = df_temp.loc[:last_idx].copy()
                else:
                    prev_rota_df = pd.DataFrame()

        if "State" in xls.sheet_names:
            df_state = pd.read_excel(xls, sheet_name="State")
            if len(df_state.columns) >= 2:
                state_dict = dict(zip(df_state.iloc[:, 0].astype(str).str.strip(), df_state.iloc[:, 1]))
                def s_str(k): 
                    val = state_dict.get(k)
                    return str(val).strip() if pd.notna(val) and str(val).strip().upper() not in ['NAN', 'NONE'] else ""
                
                master_a_seq = [x.strip().upper() for x in s_str("A_MASTER_SEQ").split(",") if x.strip()]
                b_queue = [x.strip().upper() for x in s_str("B_QUEUE").split(",") if x.strip()]
                curr_a = s_str("CURRENT_A") or None
                curr_b = s_str("CURRENT_B") or None
                
                last_a_str = s_str("LAST_A_END_DATES")
                if last_a_str:
                    for item in last_a_str.split(","):
                        if ":" in item:
                            k, d = item.split(":")
                            try: last_a_dates[k.strip()] = datetime.strptime(d.strip(), '%Y-%m-%d').date()
                            except ValueError: pass
                
        if not master_a_seq and master_a_seq_temp:
            seen = set()
            master_a_seq = [x for x in reversed(master_a_seq_temp) if not (x in seen or seen.add(x))]
            master_a_seq.reverse()

    except Exception as e:
        st.warning(f"Warning: Could not fully parse the uploaded Memory file. Proceeding safely.")
        
    memory_leave_df = pd.DataFrame.from_dict(memory_leave_dict, orient='index') if memory_leave_dict else pd.DataFrame()
    return counts, master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates, hist_b_dates, prev_rota_df, memory_leave_df

def get_leave_marker(df, d_date, ini):
    if df.empty or ini not in df.columns: return ""
    try:
        ts = pd.Timestamp(d_date).normalize()
        if ts in df.index:
            val = df.loc[ts, ini]
            if isinstance(val, pd.Series): val = val.iloc[0]
            if pd.isna(val): return ""
            val_str = str(val).strip().upper()
            if val_str in ["NAN", "NAT", "NONE", "NULL", ""]: return ""
            return val_str
    except Exception:
        pass
    return ""

def get_group(doc_ini):
    if doc_ini in ["ES", "SD", "MJ", "KS"]: return "NEURO"
    if doc_ini in ["GB", "OS", "CD", "TC", "TN"]: return "SOLID"
    return "UNKNOWN"

def generate_rota(leave_df, consultants, start_date, end_date, hist_counts, hist_master_a, hist_cycle_a, hist_b_q, hist_curr_a, hist_curr_b, hist_last_a_dates, locum_dict, hist_b_dates, ytd_s, ytd_e):
    active_initials = [str(c['Initials']).strip().upper() for c in consultants if c.get('Active', True)]
    name_map = {str(c['Initials']).strip().upper(): c['Full Name'] for c in consultants}
    
    # Extract FTE Multipliers
    fte_a_map = {}
    fte_b_map = {}
    d_map = {}
    
    for c in consultants:
        if c.get('Active', True):
            ini = str(c.get('Initials', '')).strip().upper()
            
            raw_dil = c.get('DIL Day')
            d_map[ini] = "" if pd.isna(raw_dil) or raw_dil is None or str(raw_dil).strip().upper() in ["NAN", "NONE", "NAT"] else str(raw_dil).strip().capitalize()
            
            # Floor of 0.1 prevents math crashes
            fte_a_map[ini] = max(0.1, float(c.get('FTE A', 1.0)))
            fte_b_map[ini] = max(0.1, float(c.get('FTE B', 1.0)))

    counts = {ini: {
        "A_days": hist_counts.get(ini, {}).get("A_days", 0), 
        "B_days": hist_counts.get(ini, {}).get("B_days", 0), 
        "Nights": hist_counts.get(ini, {}).get("Nights", 0)
    } for ini in active_initials}
    
    period_counts = {ini: {"A_days": 0, "B_days": 0, "Nights": 0} for ini in active_initials}
    recent_b_shifts = hist_b_dates.copy() 
    
    def get_eff_marker(d_date, ini):
        lm = get_leave_marker(leave_df, d_date, ini)
        if lm: return lm
        if d_map.get(ini) == DAY_NAMES[to_date(d_date).weekday()]: return "DIL"
        return ""

    def sync(h, a):
        nl = [i for i in h if i in a]
        for i in a:
            if i not in nl: nl.append(i)
        return nl

    if hist_master_a:
        master_a_seq = sync(hist_master_a, active_initials)
    else:
        solids = [i for i in active_initials if get_group(i) == "SOLID"]
        neuros = [i for i in active_initials if get_group(i) == "NEURO"]
        random.seed() 
        random.shuffle(solids)
        random.shuffle(neuros)
        master_a_seq = []
        while solids or neuros:
            if solids: master_a_seq.append(solids.pop(0))
            if neuros: master_a_seq.append(neuros.pop(0))
            
    cycle_a_schedule, b_queue = hist_cycle_a.copy(), sync(hist_b_q, active_initials)
    
    curr_d_start = pd.to_datetime(start_date).normalize()
    fin_d = pd.to_datetime(end_date).normalize()
    start_anc = (curr_d_start - timedelta(days=(curr_d_start.weekday() - 4) % 7)).date()
    
    cycle_a_schedule = {k: v for k, v in cycle_a_schedule.items() if k < start_anc}
    last_a_end = {k: v for k, v in hist_last_a_dates.items() if v < start_anc}

    c_fs = []
    chk = start_anc
    while chk <= to_date(fin_d) + timedelta(days=365):
        if not any(locum_dict.get(chk + timedelta(days=o)) for o in range(7)): 
            c_fs.append(chk)
        chk += timedelta(days=7)
    
    temp = {f: {'doc': None, 'locked': False} for f in c_fs}

    for f in c_fs:
        for doc in active_initials:
            req_count = sum(1 for o in range(7) if get_eff_marker(f + timedelta(days=o), doc) == "REQ")
            has_wknd = any(get_eff_marker(f + timedelta(days=o), doc) == "REQ" and (f+timedelta(days=o)).weekday() in [5,6] for o in range(7))
            is_heavy = (req_count >= 4) or (req_count >= 3 and has_wknd)
            if is_heavy and req_count == 7: 
                temp[f]['doc'] = doc
                temp[f]['locked'] = True
                break

    deck_idx = 0
    if cycle_a_schedule:
        last_base_date = max(cycle_a_schedule.keys())
        last_doc = cycle_a_schedule[last_base_date]
        if last_doc in master_a_seq:
            deck_idx = (master_a_seq.index(last_doc) + 1) % len(master_a_seq)

    # LIVE TRACKER: Follows pacing inside the loop as assignments happen
    sim_a_counts = {ini: counts[ini]["A_days"] for ini in active_initials}

    # --- THE NEW A-WEEK ENGINE: Live Pacing & Zero Swapping ---
    for f in c_fs:
        if temp[f]['locked']:
            sim_a_counts[temp[f]['doc']] += 1
            continue

        assigned = False
        attempts = 0
        
        # PASS 1: Strict Check (Respects FTE, Leave, and 28-day gap)
        while not assigned and attempts < len(master_a_seq) * 2:
            candidate = master_a_seq[deck_idx % len(master_a_seq)]
            deck_idx += 1
            attempts += 1
            
            # Pacing Tracker: If they are running 1 FTE equivalent week ahead, SKIP them.
            norm_c = sim_a_counts[candidate] / fte_a_map[candidate]
            min_norm = min(sim_a_counts[i] / fte_a_map[i] for i in active_initials)
            if norm_c - min_norm >= 1.0: 
                continue
                
            # Leave rules
            if any(get_eff_marker(f + timedelta(days=o), candidate) in ["AL", "SL", "NO", "OFF"] for o in range(7)):
                continue
                
            # 28-day rules (checks history AND current future assignments)
            too_close = False
            if (f - last_a_end.get(candidate, datetime.min.date())).days < 28:
                too_close = True
            for other_f, inf in temp.items():
                if inf['doc'] == candidate and abs((f - other_f).days) < 28:
                    too_close = True
            for past_date, past_doc in cycle_a_schedule.items():
                if past_doc == candidate and abs((f - past_date).days) < 28:
                    too_close = True
            
            if too_close:
                continue
                
            temp[f]['doc'] = candidate
            sim_a_counts[candidate] += 1
            assigned = True

        # PASS 2: Relaxed Check (If rules blocked everyone, ignore FTE pacing to guarantee coverage)
        if not assigned:
            attempts = 0
            while not assigned and attempts < len(master_a_seq):
                candidate = master_a_seq[deck_idx % len(master_a_seq)]
                deck_idx += 1
                attempts += 1
                
                if any(get_eff_marker(f + timedelta(days=o), candidate) in ["AL", "SL", "NO", "OFF"] for o in range(7)):
                    continue
                    
                too_close = False
                if (f - last_a_end.get(candidate, datetime.min.date())).days < 28:
                    too_close = True
                for other_f, inf in temp.items():
                    if inf['doc'] == candidate and abs((f - other_f).days) < 28:
                        too_close = True
                for past_date, past_doc in cycle_a_schedule.items():
                    if past_doc == candidate and abs((f - past_date).days) < 28:
                        too_close = True
                
                if too_close:
                    continue
                    
                temp[f]['doc'] = candidate
                sim_a_counts[candidate] += 1
                assigned = True

        if not assigned: 
            temp[f]['doc'] = "UNCOVERED"

    for f, inf in temp.items():
        if inf['doc']: cycle_a_schedule[f] = inf['doc']
        
    rota_data = []
    curr_d = curr_d_start
    cur_a, cur_b = hist_curr_a, hist_curr_b

    def is_clean(c, sd):
        if any(get_eff_marker(sd + timedelta(days=o), c) in ["AL", "SL", "NO", "OFF"] for o in range(7)): return False
        if (to_date(sd) - last_a_end.get(c, datetime.min.date())).days < 28: return False
        return True

    def b_penalty(doc, d):
        d_date = to_date(d)
        if (d_date - recent_b_shifts.get(doc, datetime.min.date())).days <= 21: return 1
        next_fri = d_date + timedelta(days=(4-d_date.weekday())%7)
        if cycle_a_schedule.get(next_fri) == doc or cycle_a_schedule.get(next_fri + timedelta(days=7)) == doc: return 1
        return 0

    def add_score(doc_ini, s_type, d_date):
        if doc_ini and doc_ini in period_counts:
            period_counts[doc_ini][s_type] += 1
            if ytd_s <= to_date(d_date) <= ytd_e:
                counts[doc_ini][s_type] += 1

    # Apply FTE Math to the B-Queue sorting mechanism
    norm_b = lambda x: counts[x]["B_days"] / fte_b_map[x]
    norm_n = lambda x: counts[x]["Nights"] / fte_a_map[x]

    while curr_d <= fin_d:            
        wd = curr_d.weekday()
        man_b = is_b_manual_period(curr_d)
        is_bh_today = uk_holidays is not None and to_date(curr_d) in uk_holidays
        
        w_start = curr_d - timedelta(days=wd)
        is_bh_week = False
        if uk_holidays is not None:
            for d in range(5):
                if to_date(w_start + timedelta(days=d)) in uk_holidays:
                    is_bh_week = True
                    break

        active_a_nights = [0, 2, 4, 5, 6] if is_bh_week else [1, 3, 4, 5, 6]
        active_b_nights = [1, 3] if is_bh_week else [0, 2]
        
        if wd == 4 or cur_a is None:
            anc = curr_d - timedelta(days=(curr_d.weekday() - 4) % 7)
            if any(locum_dict.get(to_date(anc + timedelta(days=o))) for o in range(7)): 
                cur_a = "LOCUM_ACADEMIC"
            else:
                cur_a = cycle_a_schedule.get(to_date(anc), "UNCOVERED")
                if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC"] and not is_clean(cur_a, anc): 
                    cur_a = "UNCOVERED"

        if wd == 0 or cur_b is None or cur_b == "MANUAL ENTRY":
            if man_b:
                cur_b = "MANUAL ENTRY"
            else:
                b_anc = curr_d - timedelta(days=curr_d.weekday())
                pr_a = cur_a 
                
                idl, part, tc = [], [], []
                for i, c in enumerate(b_queue):
                    if c != pr_a and get_eff_marker(curr_d, c) != "REQ":
                        is_tc = (to_date(b_anc) - last_a_end.get(c, datetime.min.date())).days <= 14
                        is_cl = not any(get_eff_marker(b_anc + timedelta(days=o), c) in ["AL", "SL", "NO", "OFF"] for o in range(5))
                        if is_tc: tc.append((i, c))
                        elif is_cl: idl.append((i, c))
                        else: part.append((i, c))
                
                pr_a_group = get_group(pr_a)
                
                def sort_b_queue(q_sub):
                    if not q_sub: return q_sub
                    
                    if wd in active_b_nights:
                        avg_nights = sum(norm_n(c) for c in b_queue) / len(b_queue) if b_queue else 0
                        min_nights = min([norm_n(c) for c in b_queue]) if b_queue else 0
                        min_b_days = min([norm_b(c) for c in b_queue]) if b_queue else 0
                        FAIRNESS_GAP = 1.5
                        
                        q_sub.sort(key=lambda x: (
                            b_penalty(x[1], curr_d),                      
                            1 if (norm_b(x[1]) - min_b_days) >= FAIRNESS_GAP else 0,
                            0 if (d_map.get(x[1]) in ["Monday", "Wednesday"] and norm_n(x[1]) < avg_nights) else 1,
                            1 if (norm_n(x[1]) - min_nights) >= FAIRNESS_GAP else 0,
                            0 if get_group(x[1]) != pr_a_group else 1,    
                            norm_n(x[1]),                       
                            x[0]                                          
                        ))
                    else:
                        min_b_days = min([norm_b(c) for c in b_queue]) if b_queue else 0
                        FAIRNESS_GAP = 1.5
                        q_sub.sort(key=lambda x: (
                            b_penalty(x[1], curr_d),                     
                            1 if (norm_b(x[1]) - min_b_days) >= FAIRNESS_GAP else 0,
                            0 if get_group(x[1]) != pr_a_group else 1,   
                            norm_b(x[1]),                      
                            x[0] 
                        ))
                    return q_sub
                
                if idl: 
                    idl = sort_b_queue(idl)
                    cur_b = b_queue.pop(b_queue.index(idl[0][1])); b_queue.append(cur_b)
                elif part: 
                    part = sort_b_queue(part)
                    cur_b = b_queue.pop(b_queue.index(part[0][1])); b_queue.append(cur_b)
                elif tc: 
                    tc = sort_b_queue(tc)
                    cur_b = b_queue.pop(b_queue.index(tc[0][1])); b_queue.append(cur_b)
                else: cur_b = None

        rq_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQ"), None)
        lc_d = locum_dict.get(to_date(curr_d))
        
        if lc_d: da_n, da_i = lc_d, None
        elif rq_d: da_n, da_i = name_map.get(rq_d, rq_d), rq_d
        else: da_n, da_i = name_map.get(cur_a, cur_a), (cur_a if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC", "MANUAL ENTRY"] else None)

        row = {"Date": curr_d.strftime("%d/%m/%Y"), "Day": DAY_NAMES[wd], "Consultant A": da_n, "Consultant B": "", "Nights": ""}
        if da_i: add_score(da_i, "A_days", curr_d)
        if wd in active_a_nights: 
            row["Nights"] = da_n
            if da_i: add_score(da_i, "Nights", curr_d)

        if wd in B_WORK_DAYS:
            if is_bh_today:
                row["Consultant B"] = ""
                if wd in active_b_nights:
                    if man_b: row["Nights"] = "MANUAL ENTRY"
                    else:
                        b_night_name = name_map.get(cur_b, "UNCOVERED")
                        row["Nights"] = b_night_name
                        if cur_b: add_score(cur_b, "Nights", curr_d)
            else:
                if man_b:
                    row["Consultant B"] = "MANUAL ENTRY"
                    if wd in active_b_nights: row["Nights"] = "MANUAL ENTRY"
                else:
                    if not cur_b or get_eff_marker(curr_d, cur_b) in ["AL", "SL", "NO", "OFF", "DIL", "REQ"] or cur_b == da_i:
                        avail = [i for i in active_initials if i not in [da_i, cur_b] and get_eff_marker(curr_d, i) not in ["AL", "SL", "NO", "OFF", "DIL", "REQ"]]
                        if avail:
                            safe = [c for c in avail if (to_date(curr_d) - last_a_end.get(c, datetime.min.date())).days >= 14 and c != cur_a]
                            pool = safe if safe else avail
                            a_group = get_group(da_i or cur_a)
                            
                            if wd in active_b_nights:
                                avg_nights = sum(norm_n(c) for c in pool) / len(pool) if pool else 0
                                min_nights = min([norm_n(c) for c in pool]) if pool else 0
                                min_b_days = min([norm_b(c) for c in pool]) if pool else 0
                                FAIRNESS_GAP = 1.5
                                
                                best = min(pool, key=lambda x: (
                                    b_penalty(x, curr_d),                     
                                    1 if (norm_b(x) - min_b_days) >= FAIRNESS_GAP else 0,
                                    0 if (d_map.get(x) in ["Monday", "Wednesday"] and norm_n(x) < avg_nights) else 1,
                                    1 if (norm_n(x) - min_nights) >= FAIRNESS_GAP else 0,
                                    0 if get_group(x) != a_group else 1,      
                                    norm_n(x)                      
                                ))
                            else:
                                min_b_days = min([norm_b(c) for c in pool]) if pool else 0
                                FAIRNESS_GAP = 1.5
                                best = min(pool, key=lambda x: (
                                    b_penalty(x, curr_d),                     
                                    1 if (norm_b(x) - min_b_days) >= FAIRNESS_GAP else 0,
                                    0 if get_group(x) != a_group else 1,  
                                    norm_b(x)                 
                                ))
                            
                            row["Consultant B"] = name_map.get(best)
                            add_score(best, "B_days", curr_d)
                            recent_b_shifts[best] = to_date(curr_d)   
                        else: row["Consultant B"] = "UNCOVERED"
                    else: 
                        row["Consultant B"] = name_map.get(cur_b, "UNCOVERED")
                        add_score(cur_b, "B_days", curr_d)
                        if cur_b and cur_b in period_counts: recent_b_shifts[cur_b] = to_date(curr_d)      
                    
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        b_i = next((k for k, v in name_map.items() if v == row["Consultant B"]), None)
                        if b_i: add_score(b_i, "Nights", curr_d)

        for i in active_initials: row[i] = get_eff_marker(curr_d, i)
        if da_i: last_a_end[da_i] = to_date(curr_d)
        rota_data.append(row); curr_d += timedelta(days=1)

    ex_c = fin_d + timedelta(days=1)
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
    
    # --- SELF-HEALING SESSION STATE ---
    if 'rota_result' not in st.session_state: st.session_state.rota_result = None
    
    if 'staff' not in st.session_state: 
        st.session_state.staff = pd.DataFrame(DEFAULT_STAFF).reset_index(drop=True)
    else:
        # Actively forcefully inject missing columns into stubbornly cached memory
        if 'FTE A' not in st.session_state.staff.columns:
            st.session_state.staff['FTE A'] = 1.0
        if 'FTE B' not in st.session_state.staff.columns:
            st.session_state.staff['FTE B'] = 1.0
            
    if 'locums' not in st.session_state: 
        st.session_state.locums = pd.DataFrame({"Start Date": pd.Series(dtype='datetime64[ns]'), "End Date": pd.Series(dtype='datetime64[ns]'), "Locum Name": pd.Series(dtype='str')}).reset_index(drop=True)
    
    st.sidebar.header("1. Consultant Roster")
    
    edited_staff = st.sidebar.data_editor(
        st.session_state.staff, 
        num_rows="dynamic", 
        column_config={
            "DIL Day": st.column_config.SelectboxColumn("DIL Day", options=["", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]),
            "FTE A": st.column_config.NumberColumn("FTE A", min_value=0.1, max_value=1.0, step=0.01, format="%.2f"),
            "FTE B": st.column_config.NumberColumn("FTE B", min_value=0.1, max_value=1.0, step=0.01, format="%.2f")
        }
    )
    
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
    with c1: leave_f = st.file_uploader("Upload Separate Leave Requests - Optional", type=["xlsx"])
    with c2: prev_f = st.file_uploader("Upload Previous Rota (Memory)", type=["xlsx"])

    current_year = datetime.today().year
    def_s = datetime(current_year, 1, 1).date()
    def_e = datetime(current_year, 12, 31).date()
    df_l = pd.DataFrame()
    
    if leave_f:
        try:
            df_l = pd.read_excel(leave_f); df_l.columns = [str(c).strip().upper() for c in df_l.columns]
            d_col = next((c for c in df_l.columns if "DATE" in c), df_l.columns[1])
            df_l[d_col] = pd.to_datetime(df_l[d_col], dayfirst=True, errors='coerce').dt.normalize()
            df_l = df_l.dropna(subset=[d_col]).set_index(d_col)
        except: st.error("Error loading leave file.")

    st.markdown("---")
    st.subheader("🗓️ Calendar & Fairness Settings")
    st.markdown("<p style='color:#86868b; font-size:14px;'>Set the exact dates for your new schedule, and verify the dates you want to track for total fairness balancing.</p>", unsafe_allow_html=True)
    
    c_p1, c_p2, c_y1, c_y2 = st.columns(4)
    with c_p1: sel_s = st.date_input("Start New Rota:", value=def_s, format="DD/MM/YYYY")
    with c_p2: sel_e = st.date_input("End New Rota:", value=max(def_e, sel_s), min_value=sel_s, format="DD/MM/YYYY")
    with c_y1: ytd_s = st.date_input("YTD Start Date:", value=datetime(current_year, 1, 1).date(), format="DD/MM/YYYY")
    with c_y2: ytd_e = st.date_input("YTD End Date:", value=datetime(current_year, 12, 31).date(), format="DD/MM/YYYY")
    st.markdown("---")

    counts, m_seq, c_sch, bq, ca, cb, h_last, hist_b_dates, prev_rota_df, memory_leave_df = parse_previous_rota(prev_f, edited_staff.to_dict('records'), sel_s, ytd_s, ytd_e) if prev_f else ({}, [], {}, [], None, None, {}, {}, None, pd.DataFrame())
    
    if not memory_leave_df.empty:
        if not df_l.empty:
            df_l = df_l.combine_first(memory_leave_df)
        else:
            df_l = memory_leave_df

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
        rota, summ, state = generate_rota(df_l, edited_staff.to_dict('records'), sel_s, sel_e, counts, m_seq, c_sch, bq, ca, cb, h_last, locum_dict, hist_b_dates, ytd_s, ytd_e)
        st.session_state.rota_result = (rota, summ, state, sel_e, prev_rota_df)

    if st.session_state.rota_result:
        rota, summ, state, sel_e, prev_rota_df = st.session_state.rota_result
        def highlight(row):
            dt = pd.to_datetime(row['Date'], format="%d/%m/%Y")
            if dt.weekday() >= 5 or (uk_holidays is not None and to_date(dt) in uk_holidays): return ['background-color: #e5e5ea;'] * len(row)
            return [''] * len(row)
        
        st.subheader("Rota for the Selected Period")
        try:
            target_str = pd.to_datetime(sel_e).strftime("%d/%m/%Y")
            l_idx = rota[rota['Date'] == target_str].index[-1]
            st.dataframe(rota.iloc[:l_idx+1].style.apply(highlight, axis=1), use_container_width=True, height=450)
        except: st.dataframe(rota.style.apply(highlight, axis=1), use_container_width=True, height=450)
        
        st.subheader("Fairness Calculator")
        st.dataframe(summ, use_container_width=True)
        
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='openpyxl') as writer:
            rota.to_excel(writer, index=False, sheet_name='Rota')
            
            if prev_rota_df is not None and not prev_rota_df.empty:
                prev_rota_df.to_excel(writer, index=False, sheet_name='Previous Period Rota')
                
            summ.to_excel(writer, index=False, sheet_name='Summary')
            pd.DataFrame(list(state.items()), columns=["Key", "Value"]).to_excel(writer, index=False, sheet_name='State')
            pd.DataFrame([["AL", "Annual Leave"], ["SL", "Study Leave"], ["OFF", "Off for other reasons"], ["NO", "Please do not put me on call"], ["REQ", "Please add me on this week"]], columns=["Marker", "Description"]).to_excel(writer, index=False, sheet_name='Legend')
            
            sheets_to_format = ['Rota']
            if prev_rota_df is not None and not prev_rota_df.empty:
                sheets_to_format.append('Previous Period Rota')
                
            for s_name in sheets_to_format:
                ws = writer.sheets[s_name]
                ws.freeze_panes = "C2"
                align_left = Alignment(horizontal='left', vertical='center')
                for i, col_cells in enumerate(ws.columns, 1):
                    col_name = ws.cell(row=1, column=i).value
                    vals = [len(str(cell.value)) for cell in col_cells if cell.value]
                    max_l = max(vals + [len(str(col_name)) if col_name else 5])
                    ws.column_dimensions[get_column_letter(i)].width = max_l + 2
                    for cell in col_cells: cell.alignment = align_left
            
            fill = PatternFill(start_color="E5E5EA", end_color="E5E5EA", fill_type="solid")
            
            ws_rota = writer.sheets['Rota']
            for i, r in enumerate(rota.itertuples(), start=2):
                dt = pd.to_datetime(r.Date, format="%d/%m/%Y")
                if dt.weekday() >= 5 or (uk_holidays is not None and to_date(dt) in uk_holidays):
                    for c in range(1, len(rota.columns) + 1): ws_rota.cell(row=i, column=c).fill = fill
                    
            if prev_rota_df is not None and not prev_rota_df.empty:
                ws_prev = writer.sheets['Previous Period Rota']
                
                date_col_name = next((c for c in prev_rota_df.columns if str(c).strip().upper() == "DATE"), None)
                if not date_col_name:
                    if "Day" in prev_rota_df.columns and prev_rota_df.columns.get_loc("Day") > 0:
                        date_col_name = prev_rota_df.columns[prev_rota_df.columns.get_loc("Day") - 1]
                    elif len(prev_rota_df.columns) > 1:
                        date_col_name = prev_rota_df.columns[1]

                threshold_date = None
                if date_col_name:
                    try:
                        parsed_dates = pd.to_datetime(prev_rota_df[date_col_name], format="%d/%m/%Y", errors='coerce')
                        max_date = parsed_dates.max()
                        if pd.notna(max_date):
                            threshold_date = max_date - timedelta(days=365)
                    except: pass

                for i, r in enumerate(prev_rota_df.itertuples(), start=2):
                    try:
                        if date_col_name and hasattr(r, date_col_name):
                            d_val = getattr(r, date_col_name)
                        else:
                            d_val = r[1] if len(r) > 1 else None 

                        if isinstance(d_val, str): dt = pd.to_datetime(d_val, dayfirst=True)
                        else: dt = pd.to_datetime(d_val)
                        
                        if dt.weekday() >= 5 or (uk_holidays is not None and to_date(dt) in uk_holidays):
                            for c in range(1, len(prev_rota_df.columns) + 1): ws_prev.cell(row=i, column=c).fill = fill
                            
                        if threshold_date and dt < threshold_date:
                            ws_prev.row_dimensions[i].hidden = True
                    except Exception:
                        pass
        
        st.download_button("Download Rota", out.getvalue(), f"Rota_{datetime.now().strftime('%d_%m')}.xlsx", type="primary")

if __name__ == "__main__": main()

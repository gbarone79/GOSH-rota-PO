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
    {"Initials": "GB", "Full Name": "Giuseppe Barone", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Solid"},
    {"Initials": "SD", "Full Name": "Sarita Depani", "Active": True, "DIL Day": "Monday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Neuro"},
    {"Initials": "OS", "Full Name": "Olga Slater", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Solid"},
    {"Initials": "MJ", "Full Name": "Mette Jorgensen", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Neuro"},
    {"Initials": "CD", "Full Name": "Cat Duncan", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Solid"},
    {"Initials": "KS", "Full Name": "Kavitha Srivatsa", "Active": True, "DIL Day": "Monday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Neuro"},
    {"Initials": "TC", "Full Name": "Tanzina Chowdhury", "Active": True, "DIL Day": "Tuesday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Solid"},
    {"Initials": "ES", "Full Name": "Elwira Szychot", "Active": True, "DIL Day": "Friday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Neuro"},
    {"Initials": "TN", "Full Name": "Trung Nguyen", "Active": True, "DIL Day": "Wednesday", "FTE A": 1.0, "FTE B": 1.0, "Group": "Solid"},
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

def safe_ini(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip().upper()
    if s in ["", "NAN", "NONE", "NAT", "NULL", "<NA>"]: return ""
    return s

def match_name(val, consultants):
    val = safe_ini(val)
    if not val: return None
    for c in consultants:
        ini = safe_ini(c.get('Initials'))
        if not ini: continue
        fname = str(c.get('Full Name', '')).strip().upper()
        if val == ini or val == fname: return ini
        if val in fname.split(): return ini
    return None

def get_bh_group(d):
    if uk_holidays is None or d not in uk_holidays: return None
    name = uk_holidays.get(d).upper()
    if "NEW YEAR" in name: return "NEWYEAR"
    if "EASTER" in name or "GOOD FRIDAY" in name: return "EASTER"
    if "MAY" in name or "SPRING" in name or "JUBILEE" in name or "CORONATION" in name:
        return "EARLY_MAY" if d.day <= 15 else "LATE_MAY"
    if "SUMMER" in name or "AUGUST" in name: return "AUGUST"
    if "CHRISTMAS" in name or "BOXING" in name: return "CHRISTMAS"
    return "OTHER"

def parse_previous_rota(uploaded_file, consultants, sel_s, ytd_s, ytd_e):
    counts = defaultdict(lambda: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0, "BH_groups": set()})
    master_a_seq, cycle_a_schedule, b_queue, curr_a, curr_b, last_a_dates = [], {}, [], None, None, {}
    hist_b_dates = {}
    last_b_week_start = {}
    prev_rota_df = None
    master_a_seq_temp = []
    memory_leave_dict = {}
    
    active_initials = [safe_ini(c.get('Initials')) for c in consultants if c.get('Active', True) and safe_ini(c.get('Initials'))]
    name_map = {safe_ini(c.get('Initials')): str(c.get('Full Name', '')).strip() for c in consultants if safe_ini(c.get('Initials'))}

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
            date_col = next((c for c in df_rota.columns if str(c).strip().upper() == "DATE"), df_rota.columns[0])
            col_a = next((c for c in df_rota.columns if str(c).strip().upper() in ["CONSULTANT A", "ATTENDING"]), None)
            col_b = next((c for c in df_rota.columns if str(c).strip().upper() in ["CONSULTANT B", "B CONSULTANT"]), None)
            col_n = next((c for c in df_rota.columns if str(c).strip().upper() in ["NIGHTS", "NIGHT"]), None)
            
            for _, row in df_rota.iterrows():
                try:
                    d_val = row[date_col]
                    if pd.isna(d_val): continue
                    r_date = pd.to_datetime(d_val, dayfirst=True).date()
                    a_ini = match_name(row.get(col_a), consultants)
                    b_ini = match_name(row.get(col_b), consultants)
                    n_ini = match_name(row.get(col_n), consultants)
                    
                    is_bh = uk_holidays is not None and r_date in uk_holidays and r_date.weekday() < 5
                    bh_grp = get_bh_group(r_date) if is_bh else None
                    
                    if r_date.weekday() == 0 and b_ini:
                        last_b_week_start[b_ini] = r_date
                        
                    if ytd_s <= r_date <= ytd_e and r_date < sel_s:
                        if a_ini: counts[a_ini]["A_days"] += 1
                        if b_ini: counts[b_ini]["B_days"] += 1
                        if n_ini: counts[n_ini]["Nights"] += 1
                        if is_bh:
                            if a_ini:
                                counts[a_ini]["BH_days"] += 1
                                if bh_grp: counts[a_ini]["BH_groups"].add(bh_grp)
                            if b_ini:
                                counts[b_ini]["BH_days"] += 1
                                if bh_grp: counts[b_ini]["BH_groups"].add(bh_grp)
                                
                    if r_date < sel_s:
                        if a_ini:
                            cycle_a_schedule[r_date] = a_ini
                            last_a_dates[a_ini] = max(last_a_dates.get(a_ini, datetime.min.date()), r_date)
                            if not master_a_seq_temp or master_a_seq_temp[-1] != a_ini: master_a_seq_temp.append(a_ini)
                        if b_ini: hist_b_dates[b_ini] = max(hist_b_dates.get(b_ini, datetime.min.date()), r_date)
                        if n_ini: hist_b_dates[n_ini] = max(hist_b_dates.get(n_ini, datetime.min.date()), r_date)
                    
                    l_data = {ini: safe_raw(row.get(ini)) for ini in active_initials}
                    memory_leave_dict[pd.to_datetime(r_date)] = l_data
                    clean_row = {"Date": r_date.strftime("%d/%m/%Y"), "Day": DAY_NAMES[r_date.weekday()],
                                "Consultant A": name_map.get(a_ini, safe_raw(row.get(col_a))),
                                "Consultant B": name_map.get(b_ini, safe_raw(row.get(col_b))),
                                "Nights": name_map.get(n_ini, safe_raw(row.get(col_n))), **l_data}
                    clean_rota_data.append(clean_row)
                except: continue
            if clean_rota_data:
                df_temp = pd.DataFrame(clean_rota_data)
                mask = (df_temp['Consultant A'] != '') | (df_temp['Consultant B'] != '') | (df_temp['Nights'] != '')
                if mask.any(): prev_rota_df = df_temp.loc[:mask[::-1].idxmax()].copy()

        if "State" in xls.sheet_names:
            df_state = pd.read_excel(xls, sheet_name="State")
            state_dict = dict(zip(df_state.iloc[:,0].astype(str).str.strip(), df_state.iloc[:,1]))
            master_a_seq = [x.strip().upper() for x in str(state_dict.get("A_MASTER_SEQ","")).split(",") if x.strip() and x.strip().upper() != "NONE"]
            curr_a = state_dict.get("CURRENT_A")
            curr_b = state_dict.get("CURRENT_B")
            
        if not master_a_seq and master_a_seq_temp:
            seen = set()
            master_a_seq = [x for x in reversed(master_a_seq_temp) if not (x in seen or seen.add(x)) and x != "NONE"][::-1]
    except Exception as e:
        st.warning(f"Memory parsing warning: {e}")
        
    return counts, master_a_seq, cycle_a_schedule, [], curr_a, curr_b, last_a_dates, hist_b_dates, prev_rota_df, pd.DataFrame.from_dict(memory_leave_dict, orient='index') if memory_leave_dict else pd.DataFrame(), last_b_week_start

def get_group(doc_ini):
    if doc_ini in ["ES", "SD", "MJ", "KS"]: return "NEURO"
    if doc_ini in ["GB", "OS", "CD", "TC", "TN"]: return "SOLID"
    return "UNKNOWN"

def generate_rota(leave_df, consultants, start_date, end_date, hist_counts, hist_master_a, hist_cycle_a, hist_b_q, hist_curr_a, hist_curr_b, hist_last_a_dates, locum_dict, hist_b_dates, ytd_s, ytd_e, hist_last_b_week):
    active_initials = [safe_ini(c.get('Initials')) for c in consultants if c.get('Active', True) and safe_ini(c.get('Initials'))]
    name_map = {safe_ini(c.get('Initials')): c.get('Full Name', '') for c in consultants if safe_ini(c.get('Initials'))}
    
    def safe_float(val, default=1.0):
        try:
            f = float(val)
            return f if not pd.isna(f) else default
        except:
            return default

    # Date sanitization to strictly native python datetime.date
    curr_d_start = pd.to_datetime(start_date).date()
    fin_d = pd.to_datetime(end_date).date()
    ytd_s_d = pd.to_datetime(ytd_s).date()
    ytd_e_d = pd.to_datetime(ytd_e).date()

    d_map = {}
    fte_a_map = {}
    fte_b_map = {}
    group_map = {}
    for c in consultants:
        ini = safe_ini(c.get('Initials'))
        if c.get('Active', True) and ini:
            raw_dil = c.get('DIL Day')
            d_map[ini] = "" if pd.isna(raw_dil) or raw_dil is None or str(raw_dil).strip().upper() in ["NAN", "NONE", "NAT"] else str(raw_dil).strip().capitalize()
            fte_a_map[ini] = max(0.1, safe_float(c.get('FTE A', 1.0)))
            fte_b_map[ini] = max(0.1, safe_float(c.get('FTE B', 1.0)))
            group_map[ini] = str(c.get('Group', 'Others')).strip().capitalize()

    def get_eff_marker(d_date, ini):
        ts = pd.Timestamp(d_date).normalize()
        if not leave_df.empty and ts in leave_df.index and ini in leave_df.columns:
            val = str(leave_df.loc[ts, ini]).strip().upper()
            if val not in ["NAN", "NAT", "NONE", ""]: return val
        if d_map.get(ini) == DAY_NAMES[d_date.weekday()]: return "DIL"
        return ""

    def is_blocked_for_week(d_date, ini):
        w_start = d_date - timedelta(days=d_date.weekday())
        block_count = 0
        for i in range(5):
            if get_eff_marker(w_start + timedelta(days=i), ini) in ["AL", "SL", "NO", "OFF"]:
                block_count += 1
        return block_count >= 3

    def sync(h, a):
        nl = [i for i in h if i in a]
        for i in a:
            if i not in nl: nl.append(i)
        return nl

    if hist_master_a:
        master_a_seq = sync(hist_master_a, active_initials)
    else:
        solids = [i for i in active_initials if group_map.get(i) == "Solid"]
        neuros = [i for i in active_initials if group_map.get(i) == "Neuro"]
        others = [i for i in active_initials if group_map.get(i) not in ["Solid", "Neuro"]]
        random.seed()
        random.shuffle(solids); random.shuffle(neuros); random.shuffle(others)
        master_a_seq = []
        while solids or neuros or others:
            if solids: master_a_seq.append(solids.pop(0))
            if neuros: master_a_seq.append(neuros.pop(0))
            if others: master_a_seq.append(others.pop(0))

    start_anc = curr_d_start - timedelta(days=(curr_d_start.weekday() - 4) % 7)
    
    cycle_a_schedule = {k: v for k, v in hist_cycle_a.items() if k < start_anc}
    last_a_end = {k: v for k, v in hist_last_a_dates.items() if v < start_anc}

    c_fs = []
    chk = start_anc
    # STRICT SHIFT LIMIT: A-week calculation stops perfectly at the fin_d, no phantom assignments.
    while chk <= fin_d:
        if not any(locum_dict.get(chk + timedelta(days=o)) for o in range(7)): 
            c_fs.append(chk)
        chk += timedelta(days=7)
    
    temp = {f: {'doc': None, 'locked': False} for f in c_fs}
    for f in c_fs:
        for doc in active_initials:
            if get_eff_marker(f, doc) == "REQ":
                temp[f]['doc'], temp[f]['locked'] = doc, True
                break

    sim_a_counts = {ini: hist_counts.get(ini, {}).get("A_days", 0) for ini in active_initials}
    for ini in active_initials:
        if sim_a_counts[ini] == 0: 
            others_a = [v for v in sim_a_counts.values() if v > 0]
            sim_a_counts[ini] = min(others_a) if others_a else 0

    sim_bh_groups = {ini: set(hist_counts.get(ini, {}).get("BH_groups", set())) for ini in active_initials}
    
    def week_has_bh(slot_f):
        return any(uk_holidays is not None and (slot_f + timedelta(days=i)) in uk_holidays and (slot_f + timedelta(days=i)).weekday() < 5 for i in range(7))
        
    def get_week_bh_group(slot_f):
        for i in range(7):
            d = slot_f + timedelta(days=i)
            if uk_holidays is not None and d in uk_holidays and d.weekday() < 5:
                return get_bh_group(d)
        return None

    # --- THE UNIFIED A-WEEK SCORING ENGINE ---
    for f in c_fs:
        has_bh = week_has_bh(f)
        
        if temp[f]['locked']: 
            sim_a_counts[temp[f]['doc']] += 7
            last_a_end[temp[f]['doc']] = f
            if has_bh:
                bg = get_week_bh_group(f)
                if bg: sim_bh_groups[temp[f]['doc']].add(bg)
            continue
            
        avail = []
        for cand in active_initials:
            if not any(get_eff_marker(f + timedelta(days=o), cand) in ["AL", "SL", "NO", "OFF"] for o in range(7)):
                valid = True
                for t_f, t_info in temp.items():
                    if t_info['doc'] == cand and t_f != f and abs((f - t_f).days) < 21:
                        valid = False
                        break
                if valid and (f - last_a_end.get(cand, datetime.min.date())).days >= 21:
                    avail.append(cand)
                    
        if not avail:
            temp[f]['doc'] = "UNCOVERED"
            continue
            
        def a_score(x):
            bh_pen = 0
            if has_bh:
                bh_pen = len(sim_bh_groups[x]) * 10000
                
            rest_pen = 0
            if (f - last_a_end.get(x, datetime.min.date())).days < 28: 
                rest_pen += 1000
            for t_f, t_info in temp.items():
                if t_info['doc'] == x and t_f != f and abs((f - t_f).days) < 28:
                    rest_pen += 1000
                    
            fairness = sim_a_counts[x] / fte_a_map[x]
            return (bh_pen, rest_pen, fairness)
            
        best = min(avail, key=a_score)
        temp[f]['doc'] = best
        sim_a_counts[best] += 7
        last_a_end[best] = f
        if has_bh:
            bg = get_week_bh_group(f)
            if bg: sim_bh_groups[best].add(bg)

    for f, info in temp.items():
        if info['doc']: cycle_a_schedule[f] = info['doc']
        
    rota_data = []
    curr_d = curr_d_start
    cur_a = cycle_a_schedule.get(start_anc, "UNCOVERED") if hist_curr_a is None else hist_curr_a
    cur_b = hist_curr_b if hist_curr_b else None
    
    final_cur_a = cur_a
    final_cur_b = cur_b
    
    counts = {ini: {
        "A_days": hist_counts.get(ini, {}).get("A_days", 0), 
        "B_days": hist_counts.get(ini, {}).get("B_days", 0), 
        "Nights": hist_counts.get(ini, {}).get("Nights", 0),
        "BH_days": hist_counts.get(ini, {}).get("BH_days", 0),
        "BH_groups": set(hist_counts.get(ini, {}).get("BH_groups", set()))
    } for ini in active_initials}
    
    period_counts = {ini: {"A_days": 0, "B_days": 0, "Nights": 0, "BH_days": 0} for ini in active_initials}
    
    recent_b_shifts = hist_b_dates.copy()
    last_b_week_start = hist_last_b_week.copy()

    others_b = [v for v in counts.values() if v["B_days"] > 0]
    base_b = min([v["B_days"] for v in others_b]) if others_b else 0
    others_n = [v for v in counts.values() if v["Nights"] > 0]
    base_n = min([v["Nights"] for v in others_n]) if others_n else 0
    
    for ini in active_initials:
        if counts[ini]["B_days"] == 0: counts[ini]["B_days"] = base_b
        if counts[ini]["Nights"] == 0: counts[ini]["Nights"] = base_n

    sim_b_counts = {ini: counts[ini]["B_days"] for ini in active_initials}
    sim_n_counts = {ini: counts[ini]["Nights"] for ini in active_initials}

    def add_score(doc_ini, s_type, d_date):
        if doc_ini and doc_ini in period_counts:
            # Active Period ONLY
            if curr_d_start <= d_date <= fin_d:
                period_counts[doc_ini][s_type] += 1
            if ytd_s_d <= d_date <= ytd_e_d:
                counts[doc_ini][s_type] += 1
            if s_type == "B_days": sim_b_counts[doc_ini] += 1
            if s_type == "Nights": sim_n_counts[doc_ini] += 1

    norm_b = lambda x: sim_b_counts[x] / fte_b_map[x]
    norm_n = lambda x: sim_n_counts[x] / fte_a_map[x]

    def b_penalty(doc, d_date):
        past_a = [f for f, cand in cycle_a_schedule.items() if cand == doc and f <= d_date]
        if past_a:
            last_a = max(past_a)
            if (d_date - last_a).days <= 21: return 1
            
        for f_date, cand in cycle_a_schedule.items():
            if cand == doc and 0 <= (f_date - d_date).days <= 14: return 1
            
        return 0

    def group_penalty(a_ini, b_ini):
        if not a_ini or not b_ini: return 0
        g_a = group_map.get(a_ini, "Others")
        g_b = group_map.get(b_ini, "Others")
        if g_a == "Others" or g_b == "Others": return 0
        return 1 if g_a == g_b else 0

    # THE EXTENDED LOOP: Calculates Active period normally, builds empty Skeleton Buffer for the extra 365 days
    while curr_d <= fin_d + timedelta(days=365):            
        wd = curr_d.weekday()
        
        # Initialize Skeleton Row
        row = {"Date": curr_d.strftime("%d/%m/%Y"), "Day": DAY_NAMES[wd], "Consultant A": "", "Consultant B": "", "Nights": ""}
        
        if curr_d <= fin_d:
            man_b = is_b_manual_period(curr_d)
            is_bh_today = uk_holidays is not None and curr_d in uk_holidays and curr_d.weekday() < 5
            
            w_start = curr_d - timedelta(days=wd)
            is_bh_week = False
            if uk_holidays is not None:
                for d in range(5):
                    if (w_start + timedelta(days=d)) in uk_holidays:
                        is_bh_week = True; break

            active_a_nights = [0, 2, 4, 5, 6] if is_bh_week else [1, 3, 4, 5, 6]
            active_b_nights = [1, 3] if is_bh_week else [0, 2]
            
            if wd == 4:
                anc = curr_d
                if any(locum_dict.get(anc + timedelta(days=o)) for o in range(7)): cur_a = "LOCUM_ACADEMIC"
                else: cur_a = cycle_a_schedule.get(anc, "UNCOVERED")

            if wd == 0 or cur_b is None or cur_b == "UNCOVERED":
                if man_b: 
                    cur_b = "MANUAL ENTRY"
                else:
                    avail_week = [c for c in active_initials if c != cur_a and not is_blocked_for_week(curr_d, c)]
                    if avail_week:
                        def week_score(x):
                            p_box = b_penalty(x, curr_d)
                            g_pen = group_penalty(cur_a, x)
                            bb_pen = 1 if (curr_d - last_b_week_start.get(x, datetime.min.date())).days <= 14 else 0
                            
                            score = norm_b(x)
                            if g_pen == 1: score += 3.0 
                            if bb_pen == 1: score += 500.0 
                            if p_box == 1: score += 1000.0 
                            return score
                            
                        cur_b = min(avail_week, key=week_score)
                        if cur_b != "UNCOVERED":
                            last_b_week_start[cur_b] = curr_d
                    else:
                        cur_b = "UNCOVERED"

            rq_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQ"), None)
            rqb_d = next((i for i in active_initials if get_eff_marker(curr_d, i) == "REQB"), None)
            lc_d = locum_dict.get(curr_d)
            
            da_n = lc_d if lc_d else (name_map.get(rq_d, rq_d) if rq_d else name_map.get(cur_a, cur_a))
            da_i = (cur_a if cur_a not in ["UNCOVERED", "LOCUM_ACADEMIC"] else None) if not rq_d and not lc_d else rq_d

            # Apply Assignments to Row
            row["Consultant A"] = da_n
            if da_i: add_score(da_i, "A_days", curr_d)
            if wd in active_a_nights: 
                row["Nights"] = da_n
                if da_i: add_score(da_i, "Nights", curr_d)

            if wd in B_WORK_DAYS:
                if is_bh_today:
                    row["Consultant B"] = ""
                elif man_b:
                    row["Consultant B"] = "MANUAL ENTRY"
                elif rqb_d: 
                    row["Consultant B"] = name_map.get(rqb_d, rqb_d)
                    add_score(rqb_d, "B_days", curr_d)
                    recent_b_shifts[rqb_d] = curr_d
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        add_score(rqb_d, "Nights", curr_d)
                else:
                    if not cur_b or get_eff_marker(curr_d, cur_b) in ["AL", "SL", "NO", "OFF", "DIL", "REQ", "REQB"] or cur_b == da_i:
                        avail = [i for i in active_initials if i not in [da_i, cur_b] and get_eff_marker(curr_d, i) not in ["AL", "SL", "NO", "OFF", "DIL", "REQ", "REQB"]]
                        if avail:
                            def daily_score(x):
                                p_box = b_penalty(x, curr_d)
                                g_pen = group_penalty(da_i, x)
                                recent_b = (curr_d - recent_b_shifts.get(x, datetime.min.date())).days
                                bb_pen = 1 if recent_b < 7 else 0
                                
                                if wd in active_b_nights:
                                    score = norm_b(x) + norm_n(x)
                                else:
                                    score = norm_b(x)
                                    
                                if g_pen == 1: score += 1.5
                                if bb_pen == 1: score += 500.0 
                                
                                if wd in active_b_nights:
                                    other_b_nights = [n for n in active_b_nights if n != wd]
                                    x_dil = d_map.get(x)
                                    if x_dil and x_dil in DAY_NAMES and DAY_NAMES.index(x_dil) in other_b_nights:
                                        score -= 2.0 
                                        
                                if p_box == 1: score += 1000.0
                                return score

                            best = min(avail, key=daily_score)
                            row["Consultant B"] = name_map.get(best)
                            add_score(best, "B_days", curr_d)
                            recent_b_shifts[best] = curr_d
                        else: row["Consultant B"] = "UNCOVERED"
                    else: 
                        row["Consultant B"] = name_map.get(cur_b, "UNCOVERED")
                        add_score(cur_b, "B_days", curr_d)
                        recent_b_shifts[cur_b] = curr_d
                    
                    if wd in active_b_nights:
                        row["Nights"] = row["Consultant B"]
                        b_ini = next((k for k, v in name_map.items() if v == row["Consultant B"]), None)
                        if b_ini: add_score(b_ini, "Nights", curr_d)

            if is_bh_today:
                b_ini = next((k for k, v in name_map.items() if v == row["Consultant B"]), None)
                bh_grp = get_bh_group(curr_d)
                for ini in [da_i, b_ini]:
                    if ini and ini in period_counts:
                        if curr_d_start <= curr_d <= fin_d:
                            period_counts[ini]["BH_days"] += 1
                        if ytd_s_d <= curr_d <= ytd_e_d:
                            counts[ini]["BH_days"] += 1
                            if bh_grp: counts[ini]["BH_groups"].add(bh_grp)

            if da_i: last_a_end[da_i] = curr_d
            
            # Snapshots Memory exactly on End Date
            if curr_d == fin_d:
                final_cur_a = cur_a
                final_cur_b = cur_b

        # Apply specific DIL and Leave markers for everyone ALWAYS (including Buffer)
        for i in active_initials: row[i] = get_eff_marker(curr_d, i)
        rota_data.append(row)
        curr_d += timedelta(days=1)

    state = {"A_MASTER_SEQ": ",".join(master_a_seq), "B_QUEUE": ",".join(active_initials), "CURRENT_A": final_cur_a or "", "CURRENT_B": final_cur_b or ""}
    
    summ = [{"Initials": k, "Period A": period_counts[k]["A_days"], "Period B": period_counts[k]["B_days"], "Period Nights": period_counts[k]["Nights"], "Period Total (A+N)": period_counts[k]["A_days"] + period_counts[k]["Nights"], "YTD A": counts[k]["A_days"], "YTD B": counts[k]["B_days"], "YTD Nights": counts[k]["Nights"], "YTD Total (A+N)": counts[k]["A_days"] + counts[k]["Nights"], "YTD BH": counts[k]["BH_days"]} for k in active_initials]
    
    for row in summ:
        ini = row["Initials"]
        if hist_counts.get(ini, {}).get("B_days", 0) == 0 and row["YTD B"] >= base_b:
            row["YTD B"] -= base_b
        if hist_counts.get(ini, {}).get("Nights", 0) == 0 and row["YTD Nights"] >= base_n:
            row["YTD Nights"] -= base_n

    return pd.DataFrame(rota_data), pd.DataFrame(summ), state

def main():
    st.set_page_config(layout="wide", page_title="Consultant Rota Tool")
    st.markdown("<style>.stApp {background-color: #fbfbfd !important;} p, h1, h2, h3, span, label, div {color: #1d1d1f; font-family: -apple-system, sans-serif;} div.stButton > button {background-color: #0071e3 !important; color: #fff !important; border-radius: 20px !important;}</style>", unsafe_allow_html=True)
    st.title("Consultant Rota Dashboard")
    
    if 'rota_result' not in st.session_state: st.session_state.rota_result = None
    
    if 'staff' not in st.session_state: 
        st.session_state.staff = pd.DataFrame(DEFAULT_STAFF).reset_index(drop=True)
    else:
        if 'Group' not in st.session_state.staff.columns:
            def assign_group(ini):
                ini_str = str(ini).strip().upper()
                if ini_str in ["GB", "OS", "CD", "TC", "TN"]: return "Solid"
                if ini_str in ["ES", "SD", "MJ", "KS"]: return "Neuro"
                return "Others"
            st.session_state.staff['Group'] = st.session_state.staff['Initials'].apply(assign_group)
            
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
            "FTE B": st.column_config.NumberColumn("FTE B", min_value=0.1, max_value=1.0, step=0.01, format="%.2f"),
            "Group": st.column_config.SelectboxColumn("Group", options=["Solid", "Neuro", "Others"])
        }
    )
    
    st.sidebar.header("2. Locum / Academic Overrides")
    st.sidebar.markdown("<p style='color:#86868b; font-size:14px;'>Leave 'End Date' blank for a 1-day assignment.</p>", unsafe_allow_html=True)
    edited_locums = st.sidebar.data_editor(
        st.session_state.locums, 
        num_rows="dynamic", 
        column_config={
            "Start Date": st.column_config.DateColumn("Start Date", format="DD/MM/YYYY"), 
            "End Date": st.column_config.DateColumn("End Date", format="DD/MM/YYYY")
        }
    )

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

    clean_staff = edited_staff.copy()
    if not clean_staff.empty:
        clean_staff['Initials'] = clean_staff['Initials'].astype(str).str.strip().str.upper()
        clean_staff = clean_staff[~clean_staff['Initials'].isin(['', 'NAN', 'NONE', 'NAT', 'NULL', '<NA>'])]
        clean_staff['Active'] = clean_staff['Active'].fillna(True).astype(bool)
        clean_staff['FTE A'] = pd.to_numeric(clean_staff['FTE A'], errors='coerce').fillna(1.0)
        clean_staff['FTE B'] = pd.to_numeric(clean_staff['FTE B'], errors='coerce').fillna(1.0)
        clean_staff['Group'] = clean_staff['Group'].fillna("Others")

    counts, m_seq, c_sch, bq, ca, cb, h_last, hist_b_dates, prev_rota_df, memory_leave_df, hist_last_b_week = parse_previous_rota(prev_f, clean_staff.to_dict('records'), sel_s, ytd_s, ytd_e) if prev_f else ({}, [], {}, [], None, None, {}, {}, None, pd.DataFrame(), {})
    
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
        rota, summ, state = generate_rota(df_l, clean_staff.to_dict('records'), sel_s, sel_e, counts, m_seq, c_sch, bq, ca, cb, h_last, locum_dict, hist_b_dates, ytd_s, ytd_e, hist_last_b_week)
        st.session_state.rota_result = (rota, summ, state, sel_e, prev_rota_df)

    if st.session_state.rota_result:
        rota, summ, state, sel_e, prev_rota_df = st.session_state.rota_result
        def highlight(row):
            try:
                dt = pd.to_datetime(row['Date'], format="%d/%m/%Y").date()
                if dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays): return ['background-color: #e5e5ea;'] * len(row)
            except: pass
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
            pd.DataFrame([["AL", "Annual Leave"], ["SL", "Study Leave"], ["OFF", "Off for other reasons"], ["NO", "Please do not put me on call"], ["REQ", "Please add me on this week"], ["REQB", "Please add me as B Consultant on this day"]], columns=["Marker", "Description"]).to_excel(writer, index=False, sheet_name='Legend')
            
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
                try:
                    dt = pd.to_datetime(r.Date, format="%d/%m/%Y").date()
                    if dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays):
                        for c in range(1, len(rota.columns) + 1): ws_rota.cell(row=i, column=c).fill = fill
                except: pass
                    
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
                            threshold_date = max_date.date() - timedelta(days=365)
                    except: pass

                for i, r in enumerate(prev_rota_df.itertuples(), start=2):
                    try:
                        if date_col_name and hasattr(r, date_col_name):
                            d_val = getattr(r, date_col_name)
                        else:
                            d_val = r[1] if len(r) > 1 else None 

                        if isinstance(d_val, str): dt = pd.to_datetime(d_val, dayfirst=True).date()
                        else: dt = pd.to_datetime(d_val).date()
                        
                        if dt.weekday() >= 5 or (uk_holidays is not None and dt in uk_holidays):
                            for c in range(1, len(prev_rota_df.columns) + 1): ws_prev.cell(row=i, column=c).fill = fill
                            
                        if threshold_date and dt < threshold_date:
                            ws_prev.row_dimensions[i].hidden = True
                    except Exception:
                        pass
        
        st.download_button("Download Rota", out.getvalue(), f"Rota_{datetime.now().strftime('%d_%m')}.xlsx", type="primary")

if __name__ == "__main__": main()

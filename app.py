import streamlit as st
import json
import os
import io
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Agrynox Multi-Enterprise Portal",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DB_FILE = "enterprise_travel_db.json"
USERS_FILE = "users_db.json"

# ----------------------------------------------------------------
# DATABASE INITIALIZATION
# ----------------------------------------------------------------
DEFAULT_USERS = {
    "sunildeshmukh@gmail.com": {
        "name": "SUNIL DESHMUKH",
        "company": "Ajeet Seeds Pvt Ltd",
        "role": "employee",
        "designation": "EXECUTIVE(MKG.VEG)",
        "hq": "BADNOOR",
        "grade": "S1",
        "vehicle_no": "MP28ZC9464"
    },
    "admin@agrynox.com": {
        "name": "Agrynox Master Admin",
        "company": "Agrynox",
        "role": "admin",
        "designation": "SUPER ADMINISTRATOR",
        "hq": "HEAD OFFICE",
        "grade": "M1",
        "vehicle_no": "-"
    },
    "staff@agrynox.com": {
        "name": "AGRYNOX FIELD OFFICER",
        "company": "Agrynox",
        "role": "employee",
        "designation": "AGRI OFFICER",
        "hq": "CHHINDWARA",
        "grade": "A1",
        "vehicle_no": "MP28AB1234"
    }
}

def load_json(filepath, default):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
                if data:
                    return data
        except Exception:
            pass
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(default, f, indent=2, ensure_ascii=False)
    return default

def save_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

users_db = load_json(USERS_FILE, DEFAULT_USERS)
travel_db = load_json(DB_FILE, [])

# ----------------------------------------------------------------
# 3D WARM-CREAM UI & FORCED WHITE TEXT ON ALL BUTTONS
# ----------------------------------------------------------------
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@600;700;800;900&display=swap');
    * { font-family: 'Plus Jakarta Sans', sans-serif !important; }
    
    .stApp { 
        background-color: #ede6db !important; 
        color: #1c1917 !important;
    }
    
    .block-container { 
        max-width: 920px !important; 
        padding-top: 1.2rem !important; 
        padding-bottom: 4rem !important;
    }
    
    .agrynox-header {
        background: #1c1917;
        color: #f5f5f4;
        padding: 16px 22px;
        border-radius: 12px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
        border: 2px solid #44403c;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    
    .card-box {
        background: #faf7f2;
        border: 2px solid #b8ab99;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 10px rgba(87, 83, 78, 0.1);
    }
    
    .box-title {
        font-size: 13px;
        font-weight: 900;
        color: #9a3412;
        text-transform: uppercase;
        margin-bottom: 12px;
        border-bottom: 2px solid #d6cebf;
        padding-bottom: 6px;
    }
    
    .stTextInput input, .stNumberInput input, .stDateInput input, .stSelectbox select {
        background-color: #ffffff !important;
        border: 2px solid #57534e !important;
        color: #0c0a09 !important;
        font-weight: 800 !important;
        border-radius: 8px !important;
        height: 42px !important;
    }
    
    label {
        font-weight: 800 !important;
        color: #292524 !important;
        font-size: 13px !important;
    }

    div[data-testid="stButton"] button,
    div[data-testid="stButton"] button *,
    div[data-testid="stButton"] button p,
    div[data-testid="stButton"] button span,
    div[data-testid="stButton"] button div,
    div.stDownloadButton button,
    div.stDownloadButton button * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-weight: 900 !important;
        letter-spacing: 0.3px !important;
    }

    div[data-testid="stButton"] button[kind="secondary"] {
        background: #292524 !important;
        border: 2px solid #1c1917 !important;
        box-shadow: 0 4px 0 #000000 !important;
    }
    div[data-testid="stButton"] button[kind="secondary"]:hover {
        background: #44403c !important;
        transform: translateY(-2px);
    }

    div[data-testid="stButton"] button[kind="primary"] {
        background: #0284c7 !important;
        border: 2px solid #0369a1 !important;
        box-shadow: 0 4px 0 #075985 !important;
    }
    div[data-testid="stButton"] button[kind="primary"]:hover {
        background: #0369a1 !important;
        transform: translateY(-2px);
    }

    .btn-submit button, .btn-submit button * {
        background: #15803d !important;
        border: 2px solid #14532d !important;
        box-shadow: 0 4px 0 #052e16 !important;
        width: 100% !important;
    }
    
    .btn-next button, .btn-next button * {
        background: #0284c7 !important;
        border: 2px solid #0369a1 !important;
        box-shadow: 0 4px 0 #075985 !important;
        width: 100% !important;
    }

    .btn-add button, .btn-add button * {
        background: #c2410c !important;
        border: 2px solid #9a3412 !important;
        box-shadow: 0 4px 0 #7c2d12 !important;
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------
# LOGIN & AUTHENTICATION
# ----------------------------------------------------------------
if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

if not st.session_state.auth_user:
    st.markdown("""
    <div style="text-align:center; margin-top: 40px; margin-bottom: 24px;">
        <h2 style="color:#1c1917; font-weight:900; margin:0; font-size: 30px;">🌱 AGRYNOX MULTI-ENTERPRISE</h2>
        <p style="color:#57534e; font-size:14px; font-weight:700; margin-top:4px;">Field Tour & TA/DA Automation System</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown('<div class="card-box" style="max-width: 480px; margin: auto;">', unsafe_allow_html=True)
    st.markdown('<div class="box-title">🔐 Select User Profile / लॉगिन करें</div>', unsafe_allow_html=True)
    
    emp_list = list(users_db.keys())
    selected_account = st.selectbox(
        "Select User Profile:",
        options=emp_list,
        format_func=lambda x: f"[{users_db[x].get('company', 'Ajeet Seeds')}] {users_db[x]['name']} ({x})"
    )
    
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    if st.button("🚀 Log In To System", use_container_width=True, type="primary"):
        st.session_state.auth_user = users_db[selected_account]
        st.session_state.auth_user["email"] = selected_account
        st.session_state.active_menu = "📝 Daily Tour Entry"
        st.rerun()
        
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ----------------------------------------------------------------
# LOGGED IN DASHBOARD
# ----------------------------------------------------------------
current_user = st.session_state.auth_user
is_admin = (current_user.get("role") == "admin")
user_company = current_user.get("company", "Ajeet Seeds Pvt Ltd")

if "active_menu" not in st.session_state:
    st.session_state.active_menu = "📝 Daily Tour Entry"

st.markdown(f"""
<div class="agrynox-header">
    <div>
        <h3 style="margin:0; font-size:18px; color:#ffffff;">AGRYNOX ENTERPRISE PORTAL</h3>
        <p style="margin:3px 0 0 0; color:#d6d3d1; font-size:13px; font-weight:600;">
            User: <b style="color:#fde047;">{current_user['name']}</b> | Company: <b style="color:#38bdf8;">{user_company}</b>
        </p>
    </div>
    <span style="background:{'#dc2626' if is_admin else '#15803d'}; color:white; padding:4px 12px; border-radius:6px; font-size:12px; font-weight:900;">
        {'👑 MASTER ADMIN' if is_admin else '👤 FIELD OFFICER'}
    </span>
</div>
""", unsafe_allow_html=True)

top_c1, top_c2 = st.columns([8, 2])
with top_c2:
    if st.button("🚪 Log Out", use_container_width=True, type="secondary"):
        st.session_state.auth_user = None
        st.rerun()

menu_items = ["📝 Daily Tour Entry", "📋 Trip Logbook", "📥 Export Monthly Bill"]
if is_admin:
    menu_items.append("👑 Admin Master Portal")

btn_cols = st.columns(len(menu_items))
for i, item in enumerate(menu_items):
    with btn_cols[i]:
        is_active = (st.session_state.active_menu == item)
        if st.button(
            item, 
            key=f"nav_btn_{i}", 
            use_container_width=True, 
            type="primary" if is_active else "secondary"
        ):
            st.session_state.active_menu = item
            st.rerun()

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

def generate_schedule(start_str, end_str, count):
    if count == 0:
        return []
    fmt = "%I:%M %p"
    try:
        t1 = datetime.strptime(start_str.strip().upper(), fmt)
        t2 = datetime.strptime(end_str.strip().upper(), fmt)
    except Exception:
        t1 = datetime.strptime("08:00 AM", fmt)
        t2 = datetime.strptime("08:30 PM", fmt)
    if t2 <= t1:
        t2 += timedelta(days=1)
    
    total_mins = (t2 - t1).total_seconds() / 60.0
    step = total_mins / count
    cur_t = t1
    schedule = []
    for i in range(count):
        dep = cur_t.strftime("%I:%M %p")
        arr = (t2 if i == count - 1 else cur_t + timedelta(minutes=max(round(step * 0.4), 20))).strftime("%I:%M %p")
        schedule.append({"dep": dep, "arr": arr})
        cur_t = t1 + timedelta(minutes=round((i + 1) * step))
    return schedule

# ----------------------------------------------------------------
# PURE PYTHON GENERATOR: EXACT TEMPLATE ARCHITECTURE & A4 LANDSCAPE C/F
# ----------------------------------------------------------------
def build_pure_company_workbook(target_user, month_key, user_records):
    wb = openpyxl.Workbook()
    
    # Styles
    font_company = Font(name="Arial", size=11, bold=True)
    font_sub = Font(name="Arial", size=9, bold=True)
    font_header = Font(name="Arial", size=8, bold=True)
    font_data = Font(name="Arial", size=8, bold=False)
    font_cf = Font(name="Arial", size=8, bold=True, color="9A3412")
    font_total = Font(name="Arial", size=9, bold=True)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    fill_header = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_cf = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    fill_total = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    
    thin_side = Side(style="thin", color="CBD5E1")
    grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    month_obj = datetime.strptime(month_key, "%Y-%m")
    month_str = month_obj.strftime("%B %Y").upper()

    # Flatten all trip legs
    all_legs = []
    sorted_records = sorted(user_records, key=lambda x: x.get("date", ""))
    for r in sorted_records:
        r_legs = r.get("legs", [])
        for idx, l in enumerate(r_legs):
            all_legs.append({
                "record": r,
                "leg": l,
                "is_first_leg": (idx == 0)
            })

    # ============================================================
    # SHEET 1: TE Statement (Page 1)
    # ============================================================
    ws_te = wb.active
    ws_te.title = "TE Stetment-Page 1"
    
    # Setup A4 Landscape Print Config
    ws_te.page_setup.orientation = ws_te.ORIENTATION_LANDSCAPE
    ws_te.page_setup.paperSize = ws_te.PAPERSIZE_A4
    ws_te.page_setup.fitToPage = True
    ws_te.page_setup.fitToWidth = 1
    ws_te.page_setup.fitToHeight = 0

    te_col_widths = {
        "A": 12, "B": 18, "C": 18, "D": 11, "E": 11, "F": 12,
        "G": 14, "H": 10, "I": 12, "J": 10, "K": 12, "L": 14,
        "M": 12, "N": 12, "O": 13, "P": 22
    }
    for col, width in te_col_widths.items():
        ws_te.column_dimensions[col].width = width

    def write_te_page_header(start_row, page_num):
        # Company Info Block
        ws_te.cell(row=start_row, column=1, value=f"     {target_user.get('company', 'AJEET SEEDS PVT LTD').upper()}").font = font_company
        ws_te.cell(row=start_row, column=5, value="From :").font = font_sub
        ws_te.cell(row=start_row, column=8, value="To :").font = font_sub
        ws_te.cell(row=start_row, column=11, value="Name :").font = font_sub
        ws_te.cell(row=start_row, column=12, value=target_user['name']).font = font_sub
        ws_te.cell(row=start_row, column=14, value="Head Quarter :").font = font_sub
        ws_te.cell(row=start_row, column=16, value=target_user['hq']).font = font_sub
        
        ws_te.cell(row=start_row+1, column=1, value="      Travelling Expenses Statement").font = font_company
        ws_te.cell(row=start_row+1, column=5, value="Designation :").font = font_sub
        ws_te.cell(row=start_row+1, column=7, value=target_user['designation']).font = font_sub
        ws_te.cell(row=start_row+1, column=8, value="Grade :").font = font_sub
        ws_te.cell(row=start_row+1, column=9, value=target_user['grade']).font = font_sub
        ws_te.cell(row=start_row+1, column=11, value="Vehicle No. :").font = font_sub
        ws_te.cell(row=start_row+1, column=13, value=target_user['vehicle_no']).font = font_sub
        ws_te.cell(row=start_row+1, column=15, value=f"Month : {month_str} (Page {page_num})").font = font_sub

        # Grid Table Headers
        headers_row1 = [
            (1, "Date"), (2, "From"), (3, "To"), (4, "Time"), (6, "Mode of Travel"),
            (7, "Distance (Kms.)"), (8, "Fare"), (9, "Daily Expenses"), (10, "Lodging If any"),
            (11, "Local Conveyance"), (12, "Vehicle Fuel Exp."), (13, "Vehicle repairing"),
            (14, "Other Expanses"), (15, "Total"), (16, "Remarks")
        ]
        h_r = start_row + 3
        for col_idx, text in headers_row1:
            c = ws_te.cell(row=h_r, column=col_idx, value=text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = grid_border

        # Sub-header row for Place/Arrival/Departure
        h_sub = start_row + 4
        ws_te.cell(row=h_sub, column=2, value="Place").font = font_header
        ws_te.cell(row=h_sub, column=3, value="Place").font = font_header
        ws_te.cell(row=h_sub, column=4, value="Departure").font = font_header
        ws_te.cell(row=h_sub, column=5, value="Arrival").font = font_header
        for col in range(1, 17):
            c = ws_te.cell(row=h_sub, column=col)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = grid_border
        return h_sub + 1

    LINES_PER_PAGE = 35
    total_legs = len(all_legs)
    leg_pointer = 0
    page = 1
    current_te_row = 1
    prev_cf_row = None

    while leg_pointer < total_legs or total_legs == 0:
        data_start_row = write_te_page_header(current_te_row, page)
        cf_row = data_start_row
        
        # Write C/F or B/F Row on Page 2+
        ws_te.cell(row=cf_row, column=2, value="B/F" if page > 1 else "C/F").font = font_cf
        if page > 1 and prev_cf_row:
            for col_l in ["G", "H", "I", "J", "K", "L", "M", "N", "O"]:
                c_idx = openpyxl.utils.column_index_from_string(col_l)
                ws_te.cell(row=cf_row, column=c_idx, value=f'={col_l}{prev_cf_row}').font = font_cf
        for c in range(1, 17):
            ws_te.cell(row=cf_row, column=c).border = grid_border
            ws_te.cell(row=cf_row, column=c).fill = fill_cf
        
        row_idx = cf_row + 1
        page_leg_count = 0

        while leg_pointer < total_legs and page_leg_count < LINES_PER_PAGE:
            item = all_legs[leg_pointer]
            r = item["record"]
            l = item["leg"]
            first = item["is_first_leg"]

            mode_val = l.get("mode", "Bike")
            km_val = l.get("km", 0) if mode_val == "Bike" else 0

            ws_te.cell(row=row_idx, column=1, value=r.get("date", "") if first else "").alignment = align_center
            ws_te.cell(row=row_idx, column=2, value=l.get("from", "")).alignment = align_left
            ws_te.cell(row=row_idx, column=3, value=l.get("to", "")).alignment = align_left
            ws_te.cell(row=row_idx, column=4, value=l.get("dep", "")).alignment = align_center
            ws_te.cell(row=row_idx, column=5, value=l.get("arr", "")).alignment = align_center
            ws_te.cell(row=row_idx, column=6, value=mode_val).alignment = align_center
            ws_te.cell(row=row_idx, column=7, value=km_val).alignment = align_right

            if first:
                if r.get("fare", 0) > 0: ws_te.cell(row=row_idx, column=8, value=r["fare"]).alignment = align_right
                if r.get("da", 0) > 0: ws_te.cell(row=row_idx, column=9, value=r["da"]).alignment = align_right
                if r.get("lodging", 0) > 0: ws_te.cell(row=row_idx, column=10, value=r["lodging"]).alignment = align_right
                if r.get("conv", 0) > 0 or r.get("local_conv", 0) > 0:
                    ws_te.cell(row=row_idx, column=11, value=r.get("conv", r.get("local_conv", 0))).alignment = align_right
                if r.get("other", 0) > 0 or r.get("other_exp", 0) > 0:
                    ws_te.cell(row=row_idx, column=14, value=r.get("other", r.get("other_exp", 0))).alignment = align_right

            # Live Formulas
            ws_te.cell(row=row_idx, column=12, value=f'=IF(F{row_idx}="Bike",G{row_idx}*3.5,0)').alignment = align_right
            ws_te.cell(row=row_idx, column=15, value=f'=H{row_idx}+I{row_idx}+J{row_idx}+K{row_idx}+L{row_idx}+N{row_idx}').alignment = align_right
            
            for c in range(1, 17):
                ws_te.cell(row=row_idx, column=c).border = grid_border
                ws_te.cell(row=row_idx, column=c).font = font_data

            row_idx += 1
            leg_pointer += 1
            page_leg_count += 1

        # Pad remaining lines if final page
        while page_leg_count < LINES_PER_PAGE and leg_pointer >= total_legs:
            ws_te.cell(row=row_idx, column=15, value=f'=H{row_idx}+I{row_idx}+J{row_idx}+K{row_idx}+L{row_idx}+N{row_idx}')
            for c in range(1, 17):
                ws_te.cell(row=row_idx, column=c).border = grid_border
                ws_te.cell(row=row_idx, column=c).font = font_data
            row_idx += 1
            page_leg_count += 1

        # C/F Subtotal Line
        prev_cf_row = row_idx
        ws_te.cell(row=prev_cf_row, column=2, value="C/F TOTAL").font = font_cf
        for col_l in ["G", "H", "I", "J", "K", "L", "M", "N", "O"]:
            c_idx = openpyxl.utils.column_index_from_string(col_l)
            ws_te.cell(row=prev_cf_row, column=c_idx, value=f'=SUM({col_l}{cf_row}:{col_l}{prev_cf_row-1})').font = font_cf
            ws_te.cell(row=prev_cf_row, column=c_idx).alignment = align_right
        for c in range(1, 17):
            ws_te.cell(row=prev_cf_row, column=c).border = grid_border
            ws_te.cell(row=prev_cf_row, column=c).fill = fill_cf

        # Page Footer (Signatures)
        f_row = prev_cf_row + 2
        ws_te.cell(row=f_row, column=1, value="Signature Of Employee                                Checked By                                Sanction By").font = font_sub
        ws_te.cell(row=f_row+1, column=1, value=f"Name :- {target_user['name']}         Head Quarter:- {target_user['hq']}         For Month: {month_str}").font = font_data

        current_te_row = f_row + 6
        page += 1
        if total_legs == 0:
            break

    # ============================================================
    # SHEET 2: Bike & Car Logbook-2 (2)
    # ============================================================
    ws_log = wb.create_sheet(title="Bike & Car Logbook-2 (2)")
    ws_log.page_setup.orientation = ws_log.ORIENTATION_LANDSCAPE
    ws_log.page_setup.paperSize = ws_log.PAPERSIZE_A4
    ws_log.page_setup.fitToPage = True
    ws_log.page_setup.fitToWidth = 1
    ws_log.page_setup.fitToHeight = 0

    log_col_widths = {
        "A": 13, "B": 12, "C": 20, "D": 20, "E": 13, "F": 13,
        "G": 14, "H": 12, "I": 38, "J": 20
    }
    for col, width in log_col_widths.items():
        ws_log.column_dimensions[col].width = width

    def write_log_page_header(start_row, page_num):
        ws_log.cell(row=start_row, column=1, value=f"{target_user.get('company', 'AJEET SEEDS PVT LTD').upper()} - DAILY LOG BOOK").font = font_company
        ws_log.cell(row=start_row, column=8, value=f"Vehicle No: {target_user['vehicle_no']}").font = font_sub
        ws_log.cell(row=start_row+1, column=1, value=f"Name: {target_user['name']}   |   H.Q: {target_user['hq']}   |   Month: {month_str} (Page {page_num})").font = font_sub
        ws_log.cell(row=start_row+1, column=8, value="Type: 2WHEELER").font = font_sub

        headers_log1 = [
            (1, "Date"), (2, "Petrol"), (3, "TRAVEL"), (5, "KM READING"),
            (7, "Distance covered"), (8, "Personal use"), (9, "Purpose"), (10, "Signature")
        ]
        lh_r = start_row + 3
        for col_idx, text in headers_log1:
            c = ws_log.cell(row=lh_r, column=col_idx, value=text)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = grid_border

        lh_sub = start_row + 4
        ws_log.cell(row=lh_sub, column=2, value="in Ltrs").font = font_header
        ws_log.cell(row=lh_sub, column=3, value="From").font = font_header
        ws_log.cell(row=lh_sub, column=4, value="To").font = font_header
        ws_log.cell(row=lh_sub, column=5, value="Opening").font = font_header
        ws_log.cell(row=lh_sub, column=6, value="Closing").font = font_header
        for col in range(1, 11):
            c = ws_log.cell(row=lh_sub, column=col)
            c.font = font_header
            c.fill = fill_header
            c.alignment = align_center
            c.border = grid_border
        return lh_sub + 1

    leg_pointer = 0
    page = 1
    current_log_row = 1
    prev_log_cf_row = None

    while leg_pointer < total_legs or total_legs == 0:
        data_start_row = write_log_page_header(current_log_row, page)
        cf_row = data_start_row
        
        ws_log.cell(row=cf_row, column=3, value="B/F" if page > 1 else "C/F").font = font_cf
        if page > 1 and prev_log_cf_row:
            ws_log.cell(row=cf_row, column=7, value=f'=G{prev_log_cf_row}').font = font_cf
            ws_log.cell(row=cf_row, column=2, value=f'=B{prev_log_cf_row}').font = font_cf
        for c in range(1, 11):
            ws_log.cell(row=cf_row, column=c).border = grid_border
            ws_log.cell(row=cf_row, column=c).fill = fill_cf

        row_idx = cf_row + 1
        page_leg_count = 0

        while leg_pointer < total_legs and page_leg_count < LINES_PER_PAGE:
            item = all_legs[leg_pointer]
            r = item["record"]
            l = item["leg"]
            first = item["is_first_leg"]

            mode_val = l.get("mode", "Bike")
            km_val = l.get("km", 0) if mode_val == "Bike" else 0
            petrol_val = r.get("petrol", r.get("petrol_ltrs", 0))

            ws_log.cell(row=row_idx, column=1, value=r.get("date", "") if first else "").alignment = align_center
            if first and petrol_val > 0:
                ws_log.cell(row=row_idx, column=2, value=petrol_val).alignment = align_right
            ws_log.cell(row=row_idx, column=3, value=l.get("from", "")).alignment = align_left
            ws_log.cell(row=row_idx, column=4, value=l.get("to", "")).alignment = align_left

            # Bike rules: meter advances. Jeep/Bus: blank readings and 0 distance
            if mode_val == "Bike" and km_val > 0:
                open_km = r.get("opening_km", 0)
                ws_log.cell(row=row_idx, column=5, value=open_km).alignment = align_right
                ws_log.cell(row=row_idx, column=6, value=open_km + km_val).alignment = align_right
                ws_log.cell(row=row_idx, column=7, value=f'=F{row_idx}-E{row_idx}').alignment = align_right
            else:
                ws_log.cell(row=row_idx, column=5, value="").alignment = align_right
                ws_log.cell(row=row_idx, column=6, value="").alignment = align_right
                ws_log.cell(row=row_idx, column=7, value=0).alignment = align_right

            if first:
                ws_log.cell(row=row_idx, column=9, value=r.get("purpose", "")).alignment = align_left

            for c in range(1, 11):
                ws_log.cell(row=row_idx, column=c).border = grid_border
                ws_log.cell(row=row_idx, column=c).font = font_data

            row_idx += 1
            leg_pointer += 1
            page_leg_count += 1

        while page_leg_count < LINES_PER_PAGE and leg_pointer >= total_legs:
            ws_log.cell(row=row_idx, column=7, value=f'=F{row_idx}-E{row_idx}')
            for c in range(1, 11):
                ws_log.cell(row=row_idx, column=c).border = grid_border
                ws_log.cell(row=row_idx, column=c).font = font_data
            row_idx += 1
            page_leg_count += 1

        prev_log_cf_row = row_idx
        ws_log.cell(row=prev_log_cf_row, column=3, value="C/F TOTAL").font = font_cf
        ws_log.cell(row=prev_log_cf_row, column=2, value=f'=SUM(B{cf_row}:B{prev_log_cf_row-1})').font = font_cf
        ws_log.cell(row=prev_log_cf_row, column=7, value=f'=SUM(G{cf_row}:G{prev_log_cf_row-1})').font = font_cf
        for c in range(1, 11):
            ws_log.cell(row=prev_log_cf_row, column=c).border = grid_border
            ws_log.cell(row=prev_log_cf_row, column=c).fill = fill_cf

        f_row = prev_log_cf_row + 2
        ws_log.cell(row=f_row, column=1, value="Prepared by                                 Checked by                                 Sanctioned by                                 Account officer").font = font_sub

        current_log_row = f_row + 6
        page += 1
        if total_legs == 0:
            break

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================================================
# 1. SCREEN: DAILY TOUR ENTRY
# ============================================================
if st.session_state.active_menu == "📝 Daily Tour Entry":
    if "legs" not in st.session_state:
        st.session_state.legs = [{"from": "", "to": "", "mode": "Bike", "km": ""}]
    if "form_date" not in st.session_state:
        st.session_state.form_date = datetime.today().date()
    if "opening_km" not in st.session_state:
        st.session_state.opening_km = ""
    if "closing_km" not in st.session_state:
        st.session_state.closing_km = ""

    st.markdown('<div class="card-box"><div class="box-title">📅 1. Tour Date & Timings</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.form_date = st.date_input("Tour Date", st.session_state.form_date)
    with c2:
        start_t = st.text_input("Start Time", placeholder="08:00 AM")
    with c3:
        end_t = st.text_input("End Time", placeholder="08:30 PM")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><div class="box-title">🔢 2. Odometer Readings & Fuel</div>', unsafe_allow_html=True)
    m1, m2, m3 = st.columns(3)
    with m1:
        st.session_state.opening_km = st.text_input("Opening Meter", value=st.session_state.opening_km, placeholder="e.g. 55035")
    with m2:
        st.session_state.closing_km = st.text_input("Closing Meter", value=st.session_state.closing_km, placeholder="e.g. 55089")
    with m3:
        f_petrol = st.text_input("Petrol (Litres)", placeholder="0")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><div class="box-title">📍 3. Route Details (Village to Village)</div>', unsafe_allow_html=True)
    del_idx = None
    for idx, leg in enumerate(st.session_state.legs):
        r1, r2, r3, r4, r5 = st.columns([2.8, 2.8, 1.8, 1.4, 0.7])
        with r1:
            st.session_state.legs[idx]["from"] = st.text_input(f"From #{idx+1}", value=leg["from"], placeholder="Start Place", key=f"f_{idx}", label_visibility="collapsed" if idx>0 else "visible")
        with r2:
            st.session_state.legs[idx]["to"] = st.text_input(f"To #{idx+1}", value=leg["to"], placeholder="Destination", key=f"t_{idx}", label_visibility="collapsed" if idx>0 else "visible")
        with r3:
            modes = ["Bike", "Jeep", "Bus"]
            m_idx = modes.index(leg.get("mode", "Bike")) if leg.get("mode") in modes else 0
            st.session_state.legs[idx]["mode"] = st.selectbox(f"Mode #{idx+1}", modes, index=m_idx, key=f"m_{idx}", label_visibility="collapsed" if idx>0 else "visible")
        with r4:
            st.session_state.legs[idx]["km"] = st.text_input(f"KM #{idx+1}", value=str(leg.get("km", "")), placeholder="KM", key=f"k_{idx}", label_visibility="collapsed" if idx>0 else "visible")
        with r5:
            if idx == 0: st.write("")
            if st.button("✕", key=f"del_{idx}", type="secondary"):
                del_idx = idx

    if del_idx is not None and len(st.session_state.legs) > 1:
        st.session_state.legs.pop(del_idx)
        st.rerun()

    st.markdown("<div class='btn-add'>", unsafe_allow_html=True)
    if st.button("+ Add Next Village", use_container_width=True, type="secondary"):
        last_to = st.session_state.legs[-1]["to"] if st.session_state.legs else ""
        st.session_state.legs.append({"from": last_to, "to": "", "mode": "Bike", "km": ""})
        st.rerun()
    st.markdown("</div></div>", unsafe_allow_html=True)

    st.markdown('<div class="card-box"><div class="box-title">💰 4. Expenses & Remarks</div>', unsafe_allow_html=True)
    purpose_val = st.text_input("Purpose of Visit", placeholder="e.g. Retailer and Farmer visit")
    e1, e2, e3, e4, e5 = st.columns(5)
    with e1: f_da = st.text_input("DA (Daily)", placeholder="350")
    with e2: f_fare = st.text_input("Fare", placeholder="0")
    with e3: f_lodging = st.text_input("Lodging", placeholder="0")
    with e4: f_conv = st.text_input("Local Conv.", placeholder="0")
    with e5: f_other = st.text_input("Other Exp.", placeholder="0")
    st.markdown('</div>', unsafe_allow_html=True)

    def save_tour_record():
        valid_legs = []
        for l in st.session_state.legs:
            if l["from"].strip() and l["to"].strip():
                mode_chosen = l["mode"]
                km_val = 0.0 if mode_chosen in ["Jeep", "Bus"] else (float(l["km"]) if l["km"].strip() else 0.0)
                valid_legs.append({"from": l["from"].strip(), "to": l["to"].strip(), "mode": mode_chosen, "km": km_val})

        if not valid_legs:
            st.error("Please fill at least one route leg!")
            return False

        schedule = generate_schedule(start_t or "08:00 AM", end_t or "08:30 PM", len(valid_legs))
        for i, sch in enumerate(schedule):
            valid_legs[i]["dep"] = sch["dep"]
            valid_legs[i]["arr"] = sch["arr"]

        def p_float(v):
            try: return float(str(v).strip()) if str(v).strip() else 0.0
            except: return 0.0

        record = {
            "id": datetime.now().strftime("%Y%m%d%H%M%S"),
            "user_email": current_user["email"],
            "company": user_company,
            "date": st.session_state.form_date.strftime("%Y-%m-%d"),
            "opening_km": p_float(st.session_state.opening_km),
            "closing_km": p_float(st.session_state.closing_km),
            "petrol": p_float(f_petrol),
            "purpose": purpose_val.strip() or "Farmer and retailer visit",
            "da": p_float(f_da),
            "fare": p_float(f_fare),
            "lodging": p_float(f_lodging),
            "conv": p_float(f_conv),
            "other": p_float(f_other),
            "legs": valid_legs
        }
        travel_db.append(record)
        save_json(DB_FILE, travel_db)
        return True

    btn_c1, btn_c2 = st.columns(2)
    with btn_c1:
        st.markdown('<div class="btn-submit">', unsafe_allow_html=True)
        if st.button("💾 Save Day Tour", use_container_width=True, type="secondary"):
            if save_tour_record():
                st.success(f"Tour saved successfully for {st.session_state.form_date}!")
        st.markdown('</div>', unsafe_allow_html=True)
    with btn_c2:
        st.markdown('<div class="btn-next">', unsafe_allow_html=True)
        if st.button("➕ Save & Next Day", use_container_width=True, type="secondary"):
            prev_close = st.session_state.closing_km
            if save_tour_record():
                st.session_state.form_date += timedelta(days=1)
                st.session_state.opening_km = str(prev_close) if str(prev_close).strip() else ""
                st.session_state.closing_km = ""
                st.session_state.legs = [{"from": "", "to": "", "mode": "Bike", "km": ""}]
                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# ============================================================
# 2. SCREEN: TRIP LOGBOOK
# ============================================================
elif st.session_state.active_menu == "📋 Trip Logbook":
    st.subheader(f"Trip Logbook: {current_user['name']} ({user_company})")
    user_records = [r for r in travel_db if r.get("user_email") == current_user["email"]]
    if not user_records:
        st.info("No tour records logged yet.")
    else:
        for r in reversed(user_records):
            routes_txt = " ➔ ".join([l.get("from","") for l in r.get("legs",[])] + [r.get("legs",[])[-1].get("to","")] if r.get("legs") else [])
            tot_km = sum(l.get("km",0) for l in r.get("legs",[]))
            st.markdown(f"""
            <div class="card-box" style="padding:14px; margin-bottom:10px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <b style="font-size:15px; color:#1c1917;">📅 {r['date']}</b>
                    <span style="background:#15803d; color:#ffffff; padding:3px 10px; border-radius:6px; font-weight:800; font-size:12px;">{r['opening_km']} ➔ {r['closing_km']} ({tot_km} KM)</span>
                </div>
                <div style="margin-top:6px; font-size:14px; color:#292524; font-weight:600;"><b>Route:</b> {routes_txt}</div>
                <div style="color:#57534e; font-size:12px; font-weight:600; margin-top:2px;"><b>Purpose:</b> {r['purpose']}</div>
            </div>
            """, unsafe_allow_html=True)

# ============================================================
# 3. SCREEN: EXPORT MONTHLY BILL (PURE GENERATION - NO TEMPLATE)
# ============================================================
elif st.session_state.active_menu == "📥 Export Monthly Bill":
    st.subheader(f"Generate Official Bill - {user_company}")
    user_records = [r for r in travel_db if r.get("user_email") == current_user["email"]]
    if not user_records:
        st.warning("No tour records available to export.")
    else:
        months_avail = sorted(list(set(r.get("date","")[:7] for r in user_records if r.get("date"))))
        sel_m = st.selectbox("Select Month", months_avail, format_func=lambda x: datetime.strptime(x, "%Y-%m").strftime("%B %Y"))
        
        m_recs = [r for r in user_records if r.get("date","").startswith(sel_m)]
        
        if st.button("📄 Generate Official Spreadsheet (.xlsx)", use_container_width=True, type="primary"):
            excel_bytes = build_pure_company_workbook(current_user, sel_m, m_recs)
            if excel_bytes:
                st.download_button(
                    label="📥 Click Here to Download Official Bill",
                    data=excel_bytes,
                    file_name=f"TA_BILL_{sel_m}_{current_user['name'].replace(' ','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ============================================================
# 4. SCREEN: ADMIN MASTER PORTAL
# ============================================================
elif is_admin and st.session_state.active_menu == "👑 Admin Master Portal":
    st.markdown('<div class="box-title">👑 Agrynox Master Admin - Multi Company Control</div>', unsafe_allow_html=True)
    
    adm_tab_choice = st.radio("Select View:", ["📊 Audit & Download Employee Bills", "👥 Register New Employees"], horizontal=True)
    
    if adm_tab_choice == "📊 Audit & Download Employee Bills":
        filter_company = st.radio("🏢 Filter By Company:", ["All Companies", "Ajeet Seeds Pvt Ltd", "Agrynox"], horizontal=True)
        
        if filter_company == "All Companies":
            emp_emails = [e for e, u in users_db.items() if u.get("role") != "admin"]
        else:
            emp_emails = [e for e, u in users_db.items() if u.get("role") != "admin" and u.get("company") == filter_company]
        
        if not emp_emails:
            st.warning(f"No employees registered under {filter_company}.")
        else:
            c_a1, c_a2 = st.columns(2)
            with c_a1:
                sel_emp_email = st.selectbox("Select Employee:", emp_emails, format_func=lambda x: f"[{users_db[x].get('company','Ajeet Seeds')}] {users_db[x]['name']} ({users_db[x]['hq']})")
            
            target_emp = users_db[sel_emp_email]
            emp_records = [r for r in travel_db if r.get("user_email") == sel_emp_email]
            
            with c_a2:
                emp_months = sorted(list(set(r.get("date","")[:7] for r in emp_records if r.get("date"))))
                if not emp_months:
                    st.warning("No records logged for this employee.")
                    sel_emp_m = None
                else:
                    sel_emp_m = st.selectbox("Select Billing Month:", emp_months, format_func=lambda x: datetime.strptime(x, "%Y-%m").strftime("%B %Y"))
            
            if sel_emp_m:
                recs_for_m = [r for r in emp_records if r.get("date","").startswith(sel_emp_m)]
                st.write(f"**Total Days Logged:** {len(recs_for_m)} days | **Company:** {target_emp.get('company', 'Ajeet Seeds Pvt Ltd')}")
                if st.button("📥 Download Official Bill for this Employee", use_container_width=True, type="primary"):
                    admin_excel = build_pure_company_workbook(target_emp, sel_emp_m, recs_for_m)
                    if admin_excel:
                        st.download_button(
                            label=f"⬇️ Download {target_emp['name']} - {sel_emp_m} Bill (.xlsx)",
                            data=admin_excel,
                            file_name=f"TA_BILL_{sel_m}_{target_emp['name'].replace(' ','_')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )

    elif adm_tab_choice == "👥 Register New Employees":
        st.markdown("#### Register New Field Employee")
        with st.form("new_emp_form"):
            n_comp = st.selectbox("Select Company:", ["Ajeet Seeds Pvt Ltd", "Agrynox"])
            ne1, ne2 = st.columns(2)
            with ne1:
                n_email = st.text_input("Employee Email / Gmail", placeholder="e.g. rahul@gmail.com")
                n_name = st.text_input("Full Name", placeholder="RAHUL SHARMA")
                n_desig = st.text_input("Designation", placeholder="SALES OFFICER")
            with ne2:
                n_hq = st.text_input("HQ (Headquarter)", placeholder="CHHINDWARA")
                n_grade = st.text_input("Grade", placeholder="S2")
                n_veh = st.text_input("Vehicle No", placeholder="MP28AB1234")
            
            if st.form_submit_button("Register Employee"):
                clean_em = n_email.strip().lower()
                if clean_em and n_name:
                    users_db[clean_em] = {
                        "name": n_name.strip().upper(),
                        "company": n_comp,
                        "role": "employee",
                        "designation": n_desig.strip().upper(),
                        "hq": n_hq.strip().upper(),
                        "grade": n_grade.strip().upper(),
                        "vehicle_no": n_veh.strip().upper()
                    }
                    save_json(USERS_FILE, users_db)
                    st.success(f"Registered {n_name} under {n_comp} successfully!")
                    st.rerun()
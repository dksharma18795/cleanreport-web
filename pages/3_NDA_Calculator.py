import streamlit as st
import pandas as pd
import calendar
from datetime import datetime, date
import io
import base64
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 7TH CPC PAY MATRIX (LEVEL 1 TO 8)
# ==========================================
PAY_MATRIX = {
    "Level 1 (1800 GP)": [18000, 18500, 19100, 19700, 20300, 20900, 21500, 22100, 22800, 23500, 24200, 24900, 25600, 26400, 27200, 28000, 28800, 29700, 30600, 31500, 32400, 33400, 34400, 35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900],
    "Level 2 (1900 GP)": [19900, 20500, 21100, 21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200, 31100, 32000, 33000, 34000, 35000, 36100, 37200, 38300, 39400, 40600, 41800, 43100, 44400, 45700, 47100, 48500, 50000, 51500, 53000, 54600, 56200, 57900, 59600, 61400, 63200],
    "Level 3 (2000 GP)": [21700, 22400, 23100, 23800, 24500, 25200, 26000, 26800, 27600, 28400, 29300, 30200, 31100, 32000, 33000, 34000, 35000, 36100, 37200, 38300, 39400, 40600, 41800, 43100, 44400, 45700, 47100, 48500, 50000, 51500, 53000, 54600, 56200, 57900, 59600, 61400, 63200, 65100, 67100, 69100],
    "Level 4 (2400 GP)": [25500, 26300, 27100, 27900, 28700, 29600, 30500, 31400, 32300, 33300, 34300, 35300, 36400, 37500, 38600, 39800, 41000, 42200, 43500, 44800, 46100, 47500, 48900, 50400, 51900, 53500, 55100, 56800, 58500, 60300, 62100, 64000, 65900, 67900, 69900, 72000, 74200, 76400, 78700, 81100],
    "Level 5 (2800 GP)": [29200, 30100, 31000, 31900, 32900, 33900, 34900, 35900, 37000, 38100, 39200, 40400, 41600, 42800, 44100, 45400, 46800, 48200, 49600, 51100, 52600, 54200, 55800, 57500, 59200, 61000, 62800, 64700, 66600, 68600, 70700, 72800, 75000, 77300, 79600, 82000, 84500, 87000, 89600, 92300],
    "Level 6 (4200 GP)": [35400, 36500, 37600, 38700, 39900, 41100, 42300, 43600, 44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000, 70000, 72100, 74300, 76500, 78800, 81200, 83600, 86100, 88700, 91400, 94100, 96900, 99800, 102800, 105900, 109100, 112400],
    "Level 7 (4600 GP)": [44900, 46200, 47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000, 70000, 72100, 74300, 76500, 78800, 81200, 83600, 86100, 88700, 91400, 94100, 96900, 99800, 102800, 105900, 109100, 112400, 115800, 119300, 122900, 126600, 130400, 134300, 138300, 142400],
    "Level 8 (4800 GP)": [47600, 49000, 50500, 52000, 53600, 55200, 56900, 58600, 60400, 62200, 64100, 66000, 68000, 70000, 72100, 74300, 76500, 78800, 81200, 83600, 86100, 88700, 91400, 94100, 96900, 99800, 102800, 105900, 109100, 112400, 115800, 119300, 122900, 126600, 130400, 134300, 138300, 142400, 146700, 151100],
}

ALL_BASIC_PAYS = sorted(list(set([pay for level in PAY_MATRIX.values() for pay in level])))

# ==========================================
# 2. PAGE SETUP & CSS
# ==========================================
st.set_page_config(page_title="CleanReport NDA Register", page_icon="📝", layout="wide")

hide_st_style = """
    <style>
    [data-testid="collapsedControl"] {display: none;}
    [data-testid="stSidebar"] {display: none;}
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;}
    div[data-testid="stDataEditor"] th { font-size: 12px !important; text-align: center !important; }
    div[data-testid="stDataEditor"] td { cursor: pointer; }
    
    /* ENTERPRISE GRADE BIG CLICKABLE CARDS */
    div[data-testid="stCheckbox"] {
        background-color: #f8fafc;
        border: 1.5px solid #cbd5e1;
        border-radius: 6px;
        padding: 6px 12px;
        margin-bottom: 2px;
        transition: all 0.2s ease-in-out;
    }
    div[data-testid="stCheckbox"]:hover {
        background-color: #e2e8f0;
        border-color: #94a3b8;
    }
    div[data-testid="stCheckbox"] label {
        cursor: pointer !important;
        width: 100% !important;
        height: 100% !important;
        display: flex !important;
        align-items: center !important;
        margin: 0 !important;
    }
    div[data-testid="stCheckbox"] label span {
        font-weight: 600;
        font-size: 13px;
    }
    </style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)
if st.button("🏠 Back to Dashboard"): st.switch_page("app.py")

# ==========================================
# 3. SESSION STATE INITIALIZATION
# ==========================================
for key in ['ministry', 'department', 'office']: 
    if key not in st.session_state: st.session_state[key] = ""
if 'emp_db' not in st.session_state: st.session_state.emp_db = pd.DataFrame(columns=["Emp No", "Name", "Post", "Level", "Basic Pay (₹)"])
if 'active_months' not in st.session_state: st.session_state.active_months = [] 
if 'att_dict' not in st.session_state: st.session_state.att_dict = {} 
for key in ['in_empno', 'in_name', 'in_post']:
    if key not in st.session_state: st.session_state[key] = ""
if 'report_generated' not in st.session_state: st.session_state.report_generated = False

# ==========================================
# 4. HELPER FUNCTIONS
# ==========================================
def get_formatted_dates(year, month):
    num_days = calendar.monthrange(year, month)[1]
    return [date(year, month, d).strftime("%d %b\n(%a)") for d in range(1, num_days + 1)]
def get_month_name(m, y): return f"{calendar.month_name[m]} {y}"
def get_auto_da(m, y):
    if y <= 2020: return 17.0
    elif y == 2021: return 17.0 if m < 7 else 28.0
    elif y == 2022: return 34.0 if m < 7 else 38.0
    elif y == 2023: return 42.0 if m < 7 else 46.0
    elif y == 2024: return 50.0 if m < 7 else 53.0
    elif y == 2025: return 56.0 if m < 7 else 58.0
    elif y == 2026: return 60.0 if m < 7 else 63.0
    return 0.0

def number_to_words(n):
    if n == 0: return "Zero"
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
    def get_words(num):
        if num == 0: return ""
        elif num < 20: return ones[num] + " "
        elif num < 100: return tens[num // 10] + " " + get_words(num % 10)
        else: return ones[num // 100] + " Hundred " + get_words(num % 100)
    words = []
    crore = int(n // 10000000)
    n = n % 10000000
    lakh = int(n // 100000)
    n = n % 100000
    thous = int(n // 1000)
    n = int(n % 1000)
    if crore > 0: words.append(get_words(crore).strip() + " Crore ")
    if lakh > 0: words.append(get_words(lakh).strip() + " Lakh ")
    if thous > 0: words.append(get_words(thous).strip() + " Thousand ")
    if n > 0: words.append(get_words(n).strip())
    return "Rupees " + "".join(words).strip() + " Only"

def add_employee_callback(sel_level, sel_bp):
    emp_no_input = st.session_state.in_empno.strip()
    if emp_no_input and st.session_state.in_name:
        if emp_no_input in st.session_state.emp_db["Emp No"].values:
            st.toast("⚠️ Employee ID already exists!")
            return
        new_row = {"Emp No": emp_no_input, "Name": st.session_state.in_name.upper(), "Post": st.session_state.in_post.upper(), "Level": sel_level, "Basic Pay (₹)": sel_bp}
        st.session_state.emp_db = pd.concat([st.session_state.emp_db, pd.DataFrame([new_row])], ignore_index=True)
        st.session_state.in_empno, st.session_state.in_name, st.session_state.in_post = "", "", ""

def go_to_tab(tab_name): st.session_state.current_nav = tab_name

# --- BULLETPROOF INSTANT SYNC CALLBACK ---
def sync_cb(m_, y_, eid_, dt_, k_):
    val = st.session_state[k_]
    # Update underlying dataframe immediately
    emp_idx = st.session_state.att_dict[(m_, y_)]['Emp No'].astype(str) == str(eid_)
    st.session_state.att_dict[(m_, y_)].loc[emp_idx, dt_] = val
    # Force Grid to re-render to reflect the change
    st.session_state[f"ed_key_{m_}_{y_}"] = st.session_state.get(f"ed_key_{m_}_{y_}", 0) + 1

# ==========================================
# 5. DYNAMIC EXCEL ENGINE
# ==========================================
def generate_dynamic_excel(ministry, dept, office, period_str, active_months, att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols):
    wb = Workbook()
    wb.remove(wb.active)
    font_bold = Font(bold=True)
    font_title = Font(bold=True, size=14)
    align_c = Alignment(horizontal='center', vertical='center')
    align_l = Alignment(horizontal='left', vertical='center')
    fill_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    ws_admin = wb.create_sheet("Part I - Administration")
    ws_admin.page_setup.orientation = ws_admin.ORIENTATION_LANDSCAPE
    ws_admin.page_setup.paperSize = ws_admin.PAPERSIZE_A4
    ws_admin.sheet_properties.pageSetUpPr.fitToPage = True
    ws_admin.page_setup.fitToWidth = 1
    ws_admin.page_setup.fitToHeight = False
    ws_admin.print_options.horizontalCentered = True
    ws_admin.page_margins.left = 0.2
    ws_admin.page_margins.right = 0.2
    
    row_idx = 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = ministry.upper(); ws_admin[f'A{row_idx}'].font = font_title; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = f"{dept.upper()}, {office.upper()}"; ws_admin[f'A{row_idx}'].font = font_bold; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 1
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = f"Night Duty Allowance Calculation from {period_str}"; ws_admin[f'A{row_idx}'].font = font_bold; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 2
    ws_admin.merge_cells(f'A{row_idx}:AH{row_idx}'); ws_admin[f'A{row_idx}'] = "PART I: MONTH-WISE ATTENDANCE REGISTER"; ws_admin[f'A{row_idx}'].font = font_title; ws_admin[f'A{row_idx}'].alignment = align_c
    row_idx += 2
    
    att_cell_refs = {}
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        date_cols = get_formatted_dates(y, m)
        num_days = len(date_cols)
        att_df = att_dict[(m, y)]
        if sum(sum(1 for d in date_cols if row[d] == True) for _, row in att_df.iterrows()) == 0: continue
            
        ws_admin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=2+num_days+1)
        ws_admin.cell(row=row_idx, column=1, value=f"ATTENDANCE REGISTER: {m_name.upper()} {y}").font = font_bold
        ws_admin.cell(row=row_idx, column=1).fill = fill_hdr
        row_idx += 1
        headers = ["Emp No", "Name"] + [f"{d:02d}" for d in range(1, num_days+1)] + ["Total"]
        for c_i, h in enumerate(headers, 1):
            c = ws_admin.cell(row=row_idx, column=c_i, value=h)
            c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
        row_idx += 1
        
        att_cell_refs[(m, y)] = {}
        for _, emp_row in att_df.iterrows():
            ws_admin.cell(row=row_idx, column=1, value=emp_row['Emp No']).border = border_all
            ws_admin.cell(row=row_idx, column=1).alignment = align_c
            ws_admin.cell(row=row_idx, column=2, value=emp_row['Name']).border = border_all
            ws_admin.cell(row=row_idx, column=2).alignment = align_l
            
            for d_i, d_col in enumerate(date_cols, 3):
                val = "●" if emp_row[d_col] else ""
                c = ws_admin.cell(row=row_idx, column=d_i, value=val)
                c.border = border_all; c.alignment = align_c; c.font = font_bold
            tot_col_idx = 2 + num_days + 1
            tot_col_letter = get_column_letter(tot_col_idx)
            c_tot = ws_admin.cell(row=row_idx, column=tot_col_idx, value=f'=COUNTIF(C{row_idx}:{get_column_letter(2+num_days)}{row_idx}, "●")')
            c_tot.border = border_all; c_tot.font = font_bold; c_tot.alignment = align_c
            att_cell_refs[(m, y)][emp_row['Emp No']] = f"'Part I - Administration'!{tot_col_letter}{row_idx}"
            row_idx += 1
        row_idx += 2
        
    row_idx += 2
    ws_admin.cell(row=row_idx, column=2, value="_______________________").font = font_bold
    ws_admin.cell(row=row_idx+1, column=2, value="Section In-Charge").font = font_bold
    ws_admin.cell(row=row_idx, column=24, value="_______________________").font = font_bold
    ws_admin.cell(row=row_idx+1, column=24, value="Office Spdt/Estb In-Charge").font = font_bold
    ws_admin.cell(row=row_idx+4, column=13, value="_______________________").font = font_bold; ws_admin.cell(row=row_idx+4, column=13).alignment = align_c
    ws_admin.cell(row=row_idx+5, column=13, value="Head of Office").font = font_bold; ws_admin.cell(row=row_idx+5, column=13).alignment = align_c
    ws_admin.cell(row=row_idx+6, column=13, value="(Counter Signed)").font = font_bold; ws_admin.cell(row=row_idx+6, column=13).alignment = align_c
        
    ws_admin.column_dimensions['A'].width = 10
    ws_admin.column_dimensions['B'].width = 22
    for i in range(3, 3+31): ws_admin.column_dimensions[get_column_letter(i)].width = 3.2
    ws_admin.column_dimensions[get_column_letter(3+31)].width = 6

    ws_fin = wb.create_sheet("Part II - Finance")
    ws_fin.page_setup.orientation = ws_fin.ORIENTATION_LANDSCAPE
    ws_fin.page_setup.paperSize = ws_fin.PAPERSIZE_A4
    ws_fin.sheet_properties.pageSetUpPr.fitToPage = True
    ws_fin.page_setup.fitToWidth = 1
    ws_fin.page_setup.fitToHeight = False
    ws_fin.print_options.horizontalCentered = True
    ws_fin.page_margins.left = 0.2
    ws_fin.page_margins.right = 0.2
    
    row_idx = 1
    ws_fin.merge_cells('A1:I1'); ws_fin['A1'] = ministry.upper(); ws_fin['A1'].font = font_title; ws_fin['A1'].alignment = align_c
    row_idx += 1
    ws_fin.merge_cells('A2:I2'); ws_fin['A2'] = f"{dept.upper()}, {office.upper()}"; ws_fin['A2'].font = font_bold; ws_fin['A2'].alignment = align_c
    row_idx += 2
    ws_fin.merge_cells('A3:I3'); ws_fin['A3'] = "PART II: MONTH-WISE FINANCIAL BILL & SUMMARY"; ws_fin['A3'].font = font_title; ws_fin['A3'].alignment = align_c
    row_idx += 2
    
    fin_refs = {}
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty or month_logs["Duties"].sum() == 0: continue
        da_val = edited_da_df[(edited_da_df['Month'] == m_name) & (edited_da_df['Year'] == y)]['DA (%)'].values[0]
        
        ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
        ws_fin.cell(row=row_idx, column=1, value=f"BILL FOR {m_name.upper()} {y} (DA: {da_val}%)").font = font_bold
        ws_fin.cell(row=row_idx, column=1).fill = fill_hdr
        row_idx += 1
        headers = ["Emp No", "Name", "Post", "Basic Pay", "DA %", "Duties", "1 Hr Rate", f"{int(night_hours_per_duty*10)} min Rate", "Net NDA"]
        for c_i, h in enumerate(headers, 1):
            c = ws_fin.cell(row=row_idx, column=c_i, value=h)
            c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
        row_idx += 1
        
        fin_refs[(m, y)] = {}
        start_sum_row = row_idx
        for _, row in month_logs.iterrows():
            emp_no = row['Emp No']
            ws_fin.cell(row=row_idx, column=1, value=emp_no).border = border_all; ws_fin.cell(row=row_idx, column=1).alignment = align_c
            ws_fin.cell(row=row_idx, column=2, value=row['Name']).border = border_all; ws_fin.cell(row=row_idx, column=2).alignment = align_l
            ws_fin.cell(row=row_idx, column=3, value=row['Post']).border = border_all; ws_fin.cell(row=row_idx, column=3).alignment = align_l
            
            c_bp = ws_fin.cell(row=row_idx, column=4, value=row['Basic Pay'])
            c_bp.border = border_all; c_bp.alignment = align_c; c_bp.number_format = '#,##0'
            
            c_da = ws_fin.cell(row=row_idx, column=5, value=da_val)
            c_da.border = border_all; c_da.alignment = align_c
            
            duty_ref = att_cell_refs.get((m, y), {}).get(emp_no, "0")
            c_duty = ws_fin.cell(row=row_idx, column=6, value=f"={duty_ref}")
            c_duty.border = border_all; c_duty.font = font_bold; c_duty.alignment = align_c
            
            c_1hr = ws_fin.cell(row=row_idx, column=7, value=f"=ROUND((MIN(D{row_idx}, 43600)*(1+E{row_idx}/100))/200, 2)")
            c_1hr.border = border_all; c_1hr.alignment = align_c; c_1hr.number_format = '0.00'
            
            c_shift = ws_fin.cell(row=row_idx, column=8, value=f"=ROUND(G{row_idx}*({night_hours_per_duty}/6), 2)")
            c_shift.border = border_all; c_shift.alignment = align_c; c_shift.number_format = '0.00'
            
            c_net = ws_fin.cell(row=row_idx, column=9, value=f"=ROUND(H{row_idx}*F{row_idx}, 2)")
            c_net.border = border_all; c_net.font = font_bold; c_net.alignment = align_c; c_net.number_format = '#,##0.00'
            
            fin_refs[(m, y)][emp_no] = {"duty": f"F{row_idx}", "net": f"I{row_idx}"}
            row_idx += 1
            
        ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=8)
        ws_fin.cell(row=row_idx, column=1, value=f"TOTAL NDA FOR {m_name.upper()} {y}:").font = font_bold
        ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
        c_month_sum = ws_fin.cell(row=row_idx, column=9, value=f"=SUM(I{start_sum_row}:I{row_idx-1})")
        c_month_sum.font = font_bold; c_month_sum.alignment = align_c; c_month_sum.number_format = '#,##0.00'
        row_idx += 3

    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=9)
    ws_fin.cell(row=row_idx, column=1, value="GRAND FINAL SUMMARY").font = font_bold
    ws_fin.cell(row=row_idx, column=1).fill = fill_hdr
    row_idx += 1
    sum_headers = ["Emp No", "Name", "Post"] + month_cols + ["Total Duties", "Total Hrs", "Wt Hrs", "Total Net NDA"]
    for c_i, h in enumerate(sum_headers, 1):
        c = ws_fin.cell(row=row_idx, column=c_i, value=h)
        c.font = font_bold; c.border = border_all; c.alignment = align_c; c.fill = fill_hdr
    row_idx += 1
    
    start_sum_row = row_idx
    grand_total_all_actual = 0.0 
    for _, row in final_bill.iterrows():
        emp_no = row['Emp No']
        ws_fin.cell(row=row_idx, column=1, value=emp_no).border = border_all; ws_fin.cell(row=row_idx, column=1).alignment = align_c
        ws_fin.cell(row=row_idx, column=2, value=row['Name']).border = border_all; ws_fin.cell(row=row_idx, column=2).alignment = align_l
        ws_fin.cell(row=row_idx, column=3, value=row['Post']).border = border_all; ws_fin.cell(row=row_idx, column=3).alignment = align_l
        
        c_idx = 4
        duty_cells, net_cells = [], []
        for period in active_months:
            m, y = period['m'], period['y']
            ref = fin_refs.get((m, y), {}).get(emp_no)
            c = ws_fin.cell(row=row_idx, column=c_idx)
            c.border = border_all; c.alignment = align_c
            if ref:
                c.value = f"={ref['duty']}"
                duty_cells.append(get_column_letter(c_idx) + str(row_idx))
                net_cells.append(ref['net'])
            else: c.value = 0
            c_idx += 1
            
        c_tot_d = ws_fin.cell(row=row_idx, column=c_idx, value=f"=SUM({','.join(duty_cells)})" if duty_cells else 0)
        c_tot_d.border = border_all; c_tot_d.font = font_bold; c_tot_d.alignment = align_c
        c_tot_h = ws_fin.cell(row=row_idx, column=c_idx+1, value=f"={get_column_letter(c_idx)}{row_idx}*{night_hours_per_duty}")
        c_tot_h.border = border_all; c_tot_h.alignment = align_c
        c_wt_h = ws_fin.cell(row=row_idx, column=c_idx+2, value=f"={get_column_letter(c_idx+1)}{row_idx}/6")
        c_wt_h.border = border_all; c_wt_h.alignment = align_c; c_wt_h.number_format = '0.00'
        c_net = ws_fin.cell(row=row_idx, column=c_idx+3, value=f"=SUM({','.join(net_cells)})" if net_cells else 0)
        c_net.border = border_all; c_net.font = font_bold; c_net.alignment = align_c; c_net.number_format = '#,##0.00'
        
        grand_total_all_actual += row['Net Payable NDA (₹)']
        row_idx += 1
        
    last_col_l = get_column_letter(c_idx+3)
    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+2)
    ws_fin.cell(row=row_idx, column=1, value="GRAND TOTAL PAYOUT:").font = font_bold
    ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
    c_exact_sum = ws_fin.cell(row=row_idx, column=c_idx+3, value=f'=SUM({last_col_l}{start_sum_row}:{last_col_l}{row_idx-1})')
    c_exact_sum.font = font_bold; c_exact_sum.alignment = align_c; c_exact_sum.number_format = '#,##0.00'
    row_idx += 1

    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+2)
    ws_fin.cell(row=row_idx, column=1, value="ROUNDED GRAND TOTAL:").font = font_bold
    ws_fin.cell(row=row_idx, column=1).alignment = Alignment(horizontal='right')
    c_rnd_sum = ws_fin.cell(row=row_idx, column=c_idx+3, value=f'=ROUND({last_col_l}{row_idx-1}, 0)')
    c_rnd_sum.font = font_bold; c_rnd_sum.alignment = align_c; c_rnd_sum.number_format = '#,##0'
    row_idx += 1
    
    ws_fin.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=c_idx+3)
    ws_fin.cell(row=row_idx, column=1, value=f"({number_to_words(round(grand_total_all_actual))})").font = Font(bold=True, italic=True)
    ws_fin.cell(row=row_idx, column=1).alignment = align_c
    row_idx += 4
    
    ws_fin.cell(row=row_idx, column=2, value="_______________________").font = font_bold
    ws_fin.cell(row=row_idx+1, column=2, value="Office Spdt/Estb In-Charge").font = font_bold
    ws_fin.cell(row=row_idx, column=c_idx+2, value="_______________________").font = font_bold
    ws_fin.cell(row=row_idx+1, column=c_idx+2, value="Accountant/Accounts Officer").font = font_bold
    ws_fin.cell(row=row_idx+4, column=(c_idx+4)//2, value="_______________________").font = font_bold; ws_fin.cell(row=row_idx+4, column=(c_idx+4)//2).alignment = align_c
    ws_fin.cell(row=row_idx+5, column=(c_idx+4)//2, value="Head of Office").font = font_bold; ws_fin.cell(row=row_idx+5, column=(c_idx+4)//2).alignment = align_c
    ws_fin.cell(row=row_idx+6, column=(c_idx+4)//2, value="(Counter Signed)").font = font_bold; ws_fin.cell(row=row_idx+6, column=(c_idx+4)//2).alignment = align_c

    ws_fin.column_dimensions['A'].width = 10
    ws_fin.column_dimensions['B'].width = 22
    ws_fin.column_dimensions['C'].width = 10
    ws_fin.column_dimensions['D'].width = 10
    ws_fin.column_dimensions['G'].width = 10
    ws_fin.column_dimensions['H'].width = 12
    ws_fin.column_dimensions['I'].width = 12
    ws_fin.column_dimensions[last_col_l].width = 12

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ==========================================
# 6. WORD (.DOC) GENERATION FUNCTION
# ==========================================
def generate_word_html(ministry, dept, office, period_str, active_months, att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols):
    rate_header = f"{int(night_hours_per_duty * 10)} min Rate"
    html = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head><style>
        body {{ font-family: Arial, sans-serif; font-size: 10pt; }}
        h3, h4, h5 {{ text-align: center; margin: 5px 0; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 10px; margin-bottom: 25px; border: 1px solid black; }}
        th, td {{ border: 1px solid black; padding: 4px; text-align: center; font-size: 9pt; vertical-align: middle; }}
        th {{ background-color: #d9e1f2; font-weight: bold; }}
        .header-row {{ background-color: #2f5597; color: white; font-weight: bold; padding: 5px; text-align: left; }}
    </style></head><body>
    <h3>{ministry.upper()}</h3><h4>{dept.upper()}, {office.upper()}</h4><h5>Night Duty Allowance Calculation from {period_str}</h5><br>"""
    
    html += "<h4>PART I: MONTH-WISE ATTENDANCE REGISTER</h4>"
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        date_cols = get_formatted_dates(y, m)
        att_df = att_dict[(m, y)]
        if sum(sum(1 for d in date_cols if row[d] == True) for _, row in att_df.iterrows()) == 0: continue
            
        html += f"<div class='header-row'>ATTENDANCE REGISTER: {m_name.upper()} {y}</div>"
        html += "<table><tr><th>Emp No</th><th>Name</th>"
        for d in range(1, len(date_cols) + 1): html += f"<th>{d:02d}</th>"
        html += "<th>Total</th></tr>"
        for _, row in att_df.iterrows():
            duties_count = sum(1 for d_col in date_cols if row[d_col])
            if duties_count == 0: continue
            html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name'][:20]}</td>"
            for d_col in date_cols:
                mark = "<span style='font-size:14pt;'>&#9679;</span>" if row[d_col] else ""
                html += f"<td>{mark}</td>"
            html += f"<td style='font-weight:bold;'>{duties_count}</td></tr>"
        html += "</table><br>"

    html += """<br><table style="border: none; width: 100%;"><tr style="border: none;">
        <td style="border: none; text-align: left; width: 50%;">_______________________<br>Section In-Charge</td>
        <td style="border: none; text-align: right; width: 50%;">_______________________<br>Office Spdt/Estb In-Charge</td></tr></table>
        <br><br><div style="text-align: center;">_______________________<br>Head of Office<br>(Counter Signed)</div>"""

    html += f"<h3 style='page-break-before: always;'>{ministry.upper()}</h3>"
    html += f"<h4>{dept.upper()}, {office.upper()}</h4>"
    html += "<h4>PART II: MONTH-WISE FINANCIAL BILL & GRAND SUMMARY</h4>"
    
    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty or month_logs["Duties"].sum() == 0: continue
        da_val = edited_da_df[(edited_da_df['Month'] == m_name) & (edited_da_df['Year'] == y)]['DA (%)'].values[0]
        
        html += f"<div class='header-row'>BILL FOR {m_name.upper()} {y} (DA: {da_val}%)</div>"
        html += f"<table><tr><th>Emp No</th><th>Name</th><th>Post</th><th>Basic Pay</th><th>DA %</th><th>Duties</th><th>1 Hr Rate</th><th>{rate_header}</th><th>Net NDA (&#8377;)</th></tr>"
        month_total_amt = 0.0
        for _, row in month_logs.iterrows():
            bp = row["Basic Pay"]
            hr_rate = (min(bp, 43600) * (1 + da_val/100)) / 200
            day_rate = hr_rate * (night_hours_per_duty / 6)
            net_amt = row["NDA Paid (₹)"]
            month_total_amt += net_amt
            html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name']}</td><td style='text-align:left;'>{row['Post']}</td><td>{bp:,.0f}</td><td>{int(da_val)}</td><td style='font-weight:bold;'>{row['Duties']}</td><td>{hr_rate:,.2f}</td><td>{day_rate:,.2f}</td><td style='font-weight:bold;'>{net_amt:,.2f}</td></tr>"
        html += f"<tr><td colspan='8' style='text-align:right; font-weight:bold;'>TOTAL NDA FOR {m_name.upper()} {y}:</td><td style='font-weight:bold;'>&#8377; {month_total_amt:,.2f}</td></tr></table>"

    html += "<br><h4>GRAND FINAL SUMMARY</h4>"
    html += "<table><tr><th>Emp No</th><th>Name</th><th>Post</th>"
    for mc in month_cols: html += f"<th>{mc}</th>"
    html += "<th>Total Duties</th><th>Total Hrs</th><th>Wt Hrs</th><th>Total Net NDA (&#8377;)</th></tr>"
    grand_total_all = 0.0
    for _, row in final_bill.iterrows():
        total_hrs = row['Total Night Duties'] * night_hours_per_duty
        wt_hrs = total_hrs / 6
        html += f"<tr><td>{row['Emp No']}</td><td style='text-align:left;'>{row['Name']}</td><td style='text-align:left;'>{row['Post']}</td>"
        for mc in month_cols: html += f"<td>{row[mc]}</td>"
        html += f"<td style='font-weight:bold;'>{row['Total Night Duties']}</td><td>{total_hrs}</td><td>{wt_hrs:,.2f}</td><td style='font-weight:bold;'>{row['Net Payable NDA (₹)']:,.2f}</td></tr>"
        grand_total_all += row['Net Payable NDA (₹)']
        
    rounded_grand_total = round(grand_total_all)
    amount_words = number_to_words(rounded_grand_total)
    
    html += f"<tr><td colspan='{6 + len(month_cols)}' style='text-align:right; font-weight:bold;'>GRAND TOTAL PAYOUT:</td><td style='font-weight:bold; font-size:11pt;'>&#8377; {grand_total_all:,.2f}</td></tr>"
    html += f"<tr><td colspan='{6 + len(month_cols)}' style='text-align:right; font-weight:bold;'>ROUNDED GRAND TOTAL:</td><td style='font-weight:bold; font-size:11pt;'>&#8377; {rounded_grand_total:,.0f}</td></tr>"
    html += f"<tr><td colspan='{7 + len(month_cols)}' style='text-align:center; font-weight:bold; font-style:italic;'>({amount_words})</td></tr></table>"
        
    html += """<br><br><table style="border: none; width: 100%;"><tr style="border: none;">
        <td style="border: none; text-align: left; width: 50%;">_______________________<br>Office Spdt/Estb In-Charge</td>
        <td style="border: none; text-align: right; width: 50%;">_______________________<br>Accountant/Accounts Officer</td></tr></table>
        <br><br><div style="text-align: center;">_______________________<br>Head of Office<br>(Counter Signed)</div>
    </body></html>"""
    return html.encode('utf-8')

# ==========================================
# 7. PDF GENERATION FUNCTION
# ==========================================
def generate_official_pdf(ministry, dept, office, period_str, active_months, att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols):
    pdf = FPDF(orientation='L', unit='mm', format='A4') 
    
    # ---------------------------------------------------------
    # PART I: VISUAL ATTENDANCE REGISTER
    # ---------------------------------------------------------
    pdf.add_page()
    pdf.set_font("Arial", 'B', 15)
    pdf.cell(0, 7, txt=ministry.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, txt=dept.upper() + ", " + office.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'BU', 11)
    pdf.cell(0, 7, txt=f"Night Duty Allowance Calculation from {period_str}", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 8, txt="PART I: MONTH-WISE ATTENDANCE REGISTER", ln=True, align='C')
    pdf.ln(3)

    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        date_cols = get_formatted_dates(y, m)
        att_df = att_dict[(m, y)]
        
        has_duties = sum(sum(1 for d in date_cols if row[d] == True) for _, row in att_df.iterrows())
        if has_duties == 0: continue
            
        if pdf.get_y() > 140: pdf.add_page()
            
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(220, 230, 240) 
        pdf.cell(0, 7, txt=f" ATTENDANCE REGISTER: {m_name.upper()} {y} ", ln=True, align='L', fill=True)
        
        pdf.set_font("Arial", 'B', 8)
        pdf.set_fill_color(240, 240, 240)
        num_days = len(date_cols)
        day_w = 6.2
        pdf.cell(14, 6, "Emp No", 1, 0, 'C', fill=True)
        pdf.cell(45, 6, "Name", 1, 0, 'C', fill=True)
        for d in range(1, num_days + 1): pdf.cell(day_w, 6, f"{d:02d}", 1, 0, 'C', fill=True)
        pdf.cell(12, 6, "Total", 1, 1, 'C', fill=True)
        
        for _, row in att_df.iterrows():
            duties_count = sum(1 for d_col in date_cols if row[d_col])
            if duties_count == 0: continue
            
            pdf.set_font("Arial", '', 8)
            pdf.cell(14, 6, str(row['Emp No']), 1, 0, 'C')
            pdf.cell(45, 6, str(row['Name'])[:28], 1, 0, 'L')
            for d_col in date_cols:
                if row[d_col]:
                    pdf.set_font("Arial", '', 13) 
                    pdf.cell(day_w, 6, chr(149), 1, 0, 'C')
                    pdf.set_font("Arial", '', 8) 
                else:
                    pdf.cell(day_w, 6, "", 1, 0, 'C')
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(12, 6, str(duties_count), 1, 1, 'C')
        pdf.ln(5)
        
    if pdf.get_y() > 155: pdf.add_page()
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(138, 5, "___________________________", align='L')
    pdf.cell(0, 5, "___________________________", align='R', ln=True)
    pdf.cell(138, 5, "Section In-Charge", align='L')
    pdf.cell(0, 5, "Office Spdt/Estb In-Charge", align='R', ln=True)
    pdf.ln(10)
    pdf.cell(0, 5, "___________________________", align='C', ln=True)
    pdf.cell(0, 5, "Head of Office", align='C', ln=True)
    pdf.cell(0, 5, "(Counter Signed)", align='C', ln=True)

    # ---------------------------------------------------------
    # PART II: MONTH-WISE FINANCIAL BILL & GRAND SUMMARY
    # ---------------------------------------------------------
    pdf.add_page()
    pdf.set_font("Arial", 'B', 15)
    pdf.cell(0, 7, txt=ministry.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 6, txt=dept.upper() + ", " + office.upper(), ln=True, align='C')
    pdf.set_font("Arial", 'BU', 12)
    pdf.cell(0, 8, txt="PART II: MONTH-WISE FINANCIAL BILL & GRAND SUMMARY", ln=True, align='C')
    pdf.ln(3)

    for period in active_months:
        m, y = period['m'], period['y']
        m_name = calendar.month_name[m]
        month_logs = logs_df[(logs_df["Year"] == y) & (logs_df["Month"] == m_name)]
        if month_logs.empty or month_logs["Duties"].sum() == 0: continue
            
        da_val = edited_da_df[(edited_da_df['Month'] == m_name) & (edited_da_df['Year'] == y)]['DA (%)'].values[0]
        
        if pdf.get_y() > 150: pdf.add_page()
        
        pdf.set_font("Arial", 'B', 10)
        pdf.set_fill_color(220, 230, 240)
        pdf.cell(0, 7, txt=f" BILL FOR {m_name.upper()} {y} (DA: {da_val}%) ", ln=True, align='L', fill=True)
        
        pdf.set_font("Arial", 'B', 9)
        pdf.set_fill_color(240, 240, 240)
        bw = [18, 55, 25, 25, 15, 20, 25, 30, 30] 
        
        pdf.cell(bw[0], 6, "Emp No", 1, 0, 'C', fill=True)
        pdf.cell(bw[1], 6, "Name", 1, 0, 'C', fill=True)
        pdf.cell(bw[2], 6, "Post", 1, 0, 'C', fill=True)
        pdf.cell(bw[3], 6, "Basic Pay", 1, 0, 'C', fill=True)
        pdf.cell(bw[4], 6, "DA%", 1, 0, 'C', fill=True)
        pdf.cell(bw[5], 6, "Duties", 1, 0, 'C', fill=True)
        pdf.cell(bw[6], 6, "1 Hr Rate", 1, 0, 'C', fill=True)
        pdf.cell(bw[7], 6, f"{int(night_hours_per_duty * 10)} min Rate", 1, 0, 'C', fill=True)
        pdf.cell(bw[8], 6, "Net NDA", 1, 1, 'C', fill=True)
        
        pdf.set_font("Arial", '', 9)
        month_total_amt = 0.0
        for _, row in month_logs.iterrows():
            bp = row["Basic Pay"]
            hr_rate = (min(bp, 43600) * (1 + da_val/100)) / 200
            day_rate = hr_rate * (night_hours_per_duty / 6)
            net_amt = row["NDA Paid (₹)"]
            month_total_amt += net_amt
            
            pdf.cell(bw[0], 7, str(row['Emp No']), 1, 0, 'C')
            pdf.cell(bw[1], 7, str(row['Name'])[:30], 1, 0, 'L')
            pdf.cell(bw[2], 7, str(row['Post']), 1, 0, 'L')
            pdf.cell(bw[3], 7, f"{bp:,.0f}", 1, 0, 'C')
            pdf.cell(bw[4], 7, str(int(da_val)), 1, 0, 'C')
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(bw[5], 7, str(row['Duties']), 1, 0, 'C')
            pdf.set_font("Arial", '', 9)
            pdf.cell(bw[6], 7, f"{hr_rate:,.2f}", 1, 0, 'C')
            pdf.cell(bw[7], 7, f"{day_rate:,.2f}", 1, 0, 'C')
            pdf.set_font("Arial", 'B', 9)
            pdf.cell(bw[8], 7, f"Rs. {net_amt:,.2f}", 1, 1, 'C')
            pdf.set_font("Arial", '', 9)
            
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(sum(bw[:-1]), 8, f"TOTAL NDA FOR {m_name.upper()} {y}:   ", 1, 0, 'R')
        pdf.cell(bw[-1], 8, f"Rs. {month_total_amt:,.2f}", 1, 1, 'C')
        pdf.ln(5) 
        
    # --- GRAND SUMMARY ---
    if pdf.get_y() > 130: pdf.add_page()
    pdf.ln(3)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(0, 8, txt="GRAND FINAL SUMMARY", ln=True, align='C')
    
    pdf.set_font("Arial", 'B', 9)
    pdf.set_fill_color(220, 230, 240)
    sw_base = [15, 45, 15] 
    sw_end = [16, 20, 20, 25] 
    usable_width = 277 - sum(sw_base) - sum(sw_end)
    m_w = usable_width / len(month_cols) if month_cols else 15
    
    pdf.cell(sw_base[0], 7, "Emp No", 1, 0, 'C', fill=True)
    pdf.cell(sw_base[1], 7, "Name", 1, 0, 'C', fill=True)
    pdf.cell(sw_base[2], 7, "Post", 1, 0, 'C', fill=True)
    for mc in month_cols: pdf.cell(m_w, 7, mc, 1, 0, 'C', fill=True)
    pdf.cell(sw_end[0], 7, "Tot Duties", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[1], 7, "Total Hrs", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[2], 7, "Wt Hrs", 1, 0, 'C', fill=True)
    pdf.cell(sw_end[3], 7, "Net NDA", 1, 1, 'C', fill=True)
    
    pdf.set_font("Arial", '', 9)
    grand_total_all = 0.0
    for _, row in final_bill.iterrows():
        total_hrs = row['Total Night Duties'] * night_hours_per_duty
        wt_hrs = total_hrs / 6
        pdf.cell(sw_base[0], 7, str(row['Emp No']), 1, 0, 'C')
        pdf.cell(sw_base[1], 7, row['Name'][:25], 1, 0, 'L')
        pdf.cell(sw_base[2], 7, row['Post'], 1, 0, 'L')
        
        pdf.set_font("Arial", 'B', 9)
        for mc in month_cols: pdf.cell(m_w, 7, str(row[mc]), 1, 0, 'C')
        
        pdf.cell(sw_end[0], 7, str(row['Total Night Duties']), 1, 0, 'C')
        pdf.set_font("Arial", '', 9)
        pdf.cell(sw_end[1], 7, str(total_hrs), 1, 0, 'C')
        pdf.cell(sw_end[2], 7, f"{wt_hrs:,.2f}", 1, 0, 'C')
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(sw_end[3], 7, f"Rs. {row['Net Payable NDA (₹)']:,.2f}", 1, 1, 'C')
        pdf.set_font("Arial", '', 9)
        grand_total_all += row['Net Payable NDA (₹)']
        
    pdf.set_font("Arial", 'B', 11)
    total_w_before_amt = sum(sw_base) + (m_w * len(month_cols)) + sum(sw_end[:-1])
    
    rounded_grand_total = round(grand_total_all)
    amount_words = number_to_words(rounded_grand_total)
    
    pdf.cell(total_w_before_amt, 9, "GRAND TOTAL PAYOUT:   ", 1, 0, 'R')
    pdf.cell(sw_end[-1], 9, f"Rs. {grand_total_all:,.2f}", 1, 1, 'C')
    
    pdf.cell(total_w_before_amt, 9, "ROUNDED GRAND TOTAL:   ", 1, 0, 'R')
    pdf.cell(sw_end[-1], 9, f"Rs. {rounded_grand_total:,.0f}", 1, 1, 'C')
    
    pdf.set_font("Arial", 'BI', 10)
    pdf.cell(0, 8, f"({amount_words})", 0, 1, 'C')

    if pdf.get_y() > 155: pdf.add_page()
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(138, 5, "___________________________", align='L')
    pdf.cell(0, 5, "___________________________", align='R', ln=True)
    pdf.cell(138, 5, "Office Spdt/Estb In-Charge", align='L')
    pdf.cell(0, 5, "Accountant/Accounts Officer", align='R', ln=True)
    pdf.ln(10)
    pdf.cell(0, 5, "___________________________", align='C', ln=True)
    pdf.cell(0, 5, "Head of Office", align='C', ln=True)
    pdf.cell(0, 5, "(Counter Signed)", align='C', ln=True)

    return pdf.output(dest='S').encode('latin1')

# ==========================================
# 8. MAIN UI NAVIGATION
# ==========================================
st.title("📝 CleanReport NDA Register")
st.markdown("Automated 7th CPC Multi-Month Night Duty Allowance Generator (Built-in DA & 10-Min Weightage Formula).")
st.markdown("---")

nav_options = ["🏛️ 1. Employee Master", "📅 2. Mark Attendance", "📊 3. Generate Reports"]
if 'current_nav' not in st.session_state: st.session_state.current_nav = nav_options[0]
selected_tab = st.radio("Navigation", nav_options, key="current_nav", horizontal=True, label_visibility="collapsed")
st.markdown("<br>", unsafe_allow_html=True)

# --- TAB 1: EMPLOYEE MASTER ---
if selected_tab == nav_options[0]:
    st.markdown("### 🏢 Office Details (For Report Header)")
    c1, c2, c3 = st.columns(3)
    with c1: st.session_state.ministry = st.text_input("Name of Ministry", value=st.session_state.ministry)
    with c2: st.session_state.department = st.text_input("Name of Department", value=st.session_state.department)
    with c3: st.session_state.office = st.text_input("Name of Office", value=st.session_state.office)
    
    st.markdown("---")
    st.markdown("### 👥 Add New Employee (Initial Data)")
    col1, col2, col3, col4, col5, col6 = st.columns([1.5, 2.5, 1.5, 2, 2, 1])
    with col1: st.text_input("Personal/Emp No", key="in_empno")
    with col2: st.text_input("Full Name", key="in_name")
    with col3: st.text_input("Post Held", key="in_post")
    with col4: sel_level = st.selectbox("7th CPC Pay Level", list(PAY_MATRIX.keys()))
    with col5: sel_bp = st.selectbox("Basic Pay (₹)", PAY_MATRIX[sel_level])
    with col6:
        st.markdown("<br>", unsafe_allow_html=True)
        st.button("➕ Add", type="primary", use_container_width=True, on_click=add_employee_callback, args=(sel_level, sel_bp))
                
    st.markdown("### ✏️ Current Employee Database")
    if not st.session_state.emp_db.empty:
        df_display = st.session_state.emp_db.copy()
        df_display["🗑️ Delete"] = False
        col_cfg_master = {
            "Basic Pay (₹)": st.column_config.SelectboxColumn("Basic Pay (₹)", options=ALL_BASIC_PAYS, required=True),
            "Level": st.column_config.SelectboxColumn("Level", options=list(PAY_MATRIX.keys()), required=True)
        }
        edited_df = st.data_editor(df_display, use_container_width=True, hide_index=True, column_config=col_cfg_master)
        if st.button("🗑️ Remove Selected Employees"):
            st.session_state.emp_db = edited_df[~edited_df["🗑️ Delete"]].drop(columns=["🗑️ Delete"])
            st.rerun()
    else: st.info("No employees added yet.")
    st.markdown("<br><hr>", unsafe_allow_html=True)
    st.button("Next: Mark Attendance ➡️", on_click=go_to_tab, args=(nav_options[1],), type="secondary")

# --- TAB 2: MARK ATTENDANCE ---
elif selected_tab == nav_options[1]:
    if st.session_state.emp_db.empty: st.warning("⚠️ Please add at least one employee in Step 1 first.")
    else:
        if not st.session_state.active_months:
            c1, c2, c3 = st.columns([2, 2, 4])
            with c1: start_month = st.selectbox("Select Start Month", range(1, 13), format_func=lambda x: calendar.month_name[x], index=datetime.now().month-1)
            with c2: start_year = st.selectbox("Select Start Year", range(2020, 2031), index=datetime.now().year - 2020)
            with c3:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("🚀 Start Register"):
                    st.session_state.active_months.append({'m': start_month, 'y': start_year})
                    st.rerun()
        
        if st.session_state.active_months:
            for period in st.session_state.active_months:
                m, y = period['m'], period['y']
                date_cols = get_formatted_dates(y, m)
                
                if (m, y) not in st.session_state.att_dict:
                    new_df = st.session_state.emp_db.copy().drop_duplicates("Emp No")
                    for col in date_cols: new_df[col] = False
                    st.session_state.att_dict[(m, y)] = new_df
                else:
                    existing_att = st.session_state.att_dict[(m, y)].copy().drop_duplicates("Emp No")
                    base_df = st.session_state.emp_db.copy().drop_duplicates("Emp No")
                    
                    # --- FIXED AUTO-WIPE BUG ---
                    # Merge securely to protect existing date data
                    merged_df = pd.merge(base_df, existing_att.drop(columns=["Name", "Post", "Level", "Basic Pay (₹)"], errors='ignore'), on="Emp No", how="left")
                    for col in date_cols:
                        if col not in merged_df.columns:
                            merged_df[col] = False
                        merged_df[col] = merged_df[col].fillna(False).astype(bool)
                    st.session_state.att_dict[(m, y)] = merged_df

            for idx, period in enumerate(st.session_state.active_months):
                m, y = period['m'], period['y']
                st.markdown("---")
                st.markdown(f"#### 📅 {get_month_name(m, y)} &nbsp;&nbsp; *(Applicable DA: {get_auto_da(m, y)}%)*")
                
                date_cols = get_formatted_dates(y, m)
                full_df = st.session_state.att_dict[(m, y)]

                idx_key = f"idx_{m}_{y}"
                if idx_key not in st.session_state: st.session_state[idx_key] = 0
                if st.session_state[idx_key] >= len(date_cols): st.session_state[idx_key] = 0
                
                ed_key_var = f"ed_key_{m}_{y}"
                
                # --- ⚡ SMART HORIZONTAL ROSTER (FAST ENTRY) ---
                st.markdown("##### ⚡ Daily Duty Roster (Fast Entry)")
                
                c_date, c_cards, c_btn = st.columns([2, 6, 2])
                
                with c_date:
                    sel_date = st.selectbox("📅 Date:", date_cols, index=st.session_state[idx_key], key=f"sb_dummy_{m}_{y}_{st.session_state.get(ed_key_var, 0)}")
                    st.session_state[idx_key] = date_cols.index(sel_date)
                    
                with c_cards:
                    st.markdown("<div style='margin-top:2px;'></div>", unsafe_allow_html=True)
                    cb_cols = st.columns(3) # 3 columns for better spacing
                    
                    for i, (_, row) in enumerate(full_df.iterrows()):
                        emp_id = str(row['Emp No'])
                        emp_name = str(row['Name']).split()[0] 
                        current_val = bool(row[sel_date])
                        cb_key = f"cb_{m}_{y}_{emp_id}_{sel_date}"
                        
                        if cb_key not in st.session_state:
                            st.session_state[cb_key] = current_val
                        
                        with cb_cols[i % 3]:
                            st.checkbox(f"🧑‍💼 {emp_id} - {emp_name}", value=current_val, key=cb_key, on_change=sync_cb, args=(m, y, emp_id, sel_date, cb_key))
                
                with c_btn:
                    st.markdown("<div style='margin-top:2px;'></div>", unsafe_allow_html=True)
                    if st.button("➡️ Next Day", key=f"next_btn_{m}_{y}", type="primary", use_container_width=True):
                        curr_idx = date_cols.index(sel_date)
                        if curr_idx < len(date_cols) - 1:
                            st.session_state[idx_key] = curr_idx + 1
                        else:
                            st.session_state[idx_key] = 0
                            st.toast("Reached the end of the month!")
                        
                        st.session_state[ed_key_var] = st.session_state.get(ed_key_var, 0) + 1
                        st.rerun()

                # --- ✏️ MANUAL GRID OVERRIDE ---
                st.markdown("##### ✏️ Manual Grid Override (Optional)")
                
                display_df = full_df.copy()
                display_df.index = display_df["Emp No"].astype(str) + " | " + display_df["Name"].astype(str)
                display_df.index.name = "Emp No | Name"
                display_df.rename(columns={"Basic Pay (₹)": "Basic Pay (₹) ✏️▾"}, inplace=True)
                cols_to_show = ["Basic Pay (₹) ✏️▾"] + date_cols
                display_df = display_df[cols_to_show] 
                
                col_cfg = {col: st.column_config.CheckboxColumn(width="small") for col in date_cols}
                col_cfg["Basic Pay (₹) ✏️▾"] = st.column_config.SelectboxColumn("Basic Pay (₹) ✏️▾", options=ALL_BASIC_PAYS, required=True, width="medium")
                
                editor_key = f"grid_{m}_{y}_{st.session_state.get(ed_key_var, 0)}"
                
                edited_display_df = st.data_editor(
                    display_df, key=editor_key, hide_index=False, use_container_width=True,
                    column_config=col_cfg, height=200 + (len(display_df) * 35)
                )
                
                for col in date_cols: st.session_state.att_dict[(m, y)][col] = edited_display_df[col].values
                st.session_state.att_dict[(m, y)]["Basic Pay (₹)"] = edited_display_df["Basic Pay (₹) ✏️▾"].values

            st.markdown("---")
            last_m, last_y = st.session_state.active_months[-1]['m'], st.session_state.active_months[-1]['y']
            next_m, next_y = last_m % 12 + 1, last_y + (last_m // 12)
            
            c_add, c_del, c_reset = st.columns([1, 1, 1])
            with c_add:
                if st.button(f"➕ Add Next Month ({get_month_name(next_m, next_y)})", use_container_width=True):
                    st.session_state.active_months.append({'m': next_m, 'y': next_y})
                    last_att = st.session_state.att_dict[(last_m, last_y)]
                    for _, row in last_att.iterrows():
                        st.session_state.emp_db.loc[st.session_state.emp_db["Emp No"].astype(str) == str(row["Emp No"]), "Basic Pay (₹)"] = row["Basic Pay (₹)"]
                    st.rerun()
            with c_del:
                if st.button("➖ Delete Last Month", type="secondary", use_container_width=True):
                    if len(st.session_state.active_months) > 1:
                        removed = st.session_state.active_months.pop()
                        st.session_state.att_dict.pop((removed['m'], removed['y']), None)
                        st.rerun()
            with c_reset:
                if st.button("🗑️ Reset Calendar", type="secondary", use_container_width=True):
                    st.session_state.active_months = []
                    st.session_state.att_dict = {}
                    st.session_state.report_generated = False
                    st.rerun()

        st.markdown("<br><hr>", unsafe_allow_html=True)
        st.button("Next: Generate Reports ➡️", on_click=go_to_tab, args=(nav_options[2],), type="secondary")

# --- TAB 3: GENERATE REPORTS ---
elif selected_tab == nav_options[2]:
    if not st.session_state.active_months: st.warning("⚠️ Please mark attendance in Step 2 first.")
    else:
        st.markdown("### ⚙️ DA Rate Configuration")
        st.info("💡 **For months beyond Dec 2026:** The auto DA is set to 0 by default. Double-click the cell to manually enter the DA %.")
        da_df = pd.DataFrame([{"Month": calendar.month_name[p['m']], "Year": p['y'], "DA (%)": get_auto_da(p['m'], p['y'])} for p in st.session_state.active_months])
        edited_da_df = st.data_editor(da_df, hide_index=True, use_container_width=True)
        st.markdown("---")
        night_hours_per_duty = st.number_input("Actual Night Hours per Shift (e.g. 22:00 to 06:00 = 8 hrs)", min_value=1.0, value=8.0, step=0.5)
            
        if st.button("🚀 Generate Final Bill & Comprehensive Reports", type="primary"):
            detailed_logs = []
            
            for (m, y), att_df in st.session_state.att_dict.items():
                da_val = edited_da_df[(edited_da_df['Month'] == calendar.month_name[m]) & (edited_da_df['Year'] == y)]['DA (%)'].values[0]
                date_cols = get_formatted_dates(y, m)
                for index, emp_row in att_df.iterrows():
                    duties_this_month = sum([1 for d_col in date_cols if emp_row[d_col] == True])
                    if duties_this_month > 0:
                        capped_bp = min(emp_row["Basic Pay (₹)"], 43600)
                        hourly_rate = (capped_bp * (1 + da_val/100)) / 200
                        nda_amount = hourly_rate * (night_hours_per_duty / 6.0) * duties_this_month
                        dates_str = ", ".join([d_col.split(' ')[0] for d_col in date_cols if emp_row[d_col] == True])
                        
                        detailed_logs.append({
                            "Emp No": emp_row["Emp No"], "Name": emp_row["Name"], "Post": emp_row["Post"],
                            "Basic Pay": emp_row["Basic Pay (₹)"], "Month": calendar.month_name[m], "Year": y, 
                            "DA %": da_val, "Dates_Str": dates_str, "Duties": duties_this_month, "NDA Paid (₹)": nda_amount
                        })
                            
            if not detailed_logs: st.error("No Night Duties marked! Please tick checkboxes in Step 2.")
            else:
                logs_df = pd.DataFrame(detailed_logs)
                final_bill = logs_df.groupby(["Emp No"]).agg(Total_Duties=("Duties", "sum"), Total_NDA_Amount=("NDA Paid (₹)", "sum")).reset_index()
                
                month_cols = []
                for period in st.session_state.active_months:
                    m, y = period['m'], period['y']
                    m_name = calendar.month_name[m]
                    col_name = f"{calendar.month_abbr[m]} '{str(y)[-2:]}"
                    month_cols.append(col_name)
                    month_duties = logs_df[(logs_df["Month"] == m_name) & (logs_df["Year"] == y)].groupby("Emp No")["Duties"].sum().reset_index()
                    month_duties.rename(columns={"Duties": col_name}, inplace=True)
                    final_bill = pd.merge(final_bill, month_duties, on="Emp No", how="left").fillna({col_name: 0})
                    final_bill[col_name] = final_bill[col_name].astype(int)

                final_bill = pd.merge(final_bill, st.session_state.emp_db[["Emp No", "Name", "Post"]], on="Emp No", how="left")
                final_col_order = ["Emp No", "Name", "Post"] + month_cols + ["Total_Duties", "Total_NDA_Amount"]
                final_bill = final_bill[final_col_order]
                final_bill.rename(columns={"Total_Duties": "Total Night Duties", "Total_NDA_Amount": "Net Payable NDA (₹)"}, inplace=True)
                
                period_str = f"{get_month_name(st.session_state.active_months[0]['m'], st.session_state.active_months[0]['y'])} to {get_month_name(st.session_state.active_months[-1]['m'], st.session_state.active_months[-1]['y'])}"
                
                st.session_state.final_bill_df = final_bill
                st.session_state.period_str = period_str
                st.session_state.excel_data = generate_dynamic_excel(st.session_state.ministry, st.session_state.department, st.session_state.office, period_str, st.session_state.active_months, st.session_state.att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols)
                st.session_state.pdf_bytes = generate_official_pdf(st.session_state.ministry, st.session_state.department, st.session_state.office, period_str, st.session_state.active_months, st.session_state.att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols)
                st.session_state.word_bytes = generate_word_html(st.session_state.ministry, st.session_state.department, st.session_state.office, period_str, st.session_state.active_months, st.session_state.att_dict, edited_da_df, final_bill, logs_df, night_hours_per_duty, month_cols)
                st.session_state.report_generated = True

        if st.session_state.get('report_generated', False):
            st.markdown("---")
            st.markdown("### Grand Final Summary")
            st.dataframe(st.session_state.final_bill_df, use_container_width=True, hide_index=True)
            
            st.markdown("### 📥 Download Center")
            c1, c2, c3 = st.columns(3)
            with c1: st.download_button("📊 Download CleanReport (.xlsx)", data=st.session_state.excel_data, file_name=f"CleanReport_NDA_{st.session_state.period_str}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with c2: st.download_button("📄 Download CleanReport (.pdf)", data=st.session_state.pdf_bytes, file_name=f"CleanReport_NDA_{st.session_state.period_str}.pdf", mime="application/pdf", use_container_width=True)
            with c3: st.download_button("📝 Download CleanReport (.doc)", data=st.session_state.word_bytes, file_name=f"CleanReport_NDA_{st.session_state.period_str}.doc", mime="application/msword", use_container_width=True)
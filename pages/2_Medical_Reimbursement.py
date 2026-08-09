import streamlit as st
import pandas as pd
import datetime
from fpdf import FPDF
import base64
import os

st.set_page_config(page_title="Medical Reimbursement Calculator", page_icon="🏥", layout="wide")

# ==========================================
# 1. INDIAN STATES & DISTRICTS DICTIONARY
# ==========================================
india_locations = {
    "Rajasthan": ["Ajmer", "Alwar", "Banswara", "Baran", "Barmer", "Bharatpur", "Bhilwara", "Bikaner", "Bundi", "Chittorgarh", "Churu", "Dausa", "Dholpur", "Dungarpur", "Hanumangarh", "Jaipur", "Jaisalmer", "Jalore", "Jhalawar", "Jhunjhunu", "Jodhpur", "Karauli", "Kota", "Kotputli", "Nagaur", "Pali", "Pratapgarh", "Rajsamand", "Sawai Madhopur", "Sikar", "Sirohi", "Sri Ganganagar", "Tonk", "Udaipur"],
    "Delhi": ["Central Delhi", "East Delhi", "New Delhi", "North Delhi", "North East Delhi", "North West Delhi", "Shahdara", "South Delhi", "South East Delhi", "South West Delhi", "West Delhi"],
    "Haryana": ["Ambala", "Bhiwani", "Charkhi Dadri", "Faridabad", "Fatehabad", "Gurugram", "Hisar", "Jhajjar", "Jind", "Kaithal", "Karnal", "Kurukshetra", "Mahendragarh", "Nuh", "Palwal", "Panchkula", "Panipat", "Rewari", "Rohtak", "Sirsa", "Sonipat", "Yamunanagar"],
    "Uttar Pradesh": ["Agra", "Aligarh", "Allahabad", "Ayodhya", "Bareilly", "Ghaziabad", "Gorakhpur", "Jhansi", "Kanpur", "Lucknow", "Mathura", "Meerut", "Moradabad", "Noida (GB Nagar)", "Saharanpur", "Varanasi"],
    "Maharashtra": ["Ahmednagar", "Amravati", "Aurangabad", "Jalgaon", "Kolhapur", "Mumbai City", "Mumbai Suburban", "Nagpur", "Nashik", "Pune", "Solapur", "Thane"],
    "Other State": ["Other District"]
}

# ==========================================
# 2. LOAD CGHS EXCEL FILE & FORMAT DROPDOWN
# ==========================================
@st.cache_data
def load_cghs_data():
    try:
        df = pd.read_excel('CGHS_Rates_Clean_Excel_Tier_I.xlsx')
        df = df.dropna(subset=['Treatment / Investigation'])
        df['Display_Name'] = df['CGHS Code'].astype(str) + " - " + df['Treatment / Investigation'].astype(str)
        return df
    except Exception as e:
        st.error(f"⚠️ Error loading CGHS database: {e}")
        return pd.DataFrame()

cghs_df = load_cghs_data()
procedure_list = cghs_df['Display_Name'].tolist() if not cghs_df.empty else []

# ==========================================
# 3. SIDEBAR - IMPORTANT DOWNLOAD LINKS
# ==========================================
st.sidebar.markdown("### 📥 Important Govt Forms")
st.sidebar.info("Click the buttons below to download essential reimbursement forms.")

try:
    with open("form_97.pdf", "rb") as pdf_file1:
        st.sidebar.download_button(label="📄 Download Med-97 Form", data=pdf_file1, file_name="Med-97_Essentiality_Certificate.pdf", mime="application/pdf", use_container_width=True)
except FileNotFoundError:
    st.sidebar.warning("⚠️ form_97.pdf not found in folder. Please add it to enable download.")

try:
    with open("cghs_mrc.pdf", "rb") as pdf_file2:
        st.sidebar.download_button(label="📄 Download CGHS MRC Form", data=pdf_file2, file_name="CGHS_Reimbursement_Form.pdf", mime="application/pdf", use_container_width=True)
except FileNotFoundError:
    st.sidebar.warning("⚠️ cghs_mrc.pdf not found in folder. Please add it to enable download.")

try:
    with open("cghs_order.pdf", "rb") as pdf_file3:
        st.sidebar.download_button(label="📄 Download CGHS Official Order", data=pdf_file3, file_name="CGHS_Authorized_Rate_List.pdf", mime="application/pdf", use_container_width=True)
except FileNotFoundError:
    st.sidebar.warning("⚠️ cghs_order.pdf not found in folder. Please add it to enable download.")

st.sidebar.markdown("---")
st.sidebar.markdown("**Note:** Please attach the filled forms along with the final generated bill and original receipts.")

# ==========================================
# 4. SESSION STATE SETUP
# ==========================================
if 'bill_items' not in st.session_state:
    st.session_state.bill_items = []
if 'proc_count' not in st.session_state:
    st.session_state.proc_count = 1
if 'form_id' not in st.session_state:
    st.session_state.form_id = 1 

def add_procedure_row():
    st.session_state.proc_count += 1

def sync_dates():
    fid = st.session_state.form_id
    if f"cons_date_{fid}" in st.session_state:
        new_date = st.session_state[f"cons_date_{fid}"]
        for i in range(st.session_state.proc_count):
            if f"p_date_{fid}_{i}" in st.session_state:
                st.session_state[f"p_date_{fid}_{i}"] = new_date

# ==========================================
# 5. UI - PATIENT & HOSPITAL DETAILS
# ==========================================
st.title("🏥 Govt. Medical Reimbursement Calculator")
st.markdown("Calculate admissible claims strictly as per CGHS rates and generate official audit-ready reports.")
st.markdown("---")
st.error("★ **Note:** Medical Procedure/Treatment rates are calculated as per GOI order no. F. No. 5-16/CGHS(HQ)/HEC/2024(Part 1) dt 03.10.2025")

st.subheader("👤 1. Patient & Hospital- General Details")
col1, col2, col3, col4 = st.columns(4)
with col1:
    dept_name = st.text_input("Department/Office Name")
    emp_name = st.text_input("Employee Name")
    emp_id = st.text_input("Employee ID (Optional)")
with col2:
    patient_name = st.text_input("Patient Name")
    relation = st.selectbox("Patient's relation with employee", ["Self", "Spouse", "Son", "Daughter", "Mother", "Father", "Dependent"])
    cghs_no = st.text_input("CGHS Card No. (Optional)")
with col3:
    hospital_name = st.text_input("Hospital's Name")
    doctor_name = st.text_input("Doctor's Name", value="Dr. ")
    hospital_type = st.selectbox("Hospital Category", ["Non-NABH", "NABH", "Super Speciality"], key="hosp_category")
with col4:
    selected_state = st.selectbox("State", list(india_locations.keys()))
    selected_district = st.selectbox("District", india_locations[selected_state])
    treatment_type = st.selectbox("Treatment Type", ["OPD", "OPD - Super Speciality/Psychiatry", "Day Care", "IPD"])

adm_date, dis_date = None, None
if treatment_type in ["Day Care", "IPD"]:
    ca1, ca2 = st.columns(2)
    with ca1:
        adm_date = st.date_input("Admission Date", format="DD/MM/YYYY")
    with ca2:
        dis_date = st.date_input("Discharge Date", format="DD/MM/YYYY")

st.markdown("---")

# ==========================================
# 6. DISPLAY SAVED BILL SUMMARY (SORTED)
# ==========================================
if len(st.session_state.bill_items) > 0:
    st.subheader("📋 2. Final Bill Summary (Chronologically Sorted)")
    
    bill_df = pd.DataFrame(st.session_state.bill_items)
    bill_df['Date_obj'] = pd.to_datetime(bill_df['Date'], format='%d/%m/%Y')
    
    event_dates = bill_df.groupby('Event_ID')['Date_obj'].min().sort_values()
    
    sorted_df_list = []
    display_count = 1
    for event in event_dates.index:
        edf = bill_df[bill_df['Event_ID'] == event].copy()
        edf = edf.sort_values('Date_obj') 
        edf['Event_Phase'] = f"Phase #{display_count}"
        sorted_df_list.append(edf)
        display_count += 1
        
    final_sorted_df = pd.concat(sorted_df_list)
    display_df = final_sorted_df[['Event_Phase', 'Date', 'CGHS_Code', 'Item', 'Billed', 'Admissible']]
    
    st.dataframe(display_df, use_container_width=True)
    
    total_actual = final_sorted_df['Billed'].sum()
    total_admissible = final_sorted_df['Admissible'].sum()
    
    colA, colB = st.columns(2)
    colA.metric("Total Amount Paid (Market Rate)", f"₹ {total_actual}")
    colB.metric("Total Admissible Amount (CGHS Rate)", f"₹ {total_admissible}")
    
    # ----------------------------------------
    # BRANDED PDF GENERATION FUNCTION
    # ----------------------------------------
    def create_pdf():
        pdf = FPDF()
        pdf.add_page()
        
        pdf.set_font("Arial", size=9)
        current_date = datetime.date.today().strftime('%d/%m/%Y')
        pdf.cell(0, 5, txt=f"Date: {current_date}", ln=True, align='R')
        pdf.ln(2)
        
        pdf.set_font("Arial", 'BU', 15)
        pdf.cell(0, 8, txt="Summary of Medical Reimbursement Claim", ln=True, align='C')
        pdf.ln(2)
        
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(0, 5, txt="* Medical Procedure/Treatment rates are calculated as per GOI order no. F. No. 5-16/CGHS(HQ)/HEC/2024(Part 1) dt 03.10.2025", ln=True, align='C')
        pdf.ln(5)
        
        pdf.set_font("Arial", size=9)
        pdf.cell(100, 5, txt=f"Department/Office: {dept_name}", ln=False)
        pdf.cell(100, 5, txt=f"Location: {selected_district}, {selected_state}", ln=True)
        
        pdf.cell(100, 5, txt=f"Employee Name: {emp_name} (ID: {emp_id})", ln=False)
        pdf.cell(100, 5, txt=f"Patient Name: {patient_name} ({relation})", ln=True)
        
        pdf.cell(100, 5, txt=f"Hospital: {hospital_name} ({hospital_type})", ln=False)
        pdf.cell(100, 5, txt=f"Doctor: {doctor_name}", ln=True)
        
        pdf.cell(100, 5, txt=f"CGHS Card No: {cghs_no}", ln=False)
        
        if treatment_type in ["Day Care", "IPD"] and adm_date and dis_date:
            pdf.cell(100, 5, txt=f"Treatment: {treatment_type} ({adm_date.strftime('%d/%m/%Y')} to {dis_date.strftime('%d/%m/%Y')})", ln=True)
        else:
            pdf.cell(100, 5, txt=f"Treatment: {treatment_type}", ln=True)
        pdf.ln(6)
        
        display_event_count_pdf = 1
        for event in event_dates.index:
            event_df = bill_df[bill_df['Event_ID'] == event]
            event_df = event_df.sort_values(by='Date_obj')
            
            pdf.set_font("Arial", 'B', 10)
            pdf.set_fill_color(220, 220, 220)
            pdf.cell(190, 8, txt=f" Treatment Phase #{display_event_count_pdf} ", border=1, ln=True, align='L', fill=True)
            display_event_count_pdf += 1
            
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(20, 8, "Date", border=1, align='C')
            pdf.cell(20, 8, "CGHS Code", border=1, align='C')
            pdf.cell(80, 8, "Item/Procedure", border=1, align='C')
            pdf.cell(20, 8, "CGHS Cap", border=1, align='C')
            pdf.cell(25, 8, "Billed (Rs)", border=1, align='C')
            pdf.cell(25, 8, "Admissible", border=1, align='C', ln=True)
            
            pdf.set_font("Arial", size=8)
            event_billed = 0.0
            event_admissible = 0.0
            
            for index, row in event_df.iterrows():
                pdf.cell(20, 7, str(row['Date']), border=1, align='C')
                pdf.cell(20, 7, str(row['CGHS_Code']), border=1, align='C')
                
                item_name = str(row['Item'])
                if len(item_name) > 45:
                    item_name = item_name[:42] + "..."
                pdf.cell(80, 7, item_name, border=1)
                
                pdf.cell(20, 7, str(row['CGHS_Cap']), border=1, align='C')
                pdf.cell(25, 7, f"{float(row['Billed']):.2f}", border=1, align='C')
                pdf.cell(25, 7, f"{float(row['Admissible']):.2f}", border=1, align='C', ln=True)
                
                event_billed += float(row['Billed'])
                event_admissible += float(row['Admissible'])
            
            pdf.set_font("Arial", 'B', 8)
            pdf.cell(140, 7, "Subtotal", border=1, align='R')
            pdf.cell(25, 7, f"{event_billed:.2f}", border=1, align='C')
            pdf.cell(25, 7, f"{event_admissible:.2f}", border=1, align='C', ln=True)
            pdf.ln(4)

        total_actual_pdf = bill_df['Billed'].sum()
        total_admissible_pdf = bill_df['Admissible'].sum()
        
        pdf.ln(2)
        pdf.set_font("Arial", 'B', 11)
        pdf.cell(190, 8, txt=f"Grand Total Billed Amount: Rs. {total_actual_pdf:.2f}", ln=True, align='R')
        pdf.cell(190, 8, txt=f"Final Admissible Amount: Rs. {total_admissible_pdf:.2f}", ln=True, align='R')
        
        pdf.ln(20)
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(95, 10, "___________________________", 0, 0, 'L')
        pdf.cell(95, 10, "___________________________", 0, 1, 'R')
        pdf.cell(95, 5, "Signature of Dealing Clerk / Section-In-Charge", 0, 0, 'L')
        pdf.cell(95, 5, "Signature of Controlling Authority / HOO", 0, 1, 'R')

        return pdf.output(dest='S').encode('latin1')

    if st.button("📄 Process your Clean Report Now", use_container_width=True, type="primary"):
        pdf_bytes = create_pdf()
        b64 = base64.b64encode(pdf_bytes).decode()
        
        file_identifier = emp_id if emp_id else (patient_name if patient_name else "Summary")
        
        href = f'<br><a href="data:application/pdf;base64,{b64}" download="CleanReport_Reimbursement_{file_identifier}.pdf" style="text-decoration:none; padding:10px; background-color:#4CAF50; color:white; border-radius:5px; font-weight:bold; display:block; text-align:center; margin-top:10px;">⬇️ Download your Clean Report</a>'
        st.markdown(href, unsafe_allow_html=True)
    st.markdown("---")

# ==========================================
# 7. ADD NEW EVENT BUILDER
# ==========================================
st.subheader(f"📅 3. Add Treatment Event (Consultation Entry)")
st.info("Search procedures by CGHS Code or Name. Changing Consultation Date will auto-update procedure dates.")

fid = st.session_state.form_id

c1, c2 = st.columns(2)
with c1:
    c_date = st.date_input("Date of Consultation", datetime.date.today(), format="DD/MM/YYYY", key=f"cons_date_{fid}", on_change=sync_dates)
with c2:
    c_fee = st.number_input("Consultation Fee Paid (₹)", min_value=0.0, step=10.0, key=f"cons_fee_{fid}")

st.markdown("##### 🔬 Diagnostics & Procedures")
for i in range(st.session_state.proc_count):
    cc1, cc2, cc3 = st.columns([1.5, 3, 1.5])
    with cc1:
        p_date = st.date_input(f"Date of Test #{i+1}", value=st.session_state.get(f"cons_date_{fid}", datetime.date.today()), format="DD/MM/YYYY", key=f"p_date_{fid}_{i}")
    with cc2:
        p_name = st.selectbox(f"Search Procedure/Code #{i+1}", ["None"] + procedure_list, key=f"p_name_{fid}_{i}")
    with cc3:
        p_fee = st.number_input(f"Amount Paid (₹) #{i+1}", min_value=0.0, step=10.0, key=f"p_fee_{fid}_{i}")

st.button("➕ Add Another Procedure", on_click=add_procedure_row)

st.markdown("##### 💊 Medicines")
m_fee = st.number_input("Total Amount Paid for Medicines (₹) during this event", min_value=0.0, step=10.0, key=f"med_fee_{fid}")
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 8. SAVE EVENT LOGIC
# ==========================================
if st.button("➡️ Continue to add Next Event", use_container_width=True):
    hosp_type = st.session_state["hosp_category"]
    
    if st.session_state[f"cons_fee_{fid}"] > 0:
        c_amount = st.session_state[f"cons_fee_{fid}"]
        
        # Super Speciality/Psychiatry check
        cghs_c_cap = 700.0 if treatment_type == "OPD - Super Speciality/Psychiatry" else 350.0 
        
        cghs_code = 'CN003' if treatment_type == "OPD - Super Speciality/Psychiatry" else 'CN001'
        item_name = 'Consultation OPD - Super Speciality/Psychiatry' if treatment_type == "OPD - Super Speciality/Psychiatry" else 'Consultation OPD'
        
        st.session_state.bill_items.append({
            'Event_ID': fid,
            'Date': st.session_state[f"cons_date_{fid}"].strftime('%d/%m/%Y'),
            'CGHS_Code': cghs_code,
            'Item': item_name,
            'CGHS_Cap': cghs_c_cap,
            'Billed': c_amount,
            'Admissible': min(c_amount, cghs_c_cap)
        })
        
    for i in range(st.session_state.proc_count):
        p_name_val = st.session_state[f"p_name_{fid}_{i}"]
        p_fee_val = st.session_state[f"p_fee_{fid}_{i}"]
        p_date_val = st.session_state[f"p_date_{fid}_{i}"]
        
        if p_name_val != "None" and p_fee_val > 0:
            cghs_p_cap = 0.0
            p_code = "N/A"
            p_clean_name = p_name_val
            
            try:
                row = cghs_df[cghs_df['Display_Name'] == p_name_val].iloc[0]
                p_code = str(row['CGHS Code'])
                p_clean_name = str(row['Treatment / Investigation'])
                
                if hosp_type == "Non-NABH": cghs_p_cap = float(row['Non-NABH (₹)'])
                elif hosp_type == "NABH": cghs_p_cap = float(row['NABH (₹)'])
                else: cghs_p_cap = float(row['Super Speciality (₹)'])
            except:
                pass
                
            st.session_state.bill_items.append({
                'Event_ID': fid,
                'Date': p_date_val.strftime('%d/%m/%Y'),
                'CGHS_Code': p_code,
                'Item': p_clean_name,
                'CGHS_Cap': cghs_p_cap,
                'Billed': p_fee_val,
                'Admissible': min(p_fee_val, cghs_p_cap)
            })
            
    med_amount = st.session_state[f"med_fee_{fid}"]
    if med_amount > 0:
        st.session_state.bill_items.append({
            'Event_ID': fid,
            'Date': st.session_state[f"cons_date_{fid}"].strftime('%d/%m/%Y'),
            'CGHS_Code': 'MED',
            'Item': 'Medicines (100% Claimable)',
            'CGHS_Cap': 'Full',
            'Billed': med_amount,
            'Admissible': med_amount
        })
        
    st.session_state.form_id += 1
    st.session_state.proc_count = 1
    st.rerun()
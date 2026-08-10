import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import re
import difflib
import docx

st.set_page_config(
    page_title="CleanReport - Ultimate Data Matcher", 
    page_icon="⚡", 
    layout="wide",
    initial_sidebar_state="collapsed"
)

# HIDE SIDEBAR CSS
st.markdown("""
    <style>
    [data-testid="collapsedControl"] {display: none;}
    [data-testid="stSidebar"] {display: none;}
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;}    
    </style>
""", unsafe_allow_html=True)

# BACK BUTTON
if st.button("🏠 Back to Dashboard"):
    st.switch_page("app.py")

# ==========================================
# --- CORE MATCHING ENGINE ---
# ==========================================
def get_pure_core(name):
    if not isinstance(name, str) or str(name).lower() == 'nan': return ""
    n = name.upper()
    n = n.replace("M/S.", " ").replace("M/S", " ").replace("MS.", " ")
    n = re.sub(r'\b(PVT|PRIVATE|LTD|LIMITED|P LTD|CO|CORP|CORPORATION|INC|AND|LLP|INDIA|THE|ENTERPRISES|INDUSTRIES|PRODUCTS|COMPANY)\b', ' ', n)
    return re.sub(r'[^A-Z0-9]', '', n)

def standardize_company(raw_name, master_standards, master_core_dict, master_core_list):
    if not raw_name or str(raw_name).strip() == "": return None
    orig_str = str(raw_name).strip().upper()
    for m in master_standards:
        if orig_str == m.upper(): return m
    core_orig = get_pure_core(orig_str)
    if not core_orig: return None
    if core_orig in master_core_dict: return master_core_dict[core_orig]
    for m_core, m_orig in master_core_dict.items():
        if len(m_core) >= 5 and len(core_orig) >= 5:
            if m_core in core_orig or core_orig in m_core:
                if difflib.SequenceMatcher(None, core_orig, m_core).ratio() > 0.6:
                    return m_orig
    matches = difflib.get_close_matches(core_orig, master_core_list, n=1, cutoff=0.75)
    if matches: return master_core_dict[matches[0]]
    return None

def extract_text_from_file(uploaded_file):
    text = ""
    try:
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            engine = 'calamine' if uploaded_file.name.endswith('.xls') else 'openpyxl'
            df = pd.read_excel(uploaded_file, header=None, engine=engine)
            text = "\n".join(df.iloc[:, 0].dropna().astype(str).tolist())
        elif uploaded_file.name.endswith('.docx'):
            doc = docx.Document(uploaded_file)
            text = "\n".join([p.text for p in doc.paragraphs])
        elif uploaded_file.name.endswith('.txt'):
            text = uploaded_file.read().decode('utf-8')
    except Exception as e:
        st.error(f"File reading error: {e}")
    return text

def read_master_file(uploaded_file):
    engine = 'calamine' if uploaded_file.name.endswith('.xls') else 'openpyxl'
    return pd.read_excel(uploaded_file, engine=engine)

# ==========================================
# --- MAIN AREA ---
# ==========================================
st.title("⚡ CleanReport: Ultimate Excel Data Matcher")
st.markdown("---")

st.subheader("⚙️ 1. Upload Master Files")
c1, c2 = st.columns(2)
with c1:
    sales_file = st.file_uploader("1. Upload Sales File (GC SALE)", type=["xlsx", "xls"])
with c2:
    stock_file = st.file_uploader("2. Upload Stock File (MSTC)", type=["xlsx", "xls"])

st.markdown("<br>", unsafe_allow_html=True)
st.subheader("🎯 2. Input Target Companies")
col1, col2 = st.columns(2)
with col1:
    target_file = st.file_uploader("Option A: Upload List (Excel, Word, TXT)", type=["xlsx", "xls", "docx", "txt"])
with col2:
    raw_text_input = st.text_area("Option B: Paste Raw Names (One name per line)", height=150)

st.markdown("<br>", unsafe_allow_html=True)

if st.button("🚀 Process & Generate Clean Report", use_container_width=True, type="primary"):
    if sales_file and stock_file and (target_file or raw_text_input.strip()):
        with st.spinner("Processing data, calculating metrics, and sorting records..."):
            try:
                df_sales = read_master_file(sales_file)
                df_stock = read_master_file(stock_file)
                
                sales_companies_series = df_sales.iloc[:, 1].astype(str).str.strip().str.upper()
                stock_companies_series = df_stock.iloc[:, 0].astype(str).str.strip().str.upper()
                
                sales_unique_raw = set(sales_companies_series.unique())
                stock_unique_raw = set(stock_companies_series.unique())
                
                master_standards = list(set(list(sales_unique_raw) + list(stock_unique_raw)))
                master_standards = [str(x).strip() for x in master_standards if str(x).strip() != 'NAN']
                
                master_core_dict = {}
                for m_name in master_standards:
                    core = get_pure_core(m_name)
                    if core and core not in master_core_dict:
                        master_core_dict[core] = m_name
                master_core_list = list(master_core_dict.keys())

                combined_text = ""
                if target_file:
                    combined_text += extract_text_from_file(target_file) + "\n"
                if raw_text_input:
                    combined_text += raw_text_input
                
                raw_names_list = [name.strip() for name in combined_text.split('\n') if name.strip()]
                total_provided = len(raw_names_list)
                
                valid_target_names = set()
                not_found_names = []

                for raw_name in raw_names_list:
                    standard_name = standardize_company(raw_name, master_standards, master_core_dict, master_core_list)
                    if standard_name:
                        valid_target_names.add(standard_name)
                    else:
                        not_found_names.append(raw_name)

                found_in_sales = [c for c in valid_target_names if c in sales_unique_raw]
                found_in_stock = [c for c in valid_target_names if c in stock_unique_raw]
                missing_in_sales = [c for c in valid_target_names if c not in sales_unique_raw]
                missing_in_stock = [c for c in valid_target_names if c not in stock_unique_raw]
                
                sales_filtered = df_sales[sales_companies_series.isin(valid_target_names)].copy()
                stock_filtered = df_stock[stock_companies_series.isin(valid_target_names)].copy()
                
                columns_to_drop = ['INDENT_NUM', 'STATE_NAME', 'HSN_CODE']
                sales_filtered = sales_filtered.drop(columns=[col for col in columns_to_drop if col in sales_filtered.columns])
                
                if 'INDENT_DTE' in sales_filtered.columns:
                    sales_filtered['INDENT_DTE_TEMP'] = pd.to_datetime(sales_filtered['INDENT_DTE'], errors='coerce')
                    sales_filtered = sales_filtered.sort_values(by='INDENT_DTE_TEMP', ascending=True)
                    sales_filtered = sales_filtered.drop(columns=['INDENT_DTE_TEMP'])
                
                df_stats = pd.DataFrame({
                    "DATA ANALYTICS METRIC": [
                        "Total Companies Provided (Raw Input)",
                        "Successfully Standardized (Found in Masters)",
                        "Companies Present in SALES File",
                        "Companies Present in STOCK File",
                        "Missing from SALES File",
                        "Missing from STOCK File",
                        "NOT FOUND in Any Master File"
                    ],
                    "COUNT": [
                        total_provided, len(valid_target_names), len(found_in_sales),
                        len(found_in_stock), len(missing_in_sales), len(missing_in_stock), len(not_found_names)
                    ]
                })

                max_len = max(len(missing_in_sales), len(missing_in_stock), len(not_found_names))
                df_lists = pd.DataFrame({
                    "Missing in Sales (List)": missing_in_sales + [""] * (max_len - len(missing_in_sales)),
                    "Missing in Stock (List)": missing_in_stock + [""] * (max_len - len(missing_in_stock)),
                    "Not Found in Any File (List)": not_found_names + [""] * (max_len - len(not_found_names))
                })
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    sales_filtered.to_excel(writer, sheet_name="Sales_Data", index=False)
                    stock_filtered.to_excel(writer, sheet_name="Stock_Data", index=False)
                    df_stats.to_excel(writer, sheet_name="Summary_Analytics", index=False, startrow=0)
                    start_row_for_lists = len(df_stats) + 4
                    df_lists.to_excel(writer, sheet_name="Summary_Analytics", index=False, startrow=start_row_for_lists)
                
                excel_data = output.getvalue()
                
                st.success("✅ Report Generated Successfully!")
                st.subheader("📊 Quick Analytics")
                colA, colB, colC = st.columns(3)
                colA.metric("Provided Companies", total_provided)
                colB.metric("Found in Sales", len(found_in_sales))
                colC.metric("Found in Stock", len(found_in_stock))
                
                st.download_button("📥 Download Final Excel Report", data=excel_data, file_name="Advanced_Automated_Report.xlsx", use_container_width=True)
                
            except Exception as e:
                st.error(f"Error during processing: {e}")
    else:
        st.warning("⚠️ Please upload both Master files and provide target companies!")
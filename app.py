import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Dashboard - CleanReport Workspace",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed" # Sidebar by default band rahega
)

# GOOGLE ANALYTICS INTEGRATION
ga_code = """
<!-- Google tag (gtag.js) -->
<script async src="https://www.googletagmanager.com/gtag/js?id=G-VEH0V9QEEV"></script>
<script>
  window.dataLayer = window.dataLayer || [];
  function gtag(){dataLayer.push(arguments);}
  gtag('js', new Date());

  gtag('config', 'G-VEH0V9QEEV');
</script>
"""
components.html(ga_code, height=0, width=0)

# CSS TO COMPLETELY HIDE SIDEBAR & DEFAULT MENU
hide_st_style = """
    <style>
    /* Hide the sidebar toggle button (hamburger/arrow) */
    [data-testid="collapsedControl"] {display: none;}
    /* Hide the sidebar itself completely */
    [data-testid="stSidebar"] {display: none;}
    
    #MainMenu {visibility: hidden;} 
    footer {visibility: hidden;}    
    
    /* Main Background & Font Styling */
    .stApp { background-color: #F8FAFC; }
    h1, h2, h3 { color: #1E293B; font-family: 'Inter', sans-serif; }
    .stButton>button { border-radius: 8px; font-weight: 600; transition: all 0.3s ease; }
    </style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

# MAIN LANDING PAGE
st.title("⚡ Welcome to CleanReport Workspace")
st.markdown("Your centralized professional portal for Data Analytics and Government Calculations.")
st.markdown("---")

col1, col2 = st.columns(2, gap="large")
with col1:
    st.markdown("""
        <div style="background-color: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02); height: 160px;">
            <h3 style="color: #2563EB; margin-top: 0;">📊 Ultimate Data Matcher</h3>
            <p style="color: #64748B;">Extract, clean, and cross-reference complex corporate Excel inventories with high precision fuzzy matching algorithms.</p>
        </div>
        <br>
    """, unsafe_allow_html=True)
    if st.button("🚀 Launch Data Matcher", use_container_width=True, type="primary"):
        st.switch_page("pages/1_Data_Matcher.py")

with col2:
    st.markdown("""
        <div style="background-color: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.02); height: 160px;">
            <h3 style="color: #2563EB; margin-top: 0;">🏥 Medical Reimbursement</h3>
            <p style="color: #64748B;">Automate CGHS medical claims calculation consultation-wise and instantly generate audit-ready official PDF reports.</p>
        </div>
        <br>
    """, unsafe_allow_html=True)
    if st.button("🚀 Launch Medical Calculator", use_container_width=True, type="primary"):
        st.switch_page("pages/2_Medical_Reimbursement.py")
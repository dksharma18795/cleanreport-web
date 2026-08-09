import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="CleanReport Workspace",
    page_icon="⚡",
    layout="wide"
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

# CSS TO HIDE DEFAULT MENU
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;} 
            footer {visibility: hidden;}    
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

# MAIN LANDING PAGE
st.title("⚡ Welcome to CleanReport Workspace")
st.markdown("Your centralized professional portal for Data Analytics and Government Calculations.")
st.markdown("---")

st.info("👈 Please select a tool from the left sidebar menu to get started.")

st.markdown("<br>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.subheader("📊 1. Ultimate Data Matcher")
    st.write("Extract, compare, and standardize company data from master Sales and Stock Excel files using advanced fuzzy logic.")
    
with col2:
    st.subheader("🏥 2. Medical Reimbursement")
    st.write("Calculate admissible CGHS medical claims consultation-wise and instantly generate audit-ready PDF reports.")
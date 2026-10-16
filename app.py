import streamlit as st
import pandas as pd
import gspread
import time
import unicodedata
from google.oauth2.service_account import Credentials
from datetime import datetime
import io
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- ΡΥΘΜΙΣΕΙΣ ΒΑΣΗΣ & SECRETS ---
OLD_SHEET_NAME = "Database_Center" 
NEW_SHEET_ID = "1w3PUqUrXbzxYnB0O4JSuz0TX7WQc7lHXAti6GN27dqY" 

# Αρχικοποίηση session states
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "role" not in st.session_state:
    st.session_state.role = ""
if "page" not in st.session_state:
    st.session_state.page = "menu"
if "last_activity" not in st.session_state:
    st.session_state.last_activity = time.time()
if "is_submitting" not in st.session_state:
    st.session_state.is_submitting = False

TIMEOUT_SECONDS = 600  # 10 λεπτά

if st.session_state.logged_in:
    if time.time() - st.session_state.last_activity > TIMEOUT_SECONDS:
        st.session_state.logged_in = False
        st.session_state.role = ""
        st.session_state.page = "menu"
        st.warning("⏳ Η συνεδρία έληξε λόγω αδράνειας. Παρακαλώ συνδεθείτε ξανά.")
    else:
        st.session_state.last_activity = time.time()

@st.cache_resource
def get_gspread_client():
    creds_dict = {
        "type": st.secrets["connections"]["gsheets"]["type"],
        "project_id": st.secrets["connections"]["gsheets"]["project_id"],
        "private_key_id": st.secrets["connections"]["gsheets"]["private_key_id"],
        "private_key": st.secrets["connections"]["gsheets"]["private_key"],
        "client_email": st.secrets["connections"]["gsheets"]["client_email"],
        "client_id": st.secrets["connections"]["gsheets"]["client_id"],
        "auth_uri": st.secrets["connections"]["gsheets"]["auth_uri"],
        "token_uri": st.secrets["connections"]["gsheets"]["token_uri"],
        "auth_provider_x509_cert_url": st.secrets["connections"]["gsheets"]["auth_provider_x509_cert_url"],
        "client_x509_cert_url": st.secrets["connections"]["gsheets"]["client_x509_cert_url"]
    }
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

@st.cache_data(ttl=36000)
def load_specific_db(sheet_name):
    try:
        gc = get_gspread_client()
        ws = gc.open(OLD_SHEET_NAME).worksheet(sheet_name)
        rows = ws.get_all_values()
        if not rows: return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Σφάλμα φόρτωσης {sheet_name}: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=36000)
def load_master_db(sheet_name):
    try:
        gc = get_gspread_client()
        ws = gc.open_by_key(NEW_SHEET_ID).worksheet(sheet_name)
        rows = ws.get_all_values()
        if not rows: return pd.DataFrame()
        df = pd.DataFrame(rows[1:], columns=rows[0])
        df.columns = df.columns.astype(str).str.strip()
        return df
    except Exception as e:
        st.error(f"Σφάλμα φόρτωσης Master {sheet_name}: {e}")
        return pd.DataFrame()

st.set_page_config(page_title="ROBOLAND DB", page_icon="🤖", layout="wide")

def uppercase_clean(val):
    s = str(val).strip().upper()
    tones = {
        'Á': 'Α', 'È': 'Ε', 'Ή': 'Η', 'Ί': 'Ι',
        'Ό': 'Ο', 'Ύ': 'Υ', 'Ώ': 'Ω', 'Ϊ': 'Ι', 'Ϋ': 'Υ',
        'Ά': 'Α', 'Έ': 'Ε', 'Ή': 'Η', 'Ί': 'Ι', 'Ό': 'Ο', 'Ύ': 'Υ', 'Ώ': 'Ω'
    }
    for t, letter in tones.items():
        s = s.replace(t, letter)
    return s

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def check_and_update_status_to_finish(c_id, school_year, program):
    try:
        gc = get_gspread_client()
        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
        
        raw_data = sh.get_all_records()
        df = pd.DataFrame(raw_data)
        df.columns = df.columns.astype(str).str.strip()
        
        mask = (df['id_customer'].astype(str) == str(c_id)) & \
               (df['school_year'].astype(str) == str(school_year)) & \
               (df['program'].astype(str) == str(program))
        
        target_indices = df[mask].index.tolist()
        if not target_indices: return

        for idx in target_indices:
            row_idx = idx + 2 
            current_status = str(df.loc[idx, 'status']).strip().upper()
            adj_val = df.loc[idx, 'adjusted_cost']
            total_cost = float(str(adj_val).replace(',', '.')) if pd.notna(adj_val) and str(adj_val).strip() != '' else 0.0
            
            pays_sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
            pays_df = pd.DataFrame(pays_sh.get_all_records())
            pays_df.columns = pays_df.columns.astype(str).str.strip()
            
            mask_pay = (pays_df['id_customer'].astype(str) == str(c_id)) & \
                       (pays_df['school_year'].astype(str) == str(school_year)) & \
                       (pays_df['program'].astype(str) == str(program))
            paid = pd.to_numeric(pays_df.loc[mask_pay, 'amount_paid'], errors='coerce').sum()
            balance = total_cost - paid

            if current_status != "FINISH":
                if balance <= 0 and total_cost > 0:
                    sh.update_cell(row_idx, 8, "FINISH")
    except Exception as e:
        st.error(f"Σφάλμα κατά την ενημέρωση status: {e}")

def log_to_finance_master(trans_type, category, partner, invoice, amount, method, year, source):
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
        records = sh.get_all_records()
        next_id = int(pd.DataFrame(records)['id'].max()) + 1 if records else 1
        
        final_type = trans_type.strip().upper().replace("ΕΣΟΔΑ", "ΕΣΟΔΟ").replace("ΕΞΟΔΑ", "ΕΞΟΔΑ")
        sh.append_row([
            next_id,
            datetime.today().strftime('%Y-%m-%d'),
            final_type, 
            category,
            partner,
            str(invoice),
            float(amount),
            method,
            year,
            source
        ])
    except Exception as e:
        st.error(f"Σφάλμα στο Master Log: {e}")

# ==================== LOGIN ====================
if not st.session_state.logged_in:
    st.markdown("<div style='font-size:20px; font-weight:bold; text-align: center; margin-bottom: 20px;'>Είσοδος στο ROBOLAND DB</div>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        username_input = st.text_input("Username").strip()
        password_input = st.text_input("Password", type="password").strip()
        submit_login = st.form_submit_button("Σύνδεση", disabled=st.session_state.is_submitting)

    if submit_login:
        try:
            gc = get_gspread_client()
            ws_users = gc.open(OLD_SHEET_NAME).worksheet("db_users")
            users_records = ws_users.get_all_records()
            
            if users_records:
                users_df = pd.DataFrame(users_records)
                users_df.columns = users_df.columns.astype(str).str.strip()
                users_df['username'] = users_df['username'].astype(str).str.strip()
                users_df['password'] = users_df['password'].astype(str).str.strip()
                users_df['role'] = users_df['role'].astype(str).str.strip().str.lower()
                
                user_row = users_df[(users_df['username'] == username_input) & (users_df['password'] == password_input)]
                
                if not user_row.empty:
                    st.session_state.logged_in = True
                    st.session_state.username = username_input
                    st.session_state.role = user_row.iloc[0]['role']
                    st.session_state.page = "menu"
                    st.rerun()  
                else:
                    st.error("❌ Λάθος όνομα χρήστη ή κωδικός πρόσβασης.")
            else:
                st.error("❌ Το φύλλο db_users είναι κενό.")
        except Exception as e:
            st.error(f"Σφάλμα σύνδεσης με τη βάση χρηστών: {e}")       

# ==================== MAIN APPLICATION ====================
else:
    st.sidebar.markdown(f"👤 **{st.session_state.username}** ({st.session_state.role.upper()})")
    if st.sidebar.button("🚪 Έξοδος", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.page = "menu"
        st.rerun()

    if st.button("🔄 Ενημέρωση Δεδομένων", use_container_width=True, disabled=st.session_state.is_submitting):
        st.session_state.is_submitting = True
        load_specific_db.clear()
        load_master_db.clear()
        st.success("Τα δεδομένα ενημερώθηκαν επιτυχώς από τη Google!")
        time.sleep(1)
        st.session_state.is_submitting = False
        st.rerun()
    
    # 1. ΚΕΝΤΡΙΚΟ ΜΕΝΟΥ
    if st.session_state.page == "menu":
        st.markdown("<h1 style='text-align: center;'>🤖 ROBOLAND DB</h1>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center;'>Κεντρικό Μενού Διαχείρισης</h2>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.session_state.role == "master":
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                if st.button("📈 Master Οικονομικά", use_container_width=True):
                    st.session_state.page = "finance_master"; st.rerun()
            with col_m2:
                if st.button("🛠️ Χειροκίνητες Κινήσεις", use_container_width=True):
                    st.session_state.page = "edit_master_manual"; st.rerun()
        
        st.subheader("Διαχείριση Μαθητών (Παιδιών)")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            if st.button("➕ Νέο Παιδί (Πρώτη Εγγραφή)", use_container_width=True):
                st.session_state.page = "register"; st.rerun()
            if st.button("📋 Επεξεργασία Στοιχείων Μαθητών", use_container_width=True):
                st.session_state.page = "customer_registry"; st.rerun()
        with col_c2:
            if st.button("👥 Διαχείριση Κατάστασης (STOP/Επαναφορά)", use_container_width=True):
                st.session_state.page = "view_customers"; st.rerun()
            if st.button("🔄 Εισαγωγή Μαθητή σε Τμήμα", use_container_width=True):
                st.session_state.page = "renew_registration"; st.rerun()

        st.markdown("---")
        st.subheader("Πληρωμές")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("💳 Καταχώρηση Πληρωμής (Συνδρομή)", use_container_width=True):
                st.session_state.page = "add_payment"; st.rerun()
        with col_p2:
            if st.session_state.role != "super":
                if st.button("💳 Καταχώρηση Πληρωμής (Επιπρόσθετες)", use_container_width=True):
                    st.session_state.page = "extra_payments_view"; st.rerun()
                
        st.markdown("---")
        st.subheader("Καθηγητές")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            if st.button("👨‍🏫 Εγγραφή/Επεξεργασία Καθηγητή", use_container_width=True):
                st.session_state.page = "register_teacher"; st.rerun()        
        with col_t2:
            if st.button("👩‍🏫 Στοιχεία Καθηγητών", use_container_width=True):
                st.session_state.page = "teachers_secretariat"; st.rerun()

        st.markdown("---")
        st.subheader("Διαχείριση Κέντρου")
        if st.button("📚 Προγράμματα - Τμήματα - Καθηγητές", use_container_width=True):
            st.session_state.page = "add_structure"; st.rerun()

        st.markdown("---")
        st.subheader("Αναφορές")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("📋 Μητρώο Μαθητών", use_container_width=True):
                st

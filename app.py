import streamlit as st
import pandas as pd
import gspread
import time
import unicodedata
from oauth2client.service_account import ServiceAccountCredentials
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
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
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

query_params = st.query_params
if query_params.get("ref77077") != st.secrets["URL_ACCESS_KEY"]:
    st.error("🚫 Δεν έχετε εξουσιοδότηση πρόσβασης. Παρακαλώ επικοινωνήστε με τον διαχειριστή.")
    st.stop()

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
        
        final_type = trans_type.strip().upper().replace("ΕΣΟΔΑ", "ΕΣΟΔΟ").replace("ΕΞΟΔΑ", "ΕΞΟΔΟ")
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
            df_main = load_specific_db("db_users")
            sh_second = gc.open_by_key(NEW_SHEET_ID).worksheet("db_users")
            records_second = sh_second.get_all_records()
            df_second = pd.DataFrame(records_second) if records_second else pd.DataFrame()
            
            if not df_main.empty and not df_second.empty:
                df_main.columns = df_main.columns.astype(str).str.strip()
                df_second.columns = df_second.columns.astype(str).str.strip()
                users_df = pd.concat([df_main, df_second], ignore_index=True).drop_duplicates()
            elif not df_main.empty:
                users_df = df_main
            else:
                users_df = df_second

            users_df = users_df.astype(str).apply(lambda x: x.str.strip())
            user_row = users_df[(users_df['username'] == username_input) & (users_df['password'] == password_input)]
            
            if not user_row.empty:
                st.session_state.logged_in = True
                st.session_state.username = username_input
                st.session_state.role = user_row.iloc[0]['role'].lower()
                st.session_state.page = "menu"
                st.rerun()  
            else:
                st.error("❌ Λάθος στοιχεία σύνδεσης.")
        except Exception as e:
            st.error(f"Σφάλμα σύνδεσης: {e}")       

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
                st.session_state.page = "anafora_pelaton"; st.rerun()
            if st.button("🏫 Αναφορά Καθηγητών & Τμημάτων", use_container_width=True):
                st.session_state.page = "teacher_report"; st.rerun()
        with col_r2:
            if st.button("🚫 Εκδήλωση Ενδιαφέροντος", use_container_width=True):
                st.session_state.page = "unassigned_customers"; st.rerun()
            if st.session_state.role in ["admin", "master"]:
                if st.button("📊 Οικονομική Αναφορά", use_container_width=True):
                    st.session_state.page = "advanced_reports"; st.rerun()

        st.markdown("---")
        st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
        if st.button("🛠️ Εισαγωγή/ Τροποποίηση Παραπόνων & Σχολίων", use_container_width=True):
            st.session_state.page = "complaints_page"; st.rerun()

    # --- ΣΕΛΙΔΑ: ΠΡΩΤΗ ΚΑΤΑΧΩΡΗΣΗ ΠΑΙΔΙΟΥ ---
    elif st.session_state.page == "register":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Νέα Εγγραφή Μαθητή (Παιδιού)")

        with st.form("reg_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                cf = st.text_input("Όνομα Παιδιού").strip()
                cl = st.text_input("Επώνυμο Παιδιού").strip()
                bd = st.date_input("Ημ. Γέννησης", value=datetime(2015,1,1), min_value=datetime(2000,1,1))
                tel = st.text_input("Τηλέφωνο Επικοινωνίας").strip()
                pf = st.text_input("Όνομα Πατρός")
                pl = st.text_input("Επώνυμο Πατρός")
            with c2:
                em = st.text_input("Email").strip()
                adr = st.text_input("Διεύθυνση").strip()

            sub = st.form_submit_button("💾 Αποθήκευση", disabled=st.session_state.is_submitting)

        if sub:
            st.session_state.is_submitting = True
            if not cf or not cl:
                st.error("⚠️ Το Όνομα και το Επώνυμο είναι υποχρεωτικά!")
                st.session_state.is_submitting = False
            else:
                try:
                    gc = get_gspread_client()
                    sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                    records = sh.get_all_records()
                    
                    df_records = pd.DataFrame(records)
                    next_id = int(df_records['id_customer'].max()) + 1 if not df_records.empty else 1

                    row_data = [
                        next_id, uppercase_clean(cf), uppercase_clean(cl),
                        uppercase_clean(pf), uppercase_clean(pl), tel,
                        str(em).strip().lower(), uppercase_clean(adr),
                        bd.strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'),
                        "Μαθητής", "", "", "", "", "", "", "", "", ""
                    ]
                    sh.append_row(row_data)
                    st.success("✅ Επιτυχής καταχώρηση μαθητή!")
                    time.sleep(1.5)
                    st.session_state.is_submitting = False
                    st.rerun()
                except Exception as e:
                    st.session_state.is_submitting = False
                    st.error(f"Σφάλμα κατά την αποθήκευση: {e}")     

    # --- ΣΕΛΙΔΑ: ΕΙΣΑΓΩΓΗ ΣΕ ΤΜΗΜΑ ---
    elif st.session_state.page == "renew_registration":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Εισαγωγή ή Τροποποίηση Μαθητή σε Τμήμα")
        cust_df = load_specific_db("db_customers")
        progs_df = load_specific_db("db_programs")
        classes_df = load_specific_db("db_class")
    
        if not cust_df.empty:
            cust_df['display'] = (cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")")
            sel_cust = st.selectbox("🔍 Επιλογή Μαθητή", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
        
            if sel_cust != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = str(row['id_customer']).strip()
            
                regs_df = load_specific_db("db_registrations")
                user_regs = regs_df[regs_df['id_customer'].astype(str).str.strip() == c_id]
                
                modes = ["Νέα Εγγραφή"]
                if not user_regs.empty:
                    modes.append("Τροποποίηση")
                
                mode = st.radio("Ενέργεια", modes)
                reg_to_edit = None
                
                if mode == "Τροποποίηση":
                    user_regs['display_reg'] = user_regs['school_year'].astype(str) + " - " + user_regs['program'].astype(str) + " (" + user_regs['class'].astype(str) + ")"
                    sel_reg = st.selectbox("Επιλογή εγγραφής", user_regs['display_reg'].tolist())
                    reg_to_edit = user_regs[user_regs['display_reg'] == sel_reg].iloc[0]
                
                s_year = st.selectbox("📅 Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2030)])
                prog_list = progs_df['program_name'].tolist() if not progs_df.empty else []
                s_prog = st.selectbox("📚 Πρόγραμμα", options=prog_list)
                
                default_cost = float(reg_to_edit.get('adjusted_cost', 0)) if reg_to_edit is not None else 0.0
                adj_cost = st.number_input("💰 Συμφωνηθέν Ποσό (€)", value=default_cost, step=10.0)
                
                class_options = ["Xωρίς Τμήμα"]
                if not classes_df.empty:
                    class_options += classes_df['class_name'].unique().tolist()
                sel_class = st.selectbox("🏫 Τμήμα", options=class_options)
                
                if st.button("🚀 Οριστικοποίηση Εγγραφής", disabled=st.session_state.is_submitting):
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                        
                        if mode == "Νέα Εγγραφή":
                            records = sh.get_all_records()
                            next_reg = int(pd.DataFrame(records)['id_registration'].max()) + 1 if records else 1
                            sh.append_row([next_reg, int(c_id), str(row['customer_firstname']), str(row['customer_lastname']), str(s_year), str(s_prog), str(sel_class), "ACTIVE", "", adj_cost, ""])
                        else:
                            target_reg_id = str(reg_to_edit['id_registration']).strip()
                            all_reg_values = sh.get_all_values()
                            for r_idx, sheet_row in enumerate(all_reg_values[1:], start=2):
                                if str(sheet_row[0]).strip() == target_reg_id:
                                    sh.update_cell(r_idx, 5, str(s_year))
                                    sh.update_cell(r_idx, 6, str(s_prog))
                                    sh.update_cell(r_idx, 7, str(sel_class))
                                    sh.update_cell(r_idx, 10, adj_cost)
                                    break
                        
                        st.success("✅ Η εγγραφή ενημερώθηκε επιτυχώς!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΠΛΗΡΩΜΕΣ ---
    elif st.session_state.page == "add_payment":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Διαχείριση Πληρωμών Μαθητών")
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        progs_df = load_specific_db("db_programs")
        pays_df = load_specific_db("db_payments")
        
        if not cust_df.empty:
            cust_df['display'] = (cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")")
            sel_cust = st.selectbox("🔍 Επιλογή Μαθητή", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
            
            if sel_cust != "-- Επιλέξτε --":
                c_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = int(c_row['id_customer'])
                
                user_regs = reg_df[reg_df['id_customer'].astype(int) == c_id]
                if user_regs.empty:
                    st.error("⚠️ Ο μαθητής δεν έχει ενεργή εγγραφή σε πρόγραμμα.")
                else:
                    sel_y = st.selectbox("📅 Έτος:", options=user_regs['school_year'].unique().tolist())
                    sel_prog = st.selectbox("📚 Πρόγραμμα:", options=user_regs['program'].unique().tolist())
                    
                    rec = st.text_input("Αριθμός Απόδειξης").strip()
                    amt = st.number_input("Ποσό (€)", min_value=0.0, format="%.2f")
                    pay_method = st.selectbox("Τρόπος", ["Μετρητά", "Εθνική", "Eurobank", "POS"])
                    
                    if st.button("💳 Αποθήκευση Πληρωμής"):
                        if not rec or amt <= 0:
                            st.warning("⚠️ Συμπληρώστε απόδειξη και ποσό!")
                        else:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                records = sh.get_all_records()
                                next_id = int(pd.DataFrame(records)['id_payment'].max()) + 1 if records else 1
                                
                                sh.append_row([
                                    next_id, c_id, str(c_row['customer_lastname']), str(c_row['customer_firstname']),
                                    datetime.today().strftime('%Y-%m-%d'), rec, amt, str(sel_y), pay_method, str(sel_prog)
                                ])
                                check_and_update_status_to_finish(c_id, sel_y, sel_prog)
                                log_to_finance_master("ΕΣΟΔΟ", sel_prog, f"{c_row['customer_lastname']} {c_row['customer_firstname']}", rec, amt, pay_method, sel_y, "Payments")
                                
                                st.success("✅ Η πληρωμή καταχωρήθηκε!")
                                time.sleep(1.2)
                            except Exception as e:
                                st.error(f"Σφάλμα: {e}")

    # --- ΑΝΑΦΟΡΕΣ & ΛΟΙΠΕΣ ΣΕΛΙΔΕΣ ---
    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
        st.subheader("Μητρώο Μαθητών")
        reg_df = load_specific_db("db_registrations")
        if not reg_df.empty:
            st.dataframe(reg_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν εγγραφές.")

    elif st.session_state.page == "finance_master":
        if st.session_state.role not in ["admin", "master"]:
            st.session_state.page = "menu"; st.rerun() 
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
        st.subheader("Master Οικονομικά")
        master_df = load_master_db("db_finance_master")
        if not master_df.empty:
            st.dataframe(master_df, use_container_width=True, hide_index=True)
        else:
            st.info("Το Master είναι κενό.")

    else:
        st.session_state.page = "menu"
        st.rerun()

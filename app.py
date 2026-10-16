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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- ΡΥΘΜΙΣΕΙΣ ΒΑΣΗΣ ---
OLD_SHEET_NAME = "Database_Center" # Το μοναδικό σου Google Sheet

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

# --- ΣΥΝΔΕΣΗ ΜΕ GOOGLE SHEETS ---
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

st.set_page_config(page_title="ROBOLAND DB", page_icon="🤖", layout="wide")

def process_teachers_df(df):
    df = df.copy()
    if 'teacher_lastname' in df.columns and 'teacher_firstname' in df.columns:
        df['full_name'] = df['teacher_lastname'].astype(str) + " " + df['teacher_firstname'].astype(str)
    else:
        df['full_name'] = "Άγνωστος"
    return df

CUSTOMER_COLUMNS = ['id_customer', 'customer_firstname', 'customer_lastname', 'parent_firstname', 'parent_lastname', 'telephone', 'email', 'address', 'birthday_date', 'register_date']
PROGRAM_COLUMNS = ['id_program', 'program_name', 'program_cost']
REGISTRATION_COLUMNS = ['id_registration', 'id_customer', 'customer_firstname', 'customer_lastname', 'school_year', 'program', 'class', 'status', 'stop_date', 'adjusted_cost']
PAYMENT_COLUMNS = ['id_payment', 'id_customer', 'customer_lastname', 'customer_firstname', 'payment_date', 'receipt_number', 'amount_paid', 'school_year']
CLASS_COLUMNS = ['id_class', 'class_name', 'program_name']
TEACHER_COLUMNS = ['teacher_id', 'teacher_firstname', 'teacher_lastname', 'teacher_email', 'teacher_telephone', 'teacher_bank', 'teacher_iban']

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

def clean_text(val):
    s = str(val).strip().lower()
    tones = {'ά': 'α', 'έ': 'ε', 'ή': 'η', 'ί': 'ι', 'ό': 'ο', 'ύ': 'υ', 'ώ': 'ω', 'ΐ': 'ι', 'ΰ': 'υ'}
    for t, letter in tones.items():
        s = s.replace(t, letter)
    if s.endswith('ς') or s.endswith('σ'):
        s = s[:-1]
    return s

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

# --- PDF ΣΥΝΑΡΤΗΣΕΙΣ ---
def generate_report_pdf_customers(report_df, selected_year="ΟΛΑ", selected_student="ΟΛΟΙ"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, alignment=1, textColor=colors.HexColor('#1F2937'), fontName='Helvetica-Bold')
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], fontSize=11, alignment=1, spaceAfter=20, textColor=colors.HexColor('#4B5563'), fontName='Helvetica-Bold')
    label_style = ParagraphStyle('LabelStyle', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'))
    value_style = ParagraphStyle('ValueStyle', parent=styles['Normal'], fontSize=9.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'))
    name_style = ParagraphStyle('NameStyle', parent=styles['Normal'], fontSize=11, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'))

    story.append(Paragraph("ΜΗΤΡΩΟ ΠΕΛΑΤΩΝ", title_style))
    filter_text = f"ΦΙΛΤΡΑ: ΕΤΟΣ: {remove_accents(str(selected_year).upper())} | ΜΑΘΗΤΗΣ: {remove_accents(str(selected_student).upper())}"
    story.append(Paragraph(filter_text, subtitle_style))
    story.append(Spacer(1, 10))

    for index, row in report_df.iterrows():
        lname = remove_accents(str(row.get('customer_lastname', '')).upper())
        fname = remove_accents(str(row.get('customer_firstname', '')).upper())
        year = remove_accents(str(row.get('school_year', '')).upper())
        prog = remove_accents(str(row.get('program', '')).upper())
        u_class = remove_accents(str(row.get('class', '')).upper())
        status = remove_accents(str(row.get('status', '')).upper())

        row_data = [
            [Paragraph("ΜΑΘΗΤΗΣ:", label_style), Paragraph("ΣΧΟΛΙΚΟ ΕΤΟΣ:", label_style), Paragraph("ΚΑΤΑΣΤΑΣΗ:", label_style)],
            [Paragraph(f"{lname} {fname}", name_style), Paragraph(year, value_style), Paragraph(status, value_style)],
            [Paragraph("ΠΡΟΓΡΑΜΜΑ:", label_style), Paragraph("ΤΜΗΜΑ:", label_style), Paragraph("", label_style)],
            [Paragraph(prog, value_style), Paragraph(u_class, value_style), Paragraph("", value_style)]
        ]
        
        t = Table(row_data, colWidths=[270, 135, 135])
        t.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#F3F4F6')),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        story.append(t)
        story.append(Spacer(1, 10))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_teacher_pdf(report_df, selected_year, selected_teacher):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=15, alignment=1, fontName='Helvetica-Bold')
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=11, alignment=1, spaceAfter=20, fontName='Helvetica-Bold')
    label_style = ParagraphStyle('Label', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold')
    val_style = ParagraphStyle('Val', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold')

    story.append(Paragraph("ΑΝΑΦΟΡΑ ΚΑΘΗΓΗΤΩΝ", title_style))
    story.append(Paragraph(f"ΕΤΟΣ: {remove_accents(str(selected_year).upper())} | ΚΑΘΗΓΗΤΗΣ: {remove_accents(str(selected_teacher).upper())}", sub_style))

    for _, row in report_df.iterrows():
        data = [
            [Paragraph("ΤΜΗΜΑ:", label_style), Paragraph(remove_accents(str(row.get('class_name', '')).upper()), val_style)],
            [Paragraph("ΠΡΟΓΡΑΜΜΑ:", label_style), Paragraph(remove_accents(str(row.get('program_name', '')).upper()), val_style)],
            [Paragraph("ΚΑΘΗΓΗΤΗΣ:", label_style), Paragraph(remove_accents(str(row.get('teachers', '')).upper()), val_style)],
            [Paragraph("ΜΑΘΗΤΕΣ:", label_style), Paragraph(str(row.get('Μαθητές', '0')), val_style)]
        ]
        t = Table(data, colWidths=[150, 350])
        t.setStyle(TableStyle([('BOX', (0,0), (-1,-1), 0.5, colors.black), ('GRID', (0,0), (-1,-1), 0.2, colors.grey)]))
        story.append(t)
        story.append(Spacer(1, 10))
        
    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_class_pdf(students_df, class_name, teacher_name, program_name, is_paid):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=14, fontName='Helvetica-Bold')
    info_style = ParagraphStyle('Info', parent=styles['Normal'], fontSize=10, spaceAfter=10, fontName='Helvetica-Bold')

    story.append(Paragraph(f"ΤΜΗΜΑ: {remove_accents(class_name.upper())}", title_style))
    story.append(Paragraph(f"ΚΑΘΗΓΗΤΗΣ: {remove_accents(teacher_name.upper())} | ΠΡΟΓΡΑΜΜΑ: {remove_accents(program_name.upper())}", info_style))
    story.append(Paragraph(f"ΑΠΟΠΛΗΡΩΜΗ: {'ΝΑΙ' if is_paid else 'ΟΧΙ'}", info_style))
    story.append(Spacer(1, 10))

    table_data = [["ΕΠΩΝΥΜΟ", "ΟΝΟΜΑ", "ΤΗΛΕΦΩΝΟ"]]
    for _, row in students_df.iterrows():
        table_data.append([
            remove_accents(str(row.get('Επώνυμο', '')).upper()),
            remove_accents(str(row.get('Όνομα', '')).upper()),
            str(row.get('Τηλέφωνο', ''))
        ])
    
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 0.5, colors.black), ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold')]))
    story.append(t)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

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
        st.success("Τα δεδομένα ενημερώθηκαν επιτυχώς από τη Google!")
        time.sleep(1)
        st.session_state.is_submitting = False
        st.rerun()
    
    # 1. ΚΕΝΤΡΙΚΟ ΜΕΝΟΥ
    if st.session_state.page == "menu":
        st.markdown("<h1 style='text-align: center;'>🤖 ROBOLAND DB</h1>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center;'>Κεντρικό Μενού Διαχείρισης</h2>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.subheader("Διαχείριση Πελατών / Μαθητών")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            if st.button("➕ Νέος Πελάτης (Πρώτη Εγγραφή)", use_container_width=True):
                st.session_state.page = "register"; st.rerun()
            if st.button("📋 Επεξεργασία Στοιχείων Πελατών", use_container_width=True):
                st.session_state.page = "customer_registry"; st.rerun()
        with col_c2:
            if st.button("👥 Διαχείριση Κατάστασης Πελατών (STOP/Επαναφορά)", use_container_width=True):
                st.session_state.page = "view_customers"; st.rerun()
            if st.button("🔄 Εισαγωγή Πελάτη σε Τμήμα", use_container_width=True):
                st.session_state.page = "renew_registration"; st.rerun()

        st.markdown("---")
        st.subheader("Πληρωμές")
        if st.button("💳 Καταχώρηση Πληρωμής (Συνδρομή)", use_container_width=True):
            st.session_state.page = "add_payment"; st.rerun()
                
        st.markdown("---")
        st.subheader("Καθηγητές")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("👨‍🏫 Εγγραφή/Επεξεργασία Καθηγητή", use_container_width=True):
                st.session_state.page = "register_teacher"; st.rerun()        
        with col_p2:
            if st.button("👩‍🏫 Στοιχεία Καθηγητών", use_container_width=True):
                st.session_state.page = "teachers_secretariat"; st.rerun()

        st.markdown("---")
        st.subheader("Διαχείριση Κέντρου")
        if st.button("📚 Προγράμματα - Τμήματα - Καθηγητές", use_container_width=True):
            st.session_state.page = "add_structure"; st.rerun()

        st.markdown("---")
        st.subheader("Εταιρικές Συναλλαγές")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            if st.button("📝 Καταχώρηση Στοιχείων Εταιρείας", use_container_width=True):
                st.session_state.page = "corp_register"; st.rerun()
        with col_e2:
            if st.button("💸 Διαχείριση Συναλλαγών", use_container_width=True):
                st.session_state.page = "corp_payments"; st.rerun()
                
        st.markdown("---")
        st.subheader("Αναφορές")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.session_state.role in ["employee", "admin", "master"]:
                if st.button("📋 Μητρώο Πελατών", use_container_width=True):
                    st.session_state.page = "anafora_pelaton"; st.rerun()
                if st.button("🏫 Αναφορά Καθηγητών & Τμημάτων", use_container_width=True):
                    st.session_state.page = "teacher_report"; st.rerun()
        with col_r2:
            if st.button("🚫 Εκδήλωση Ενδιαφέροντος", use_container_width=True):
                st.session_state.page = "unassigned_customers"; st.rerun()

        st.markdown("---")
        st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
        if st.button("🛠️ Εισαγωγή/ Τροποποίηση Παραπόνων & Σχολίων", use_container_width=True):
            st.session_state.page = "complaints_page"; st.rerun()

    # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ ΕΤΑΙΡΕΙΑΣ ---
    elif st.session_state.page == "corp_register":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Καταχώρηση & Τροποποίηση Εταιρικών Στοιχείων")
        comp_df = load_specific_db("db_customers_company")
        
        mode_corp = st.radio("Ενέργεια", ["Νέα Καταχώρηση", "Τροποποίηση Υπάρχουσας"], horizontal=True)
        selected_corp = None
        
        if mode_corp == "Τροποποίηση Υπάρχουσας" and not comp_df.empty:
            try:
                comp_df['display'] = comp_df['id'].astype(str) + " - " + comp_df['company_name'].astype(str) + " (ΑΦΜ: " + comp_df['vat_number'].astype(str) + ")"
                sel_c_box = st.selectbox("Επιλογή Εταιρείας για Τροποποίηση", options=["-- Επιλέξτε --"] + comp_df['display'].tolist())
                if sel_c_box != "-- Επιλέξτε --":
                    selected_corp = comp_df[comp_df['display'] == sel_c_box].iloc[0]
            except Exception:
                st.warning("Υπήρξε πρόβλημα στην ανάγνωση των στηλών της εταιρείας.")

        def_name = str(selected_corp['company_name']) if selected_corp is not None and 'company_name' in selected_corp else ""
        def_vat = str(selected_corp['vat_number']) if selected_corp is not None and 'vat_number' in selected_corp else ""
        def_tax = str(selected_corp['tax_office']) if selected_corp is not None and 'tax_office' in selected_corp else ""
        def_addr = str(selected_corp['address']) if selected_corp is not None and 'address' in selected_corp else ""
        def_email = str(selected_corp['email']) if selected_corp is not None and 'email' in selected_corp else ""
        def_phone = str(selected_corp['phone']) if selected_corp is not None and 'phone' in selected_corp else ""
        def_year = str(selected_corp['school_year']) if selected_corp is not None and 'school_year' in selected_corp else "2026-2027"

        with st.form("corp_reg_form"):
            name = st.text_input("Επωνυμία", value=def_name)
            vat = st.text_input("ΑΦΜ", value=def_vat)
            tax_office = st.text_input("ΔΟΥ", value=def_tax)
            addr = st.text_input("Διεύθυνση", value=def_addr)
            email = st.text_input("Email", value=def_email)
            phone = st.text_input("Τηλέφωνο", value=def_phone)
            years_list = [f"{y}-{y+1}" for y in range(2026, 2041)]
            default_year_idx = years_list.index(def_year) if def_year in years_list else 0
            year = st.selectbox("Σχολικό Έτος", options=years_list, index=default_year_idx)
            
            btn_label = "💾 Αποθήκευση Τροποποίησης" if mode_corp == "Τροποποίηση Υπάρχουσας" and selected_corp is not None else "Αποθήκευση"
            
            if st.form_submit_button(btn_label):
                if not name or not vat:
                    st.warning("⚠️ Συμπληρώστε τουλάχιστον την επωνυμία και το ΑΦΜ!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers_company")
                        if mode_corp == "Τροποποίηση Υπάρχουσας" and selected_corp is not None:
                            corp_id = str(selected_corp['id']).strip()
                            all_rows = sh.get_all_values()
                            target_row = None
                            for idx, row_vals in enumerate(all_rows[1:], start=2):
                                if str(row_vals[0]).strip() == corp_id:
                                    target_row = idx
                                    break
                            if target_row:
                                sh.update_cell(target_row, 2, name)
                                sh.update_cell(target_row, 3, vat)
                                sh.update_cell(target_row, 4, tax_office)
                                sh.update_cell(target_row, 5, addr)
                                sh.update_cell(target_row, 6, email)
                                sh.update_cell(target_row, 7, phone)
                                sh.update_cell(target_row, 8, year)
                                st.success("✅ Τα στοιχεία της εταιρείας τροποποιήθηκαν επιτυχώς!")
                        else:
                            records = sh.get_all_records()
                            next_id = int(pd.DataFrame(records)['id'].max()) + 1 if records and 'id' in pd.DataFrame(records).columns else 1
                            sh.append_row([next_id, name, vat, tax_office, addr, email, phone, year])
                            st.success("✅ Καταχωρήθηκε!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΠΛΗΡΩΜΕΣ ΕΤΑΙΡΕΙΑΣ ---
    elif st.session_state.page == "corp_payments":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Πληρωμές & Συναλλαγές Εταιρειών (Έσοδα/Έξοδα)")
        comp_df = load_specific_db("db_customers_company")
        trans_df = load_specific_db("db_payments_company")
        
        with st.form("corp_pay_form", clear_on_submit=True):
            trans_type = st.selectbox("Τύπος", ["Έσοδα", "Έξοδα"])
            inv_num = st.text_input("Αριθμός Τιμολογίου")
            comp = st.selectbox("Εταιρεία", comp_df['company_name'].tolist() if not comp_df.empty and 'company_name' in comp_df.columns else [])
            year = st.selectbox("Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2041)])
            inv_date = st.date_input("Ημερομηνία Τιμολογίου")
            pay_date = st.date_input("Ημερομηνία Πληρωμής")
            amount = st.number_input("Ποσό (€)", min_value=0.0)
            pay_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "POS", "Επι πιστώση", "Εθνική", "Eurobank"])
            
            if st.form_submit_button("Αποθήκευση Συναλλαγής"):
                if not inv_num:
                    st.warning("Ο Αριθμός Τιμολογίου είναι απαραίτητος!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                        c_id = ""
                        if not comp_df.empty and 'company_name' in comp_df.columns:
                            matched_c = comp_df[comp_df['company_name'] == comp]
                            if not matched_c.empty:
                                c_id = matched_c['id'].iloc[0]
                                
                        records = sh.get_all_records()
                        next_id = int(pd.DataFrame(records)['id'].max()) + 1 if records and 'id' in pd.DataFrame(records).columns else 1
                        
                        sh.append_row([
                            next_id, inv_num, trans_type, c_id, comp, year, 
                            inv_date.strftime('%Y-%m-%d'), pay_date.strftime('%Y-%m-%d'), 
                            datetime.today().strftime('%Y-%m-%d'), amount, pay_method
                        ])
                        st.success("✅ Η συναλλαγή αποθηκεύτηκε!")
                        time.sleep(1.2)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΔΙΑΧΕΙΡΙΣΗ ΚΑΤΑΣΤΑΣΗΣ (STOP) ---
    elif st.session_state.page == "view_customers":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun()

        st.subheader("🛑 Διαχείριση Διακοπών & Καταστάσεων")
        df_cust = load_specific_db("db_customers")
        df_reg = load_specific_db("db_registrations")

        if df_cust.empty or df_reg.empty:
            st.warning("Δεν βρέθηκαν δεδομένα πελατών ή εγγραφών.")
        else:
            df_cust['display_name'] = df_cust['customer_lastname'].astype(str) + " " + df_cust['customer_firstname'].astype(str) + " (ID: " + df_cust['id_customer'].astype(str) + ")"
            selected_student = st.selectbox("🔍 Επιλογή πελάτη", options=["-- Επιλέξτε --"] + df_cust['display_name'].tolist())

            if selected_student != "-- Επιλέξτε --" and "ID: " in selected_student:
                stud_id_str = selected_student.split("ID: ")[1].replace(")", "").strip()
                stud_regs = df_reg[df_reg['id_customer'].astype(str).str.strip() == stud_id_str]

                for idx, r in stud_regs.iterrows():
                    st.write(f"📅 Έτος: {r['school_year']} | 📚 Πρόγραμμα: {r['program']} | Κατάσταση: {r['status']}")
                    if st.button(f"⚙️ Αλλαγή Κατάστασης (ID: {r['id_registration']})", key=f"status_btn_{r['id_registration']}"):
                        st.session_state[f"show_form_{r['id_registration']}"] = True

                    if st.session_state.get(f"show_form_{r['id_registration']}", False):
                        with st.form(key=f"form_{r['id_registration']}"):
                            chosen_status = st.selectbox("Κατάσταση", ['ACTIVE', 'FINISH', 'STOP'], index=['ACTIVE', 'FINISH', 'STOP'].index(str(r['status']).strip().upper()) if str(r['status']).strip().upper() in ['ACTIVE', 'FINISH', 'STOP'] else 0)
                            stop_d = st.date_input("Ημερομηνία Διακοπής", value=datetime.today())
                            new_cost = st.number_input("Τελικό ποσό (€)", value=float(str(r.get('adjusted_cost', 0)).replace(',', '.') or 0))
                            
                            if st.form_submit_button("💾 Οριστικοποίηση"):
                                try:
                                    gc = get_gspread_client()
                                    sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                    all_rows = sh.get_all_values()
                                    for r_i, s_row in enumerate(all_rows[1:], start=2):
                                        if str(s_row[0]).strip() == str(r['id_registration']).strip():
                                            sh.update_cell(r_i, 8, chosen_status)
                                            sh.update_cell(r_i, 9, stop_d.strftime('%Y-%m-%d') if chosen_status == 'STOP' else '')
                                            sh.update_cell(r_i, 10, new_cost)
                                            break
                                    st.success("✅ Ενημερώθηκε!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΠΡΩΤΗ ΚΑΤΑΧΩΡΗΣΗ ΠΕΛΑΤΗ ---
    elif st.session_state.page == "register":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Νέα Εγγραφή Πελάτη")
        cat = st.radio("Κατηγορία Πελάτη", ["Μαθητής", "Ενήλικας"])

        with st.form("reg_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                cf = st.text_input("Όνομα").strip()
                cl = st.text_input("Επώνυμο").strip()
                bd = st.date_input("Ημ. Γέννησης", value=datetime(2015,1,1), min_value=datetime(1940,1,1))
                tel = st.text_input("Τηλέφωνο").strip()
                pf = st.text_input("Όνομα Πατρός")
                pl = st.text_input("Επώνυμο Πατρός")
                mf = st.text_input("Όνομα Μητρός")
            with c2:
                em = st.text_input("Email").strip()
                adr = st.text_input("Διεύθυνση").strip()
                afm = st.text_input("ΑΦΜ") if cat == "Ενήλικας" else ""
                town = st.text_input("Πόλη") if cat == "Ενήλικας" else ""
                zip_code = st.text_input("ΤΚ") if cat == "Ενήλικας" else ""
                tax_office = st.text_input("ΔΟΥ") if cat == "Ενήλικας" else ""
                amka = st.text_input("ΑΜΚΑ") if cat == "Ενήλικας" else ""
                iban = st.text_input("IBAN") if cat == "Ενήλικας" else ""
                bank = st.text_input("Τράπεζα") if cat == "Ενήλικας" else ""

            if st.form_submit_button("💾 Αποθήκευση"):
                if not cf or not cl:
                    st.error("⚠️ Το Όνομα και το Επώνυμο είναι υποχρεωτικά!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                        records = sh.get_all_records()
                        next_id = int(pd.DataFrame(records)['id_customer'].max()) + 1 if records and 'id_customer' in pd.DataFrame(records).columns else 1

                        row_data = [
                            next_id, uppercase_clean(cf), uppercase_clean(cl),
                            uppercase_clean(pf), uppercase_clean(pl), tel,
                            str(em).strip().lower(), uppercase_clean(adr),
                            bd.strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'),
                            cat, afm, uppercase_clean(mf), "", uppercase_clean(town),
                            zip_code, tax_office, amka, iban, uppercase_clean(bank)
                        ]
                        sh.append_row(row_data)
                        st.success("✅ Επιτυχής καταχώρηση!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΕΠΕΞΕΡΓΑΣΙΑ ΣΤΟΙΧΕΙΩΝ ΠΕΛΑΤΗ ---
    elif st.session_state.page == "customer_registry":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Επεξεργασία Στοιχείων Πελάτη")
        cust_df = load_specific_db("db_customers")
        
        if cust_df.empty:
            st.warning("⚠️ Η βάση πελατών είναι κενή.")
        else:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())

            if sel_cust != "-- Επιλέξτε --":
                selected_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                cust_id = str(selected_row['id_customer']).strip()

                with st.form("edit_customer_form_safe"):
                    new_fname = st.text_input("Όνομα", value=str(selected_row.get('customer_firstname', '')))
                    new_lname = st.text_input("Επώνυμο", value=str(selected_row.get('customer_lastname', '')))
                    new_tel = st.text_input("Τηλέφωνο", value=str(selected_row.get('telephone', '')))
                    new_email = st.text_input("Email", value=str(selected_row.get('email', '')))
                    new_adr = st.text_input("Διεύθυνση", value=str(selected_row.get('address', '')))

                    if st.form_submit_button("💾 Αποθήκευση αλλαγών"):
                        try:
                            gc = get_gspread_client()
                            sh_cust = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                            cell = sh_cust.find(cust_id)
                            if cell:
                                row_idx = cell.row
                                sh_cust.update_cell(row_idx, 2, uppercase_clean(new_fname))
                                sh_cust.update_cell(row_idx, 3, uppercase_clean(new_lname))
                                sh_cust.update_cell(row_idx, 6, new_tel.strip())
                                sh_cust.update_cell(row_idx, 7, str(new_email).strip().lower())
                                sh_cust.update_cell(row_idx, 8, uppercase_clean(new_adr))
                                st.success("✅ Ενημερώθηκε!")
                                st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΕΙΣΑΓΩΓΗ ΣΕ ΤΜΗΜΑ ---
    elif st.session_state.page == "renew_registration":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Εισαγωγή ή Τροποποίηση Πελάτη σε Τμήμα")
        cust_df = load_specific_db("db_customers")
        progs_df = load_specific_db("db_programs")
        classes_df = load_specific_db("db_class")
    
        if not cust_df.empty:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            sel_cust = st.selectbox("🔍 Επιλογή πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
        
            if sel_cust != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = str(row['id_customer']).strip()
            
                regs_df = load_specific_db("db_registrations")
                user_regs = regs_df[regs_df['id_customer'].astype(str).str.strip() == c_id]
                
                mode = st.radio("Ενέργεια", ["Νέα Εγγραφή", "Τροποποίηση"] if not user_regs.empty else ["Νέα Εγγραφή"])
                reg_to_edit = None
                
                if mode == "Τροποποίηση" and not user_regs.empty:
                    user_regs['display_reg'] = user_regs['school_year'].astype(str) + " - " + user_regs['program'].astype(str) + " (" + user_regs['class'].astype(str) + ")"
                    sel_reg = st.selectbox("Επιλογή εγγραφής", user_regs['display_reg'].tolist())
                    reg_to_edit = user_regs[user_regs['display_reg'] == sel_reg].iloc[0]
                
                s_year = st.selectbox("📅 Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2030)])
                prog_list = progs_df['program_name'].tolist() if not progs_df.empty else []
                s_prog = st.selectbox("📚 Πρόγραμμα", options=prog_list)
                
                default_cost = float(str(reg_to_edit.get('adjusted_cost', 0)).replace(',', '.') or 0) if reg_to_edit is not None else 0.0
                adj_cost = st.number_input("💰 Συμφωνηθέν Ποσό (€)", value=default_cost, step=10.0)
                
                class_options = ["Xωρίς Τμήμα"] + (classes_df['class_name'].unique().tolist() if not classes_df.empty else [])
                sel_class = st.selectbox("🏫 Τμήμα", options=class_options)
                
                if st.button("🚀 Οριστικοποίηση Εγγραφής"):
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                        if mode == "Νέα Εγγραφή":
                            records = sh.get_all_records()
                            next_reg = int(pd.DataFrame(records)['id_registration'].max()) + 1 if records and 'id_registration' in pd.DataFrame(records).columns else 1
                            sh.append_row([next_reg, int(c_id), str(row['customer_firstname']), str(row['customer_lastname']), str(s_year), str(s_prog), str(sel_class), "ACTIVE", "", adj_cost, ""])
                        else:
                            target_reg_id = str(reg_to_edit['id_registration']).strip()
                            all_rows = sh.get_all_values()
                            for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                if str(s_row[0]).strip() == target_reg_id:
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
            
        st.subheader("Διαχείριση Πληρωμών")
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        pays_df = load_specific_db("db_payments")
        
        if not cust_df.empty:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
            
            if sel_cust != "-- Επιλέξτε --":
                c_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = int(c_row['id_customer'])
                
                user_regs = reg_df[reg_df['id_customer'].astype(str).str.strip() == str(c_id)]
                if user_regs.empty:
                    st.error("⚠️ Ο πελάτης δεν έχει ενεργή εγγραφή σε πρόγραμμα.")
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
                                next_id = int(pd.DataFrame(records)['id_payment'].max()) + 1 if records and 'id_payment' in pd.DataFrame(records).columns else 1
                                
                                sh.append_row([
                                    next_id, c_id, str(c_row['customer_lastname']), str(c_row['customer_firstname']),
                                    datetime.today().strftime('%Y-%m-%d'), rec, amt, str(sel_y), pay_method, str(sel_prog)
                                ])
                                check_and_update_status_to_finish(c_id, sel_y, sel_prog)
                                st.success("✅ Η πληρωμή καταχωρήθηκε!")
                                time.sleep(1.2)
                                st.rerun()
                            except Exception as e:
                                st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΔΙΑΧΕΙΡΙΣΗ ΚΑΘΗΓΗΤΩΝ ---
    elif st.session_state.page == "register_teacher":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Διαχείριση Καθηγητών")
        df_teachers = load_specific_db("db_teachers")
        
        teach_opts = ["-- Νέα Εγγραφή --"]
        if not df_teachers.empty and 'teacher_lastname' in df_teachers.columns:
            df_teachers['full_name'] = df_teachers['teacher_lastname'].astype(str) + " " + df_teachers['teacher_firstname'].astype(str)
            teach_opts += df_teachers['full_name'].tolist()

        sel_teacher = st.selectbox("🔍 Επιλογή Καθηγητή", options=teach_opts)
        selected_t = df_teachers[df_teachers['full_name'] == sel_teacher].iloc[0] if sel_teacher != "-- Νέα Εγγραφή --" and not df_teachers.empty else None

        with st.form("teacher_form"):
            t_fname = st.text_input("Όνομα", value=selected_t.get('teacher_firstname', '') if selected_t is not None else "")
            t_lname = st.text_input("Επώνυμο", value=selected_t.get('teacher_lastname', '') if selected_t is not None else "")
            t_email = st.text_input("Email", value=selected_t.get('teacher_email', '') if selected_t is not None else "")
            t_tel = st.text_input("Τηλέφωνο", value=selected_t.get('teacher_telephone', '') if selected_t is not None else "")
            t_bank = st.selectbox("Τράπεζα", ["Εθνική", "Πειραιώς", "Alpha Bank", "Eurobank", "Attica Bank"])
            t_iban = st.text_input("IBAN", value=selected_t.get('teacher_iban', '') if selected_t is not None else "")

            if st.form_submit_button("💾 Αποθήκευση"):
                try:
                    gc = get_gspread_client()
                    sh = gc.open(OLD_SHEET_NAME).worksheet("db_teachers")
                    if selected_t is None:
                        next_id = int(pd.DataFrame(df_teachers)['teacher_id'].max()) + 1 if not df_teachers.empty and 'teacher_id' in df_teachers.columns else 1
                        sh.append_row([next_id, t_fname, t_lname, t_email, t_tel, t_bank, t_iban])
                        st.success("✅ Νέος καθηγητής καταχωρήθηκε!")
                    else:
                        row_idx = selected_t.name + 2
                        sh.update(f"B{row_idx}:G{row_idx}", [[t_fname, t_lname, t_email, t_tel, t_bank, t_iban]])
                        st.success("✅ Τα στοιχεία ενημερώθηκαν!")
                    time.sleep(1.2)
                    st.rerun()
                except Exception as e:
                    st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΣΤΟΙΧΕΙΑ ΚΑΘΗΓΗΤΩΝ ---
    elif st.session_state.page == "teachers_secretariat":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
        st.subheader("👩‍🏫 Στοιχεία Καθηγητών")
        teachers_df = load_specific_db("db_teachers")
        if not teachers_df.empty:
            st.dataframe(teachers_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν καθηγητές.")

    # --- ΣΕΛΙΔΑ: ΔΟΜΗ / ΠΡΟΓΡΑΜΜΑΤΑ / ΤΜΗΜΑΤΑ ---
    elif st.session_state.page == "add_structure":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("🛠️ Διαχείριση Προγραμμάτων & Τμημάτων")
        tab_prog, tab_class = st.tabs(["📚 Προγράμματα", "🏫 Τμήματα"])

        with tab_prog:
            prog_df = load_specific_db("db_programs")
            with st.form("form_prog_new", clear_on_submit=True):
                p_name = st.text_input("Όνομα Προγράμματος")
                p_cost = st.number_input("Κόστος (€)", min_value=0.0)
                if st.form_submit_button("💾 Αποθήκευση"):
                    if p_name:
                        try:
                            gc = get_gspread_client()
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_programs")
                            recs = sh.get_all_records()
                            next_id = int(pd.DataFrame(recs)['id_program'].max()) + 1 if recs and 'id_program' in pd.DataFrame(recs).columns else 1
                            sh.append_row([next_id, p_name, float(p_cost)])
                            st.success("✅ Αποθηκεύτηκε!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")
            if not prog_df.empty:
                st.dataframe(prog_df, use_container_width=True, hide_index=True)

        with tab_class:
            class_df = load_specific_db("db_class")
            prog_df_c = load_specific_db("db_programs")
            prog_list = prog_df_c['program_name'].tolist() if not prog_df_c.empty else []

            with st.form("form_class_new", clear_on_submit=True):
                c_name = st.text_input("Όνομα Τμήματος")
                c_prog = st.selectbox("Πρόγραμμα", prog_list)
                c_year = st.selectbox("Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2041)])
                c_start = st.date_input("Έναρξη", value=datetime.today())
                c_end = st.date_input("Λήξη", value=datetime.today())

                if st.form_submit_button("💾 Δημιουργία"):
                    if c_name:
                        try:
                            gc = get_gspread_client()
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                            recs = sh.get_all_records()
                            next_c_id = int(pd.DataFrame(recs)['id_class'].max()) + 1 if recs and 'id_class' in pd.DataFrame(recs).columns else 1
                            sh.append_row([next_c_id, c_name, c_prog, "", c_year, c_start.strftime('%Y-%m-%d'), c_end.strftime('%Y-%m-%d')])
                            st.success("✅ Δημιουργήθηκε!")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")
            if not class_df.empty:
                st.dataframe(class_df, use_container_width=True, hide_index=True)

    # --- ΣΕΛΙΔΑ: ΜΗΤΡΩΟ ΠΕΛΑΤΩΝ ---
    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun() 
        
        st.subheader("Μητρώο Πελατών")
        reg_df = load_specific_db("db_registrations")
        if not reg_df.empty:
            st.dataframe(reg_df, use_container_width=True, hide_index=True)
            if st.button("📥 Παραγωγή PDF"):
                rep_pdf = generate_report_pdf_customers(reg_df)
                if rep_pdf:
                    st.download_button("📥 Κατεβάστε το αρχείο", data=rep_pdf, file_name="mitroo.pdf", mime="application/pdf")
        else:
            st.info("Δεν βρέθηκαν εγγραφές.")

    # --- ΣΕΛΙΔΑ: ΑΝΑΦΟΡΑ ΚΑΘΗΓΗΤΩΝ & ΤΜΗΜΑΤΩΝ ---
    elif st.session_state.page == "teacher_report":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Αναφορά Καθηγητών & Τμημάτων")
        classes_df = load_specific_db("db_class")
        if not classes_df.empty:
            st.dataframe(classes_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν τμήματα.")

    # --- ΣΕΛΙΔΑ: ΕΚΔΗΛΩΣΗ ΕΝΔΙΑΦΕΡΟΝΤΟΣ (ΑΝΕΝΤΑΧΤΟΙ) ---
    elif st.session_state.page == "unassigned_customers":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun()
    
        st.subheader("🚫 Εκδήλωση Ενδιαφέροντος (Ανένταχτοι Πελάτες)")
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        
        if not cust_df.empty:
            reg_ids = set(reg_df['id_customer'].astype(str).str.strip().tolist()) if not reg_df.empty and 'id_customer' in reg_df.columns else set()
            unassigned = cust_df[~cust_df['id_customer'].astype(str).str.strip().isin(reg_ids)]
            if unassigned.empty:
                st.success("🎉 Όλοι οι πελάτες έχουν εγγραφή σε τμήμα!")
            else:
                st.dataframe(unassigned, use_container_width=True, hide_index=True)
        else:
            st.info("Η βάση πελατών είναι κενή.")

    # --- ΣΕΛΙΔΑ: ΠΑΡΑΠΟΝΑ & ΣΧΟΛΙΑ ---
    elif st.session_state.page == "complaints_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("🛠️ Διαχείριση & Παράπονα / Σχόλια Πελατών")
        complaints_df = load_specific_db("db_complaint")
        
        with st.form("new_complaint_form", clear_on_submit=True):
            c_subj = st.text_input("Θέμα").strip()
            c_text = st.text_area("Κείμενο").strip()
            if st.form_submit_button("💾 Αποθήκευση"):
                if c_subj and c_text:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_complaint")
                        recs = sh.get_all_records()
                        next_id = int(pd.DataFrame(recs)['id_complaint'].max()) + 1 if recs and 'id_complaint' in pd.DataFrame(recs).columns else 1
                        sh.append_row([next_id, 1, "Γενικός", "Πελάτης", datetime.today().strftime('%Y-%m-%d'), c_subj, c_text, "Admin", "ΕΝΕΡΓΟ", "Παράπονο"])
                        st.success("✅ Αποθηκεύτηκε!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

        if not complaints_df.empty:
            st.dataframe(complaints_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν υπάρχουν καταχωρημένα παράπονα.")

    else:
        st.session_state.page = "menu"
        st.rerun()

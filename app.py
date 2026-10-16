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
        return pd.DataFrame()

st.set_page_config(page_title="ROBOLAND DB", page_icon="🤖", layout="wide")

def uppercase_clean(val):
    s = str(val).strip().upper()
    tones = {'Á': 'Α', 'È': 'Ε', 'Ή': 'Η', 'Ί': 'Ι', 'Ό': 'Ο', 'Ύ': 'Υ', 'Ώ': 'Ω', 'Ϊ': 'Ι', 'Ϋ': 'Υ', 'Ά': 'Α', 'Έ': 'Ε', 'Ή': 'Η', 'Ί': 'Ι', 'Ό': 'Ο', 'Ύ': 'Υ', 'Ώ': 'Ω'}
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
        
        prog_col = next((c for c in df.columns if c.lower() in ['program', 'δ', 'j', 'πρόγραμμα']), 'program')
        adj_col = next((c for c in df.columns if 'adjusted' in c.lower() or 'cost' in c.lower()), 'adjusted_cost')

        mask = (df['id_customer'].astype(str) == str(c_id)) & (df['school_year'].astype(str) == str(school_year)) & (df[prog_col].astype(str) == str(program))
        target_indices = df[mask].index.tolist()
        if not target_indices: return

        for idx in target_indices:
            row_idx = idx + 2 
            current_status = str(df.loc[idx, 'status']).strip().upper()
            adj_val = df.loc[idx, adj_col] if adj_col in df.columns else 0
            total_cost = float(str(adj_val).replace(',', '.')) if pd.notna(adj_val) and str(adj_val).strip() != '' else 0.0
            
            pays_sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
            pays_df = pd.DataFrame(pays_sh.get_all_records())
            pays_df.columns = pays_df.columns.astype(str).str.strip()
            p_prog_col = next((c for c in pays_df.columns if 'prog' in c.lower()), 'program')
            
            mask_pay = (pays_df['id_customer'].astype(str) == str(c_id)) & (pays_df['school_year'].astype(str) == str(school_year)) & (pays_df[p_prog_col].astype(str) == str(program))
            paid = pd.to_numeric(pays_df.loc[mask_pay, 'amount_paid'], errors='coerce').sum()
            balance = total_cost - paid

            if current_status != "FINISH":
                if balance <= 0 and total_cost > 0:
                    sh.update_cell(row_idx, 8, "FINISH")
    except Exception as e:
        pass

# --- PDF ΣΥΝΑΡΤΗΣΕΙΣ ---
def generate_customer_payment_history_pdf(customer_name, selected_year, payments_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))
    total_style = ParagraphStyle('Total', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', alignment=2, textColor=colors.HexColor('#111827'))

    story.append(Paragraph("ΑΝΑΛΥΤΙΚΟ ΙΣΤΟΡΙΚΟ ΠΛΗΡΩΜΩΝ ΜΑΘΗΤΗ", title_style))
    filter_info = f"ΜΑΘΗΤΗΣ: {remove_accents(str(customer_name).upper())} | ΕΤΟΣ: {remove_accents(str(selected_year).upper())}"
    story.append(Paragraph(filter_info, sub_style))
    story.append(Spacer(1, 5))

    table_data = [[Paragraph("ΗΜΕΡΟΜΗΝΙΑ", th_style), Paragraph("ΑΡ. ΑΠΟΔΕΙΞΗΣ", th_style), Paragraph("ΠΡΟΓΡΑΜΜΑ", th_style), Paragraph("ΕΤΟΣ", th_style), Paragraph("ΤΡΟΠΟΣ", th_style), Paragraph("ΠΟΣΟ (€)", th_style)]]
    total_amount = 0.0
    for _, row in payments_df.iterrows():
        amt = float(row.get('Ποσό (€)', 0))
        total_amount += amt
        table_data.append([
            Paragraph(str(row.get('Ημερομηνία', '')), td_style),
            Paragraph(remove_accents(str(row.get('Αρ. Απόδειξης', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Πρόγραμμα', ''))), td_style),
            Paragraph(str(row.get('Έτος', '')), td_style),
            Paragraph(remove_accents(str(row.get('Τρόπος', ''))), td_style),
            Paragraph(f"{amt:.2f}", td_style)
        ])

    t = Table(table_data, colWidths=[70, 85, 140, 60, 85, 60], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')), ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB'))]))
    story.append(t)
    story.append(Spacer(1, 15))
    story.append(Paragraph(f"ΣΥΝΟΛΙΚΟ ΠΟΣΟ ΠΛΗΡΩΜΩΝ: {total_amount:.2f} EUR", total_style))
    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_attendances_pdf(attendances_df, selected_year):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))

    story.append(Paragraph("ΜΑΖΙΚΗ ΑΝΑΦΟΡΑ ΠΑΡΟΥΣΙΩΝ ΜΑΘΗΤΩΝ (ΤΜΗΜΑ ΙΔΙΩΤΙΚΑ)", title_style))
    story.append(Paragraph(f"ΣΧΟΛΙΚΟ ΕΤΟΣ: {selected_year}", sub_style))
    story.append(Spacer(1, 5))

    table_data = [[Paragraph("ΜΑΘΗΤΗΣ (ID)", th_style), Paragraph("ΗΜΕΡΟΜΗΝΙΑ", th_style), Paragraph("ΠΑΡΟΥΣΙΑ", th_style)]]
    for _, row in attendances_df.iterrows():
        table_data.append([
            Paragraph(remove_accents(str(row.get('student_name', ''))), td_style),
            Paragraph(str(row.get('date', '')), td_style),
            Paragraph(str(row.get('status', '')), td_style)
        ])

    t = Table(table_data, colWidths=[200, 150, 100], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')), ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB'))]))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_monthly_attendances_pdf(report_df, month_str, year_str):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))

    story.append(Paragraph(f"ΜΗΝΙΑΙΑ ΑΝΑΦΟΡΑ ΠΑΡΟΥΣΙΩΝ (ΜΗΝΑΣ: {month_str} | ΕΤΟΣ: {year_str})", title_style))
    story.append(Spacer(1, 5))

    table_data = [[Paragraph("ΟΝΟΜΑΤΕΠΩΝΥΜΟ ΜΑΘΗΤΗ", th_style), Paragraph("ΣΥΝΟΛΙΚΕΣ ΠΑΡΟΥΣΙΕΣ", th_style), Paragraph("ΗΜΕΡΕΣ ΠΑΡΟΥΣΙΑΣ", th_style)]]
    for _, row in report_df.iterrows():
        table_data.append([
            Paragraph(remove_accents(str(row.get('Μαθητής', ''))), td_style),
            Paragraph(str(row.get('Συνολικές Παρουσίες', 0)), td_style),
            Paragraph(str(row.get('Ημέρες', '')), td_style)
        ])

    t = Table(table_data, colWidths=[180, 100, 260], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')), ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB'))]))
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
                user_row = users_df[(users_df['username'].str.strip() == username_input) & (users_df['password'].str.strip() == password_input)]
                if not user_row.empty:
                    st.session_state.logged_in = True
                    st.session_state.username = username_input
                    st.session_state.role = user_row.iloc[0]['role'].strip().lower()
                    st.session_state.page = "menu"
                    st.rerun()  
                else:
                    st.error("❌ Λάθος όνομα χρήστη ή κωδικός πρόσβασης.")
            else:
                st.error("❌ Το φύλλο db_users είναι κενό.")
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
        st.subheader("Παρουσιολόγιο & Πληρωμές")
        col_pr1, col_pr2 = st.columns(2)
        with col_pr1:
            if st.button("✅ Παρουσιολόγιο (Τμήμα Ιδιωτικά)", use_container_width=True):
                st.session_state.page = "attendances_page"; st.rerun()
        with col_pr2:
            if st.button("💳 Καταχώρηση Πληρωμής (Συνδρομή)", use_container_width=True):
                st.session_state.page = "add_payment"; st.rerun()

        st.markdown("---")
        st.subheader("Οικονομικά & Έξοδα")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("📊 Αναφορά & Ιστορικό Πληρωμών Μαθητή", use_container_width=True):
                st.session_state.page = "payment_history_report"; st.rerun()
        with col_p2:
            if st.button("💸 Καταχώρηση Εξόδων", use_container_width=True):
                st.session_state.page = "corp_payments"; st.rerun()

        st.markdown("---")
        st.subheader("Διαχείριση Κέντρου")
        if st.button("📚 Προγράμματα - Τμήματα", use_container_width=True):
            st.session_state.page = "add_structure"; st.rerun()

        st.markdown("---")
        st.subheader("Αναφορές")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("📋 Μητρώο Πελατών", use_container_width=True):
                st.session_state.page = "anafora_pelaton"; st.rerun()
        with col_r2:
            if st.button("📅 Μηνιαία Αναφορά Παρουσίων", use_container_width=True):
                st.session_state.page = "monthly_attendance_report"; st.rerun()

        st.markdown("---")
        st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
        if st.button("🛠️ Εισαγωγή/ Τροποποίηση Παραπόνων & Σχολίων", use_container_width=True):
            st.session_state.page = "complaints_page"; st.rerun()

    # --- ΣΕΛΙΔΑ: ΠΑΡΟΥΣΙΟΛΟΓΙΟ ---
    elif st.session_state.page == "attendances_page":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun()

        st.subheader("✅ Παρουσιολόγιο (Τμήμα: Ιδιωτικά)")
        
        years_options = ["2026-2027", "2027-2028", "2028-2029", "2029-2030"]
        selected_year_att = st.selectbox("📅 Επιλογή Έτους", options=years_options)
        
        today_date_str = datetime.today().strftime('%Y-%m-%d')
        st.info(f"📅 Ημερομηνία Ημέρας: **{today_date_str}**")

        regs_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")
        pays_df = load_specific_db("db_payments")

        if regs_df.empty or cust_df.empty:
            st.warning("Δεν βρέθηκαν δεδομένα εγγραφών ή πελατών.")
        else:
            regs_df.columns = regs_df.columns.astype(str).str.strip()
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce')

            prog_col_name = next((c for c in regs_df.columns if c.lower() in ['program', 'δ', 'j', 'πρόγραμμα']), 'program')
            priv_regs = regs_df[(regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά') & (regs_df['school_year'].astype(str).str.strip() == selected_year_att)].copy()

            if priv_regs.empty:
                st.info(f"Δεν βρέθηκαν μαθητές στο τμήμα 'Ιδιωτικά' για το έτος {selected_year_att}.")
            else:
                att_df = load_specific_db("db_attendances")
                today_att = pd.DataFrame()
                if not att_df.empty and 'date' in att_df.columns:
                    today_att = att_df[(att_df['date'].astype(str).str.strip() == today_date_str) & (att_df['school_year'].astype(str).str.strip() == selected_year_att)]

                with st.form("attendance_form"):
                    st.markdown("#### Λίστα Μαθητών")
                    attendance_inputs = {}

                    for idx, r in priv_regs.iterrows():
                        c_id = str(r['id_customer']).strip()
                        fname = str(r['customer_firstname']).strip()
                        lname = str(r['customer_lastname']).strip()
                        prog_name = str(r.iloc[9] if len(r) > 9 else r.get(prog_col_name, '')).strip()
                        energy_type = str(r.get('energy', 'Ανοιχτή πληρωμή')).strip()
                        
                        cust_match = cust_df[cust_df['id_customer'] == pd.to_numeric(c_id, errors='coerce')]
                        cust_cost = 5.0
                        if not cust_match.empty and len(cust_match.columns) >= 11:
                            try:
                                cust_cost = float(str(cust_match.iloc[0, 10]).replace(',', '.') or 5.0)
                            except:
                                cust_cost = 5.0

                        last_pay_date = None
                        if not pays_df.empty:
                            pays_df.columns = pays_df.columns.astype(str).str.strip()
                            cust_pays = pays_df[
                                (pays_df['id_customer'].astype(str).str.strip() == c_id) &
                                (pays_df['school_year'].astype(str).str.strip() == selected_year_att) &
                                (pays_df.iloc[:, 9].astype(str).str.strip().lower() == prog_name.lower())
                            ]
                            if not cust_pays.empty and 'payment_date' in cust_pays.columns:
                                cust_pays['dt'] = pd.to_datetime(cust_pays['payment_date'], errors='coerce')
                                valid_dts = cust_pays['dt'].dropna()
                                if not valid_dts.empty:
                                    last_pay_date = valid_dts.max()

                        existing_rec = pd.DataFrame()
                        if not today_att.empty:
                            existing_rec = today_att[today_att['id_customer'].astype(str).str.strip() == c_id]
                        
                        default_val = False
                        if not existing_rec.empty:
                            status_col_name = next((c for c in existing_rec.columns if c.lower() == 'status'), None)
                            if status_col_name and str(existing_rec.iloc[0].get(status_col_name, '')).strip().upper() == 'ΝΑΙ':
                                default_val = True

                        col1, col2, col3 = st.columns([3, 2, 2])
                        col1.write(f"**{lname} {fname}** (ID: {c_id})")
                        col2.write(f"Τύπος: *{energy_type}* | Κόστος: {cust_cost}€")
                        val = col3.toggle("ΝΑΙ / ΟΧΙ", value=default_val, key=f"att_{c_id}")
                        
                        attendance_inputs[c_id] = {
                            'val': val,
                            'row_idx': idx + 2,
                            'energy': energy_type,
                            'cust_cost': cust_cost,
                            'prog': prog_name,
                            'last_pay_date': last_pay_date,
                            'fname': fname,
                            'lname': lname
                        }

                    if st.form_submit_button("💾 Αποθήκευση Παρουσιολογίου"):
                        try:
                            gc = get_gspread_client()
                            try:
                                sh_att = gc.open(OLD_SHEET_NAME).worksheet("db_attendances")
                            except:
                                sh_att = gc.open(OLD_SHEET_NAME).add_worksheet(title="db_attendances", rows="1000", cols="10")
                                sh_att.append_row(["id_attendance", "id_customer", "student_name", "school_year", "date", "status"])

                            att_records = sh_att.get_all_records()
                            next_att_id = int(pd.DataFrame(att_records)['id_attendance'].max()) + 1 if att_records and 'id_attendance' in pd.DataFrame(att_records).columns else 1
                            sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")

                            for c_id, data in attendance_inputs.items():
                                status_str = "ΝΑΙ" if data['val'] else "ΟΧΙ"
                                
                                found_att = False
                                if att_records:
                                    for ai, a_row in enumerate(att_records, start=2):
                                        if str(a_row.get('id_customer', '')).strip() == c_id and str(a_row.get('date', '')).strip() == today_date_str and str(a_row.get('school_year', '')).strip() == selected_year_att:
                                            sh_att.update_cell(ai, 6, status_str)
                                            found_att = True
                                            break
                                if not found_att:
                                    sh_att.append_row([next_att_id, c_id, f"{data['lname']} {data['fname']}", selected_year_att, today_date_str, status_str])
                                    next_att_id += 1

                                if data['energy'].lower() == 'ανοιχτή πληρωμή':
                                    count_nais = 0
                                    all_att_list = []
                                    if att_records:
                                        for ar in att_records:
                                            if str(ar.get('id_customer', '')).strip() == c_id and str(ar.get('school_year', '')).strip() == selected_year_att:
                                                st_val = ar.get('status', ar.get('Status', ar.get('STATUS', '')))
                                                all_att_list.append({'date': ar.get('date'), 'status': st_val})
                                    
                                    today_updated = False
                                    for item in all_att_list:
                                        if item['date'] == today_date_str:
                                            item['status'] = status_str
                                            today_updated = True
                                    if not today_updated:
                                        all_att_list.append({'date': today_date_str, 'status': status_str})

                                    for item in all_att_list:
                                        if str(item['status']).strip().upper() == 'ΝΑΙ':
                                            item_dt = pd.to_datetime(item['date'], errors='coerce')
                                            if data['last_pay_date'] is not None and pd.notna(item_dt):
                                                if item_dt > data['last_pay_date']:
                                                    count_nais += 1
                                            else:
                                                count_nais += 1

                                    new_adjusted_cost = data['cust_cost'] * count_nais
                                    sh_reg.update_cell(data['row_idx'], 10, new_adjusted_cost)

                            st.success("✅ Το παρουσιολόγιο αποθηκεύτηκε και οι χρεώσεις ενημερώθηκαν επιτυχώς!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα αποθήκευσης παρουσιών: {e}")

        st.markdown("---")
        st.markdown("#### 📊 Μαζική Αναφορά & PDF Παρουσίων Μαθητών")
        att_full_df = load_specific_db("db_attendances")
        if not att_full_df.empty:
            att_full_df.columns = att_full_df.columns.astype(str).str.strip()
            att_filtered = att_full_df[att_full_df['school_year'].astype(str).str.strip() == selected_year_att]
            
            if not att_filtered.empty:
                st.dataframe(att_filtered[['student_name', 'date', 'status']], use_container_width=True, hide_index=True)
                
                if st.button("📥 Παραγωγή PDF Παρουσίων"):
                    pdf_att = generate_attendances_pdf(att_filtered, selected_year_att)
                    if pdf_att:
                        st.download_button("📥 Κατεβάστε το PDF Παρουσίων", data=pdf_att, file_name="parousies_mathiton.pdf", mime="application/pdf")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένες παρουσίες για το επιλεγμένο έτος.")
        else:
            st.info("Δεν υπάρχουν δεδομένα παρουσιών.")

    # --- ΣΕΛΙΔΑ: ΜΗΝΙΑΙΑ ΑΝΑΦΟΡΑ ΠΑΡΟΥΣΙΩΝ ---
    elif st.session_state.page == "monthly_attendance_report":
        if st.button("⬅️ Μενού", use_container_width=True):
            st.session_state.page = "menu"; st.rerun()

        st.subheader("📅 Μηνιαία Αναφορά Παρουσίων Μαθητών")
        
        col_m1, col_m2 = st.columns(2)
        months_dict = {"Ιανουάριος": "01", "Φεβρουάριος": "02", "Μάρτιος": "03", "Απρίλιος": "04", "Μάιος": "05", "Ιούνιος": "06", "Ιούλιος": "07", "Αύγουστος": "08", "Σεπτέμβριος": "09", "Οκτώβριος": "10", "Νοέμβριος": "11", "Δεκέμβριος": "12"}
        sel_month_name = col_m1.selectbox("Επιλογή Μήνα", options=list(months_dict.keys()))
        sel_month_num = months_dict[sel_month_name]
        sel_year_rep = col_m2.selectbox("Επιλογή Έτους", options=["2026-2027", "2027-2028", "2028-2029"])

        att_df = load_specific_db("db_attendances")
        if att_df.empty:
            st.info("Δεν βρέθηκαν δεδομένα παρουσιών.")
        else:
            att_df.columns = att_df.columns.astype(str).str.strip()
            status_col = next((c for c in att_df.columns if c.lower() == 'status'), None)
            if not status_col:
                st.error("Δεν βρέθηκε η στήλη 'status' στο φύλλο db_attendances.")
            else:
                att_df['dt'] = pd.to_datetime(att_df['date'], errors='coerce')
                filtered_m = att_df[
                    (att_df['school_year'].astype(str).str.strip() == sel_year_rep) &
                    (att_df[status_col].astype(str).str.strip().upper() == 'ΝΑΙ') &
                    (att_df['dt'].dt.strftime('%m') == sel_month_num)
                ].copy()

                if filtered_m.empty:
                    st.info(f"Δεν βρέθηκαν παρουσίες 'ΝΑΙ' για τον μήνα {sel_month_name} ({sel_year_rep}).")
                else:
                    report_rows = []
                    grouped = filtered_m.groupby(['id_customer', 'student_name'])
                    for (c_id, s_name), g in grouped:
                        g_sorted = g.sort_values(by='dt')
                        days_list

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
        mask = (df['id_customer'].astype(str) == str(c_id)) & (df['school_year'].astype(str) == str(school_year)) & (df['program'].astype(str) == str(program))
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
            mask_pay = (pays_df['id_customer'].astype(str) == str(c_id)) & (pays_df['school_year'].astype(str) == str(school_year)) & (pays_df['program'].astype(str) == str(program))
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
                st.session_state.page = "register"
                st.rerun()
            if st.button("📋 Επεξεργασία Στοιχείων Πελατών", use_container_width=True):
                st.session_state.page = "customer_registry"
                st.rerun()
        with col_c2:
            if st.button("👥 Διαχείριση Κατάστασης Πελατών (STOP/Επαναφορά)", use_container_width=True):
                st.session_state.page = "view_customers"
                st.rerun()
            if st.button("🔄 Εισαγωγή Πελάτη σε Τμήμα", use_container_width=True):
                st.session_state.page = "renew_registration"
                st.rerun()

        st.markdown("---")
        st.subheader("Παρουσιολόγιο & Πληρωμές")
        col_pr1, col_pr2 = st.columns(2)
        with col_pr1:
            if st.button("✅ Παρουσιολόγιο (Τμήμα Ιδιωτικά)", use_container_width=True):
                st.session_state.page = "attendances_page"
                st.rerun()
        with col_pr2:
            if st.button("💳 Καταχώρηση Πληρωμής (Συνδρομή)", use_container_width=True):
                st.session_state.page = "add_payment"
                st.rerun()

        st.markdown("---")
        st.subheader("Οικονομικά & Έξοδα")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("📊 Αναφορά & Ιστορικό Πληρωμών Μαθητή", use_container_width=True):
                st.session_state.page = "payment_history_report"
                st.rerun()
        with col_p2:
            if st.button("💸 Καταχώρηση Εξόδων", use_container_width=True):
                st.session_state.page = "corp_payments"
                st.rerun()

        st.markdown("---")
        st.subheader("Διαχείριση Κέντρου")
        if st.button("📚 Τμήματα", use_container_width=True):
            st.session_state.page = "add_structure"
            st.rerun()

        st.markdown("---")
        st.subheader("Αναφορές")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            if st.button("📋 Μητρώο Πελατών", use_container_width=True):
                st.session_state.page = "anafora_pelaton"
                st.rerun()
        with col_r2:
            if st.button("📅 Μηνιαία Αναφορά Παρουσίων", use_container_width=True):
                st.session_state.page = "monthly_attendance_report"
                st.rerun()

        st.markdown("---")
        st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
        if st.button("🛠️ Εισαγωγή/ Τροποποίηση Παραπόνων & Σχολίων", use_container_width=True):
            st.session_state.page = "complaints_page"
            st.rerun()
#----------------------------------------------------------------
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
                        prog_name = str(r.get('program', '')).strip()
                        energy_type = str(r.get('energy', 'Ανοιχτή πληρωμή')).strip()
                        
                        # Κόστος από db_customers στήλη K (index 10)
                        cust_match = cust_df[cust_df['id_customer'] == pd.to_numeric(c_id, errors='coerce')]
                        cust_cost = 5.0
                        if not cust_match.empty and len(cust_match.columns) >= 11:
                            try:
                                cust_cost = float(str(cust_match.iloc[0, 10]).replace(',', '.') or 5.0)
                            except:
                                cust_cost = 5.0

                        # Βρίσκουμε την τελευταία ημερομηνία πληρωμής για αυτόν τον μαθητή/πρόγραμμα/έτος
                        last_pay_date = None
                        if not pays_df.empty:
                            pays_df.columns = pays_df.columns.astype(str).str.strip()
                            cust_pays = pays_df[
                                (pays_df['id_customer'].astype(str).str.strip() == c_id) &
                                (pays_df['school_year'].astype(str).str.strip() == selected_year_att) &
                                (pays_df['program'].astype(str).str.strip().lower() == prog_name.lower())
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
                            if str(existing_rec.iloc[0].get('status', '')).strip().upper() == 'ΝΑΙ':
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
                                
                                # Καταχώρηση παρουσίας
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

                                # Υπολογισμός adjusted_cost αν είναι 'Ανοιχτή πληρωμή' με μηδενισμό από την τελευταία πληρωμή
                                if data['energy'].lower() == 'ανοιχτή πληρωμή':
                                    # Μετράμε παρουσίες ΝΑΙ από την τελευταία πληρωμή και μετά
                                    count_nais = 0
                                    # Φτιάχνουμε μια λίστα προσωρινή με όλες τις παρουσίες συν τη σημερινή
                                    all_att_list = []
                                    if att_records:
                                        for ar in att_records:
                                            if str(ar.get('id_customer', '')).strip() == c_id and str(ar.get('school_year', '')).strip() == selected_year_att:
                                                all_att_list.append({'date': ar.get('date'), 'status': ar.get('status')})
                                    
                                    # Ενημερώνουμε ή προσθέτουμε τη σημερινή στη λίστα υπολογισμού
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
            # Φιλτράρισμα ανά έτος και μήνα (ημερομηνία μορφής YYYY-MM-DD)
            att_df['dt'] = pd.to_datetime(att_df['date'], errors='coerce')
            filtered_m = att_df[
                (att_df['school_year'].astype(str).str.strip() == sel_year_rep) &
                (att_df['status'].astype(str).str.strip().upper() == 'ΝΑΙ') &
                (att_df['dt'].dt.strftime('%m') == sel_month_num)
            ].copy()

            if filtered_m.empty:
                st.info(f"Δεν βρέθηκαν παρουσίες 'ΝΑΙ' για τον μήνα {sel_month_name} ({sel_year_rep}).")
            else:
                # Ομαδοποίηση ανά μαθητή
                report_rows = []
                grouped = filtered_m.groupby(['id_customer', 'student_name'])
                for (c_id, s_name), g in grouped:
                    g_sorted = g.sort_values(by='dt')
                    # Ημέρες σε μορφή 10/10, 12/10 κλπ.
                    days_list = [d.strftime('%d/%m') for d in g_sorted['dt'].dropna()]
                    days_str = ", ".join(days_list)
                    total_count = len(days_list)
                    report_rows.append({
                        'Μαθητής': str(s_name).strip().upper(),
                        'Συνολικές Παρουσίες': total_count,
                        'Ημέρες': days_str
                    })

                rep_df = pd.DataFrame(report_rows)
                # Αλφαβητική ταξινόμηση Ονοματεπωνύμων
                rep_df = rep_df.sort_values(by='Μαθητής', ascending=True).reset_index(drop=True)

                st.dataframe(rep_df, use_container_width=True, hide_index=True)

                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("📥 Παραγωγή PDF Μηνιαίας Αναφοράς"):
                    pdf_m = generate_monthly_attendances_pdf(rep_df, sel_month_name, sel_year_rep)
                    if pdf_m:
                        st.download_button(
                            label="📥 Κατεβάστε το PDF Μηνιαίας Αναφοράς",
                            data=pdf_m,
                            file_name=f"miniaia_anafora_parousion_{sel_month_num}.pdf",
                            mime="application/pdf"
                        )
#--------------------------------------------------------------------------------------------------
    # --- ΣΕΛΙΔΑ: ΕΙΣΑΓΩΓΗ ΣΕ ΤΜΗΜΑ ---
    elif st.session_state.page == "renew_registration":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 

        st.subheader("Εισαγωγή ή Τροποποίηση Πελάτη σε Τμήμα")
        cust_df = load_specific_db("db_customers")
        progs_df = load_specific_db("db_programs")
        classes_df = load_specific_db("db_class")
    
        if not cust_df.empty:
            # Ταξινόμηση πελατών αλφαβητικά με βάση το επώνυμο
            if 'customer_lastname' in cust_df.columns and 'customer_firstname' in cust_df.columns:
                cust_df = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])

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
                adj_cost = st.number_input("💰 Συμφωνηθέν Ποσό (adjusted_cost) (€)", value=default_cost, step=10.0)
                
                class_options = ["Xωρίς Τμήμα"] + (classes_df['class_name'].unique().tolist() if not classes_df.empty else [])
                sel_class = st.selectbox("🏫 Τμήμα", options=class_options)
                
                energy_options = ["Ανοιχτή πληρωμή", "Κλειστή πληρωμή"]
                def_energy_idx = 0
                if reg_to_edit is not None and 'energy' in reg_to_edit:
                    val_en = str(reg_to_edit['energy']).strip()
                    if val_en in energy_options:
                        def_energy_idx = energy_options.index(val_en)
                sel_energy = st.selectbox("💡 Τύπος Πληρωμής (energy)", options=energy_options, index=def_energy_idx)
                
                if st.button("🚀 Οριστικοποίηση Εγγραφής"):
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                        
                        if mode == "Νέα Εγγραφή":
                            records = sh.get_all_records()
                            next_reg = int(pd.DataFrame(records)['id_registration'].max()) + 1 if records and 'id_registration' in pd.DataFrame(records).columns else 1
                            # Σειρά πεδίων: id_registration, id_customer, customer_firstname, customer_lastname, school_year, program, class, status, stop_date, adjusted_cost, energy, xreos
                            row_data = [
                                next_reg, 
                                int(c_id), 
                                str(row['customer_firstname']), 
                                str(row['customer_lastname']), 
                                str(s_year), 
                                str(s_prog), 
                                str(sel_class), 
                                "Active", 
                                "", 
                                adj_cost, 
                                sel_energy, 
                                0.0
                            ]
                            sh.append_row(row_data)
                        else:
                            target_reg_id = str(reg_to_edit['id_registration']).strip()
                            all_rows = sh.get_all_values()
                            header = [h.strip().lower() for h in all_rows[0]]
                            
                            for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                if str(s_row[0]).strip() == target_reg_id:
                                    # Ενημέρωση πεδίων δυναμικά ή με βάση τη σειρά
                                    sh.update_cell(r_idx, 5, str(s_year))         # school_year (στήλη Ε / 5)
                                    sh.update_cell(r_idx, 6, str(s_prog))         # program (στήλη ΣΤ / 6)
                                    sh.update_cell(r_idx, 7, str(sel_class))      # class (στήλη Ζ / 7)
                                    sh.update_cell(r_idx, 10, adj_cost)       # adjusted_cost (στήλη Ι / 10)
                                    sh.update_cell(r_idx, 11, sel_energy)     # energy (στήλη Κ / 11)
                                    break
                                    
                        st.success("✅ Η εγγραφή/τροποποίηση καταχωρήθηκε επιτυχώς!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")
#---------------------------------------------------------------------------------------------------------------
    # --- ΥΠΟΛΟΙΠΕΣ ΣΕΛΙΔΕΣ ---
    # --- ΣΕΛΙΔΑ: ΝΕΑ ΕΓΓΡΑΦΗ ΠΕΛΑΤΗ ---
    elif st.session_state.page == "register":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
            
        st.subheader("Νέα Εγγραφή Πελάτη")
        with st.form("reg_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                cf = st.text_input("Όνομα").strip()
                cl = st.text_input("Επώνυμο").strip()
                pf = st.text_input("Όνομα Πατρός").strip()
                pl = st.text_input("Επώνυμο Πατρός").strip()
                tel = st.text_input("Τηλέφωνο").strip()
                em = st.text_input("Email").strip()
            with c2:
                adr = st.text_input("Διεύθυνση").strip()
                bd = st.date_input("Ημερομηνία Γέννησης", value=datetime(2015,1,1))
                tax_user = st.text_input("Taxis User").strip()
                tax_pass = st.text_input("Taxis Pass").strip()
                adt = st.text_input("ADT").strip()
                afm = st.text_input("AFM").strip()

            if st.form_submit_button("💾 Αποθήκευση"):
                if not cf or not cl:
                    st.error("⚠️ Το Όνομα και το Επώνυμο είναι υποχρεωτικά!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                        recs = sh.get_all_records()
                        next_id = int(pd.DataFrame(recs)['id_customer'].max()) + 1 if recs and 'id_customer' in pd.DataFrame(recs).columns else 1

                        today_str = datetime.today().strftime('%Y-%m-%d')
                        bd_str = bd.strftime('%Y-%m-%d') if bd else today_str

                        # Σειρά πεδίων ακριβώς όπως τη ζήτησες
                        row_data = [
                            next_id, 
                            uppercase_clean(cf), 
                            uppercase_clean(cl), 
                            uppercase_clean(pf), 
                            uppercase_clean(pl), 
                            tel, 
                            em.lower(), 
                            uppercase_clean(adr), 
                            bd_str, 
                            today_str, 
                            tax_user, 
                            tax_pass, 
                            adt, 
                            afm
                        ]
                        sh.append_row(row_data)
                        st.success("✅ Επιτυχής καταχώρηση πελάτη!")
                        time.sleep(1.2)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα αποθήκευσης: {e}")

    # --- ΣΕΛΙΔΑ: ΕΠΕΞΕΡΓΑΣΙΑ ΣΤΟΙΧΕΙΩΝ ΠΕΛΑΤΗ ---
    elif st.session_state.page == "customer_registry":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 

        st.subheader("Επεξεργασία Στοιχείων Πελάτη")
        cust_df = load_specific_db("db_customers")
        
        if cust_df.empty:
            st.warning("⚠️ Η βάση πελατών είναι κενή.")
        else:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            sel = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())

            if sel != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel].iloc[0]
                c_id_str = str(row['id_customer']).strip()

                with st.form("edit_customer_form"):
                    c1, c2 = st.columns(2)
                    with c1:
                        fn = st.text_input("Όνομα", value=str(row.get('customer_firstname', '')))
                        ln = st.text_input("Επώνυμο", value=str(row.get('customer_lastname', '')))
                        p_fn = st.text_input("Όνομα Πατρός", value=str(row.get('parent_firstname', '')))
                        p_ln = st.text_input("Επώνυμο Πατρός", value=str(row.get('parent_lastname', '')))
                        tel = st.text_input("Τηλέφωνο", value=str(row.get('telephone', '')))
                        em = st.text_input("Email", value=str(row.get('email', '')))
                    with c2:
                        adr = st.text_input("Διεύθυνση", value=str(row.get('address', '')))
                        tax_u = st.text_input("Taxis User", value=str(row.get('Taxis_User', '')))
                        tax_p = st.text_input("Taxis Pass", value=str(row.get('Taxis_Pass', '')))
                        adt_val = st.text_input("ADT", value=str(row.get('ADT', row.get('Passport', ''))))
                        afm_val = st.text_input("AFM", value=str(row.get('AFM', '')))

                    if st.form_submit_button("💾 Αποθήκευση Αλλαγών"):
                        try:
                            gc = get_gspread_client()
                            sh_cust = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                            cell = sh_cust.find(c_id_str)
                            if cell:
                                r_idx = cell.row
                                sh_cust.update_cell(r_idx, 2, uppercase_clean(fn))
                                sh_cust.update_cell(r_idx, 3, uppercase_clean(ln))
                                sh_cust.update_cell(r_idx, 4, uppercase_clean(p_fn))
                                sh_cust.update_cell(r_idx, 5, uppercase_clean(p_ln))
                                sh_cust.update_cell(r_idx, 6, tel.strip())
                                sh_cust.update_cell(r_idx, 7, em.strip().lower())
                                sh_cust.update_cell(r_idx, 8, uppercase_clean(adr))
                                sh_cust.update_cell(r_idx, 11, tax_u.strip())
                                sh_cust.update_cell(r_idx, 12, tax_p.strip())
                                sh_cust.update_cell(r_idx, 13, adt_val.strip())
                                sh_cust.update_cell(r_idx, 14, afm_val.strip())
                                
                                st.success("✅ Τα στοιχεία του πελάτη ενημερώθηκαν επιτυχώς!")
                                time.sleep(1.2)
                                st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα ενημέρωσης: {e}")
#---------------------------------------------------------------------------------------------------------
    # --- ΣΕΛΙΔΑ: ΔΙΑΧΕΙΡΙΣΗ STOP ---
    elif st.session_state.page == "view_customers":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun()

        st.subheader("Διαχείριση STOP & Κατάστασης Εγγραφών")
        reg_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")

        if reg_df.empty:
            st.info("Δεν βρέθηκαν εγγραφές στο `db_registrations`.")
        else:
            reg_df.columns = reg_df.columns.astype(str).str.strip()
            
            # Αντικατάσταση ή προσθήκη ονοματεπωνύμων από το db_customers για αλφαβητική ταξινόμηση
            if not cust_df.empty:
                cust_df.columns = cust_df.columns.astype(str).str.strip()
                cust_df['customer_lastname'] = cust_df['customer_lastname'].astype(str).str.strip()
                cust_df['customer_firstname'] = cust_df['customer_firstname'].astype(str).str.strip()
                cust_df = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])
                
                cust_df['display'] = cust_df['customer_lastname'] + " " + cust_df['customer_firstname'] + " (ID: " + cust_df['id_customer'].astype(str) + ")"
                sel_cust = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())

                if sel_cust != "-- Επιλέξτε --":
                    row_c = cust_df[cust_df['display'] == sel_cust].iloc[0]
                    c_id = str(row_c['id_customer']).strip()
                    
                    user_regs = reg_df[reg_df['id_customer'].astype(str).str.strip() == c_id]
                    if user_regs.empty:
                        st.warning("Ο συγκεκριμένος πελάτης δεν έχει καμία εγγραφή.")
                    else:
                        user_regs['display_reg'] = user_regs['school_year'].astype(str) + " - " + user_regs['program'].astype(str) + " (" + user_regs['class'].astype(str) + ")"
                        sel_reg_str = st.selectbox("📚 Επιλογή Εγγραφής / Τμήματος", options=user_regs['display_reg'].tolist())
                        reg_row = user_regs[user_regs['display_reg'] == sel_reg_str].iloc[0]
                        
                        target_reg_id = str(reg_row['id_registration']).strip()
                        
                        current_status = str(reg_row.get('status', 'Active')).strip()
                        status_options = ["Active", "Stop"]
                        def_status_idx = status_options.index(current_status) if current_status in status_options else 0
                        
                        current_cost = 0.0
                        try:
                            current_cost = float(str(reg_row.get('adjusted_cost', 0)).replace(',', '.') or 0.0)
                        except:
                            current_cost = 0.0

                        current_stop_date = str(reg_row.get('stop_date', '')).strip()
                        parsed_date = datetime.today()
                        if current_stop_date:
                            try:
                                parsed_date = datetime.strptime(current_stop_date, '%Y-%m-%d')
                            except:
                                pass

                        with st.form("stop_edit_form"):
                            st.markdown(f"**Επεξεργασία Εγγραφής ID:** {target_reg_id}")
                            new_status = st.selectbox("📌 Κατάσταση (Status)", options=status_options, index=def_status_idx)
                            new_cost = st.number_input("💰 Συμφωνηθέν Ποσό (adjusted_cost) (€)", value=current_cost, step=10.0, format="%.2f")
                            
                            # Αν επιλεγεί Stop, ενημερώνεται η stop_date
                            new_stop_date = st.date_input("📅 Ημερομηνία Σταματήματος (stop_date)", value=parsed_date)
                            
                            if st.form_submit_button("💾 Αποθήκευση Αλλαγών"):
                                try:
                                    gc = get_gspread_client()
                                    sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                    all_rows = sh_reg.get_all_values()
                                    
                                    stop_date_str = new_stop_date.strftime('%Y-%m-%d') if new_status.lower() == 'stop' else ""
                                    
                                    for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                        if str(s_row[0]).strip() == target_reg_id:
                                            # Στήλη H / index 7 -> status, Στήλη I / index 8 -> stop_date, Στήλη J / index 9 -> adjusted_cost
                                            sh_reg.update_cell(r_idx, 8, new_status)
                                            sh_reg.update_cell(r_idx, 9, stop_date_str)
                                            sh_reg.update_cell(r_idx, 10, new_cost)
                                            break
                                            
                                    st.success("✅ Η κατάσταση και τα οικονομικά στοιχεία ενημερώθηκαν επιτυχώς!")
                                    time.sleep(1.2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Σφάλμα ενημέρωσης: {e}")
            else:
                st.dataframe(reg_df, use_container_width=True, hide_index=True)
#---------------------------------------------------------------------------------------------------------
   # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ & ΕΠΕΞΕΡΓΑΣΙΑ ΠΛΗΡΩΜΗΣ ΣΥΝΔΡΟΜΗΣ ---
    elif st.session_state.page == "add_payment":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
            
        st.subheader("Καταχώρηση & Επεξεργασία Πληρωμής Συνδρομής")
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        pays_df = load_specific_db("db_payments")
        
        if not cust_df.empty:
            # Αλφαβητική ταξινόμηση πελατών βάσει επωνύμου
            if 'customer_lastname' in cust_df.columns and 'customer_firstname' in cust_df.columns:
                cust_df = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])

            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            sel = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
            
            if sel != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel].iloc[0]
                c_id = str(row['id_customer']).strip()
                
                # --- ΕΠΙΛΟΓΗ ΛΕΙΤΟΥΡΓΙΑΣ: ΝΕΑ ΠΛΗΡΩΜΗ Η ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΑΣ ---
                pay_mode = st.radio("Ενέργεια Πληρωμής", ["Νέα Καταχώρηση", "Τροποποίηση Υπάρχουσας Πληρωμής"])
                
                if pay_mode == "Τροποποίηση Υπάρχουσας Πληρωμής":
                    if pays_df.empty:
                        st.info("Δεν υπάρχουν καταχωρημένες πληρωμές.")
                    else:
                        pays_df.columns = pays_df.columns.astype(str).str.strip()
                        cust_pays = pays_df[pays_df['id_customer'].astype(str).str.strip() == c_id].copy()
                        
                        if cust_pays.empty:
                            st.info("Ο συγκεκριμένος πελάτης δεν έχει καμία καταχωρημένη πληρωμή.")
                        else:
                            cust_pays['display_pay'] = "ID: " + cust_pays['id_payment'].astype(str) + " | Ημ/νία: " + cust_pays.get('payment_date', '').astype(str) + " | Απόδειξη: " + cust_pays.get('receipt_number', '').astype(str) + " | Ποσό: " + cust_pays.get('amount_paid', '').astype(str) + "€"
                            sel_pay_str = st.selectbox("Επιλογή Πληρωμής προς Επεξεργασία", options=cust_pays['display_pay'].tolist())
                            pay_to_edit = cust_pays[cust_pays['display_pay'] == sel_pay_str].iloc[0]
                            
                            target_pay_id = str(pay_to_edit['id_payment']).strip()
                            
                            cur_rec = str(pay_to_edit.get('receipt_number', '')).strip()
                            cur_amt = 0.0
                            try:
                                cur_amt = float(str(pay_to_edit.get('amount_paid', 0)).replace(',', '.') or 0.0)
                            except:
                                cur_amt = 0.0
                                
                            cur_yr = str(pay_to_edit.get('school_year', '2026-2027')).strip()
                            cur_method = str(pay_to_edit.get('pay_method', 'Μετρητά')).strip()
                            method_options = ["Μετρητά", "Εθνική", "Eurobank", "POS"]
                            cur_method_idx = method_options.index(cur_method) if cur_method in method_options else 0

                            with st.form("edit_pay_form"):
                                st.markdown(f"**Τροποποίηση Πληρωμής ID:** {target_pay_id}")
                                edit_rec = st.text_input("Αρ. Απόδειξης", value=cur_rec).strip()
                                edit_amt = st.number_input("Ποσό (€)", min_value=0.0, value=cur_amt, format="%.2f")
                                edit_yr = st.text_input("Σχολικό Έτος (Στήλη H)", value=cur_yr).strip()
                                edit_method = st.selectbox("Τρόπος Πληρωμής", options=method_options, index=cur_method_idx)
                                
                                if st.form_submit_button("💾 Αποθήκευση Αλλαγών Πληρωμής"):
                                    try:
                                        gc = get_gspread_client()
                                        sh_pay = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                        all_pay_rows = sh_pay.get_all_values()
                                        
                                        for pr_idx, p_row in enumerate(all_pay_rows[1:], start=2):
                                            if str(p_row[0]).strip() == target_pay_id:
                                                # Στήλη F (index 5) -> receipt_number, Στήλη G (index 6) -> amount_paid, Στήλη H (index 7) -> school_year, Στήλη I (index 8) -> pay_method
                                                sh_pay.update_cell(pr_idx, 6, edit_rec)
                                                sh_pay.update_cell(pr_idx, 7, edit_amt)
                                                sh_pay.update_cell(pr_idx, 8, edit_yr)
                                                sh_pay.update_cell(pr_idx, 9, edit_method)
                                                break
                                                
                                        st.success("✅ Η πληρωμή ενημερώθηκε επιτυχώς!")
                                        time.sleep(1.2)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Σφάλμα ενημέρωσης πληρωμής: {e}")
                
                else:
                    # --- ΝΕΑ ΚΑΤΑΧΩΡΗΣΗ ΠΛΗΡΩΜΗΣ ---
                    user_regs = pd.DataFrame()
                    if not reg_df.empty:
                        reg_df.columns = reg_df.columns.astype(str).str.strip()
                        user_regs = reg_df[reg_df['id_customer'].astype(str).str.strip() == c_id]

                    if user_regs.empty:
                        st.warning("⚠️ Ο συγκεκριμένος πελάτης δεν έχει ενεργή εγγραφή σε τμήμα/πρόγραμμα.")
                    else:
                        prog_col_name = next((c for c in user_regs.columns if c.lower() in ['program', 'δ', 'j', 'πρόγραμμα']), 'program')
                        
                        yr = st.selectbox("📅 Έτος", options=user_regs['school_year'].dropna().unique().tolist())
                        prog = st.selectbox("📚 Πρόγραμμα", options=user_regs[prog_col_name].dropna().unique().tolist())
                        
                        current_reg = user_regs[(user_regs['school_year'].astype(str).str.strip() == str(yr)) & (user_regs[prog_col_name].astype(str).str.strip() == str(prog))]
                        
                        adjusted_val = 0.0
                        xreos_val = 0.0
                        
                        if not current_reg.empty:
                            r_item = current_reg.iloc[0]
                            adj_str = str(r_item.get('adjusted_cost', 0)).replace(',', '.')
                            try: adjusted_val = float(adj_str) if adj_str.strip() != '' else 0.0
                            except: adjusted_val = 0.0
                            
                            xreos_raw = r_item.iloc[11] if len(r_item) > 11 else r_item.get('xreos', 0)
                            xreos_str = str(xreos_raw).replace(',', '.')
                            try: xreos_val = float(xreos_str) if xreos_str.strip() != '' else 0.0
                            except: xreos_val = 0.0

                        st.info(f"💰 **Συμφωνηθέν Ποσό (adjusted_cost):** {adjusted_val:.2f}€ | 📉 **Τρέχον Χρέος (xreos):** {xreos_val:.2f}€")

                        with st.form("pay"):
                            rec = st.text_input("Αρ. Απόδειξης").strip()
                            amt = st.number_input("Ποσό (€)", min_value=0.0, format="%.2f")
                            pay_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "Εθνική", "Eurobank", "POS"])
                            
                            if st.form_submit_button("💾 Αποθήκευση"):
                                if not rec or amt <= 0:
                                    st.warning("⚠️ Συμπληρώστε αριθμό απόδειξης και έγκυρο ποσό!")
                                else:
                                    try:
                                        gc = get_gspread_client()
                                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                        
                                        rows_data = sh.get_all_values()
                                        nid = 1
                                        if len(rows_data) > 1:
                                            ids = []
                                            for r in rows_data[1:]:
                                                if r and r[0].isdigit():
                                                    ids.append(int(r[0]))
                                            if ids:
                                                nid = max(ids) + 1
                                        
                                        sh.append_row([
                                            nid,                              # 1: id_payment
                                            int(c_id),                        # 2: id_customer
                                            str(row['customer_lastname']),    # 3: customer_lastname
                                            str(row['customer_firstname']),   # 4: customer_firstname
                                            datetime.today().strftime('%Y-%m-%d'), # 5: payment_date
                                            rec,                              # 6: receipt_number
                                            amt,                              # 7: amount_paid
                                            str(yr),                          # 8 (Στήλη H): school_year
                                            pay_method                        # 9: pay_method
                                        ])
                                        
                                        check_and_update_status_to_finish(int(c_id), yr, prog)
                                        st.success("✅ Η πληρωμή καταχωρήθηκε επιτυχώς!")
                                        time.sleep(1.2)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Σφάλμα κατά την αποθήκευση: {e}")
#------------------------------------------------------------------------------------
    elif st.session_state.page == "payment_history_report":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun()
        st.subheader("Ιστορικό Πληρωμών")
        pays = load_specific_db("db_payments")
        if not pays.empty: st.dataframe(pays, use_container_width=True, hide_index=True)
#------------------------------------------------------------------------------------
    elif st.session_state.page == "corp_payments":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Καταχώρηση Εξόδων")
        with st.form("exp"):
            inv = st.text_input("Αρ. Παραστατικού")
            sup = st.text_input("Προμηθευτής")
            amt = st.number_input("Ποσό (€)", min_value=0.0)
            if st.form_submit_button("💾 Αποθήκευση"):
                gc = get_gspread_client()
                sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                recs = sh.get_all_records()
                nid = int(pd.DataFrame(reg['id']).max()) + 1 if recs and 'id' in pd.DataFrame(recs).columns else 1
                sh.append_row([nid, inv, "Έξοδα", "", sup, "2026-2027", datetime.today().strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'), amt, "Μετρητά"])
                st.success("✅ Το έξοδο αποθηκεύτηκε!")
                st.rerun()

    # --- ΣΕΛΙΔΑ: ΠΡΟΓΡΑΜΜΑΤΑ & ΤΜΗΜΑΤΑ ---
    elif st.session_state.page == "add_structure":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
            
        st.subheader("📚 Διαχείριση Τμημάτων")
        
        progs_df = load_specific_db("db_programs")
        prog_options = progs_df['program_name'].tolist() if not progs_df.empty else []
        c_df = load_specific_db("db_class")
        
        struct_mode = st.radio("Ενέργεια Τμημάτων", ["Δημιουργία Νέου Τμήματος", "Επεξεργασία Υπάρχοντος Τμήματος"])
        
        if struct_mode == "Δημιουργία Νέου Τμήματος":
            st.markdown("#### Δημιουργία Νέου Τμήματος")
            with st.form("new_class_form", clear_on_submit=True):
                new_class_name = st.text_input("Όνομα Νέου Τμήματος (π.χ. Τμήμα Α, Group 1)").strip()
                sel_prog_for_class = st.selectbox("📚 Επιλογή Προγράμματος", options=prog_options if prog_options else ["-- Κανένα Πρόγραμμα --"])
                
                year_options = [f"{y}-{y+1}" for y in range(2026, 2030)]
                sel_year_for_class = st.selectbox("📅 Σχολικό Έτος", options=year_options)
                
                if st.form_submit_button("💾 Αποθήκευση Τμήματος"):
                    if not new_class_name:
                        st.warning("⚠️ Συμπληρώστε το όνομα του τμήματος!")
                    else:
                        try:
                            start_year_str = sel_year_for_class.split("-")[0]
                            end_year_str = "20" + sel_year_for_class.split("-")[1] if len(sel_year_for_class.split("-")[1]) == 2 else str(int(start_year_str) + 1)
                            
                            auto_start_date = f"{start_year_str}-09-01"
                            auto_end_date = f"{end_year_str}-07-31"

                            gc = get_gspread_client()
                            sh_c = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                            c_rows = sh_c.get_all_values()
                            next_c_id = 1
                            if len(c_rows) > 1:
                                ids = [int(r[0]) for r in c_rows[1:] if r and r[0].isdigit()]
                                if ids:
                                    next_c_id = max(ids) + 1
                            
                            sh_c.append_row([
                                next_c_id, 
                                uppercase_clean(new_class_name), 
                                str(sel_prog_for_class), 
                                str(sel_year_for_class), 
                                auto_start_date, 
                                auto_end_date
                            ])
                            
                            st.success(f"✅ Το τμήμα '{new_class_name}' δημιουργήθηκε επιτυχώς!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα αποθήκευσης τμήματος: {e}")
        
        else:
            st.markdown("#### Επεξεργασία Υπάρχοντος Τμήματος")
            if c_df.empty:
                st.info("Δεν υπάρχουν διαθέσιμα τμήματα προς επεξεργασία.")
            else:
                c_df.columns = c_df.columns.astype(str).str.strip()
                
                col_id = c_df.columns[0]
                col_name = 'class_name' if 'class_name' in c_df.columns else c_df.columns[1]
                col_prog = 'program' if 'program' in c_df.columns else (c_df.columns[2] if len(c_df.columns) > 2 else c_df.columns[1])
                col_year = 'school_year' if 'school_year' in c_df.columns else (c_df.columns[3] if len(c_df.columns) > 3 else c_df.columns[0])

                c_df['display_class'] = "ID: " + c_df[col_id].astype(str) + " | Τμήμα: " + c_df[col_name].astype(str) + " | Έτος: " + c_df[col_year].astype(str)
                
                sel_class_str = st.selectbox("Επιλογή Τμήματος προς Επεξεργασία", options=c_df['display_class'].tolist())
                class_to_edit = c_df[c_df['display_class'] == sel_class_str].iloc[0]
                
                target_class_id = str(class_to_edit[col_id]).strip()
                cur_name = str(class_to_edit[col_name]).strip()
                cur_prog = str(class_to_edit[col_prog]).strip()
                cur_year = str(class_to_edit[col_year]).strip()
                
                prog_idx = prog_options.index(cur_prog) if cur_prog in prog_options else 0
                year_options = [f"{y}-{y+1}" for y in range(2026, 2030)]
                year_idx = year_options.index(cur_year) if cur_year in year_options else 0

                with st.form("edit_class_form"):
                    edit_c_name = st.text_input("Όνομα Τμήματος", value=cur_name).strip()
                    edit_prog = st.selectbox("Πρόγραμμα", options=prog_options if prog_options else [cur_prog], index=prog_idx)
                    edit_year = st.selectbox("Σχολικό Έτος", options=year_options, index=year_idx)
                    
                    if st.form_submit_button("💾 Αποθήκευση Αλλαγών Τμήματος"):
                        try:
                            start_year_str = edit_year.split("-")[0]
                            end_year_str = "20" + edit_year.split("-")[1] if len(edit_year.split("-")[1]) == 2 else str(int(start_year_str) + 1)
                            
                            new_start = f"{start_year_str}-09-01"
                            new_end = f"{end_year_str}-07-31"

                            gc = get_gspread_client()
                            sh_c = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                            all_c_rows = sh_c.get_all_values()
                            
                            for r_idx, crow in enumerate(all_c_rows[1:], start=2):
                                if str(crow[0]).strip() == target_class_id:
                                    sh_c.update_cell(r_idx, 2, uppercase_clean(edit_c_name))
                                    sh_c.update_cell(r_idx, 3, str(edit_prog))
                                    sh_c.update_cell(r_idx, 4, str(edit_year))
                                    sh_c.update_cell(r_idx, 5, new_start)
                                    sh_c.update_cell(r_idx, 6, new_end)
                                    break
                                    
                            st.success("✅ Το τμήμα ενημερώθηκε επιτυχώς!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα ενημέρωσης τμήματος: {e}")

        st.markdown("---")
        st.markdown("#### Υπάρχοντα Τμήματα")
        c_df_refreshed = load_specific_db("db_class")
        if not c_df_refreshed.empty:
            st.dataframe(c_df_refreshed, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν τμήματα.")

        st.markdown("---")
        st.markdown("#### Υπάρχοντα Προγράμματα")
        if not progs_df.empty:
            st.dataframe(progs_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν προγράμματα.")
#------------------------------------------------------------------------
    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Μητρώο Πελατών")
        reg = load_specific_db("db_registrations")
        if not reg.empty: st.dataframe(reg, use_container_width=True)
#------------------------------------------------------------------------
    elif st.session_state.page == "unassigned_customers":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun()
        st.subheader("Ανένταχτοι Πελάτοι")
        cust = load_specific_db("db_customers")
        if not cust.empty: st.dataframe(cust, use_container_width=True)
#------------------------------------------------------------------------
    elif st.session_state.page == "complaints_page":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Παράπονα & Σχόλια")
        comp = load_specific_db("db_complaint")
        if not comp.empty: st.dataframe(comp, use_container_width=True)

    else:
        st.session_state.page = "menu"
        st.rerun()

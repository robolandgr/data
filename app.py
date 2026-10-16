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
        st.subheader("Οικονομικά & Εξοδα")
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
        if st.button("📋 Μητρώο Πελατών", use_container_width=True):
            st.session_state.page = "anafora_pelaton"; st.rerun()

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
        if regs_df.empty:
            st.warning("Δεν βρέθηκαν εγγραφές.")
        else:
            regs_df.columns = regs_df.columns.astype(str).str.strip()
            # Φιλτράρισμα για τμήμα 'Ιδιωτικά' και επιλεγμένο έτος
            priv_regs = regs_df[(regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά') & (regs_df['school_year'].astype(str).str.strip() == selected_year_att)].copy()

            if priv_regs.empty:
                st.info(f"Δεν βρέθηκαν μαθητές στο τμήμα 'Ιδιωτικά' για το έτος {selected_year_att}.")
            else:
                # Φόρτωση προγραμμάτων για να ξέρουμε την αρχική τιμή αν χρειαστεί
                progs_df = load_specific_db("db_programs")
                
                # Φόρτωση υπαρχουσών παρουσιών για σημερα
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
                        energy_type = str(r.get('energy', 'Ανοιχτή πληρωμή')).strip()
                        
                        # Προεπιλογή ΝΑΙ/ΟΧΙ βάσει αποθήκευσης ημέρας αν υπάρχει
                        existing_rec = pd.DataFrame()
                        if not today_att.empty:
                            existing_rec = today_att[today_att['id_customer'].astype(str).str.strip() == c_id]
                        
                        default_val = False
                        if not existing_rec.empty:
                            if str(existing_rec.iloc[0].get('status', '')).strip().upper() == 'ΝΑΙ':
                                default_val = True

                        col1, col2, col3 = st.columns([3, 2, 2])
                        col1.write(f"**{lname} {fname}** (ID: {c_id})")
                        col2.write(f"Τύπος: *{energy_type}*")
                        val = col3.toggle("ΝΑΙ / ΟΧΙ", value=default_val, key=f"att_{c_id}")
                        attendance_inputs[c_id] = {
                            'val': val,
                            'row_idx': idx + 2, # Google sheet row index
                            'energy': energy_type,
                            'prog': str(r.get('program', '')).strip(),
                            'adj_cost': str(r.get('adjusted_cost', '')).strip(),
                            'fname': fname,
                            'lname': lname
                        }

                    if st.form_submit_button("💾 Αποθήκευση Παρουσιολογίου"):
                        try:
                            gc = get_gspread_client()
                            
                            # Δημιουργία ή έλεγχος worksheet db_attendances
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
                                
                                # 1. Καταχώρηση παρουσίας στο db_attendances
                                # Ελέγχουμε αν υπάρχει ήδη εγγραφή για σήμερα να την ενημερώσουμε ή να προσθέσουμε νέα
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

                                # 2. Υπολογισμός adjusted_cost αν είναι 'Ανοιχτή πληρωμή'
                                if data['energy'].lower() == 'ανοιχτή πληρωμή' and data['val']:
                                    # Βρίσκουμε τη βασική τιμή προγράμματος
                                    base_cost = 5.0 # προεπιλογή 5€ αν δεν βρεθεί
                                    if not progs_df.empty:
                                        p_match = progs_df[progs_df['program_name'].astype(str).str.strip().str.lower() == data['prog'].lower()]
                                        if not p_match.empty:
                                            base_cost = float(p_match.iloc[0].get('program_cost', 5.0) or 5.0)
                                    
                                    # Υπολογισμός πολλαπλασιασμού βάσει παρουσιών ΝΑΙ
                                    # Μετράμε πόσες φορές έχει γίνει ΝΑΙ συνολικά για αυτόν τον μαθητή στο έτος
                                    current_nais = 1
                                    if att_records:
                                        current_nais = sum(1 for ar in att_records if str(ar.get('id_customer', '')).strip() == c_id and str(ar.get('school_year', '')).strip() == selected_year_att and str(ar.get('status', '')).strip().upper() == 'ΝΑΙ')
                                        if not found_att and data['val']:
                                            current_nais += 1

                                    new_adjusted_cost = base_cost * current_nais
                                    sh_reg.update_cell(data['row_idx'], 10, new_adjusted_cost)

                            st.success("✅ Το παρουσιολόγιο αποθηκεύτηκε και οι χρεώσεις ενημερώθηκαν επιτυχώς!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα αποθήκευσης παρουσιών: {e}")

        # --- ΜΑΖΙΚΗ ΑΝΑΦΟΡΑ & PDF ΠΑΡΟΥΣΙΩΝ ---
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

    # --- ΣΕΛΙΔΑ: ΕΙΣΑΓΩΓΗ ΣΕ ΤΜΗΜΑ (ΜΕ ΑΝΟΙΧΤΗ/ΚΛΕΙΣΤΗ ΠΛΗΡΩΜΗ ΣΤΗ ΣΤΗΛΗ Κ) ---
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
                
                # Επιλογή τύπου πληρωμής για αποθήκευση στη στήλη K (energy)
                energy_options = ["Ανοιχτή πληρωμή", "Κλειστή πληρωμή"]
                def_energy_idx = 0
                if reg_to_edit is not None and 'energy' in reg_to_edit:
                    val_en = str(reg_to_edit['energy']).strip()
                    if val_en in energy_options:
                        def_energy_idx = energy_options.index(val_en)
                sel_energy = st.selectbox("💡 Τύπος Πληρωμής (Στήλη K - energy)", options=energy_options, index=def_energy_idx)
                
                if st.button("🚀 Οριστικοποίηση Εγγραφής"):
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                        if mode == "Νέα Εγγραφή":
                            records = sh.get_all_records()
                            next_reg = int(pd.DataFrame(records)['id_registration'].max()) + 1 if records and 'id_registration' in pd.DataFrame(records).columns else 1
                            # Στήλες: id_registration, id_customer, customer_firstname, customer_lastname, school_year, program, class, status, stop_date, adjusted_cost, energy (Στήλη K)
                            sh.append_row([next_reg, int(c_id), str(row['customer_firstname']), str(row['customer_lastname']), str(s_year), str(s_prog), str(sel_class), "ACTIVE", "", adj_cost, sel_energy])
                        else:
                            target_reg_id = str(reg_to_edit['id_registration']).strip()
                            all_rows = sh.get_all_values()
                            for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                if str(s_row[0]).strip() == target_reg_id:
                                    sh.update_cell(r_idx, 5, str(s_year))
                                    sh.update_cell(r_idx, 6, str(s_prog))
                                    sh.update_cell(r_idx, 7, str(sel_class))
                                    sh.update_cell(r_idx, 10, adj_cost)
                                    sh.update_cell(r_idx, 11, sel_energy) # Στήλη K
                                    break
                        st.success("✅ Η εγγραφή ενημερώθηκε επιτυχώς με τον τύπο πληρωμής!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΥΠΟΛΟΙΠΕΣ ΣΕΛΙΔΕΣ (Μητρώο, Πληρωμές, Εξοδα, Παράπονα κλπ.) ---
    elif st.session_state.page == "register":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Νέα Εγγραφή Πελάτη")
        cat = st.radio("Κατηγορία", ["Μαθητής", "Ενήλικας"])
        with st.form("reg_form", clear_on_submit=True):
            cf = st.text_input("Όνομα").strip()
            cl = st.text_input("Επώνυμο").strip()
            tel = st.text_input("Τηλέφωνο").strip()
            if st.form_submit_button("💾 Αποθήκευση"):
                if cf and cl:
                    gc = get_gspread_client()
                    sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                    recs = sh.get_all_records()
                    next_id = int(pd.DataFrame(recs)['id_customer'].max()) + 1 if recs and 'id_customer' in pd.DataFrame(recs).columns else 1
                    sh.append_row([next_id, uppercase_clean(cf), uppercase_clean(cl), "", "", tel, "", "", datetime.today().strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'), cat, "", "", "", "", "", "", "", "", ""])
                    st.success("✅ Αποθηκεύτηκε!")
                    st.rerun()

    elif st.session_state.page == "customer_registry":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Επεξεργασία Στοιχείων Πελάτη")
        cust_df = load_specific_db("db_customers")
        if not cust_df.empty:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str)
            sel = st.selectbox("Επιλογή", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
            if sel != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel].iloc[0]
                with st.form("edit"):
                    fn = st.text_input("Όνομα", value=row['customer_firstname'])
                    ln = st.text_input("Επώνυμο", value=row['customer_lastname'])
                    if st.form_submit_button("💾 Αποθήκευση"):
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                        cell = sh.find(str(row['id_customer']))
                        if cell:
                            sh.update_cell(cell.row, 2, uppercase_clean(fn))
                            sh.update_cell(cell.row, 3, uppercase_clean(ln))
                            st.success("✅ Ενημερώθηκε!")
                            st.rerun()

    elif st.session_state.page == "view_customers":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun()
        st.subheader("Διαχείριση STOP")
        reg_df = load_specific_db("db_registrations")
        if not reg_df.empty:
            st.dataframe(reg_df, use_container_width=True, hide_index=True)

    elif st.session_state.page == "add_payment":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Καταχώρηση Πληρωμής Συνδρομής")
        cust_df = load_specific_db("db_customers")
        if not cust_df.empty:
            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str)
            sel = st.selectbox("Πελάτης", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
            if sel != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel].iloc[0]
                with st.form("pay"):
                    rec = st.text_input("Αρ. Απόδειξης")
                    amt = st.number_input("Ποσό (€)", min_value=0.0)
                    yr = st.selectbox("Έτος", ["2026-2027", "2027-2028"])
                    prog = st.text_input("Πρόγραμμα")
                    if st.form_submit_button("💾 Αποθήκευση"):
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                        recs = sh.get_all_records()
                        nid = int(pd.DataFrame(recs)['id_payment'].max()) + 1 if recs and 'id_payment' in pd.DataFrame(recs).columns else 1
                        sh.append_row([nid, row['id_customer'], row['customer_lastname'], row['customer_firstname'], datetime.today().strftime('%Y-%m-%d'), rec, amt, yr, "Μετρητά", prog])
                        check_and_update_status_to_finish(row['id_customer'], yr, prog)
                        st.success("✅ Καταχωρήθηκε!")
                        st.rerun()

    elif st.session_state.page == "payment_history_report":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun()
        st.subheader("Ιστορικό Πληρωμών")
        pays = load_specific_db("db_payments")
        if not pays.empty:
            st.dataframe(pays, use_container_width=True, hide_index=True)

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
                nid = int(pd.DataFrame(recs)['id'].max()) + 1 if recs and 'id' in pd.DataFrame(recs).columns else 1
                sh.append_row([nid, inv, "Έξοδα", "", sup, "2026-2027", datetime.today().strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'), datetime.today().strftime('%Y-%m-%d'), amt, "Μετρητά"])
                st.success("✅ Το έξοδο αποθηκεύτηκε!")
                st.rerun()

    elif st.session_state.page == "add_structure":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Προγράμματα & Τμήματα")
        p_df = load_specific_db("db_programs")
        c_df = load_specific_db("db_class")
        st.dataframe(p_df, use_container_width=True)
        st.dataframe(c_df, use_container_width=True)

    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Μητρώο Πελατών")
        reg = load_specific_db("db_registrations")
        if not reg.empty: st.dataframe(reg, use_container_width=True)

    elif st.session_state.page == "unassigned_customers":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun()
        st.subheader("Ανένταχτοι Πελάτες")
        cust = load_specific_db("db_customers")
        if not cust.empty: st.dataframe(cust, use_container_width=True)

    elif st.session_state.page == "complaints_page":
        if st.button("⬅️ Μενού", use_container_width=True): st.session_state.page = "menu"; st.rerun() 
        st.subheader("Παράπονα & Σχόλια")
        comp = load_specific_db("db_complaint")
        if not comp.empty: st.dataframe(comp, use_container_width=True)

    else:
        st.session_state.page = "menu"
        st.rerun()

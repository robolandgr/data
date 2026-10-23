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

def safe_float_conversion(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    # Αν το string περιέχει κόμμα (π.χ. "5,2" ή "5,20"), το μετατρέπουμε σε τελεία για σωστό float
    if ',' in s and '.' not in s:
        s = s.replace(',', '.')
    elif ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    try:
        return float(s)
    except:
        return 0.0

def remove_accents(input_str):
    if not isinstance(input_str, str): return str(input_str)
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def update_customer_xreos(c_id, school_year, new_payment_amt=0.0):
    try:
        gc = get_gspread_client()
        sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
        regs_data = sh_reg.get_all_records()
        if not regs_data:
            return
            
        df_reg = pd.DataFrame(regs_data)
        df_reg.columns = df_reg.columns.astype(str).str.strip()
        
        # Καθαρισμός ID και έτους για να ταιριάζουν 100%
        target_id = str(c_id).strip()
        target_year = str(school_year).strip()
        
        df_reg['id_clean'] = df_reg['id_customer'].astype(str).str.strip()
        df_reg['year_clean'] = df_reg['school_year'].astype(str).str.strip()
        
        matched = df_reg[(df_reg['id_clean'] == target_id) & (df_reg['year_clean'] == target_year)]
        
        if matched.empty:
            # Αν δεν βρει ακριβές έτος, δοκιμάζουμε μόνο με ID για να μην κολλήσει ποτέ
            matched = df_reg[df_reg['id_clean'] == target_id]
            
        if matched.empty:
            return
            
        for idx_row in matched.index:
            reg_row_idx = idx_row + 2 # Google Sheet row index (ξεκινάει από 2 λόγω headers)
            r_item = df_reg.iloc[idx_row]
            
            # Διάβαστη τρέχοντος χρέους και total_paid
            raw_xreos = str(r_item.get('xreos', 0)).replace(',', '.').strip()
            current_xreos = float(raw_xreos) if raw_xreos and raw_xreos != '' else 0.0
            
            raw_paid = str(r_item.get('total_paid', 0)).replace(',', '.').strip()
            current_total_paid = float(raw_paid) if raw_paid and raw_paid != '' else 0.0
            
            if new_payment_amt > 0:
                new_total_paid = current_total_paid + float(new_payment_amt)
                new_xreos = current_xreos - float(new_payment_amt)
                
                # Στήλη L (12) για το xreos, Στήλη O (15) για το total_paid
                sh_reg.update_cell(reg_row_idx, 12, new_xreos)
                sh_reg.update_cell(reg_row_idx, 15, new_total_paid)

    except Exception as e:
        print(f"Error in update_customer_xreos: {e}")
        pass        
def generate_single_attendance_pdf(attendance_rows, date_str, school_year):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=13, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=8)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=9, fontName='Helvetica', textColor=colors.HexColor('#111827'))

    story.append(Paragraph("ΠΑΡΟΥΣΙΟΛΟΓΙΟ ΜΑΘΗΤΩΝ (ΤΜΗΜΑ: ΙΔΙΩΤΙΚΑ)", title_style))
    story.append(Paragraph(f"ΗΜΕΡΟΜΗΝΙΑ: {date_str} | ΣΧΟΛΙΚΟ ΕΤΟΣ: {school_year}", sub_style))
    story.append(Spacer(1, 5))

    table_data = [[Paragraph("ΟΝΟΜΑΤΕΠΩΝΥΜΟ ΜΑΘΗΤΗ", th_style), Paragraph("ΠΡΟΓΡΑΜΜΑ", th_style), Paragraph("ΠΑΡΟΥΣΙΑ (ΝΑΙ/ΟΧΙ)", th_style)]]
    for r in attendance_rows:
        table_data.append([
            Paragraph(remove_accents(str(r.get('name', ''))), td_style),
            Paragraph(remove_accents(str(r.get('prog', ''))), td_style),
            Paragraph(str(r.get('status', 'ΟΧΙ')), td_style)
        ])

    t = Table(table_data, colWidths=[220, 180, 140], repeatRows=1)
    t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')), ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB'))]))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

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

def generate_report_pdf_customers(report_df, selected_year="ΟΛΑ", selected_class="Ιδιωτικά"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=13, alignment=1, textColor=colors.HexColor('#1F2937'), fontName='Helvetica-Bold')
    subtitle_style = ParagraphStyle('SubTitle', parent=styles['Normal'], fontSize=10, alignment=1, spaceAfter=15, textColor=colors.HexColor('#4B5563'), fontName='Helvetica-Bold')
    label_style = ParagraphStyle('Label', parent=styles['Normal'], fontSize=7.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'))
    value_style = ParagraphStyle('Value', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))
    name_style = ParagraphStyle('Name', parent=styles['Normal'], fontSize=9.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'))

    story.append(Paragraph("ΜΗΤΡΩΟ ΜΑΘΗΤΩΝ", title_style))
    story.append(Paragraph(f"ΣΧΟΛΙΚΟ ΕΤΟΣ: {remove_accents(str(selected_year).upper())}", subtitle_style))
    story.append(Spacer(1, 5))

    class_lower = str(selected_class).strip().lower()

    for _, row in report_df.iterrows():
        lname = remove_accents(str(row.get('customer_lastname', '')).upper())
        fname = remove_accents(str(row.get('customer_firstname', '')).upper())
        
        p_lname = remove_accents(str(row.get('parent_lastname', '')))
        p_fname = remove_accents(str(row.get('parent_firstname', '')))
        tele = remove_accents(str(row.get('telephone', '')))
        mail = remove_accents(str(row.get('email', '')))
        status = remove_accents(str(row.get('status', '')))

        if class_lower == "ιδιωτικά":
            adj_c = remove_accents(str(row.get('adjusted_cost', '0')))
            energy_t = remove_accents(str(row.get('energy', '')))
            xreos_v = remove_accents(str(row.get('xreos', '0')))

            row_data = [
                [Paragraph("ΜΑΘΗΤΗΣ:", label_style), Paragraph("ΓΟΝΕΑΣ:", label_style), Paragraph("ΤΗΛΕΦΩΝΟ / EMAIL:", label_style)],
                [Paragraph(f"{lname} {fname}", name_style), Paragraph(f"{p_lname} {p_fname}", value_style), Paragraph(f"{tele} | {mail}", value_style)],
                [Paragraph("ΚΑΤΑΣΤΑΣΗ:", label_style), Paragraph("ΤΥΠΟΣ ΧΡΕΩΣΗΣ:", label_style), Paragraph("ΣΥΜΦΩΝΗΘΕΝ / ΧΡΕΟΣ:", label_style)],
                [Paragraph(status, value_style), Paragraph(energy_t, value_style), Paragraph(f"Κόστος: {adj_c}€ | Χρέος: {xreos_v}€", value_style)]
            ]
            t = Table(row_data, colWidths=[180, 150, 210])

        else: # Voucher
            m_bar = ""
            for k in ['m_bardia', 'M Bardia', 'bardia', 'Βάρδια', 'βαρδια']:
                if k in row and pd.notna(row[k]) and str(row[k]).strip().lower() != 'nan':
                    m_bar = remove_accents(str(row[k]).strip())
                    break
            
            stop_d = remove_accents(str(row.get('stop_date', '')))

            row_data = [
                [Paragraph("ΜΑΘΗΤΗΣ (VOUCHER):", label_style), Paragraph("ΓΟΝΕΑΣ:", label_style), Paragraph("ΤΗΛΕΦΩΝΟ / EMAIL:", label_style)],
                [Paragraph(f"{lname} {fname}", name_style), Paragraph(f"{p_lname} {p_fname}", value_style), Paragraph(f"{tele} | {mail}", value_style)],
                [Paragraph("ΚΑΤΑΣΤΑΣΗ:", label_style), Paragraph("ΒΑΡΔΙΑ:", label_style), Paragraph("ΗΜ. ΛΗΞΗΣ (STOP DATE):", label_style)],
                [Paragraph(status, value_style), Paragraph(m_bar if m_bar else "-", value_style), Paragraph(stop_d, value_style)]
            ]
            t = Table(row_data, colWidths=[180, 150, 210])

        t.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')), 
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#F3F4F6')), 
            ('TOPPADDING', (0, 0), (-1, -1), 4), 
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4)
        ]))
        story.append(t)
        story.append(Spacer(1, 6))

    doc.build(story)
    buffer.seek(0)
    return buffer

def generate_summary_payments_pdf(summary_df, school_year_filter, class_filter):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))

    story.append(Paragraph("ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΝΑΦΟΡΑ ΟΙΚΟΝΟΜΙΚΩΝ ΜΑΘΗΤΩΝ", title_style))
    story.append(Paragraph(f"ΦΙΛΤΡΑ -> ΕΤΟΣ: {school_year_filter} | ΤΜΗΜΑ: {class_filter}", sub_style))
    story.append(Spacer(1, 5))

    table_data = [[Paragraph("ΟΝΟΜΑΤΕΠΩΝΥΜΟ", th_style), Paragraph("ΣΥΜΦΩΝΗΘΕΝ (€)", th_style), Paragraph("CLASS", th_style), Paragraph("ΠΛΗΡΩΜΕΝΑ (€)", th_style), Paragraph("ΥΠΟΛΟΙΠΟ (€)", th_style)]]
    
    for _, row in summary_df.iterrows():
        table_data.append([
            Paragraph(remove_accents(str(row.get('Ονοματεπώνυμο', ''))), td_style),
            Paragraph(f"{row.get('Συμφωνηθέν (€)', 0):.2f}", td_style),
            Paragraph(remove_accents(str(row.get('Class', ''))), td_style),
            Paragraph(f"{row.get('Πληρωμένα (€)', 0):.2f}", td_style),
            Paragraph(f"{row.get('Υπόλοιπο (€)', 0):.2f}", td_style)
        ])

    t = Table(table_data, colWidths=[150, 90, 80, 100, 80], repeatRows=1)
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
    
    # Μεταφορά του κουμπιού ενημέρωσης στην αριστερή πλευρική στήλη (sidebar)
    if st.sidebar.button("🔄 Ενημέρωση Δεδομένων", use_container_width=True, disabled=st.session_state.is_submitting):
        st.session_state.is_submitting = True
        load_specific_db.clear()
        st.success("Τα δεδομένα ενημερώθηκαν επιτυχώς από τη Google!")
        time.sleep(1)
        st.session_state.is_submitting = False
        st.rerun()

    st.sidebar.markdown("---")
    
    if st.sidebar.button("🚪 Έξοδος", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.page = "menu"
        st.rerun()
    # 1. ΚΕΝΤΡΙΚΟ ΜΕΝΟΥ (ΜΟΝΤΕΡΝΟΣ ΣΧΕΔΙΑΣΜΟΣ)
    if st.session_state.page == "menu":
        # Custom CSS για ωραιότερα οπτικά εφέ και κάρτες
        st.markdown("""
            <style>
                .main-title {
                    text-align: center;
                    font-size: 2.8rem;
                    font-weight: 800;
                    color: #1E3A8A;
                    margin-bottom: 0px;
                }
                .sub-title {
                    text-align: center;
                    font-size: 1.2rem;
                    color: #6B7280;
                    margin-bottom: 30px;
                }
                .card-header {
                    font-size: 1.1rem;
                    font-weight: 700;
                    color: #374151;
                    margin-bottom: 10px;
                    border-bottom: 2px solid #E5E7EB;
                    padding-bottom: 5px;
                }
            </style>
        """, unsafe_allow_html=True)

        st.markdown("<h1 class='main-title'>🤖 ROBOLAND DB 🤖</h1>", unsafe_allow_html=True)
        st.markdown("<p class='sub-title'>Κεντρικό Σύστημα Διαχείρισης</p>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        # Χρήση στηλών για οπτικό διαχωρισμό σε "Κάρτες" (Cards)
        col_m1, col_m2, col_m3 = st.columns(3, gap="medium")

        # --- ΚΑΡΤΑ 1: ΔΙΑΧΕΙΡΙΣΗ ΠΕΛΑΤΩΝ ---
        with col_m1:
            with st.container(border=True):
                st.markdown("<div class='card-header'>👥 Μαθητές</div>", unsafe_allow_html=True)
                if st.button("➕ Νέος Πελάτης", use_container_width=True, key="btn_reg"):
                    st.session_state.page = "register"
                    st.rerun()
                if st.button("📋 Επεξεργασία Στοιχείων", use_container_width=True, key="btn_cust_reg"):
                    st.session_state.page = "customer_registry"
                    st.rerun()
                if st.button("🔄 Εισαγωγή σε Τμήμα", use_container_width=True, key="btn_renew"):
                    st.session_state.page = "renew_registration"
                    st.rerun()
                if st.button("🛑 Διαχείριση STOP/Κατάστασης", use_container_width=True, key="btn_view_cust"):
                    st.session_state.page = "view_customers"
                    st.rerun()

        # --- ΚΑΡΤΑ 2: ΠΑΡΟΥΣΙΟΛΟΓΙΟ & ΟΙΚΟΝΟΜΙΚΑ ---
        with col_m2:
            with st.container(border=True):
                st.markdown("<div class='card-header'>💳 Λειτουργία</div>", unsafe_allow_html=True)
                if st.button("✅ Παρουσιολόγιο", use_container_width=True, key="btn_att"):
                    st.session_state.page = "attendances_page"
                    st.rerun()
                if st.button("💵 Καταχώρηση Πληρωμής", use_container_width=True, key="btn_pay"):
                    st.session_state.page = "add_payment"
                    st.rerun()
                if st.button("📊 Ιστορικό Πληρωμών", use_container_width=True, key="btn_hist_rep"):
                    st.session_state.page = "payment_history_report"
                    st.rerun()
                if st.button("📚 Διαχείριση Τμημάτων", use_container_width=True, key="btn_struct"):
                    st.session_state.page = "add_structure"
                    st.rerun()

        # --- ΚΑΡΤΑ 3: ΕΞΟΔΑ, ΕΤΑΙΡΕΙΕΣ & ΑΝΑΦΟΡΕΣ ---
        with col_m3:
            with st.container(border=True):
                st.markdown("<div class='card-header'>🏢 Επιχείρηση </div>", unsafe_allow_html=True)
                if st.button("🏢 Προμηθευτές / Εταιρείες", use_container_width=True, key="btn_suppliers"):
                    st.session_state.page = "corp_suppliers_page"
                    st.rerun()
                if st.button("💸 Καταχώρηση Εξόδων", use_container_width=True, key="btn_corp_pay"):
                    st.session_state.page = "corp_payments"
                    st.rerun()
                if st.button("📊 Προβολή Εξόδων", use_container_width=True, key="btn_corp_rep"):
                    st.session_state.page = "corp_payments_report"
                    st.rerun()
                if st.button("📋 Μητρώο Πελατών", use_container_width=True, key="btn_anafora"):
                    st.session_state.page = "anafora_pelaton"
                    st.rerun()
                if st.button("📊 Ταμείο", use_container_width=True):
                    st.session_state.page = "db_final_page"
                    st.rerun()    

        st.markdown("<br>", unsafe_allow_html=True)
        
        # Μικρότερο ξεχωριστό τμήμα για τα παράπονα/σχόλια στο κάτω μέρος
        with st.container(border=True):
            col_ex1, col_ex2 = st.columns([3, 1])
            with col_ex1:
                st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
                st.write("Καταγραφή και παρακολούθηση σχολίων, παραπόνων ή παρατηρήσεων πελατών.")
            with col_ex2:
                st.write("")
                if st.button("🛠️ Μετάβαση", use_container_width=True, key="btn_complaints"):
                    st.session_state.page = "complaints_page"
                    st.rerun()
        
        # Ξεχωριστό τμήμα για τους Pre-Register μαθητές στο κάτω μέρος
        with st.container(border=True):
            col_pr_ex1, col_pr_ex2 = st.columns([3, 1])
            with col_pr_ex1:
                st.markdown("##### 📝 Διαχείριση Pre-Register Μαθητών")
                st.write("Προβολή και παρακολούθηση μαθητών σε κατάσταση προεγγραφής ανά σχολικό έτος.")
            with col_pr_ex2:
                st.write("")
                if st.button("🛠️ Μετάβαση", use_container_width=True, key="btn_prereg_bottom"):
                    st.session_state.page = "pre_register_page"
                    st.rerun()
 #----------------------------------------------------------------------   
    # --- ΣΕΛΙΔΑ: ΠΑΡΟΥΣΙΟΛΟΓΙΟ (ΤΜΗΜΑ: ΙΔΙΩΤΙΚΑ) ---
    elif st.session_state.page == "attendances_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun()

        st.subheader("✅ Παρουσιολόγιο")
        
        years_options = ["2026-2027", "2027-2028", "2028-2029", "2029-2030"]
        selected_year_att = st.selectbox("📅 Επιλογή Σχολικού Έτους", options=years_options)
        
        selected_date_obj = st.date_input("📅 Ημερομηνία", value=datetime.today())
        target_date_str = selected_date_obj.strftime('%Y-%m-%d')
        st.info(f"Επιλεγμένη Ημερομηνία Καταγραφής: **{target_date_str}** | Σχολικό Έτος: **{selected_year_att}**")

        regs_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")
        att_df = load_specific_db("db_attendances")

        if regs_df.empty or cust_df.empty:
            st.warning("⚠️ Δεν βρέθηκαν δεδομένα εγγραφών ή πελατών.")
        else:
            regs_df.columns = regs_df.columns.astype(str).str.strip()
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            
            priv_regs = regs_df[
                (regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά') & 
                (regs_df['school_year'].astype(str).str.strip() == selected_year_att)
            ].copy()

            if priv_regs.empty:
                st.info(f"Δεν βρέθηκαν μαθητές στο τμήμα για το σχολικό έτος {selected_year_att}.")
            else:
                target_date_att = pd.DataFrame()
                if not att_df.empty and 'date' in att_df.columns:
                    att_df.columns = att_df.columns.astype(str).str.strip()
                    target_date_att = att_df[
                        (att_df['date'].astype(str).str.strip() == target_date_str) & 
                        (att_df['school_year'].astype(str).str.strip() == selected_year_att)
                    ]

                with st.form("attendance_form"):
                    st.markdown(f"#### Λίστα Μαθητών - {target_date_str}")
                    attendance_inputs = {}
                    pdf_preview_rows = []

                    for idx, r in priv_regs.iterrows():
                        c_id = str(r['id_customer']).strip()
                        fname = str(r['customer_firstname']).strip()
                        lname = str(r['customer_lastname']).strip()
                        prog_name = str(r.get('program', '')).strip()
                        reg_id = str(r.get('id_registration', '')).strip()
                        
                        adj_cost_val = 0.0
                        try:
                            adj_str = str(r.get('adjusted_cost', 0)).replace(',', '.')
                            adj_cost_val = float(adj_str) if adj_str.strip() != '' else 0.0
                        except:
                            adj_cost_val = 0.0

                        existing_rec = pd.DataFrame()
                        if not target_date_att.empty:
                            existing_rec = target_date_att[target_date_att['id_customer'].astype(str).str.strip() == c_id]
                        
                        default_val = False
                        if not existing_rec.empty:
                            if str(existing_rec.iloc[0].get('status', '')).strip().upper() == 'ΝΑΙ':
                                default_val = True

                        col1, col2, col3 = st.columns([3, 2, 2])
                        col1.write(f"**{lname} {fname}** (ID: {c_id})")
                        col2.write(f"Πρόγραμμα: *{prog_name}* | Κόστος/Μάθημα: {adj_cost_val:.2f}€")
                        val = col3.toggle("ΝΑΙ / ΟΧΙ", value=default_val, key=f"att_{c_id}_{target_date_str}")
                        
                        reg_row_index = idx + 2

                        attendance_inputs[c_id] = {
                            'val': val,
                            'reg_row_index': reg_row_index,
                            'adj_cost': adj_cost_val,
                            'prog': prog_name,
                            'fname': fname,
                            'lname': lname
                        }

                        pdf_preview_rows.append({
                            'name': f"{lname} {fname} (ID: {c_id})",
                            'prog': prog_name,
                            'status': "ΝΑΙ" if val else "ΟΧΙ"
                        })

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
                            all_reg_values = sh_reg.get_all_values()

                            rows_to_append_or_update = []
                            for c_id, data in attendance_inputs.items():
                                if data['val']: 
                                    rows_to_append_or_update.append((c_id, data))

                            for c_id, data in rows_to_append_or_update:
                                found_att = False
                                already_had_nai = False  # Έλεγχος αν είχε ήδη ΝΑΙ στη βάση για σήμερα
                                
                                if att_records:
                                    for ai, a_row in enumerate(att_records, start=2):
                                        if str(a_row.get('id_customer', '')).strip() == c_id and str(a_row.get('date', '')).strip() == target_date_str and str(a_row.get('school_year', '')).strip() == selected_year_att:
                                            if str(a_row.get('status', '')).strip().upper() == 'ΝΑΙ':
                                                already_had_nai = True
                                            sh_att.update_cell(ai, 6, "ΝΑΙ")
                                            found_att = True
                                            break
                                            
                                if not found_att:
                                    sh_att.append_row([next_att_id, c_id, f"{data['lname']} {data['fname']}", selected_year_att, target_date_str, "ΝΑΙ"])
                                    next_att_id += 1

                                # Υπολογισμός χρέους ΜΟΝΟ αν ο μαθητής ΔΕΝ είχε ήδη καταχωρηθεί ως ΝΑΙ σήμερα
                                if not already_had_nai:
                                    r_idx = data['reg_row_index']
                                    reg_row_data = all_reg_values[r_idx - 1]
                                    
                                    energy_type = str(reg_row_data[10]).strip() if len(reg_row_data) > 10 else ""
                                    
                                    if energy_type.lower() == "ανοιχτή πληρωμή":
                                        current_xreos = 0.0
                                        try:
                                            current_xreos = float(str(reg_row_data[11]).replace(',', '.') or 0.0)
                                        except:
                                            current_xreos = 0.0

                                        new_xreos = current_xreos + data['adj_cost']
                                        sh_reg.update_cell(r_idx, 12, new_xreos)

                            st.success("✅ Το παρουσιολόγιο αποθηκεύτηκε επιτυχώς (χωρίς διπλοχρεώσεις)!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα αποθήκευσης παρουσιών: {e}")

                st.markdown("<br>", unsafe_allow_html=True)
                pdf_bytes = generate_single_attendance_pdf(pdf_preview_rows, target_date_str, selected_year_att)
                if pdf_bytes:
                    st.download_button(
                        label="📥 Λήψη Παρουσιολογίου",
                        data=pdf_bytes,
                        file_name=f"parousiologio_{target_date_str}.pdf",
                        mime="application/pdf"
                    )

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
                
                # --- ΕΠΙΛΟΓΗ ΤΜΗΜΑΤΟΣ (ΜΠΑΙΝΕΙ ΠΡΙΝ ΤΑ ΥΠΟΛΟΙΠΑ ΓΙΑ ΝΑ ΕΛΕΓΧΕΙ ΤΗΝ ΕΜΦΑΝΙΣΗ) ---
                class_options = ["Xωρίς Τμήμα"] + (classes_df['class_name'].unique().tolist() if not classes_df.empty else [])
                
                default_class_idx = 0
                if reg_to_edit is not None and 'class' in reg_to_edit:
                    cur_cls = str(reg_to_edit['class']).strip()
                    if cur_cls in class_options:
                        default_class_idx = class_options.index(cur_cls)
                        
                sel_class = st.selectbox("🏫 Τμήμα", options=class_options, index=default_class_idx)
                
                # Έλεγχος τύπου τμήματος (Voucher ή Ιδιωτικά)
                is_voucher = "voucher" in sel_class.lower()
                is_private = "ιδιωτικά" in sel_class.lower() or "ιδιωτικα" in sel_class.lower()

                # --- ΕΜΦΑΝΙΣΗ ΒΑΡΔΙΑΣ (ΜΟΝΟ ΑΝ ΕΠΙΛΕΓΕΙ VOUCHER) ---
                sel_bardia = ""
                if is_voucher:
                    bardia_options = ["Α Βάρδια", "Β Βάρδια"]
                    def_bardia_idx = 0
                    if reg_to_edit is not None and 'm_bardia' in reg_to_edit:
                        val_bar = str(reg_to_edit['m_bardia']).strip()
                        if val_bar in bardia_options:
                            def_bardia_idx = bardia_options.index(val_bar)
                    sel_bardia = st.selectbox("🕒 Επιλογή Βάρδιας", options=bardia_options, index=def_bardia_idx)

                # --- ΕΜΦΑΝΙΣΗ ΣΥΜΦΩΝΗΘΕΝΤΟΣ & ΤΥΠΟΥ ΠΛΗΡΩΜΗΣ (ΜΟΝΟ ΑΝ ΕΠΙΛΕΓΕΙ ΙΔΙΩΤΙΚΑ) ---
                adj_cost = 0.0
                sel_energy = ""
                if is_private:
                    default_cost = float(str(reg_to_edit.get('adjusted_cost', 0)).replace(',', '.') or 0) if reg_to_edit is not None else 0.0
                    adj_cost = st.number_input("💰 Συμφωνηθέν Ποσό (€)", value=default_cost, step=10.0)
                    
                    energy_options = ["Ανοιχτή πληρωμή", "Κλειστή πληρωμή"]
                    def_energy_idx = 0
                    if reg_to_edit is not None and 'energy' in reg_to_edit:
                        val_en = str(reg_to_edit['energy']).strip()
                        if val_en in energy_options:
                            def_energy_idx = energy_options.index(val_en)
                    sel_energy = st.selectbox("💡 Τύπος Πληρωμής", options=energy_options, index=def_energy_idx)

                default_start_date = datetime.today()
                if reg_to_edit is not None and 'start_date' in reg_to_edit:
                    try:
                        parsed_date = pd.to_datetime(reg_to_edit['start_date']).date()
                        if not pd.isna(parsed_date):
                            default_start_date = parsed_date
                    except:
                        pass
                sel_start_date = st.date_input("📅 Ημερομηνία Έναρξης", value=default_start_date)
                start_date_str = sel_start_date.strftime('%Y-%m-%d')
                
                if st.button("🚀 Οριστικοποίηση Εγγραφής"):
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                        
                        if mode == "Νέα Εγγραφή":
                            records = sh.get_all_records()
                            next_reg = int(pd.DataFrame(records)['id_registration'].max()) + 1 if records and 'id_registration' in pd.DataFrame(records).columns else 1
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
                                0.0,
                                str(sel_bardia),
                                str(start_date_str)
                            ]
                            sh.append_row(row_data)
                        else:
                            target_reg_id = str(reg_to_edit['id_registration']).strip()
                            all_rows = sh.get_all_values()
                            
                            for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                if str(s_row[0]).strip() == target_reg_id:
                                    sh.update_cell(r_idx, 5, str(s_year))
                                    sh.update_cell(r_idx, 6, str(s_prog))
                                    sh.update_cell(r_idx, 7, str(sel_class))
                                    sh.update_cell(r_idx, 10, adj_cost)
                                    sh.update_cell(r_idx, 11, sel_energy)
                                    sh.update_cell(r_idx, 13, str(sel_bardia))
                                    sh.update_cell(r_idx, 14, str(start_date_str))
                                    break
                                
                        st.success("✅ Η εγγραφή/τροποποίηση καταχωρήθηκε επιτυχώς!")
                        time.sleep(1.5)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

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
                tax_user = st.text_input("Taxis Username").strip()
                tax_pass = st.text_input("Taxis Password").strip()
                adt = st.text_input("Ταυτότητα").strip()
                afm = st.text_input("ΑΦΜ").strip()

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
                        tax_u = st.text_input("Taxis Username", value=str(row.get('Taxis_User', '')))
                        tax_p = st.text_input("Taxis Password", value=str(row.get('Taxis_Pass', '')))
                        adt_val = st.text_input("Ταυτότητα", value=str(row.get('ADT', row.get('Passport', ''))))
                        afm_val = st.text_input("ΑΦΜ", value=str(row.get('AFM', '')))

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
                                time.sleep(1.3)
                                st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα ενημέρωσης: {e}")

    # --- ΣΕΛΙΔΑ: ΔΙΑΧΕΙΡΙΣΗ STOP ---
    elif st.session_state.page == "view_customers":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun()

        st.subheader("Διαχείριση Εγγραφών")
        reg_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")

        if reg_df.empty:
            st.info("Δεν βρέθηκαν εγγραφές στο `db_registrations`.")
        else:
            reg_df.columns = reg_df.columns.astype(str).str.strip()
            
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
                            new_status = st.selectbox("📌 Κατάσταση", options=status_options, index=def_status_idx)
                            new_cost = st.number_input("💰 Συμφωνηθέν Ποσό (€)", value=current_cost, step=10.0, format="%.2f")
                            
                            new_stop_date = st.date_input("📅 Ημερομηνία Σταματήματος", value=parsed_date)
                            
                            if st.form_submit_button("💾 Αποθήκευση Αλλαγών"):
                                try:
                                    gc = get_gspread_client()
                                    sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                    all_rows = sh_reg.get_all_values()
                                    
                                    stop_date_str = new_stop_date.strftime('%Y-%m-%d') if new_status.lower() == 'stop' else ""
                                    
                                    for r_idx, s_row in enumerate(all_rows[1:], start=2):
                                        if str(s_row[0]).strip() == target_reg_id:
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

    # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ & ΕΠΕΞΕΡΓΑΣΙΑ ΠΛΗΡΩΜΗΣ ΣΥΝΔΡΟΜΗΣ ---
    elif st.session_state.page == "add_payment":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
            
        st.subheader("Καταχώρηση & Επεξεργασία Πληρωμής Συνδρομής (Ιδιωτικά)")
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        pays_df = load_specific_db("db_payments")
        
        if cust_df.empty or reg_df.empty:
            st.warning("⚠️ Δεν βρέθηκαν δεδομένα πελατών ή εγγραφών.")
        else:
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            reg_df.columns = reg_df.columns.astype(str).str.strip()
            
            # Φιλτράρισμα ΜΟΝΟ των μαθητών που έχουν εγγραφή στο τμήμα "Ιδιωτικά"
            priv_regs = reg_df[reg_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά']
            valid_customer_ids = priv_regs['id_customer'].astype(str).str.strip().unique().tolist()
            
            priv_cust_df = cust_df[cust_df['id_customer'].astype(str).str.strip().isin(valid_customer_ids)].copy()
            
            if priv_cust_df.empty:
                st.info("Δεν βρέθηκαν μαθητές στο τμήμα Ιδιωτικά.")
            else:
                if 'customer_lastname' in priv_cust_df.columns and 'customer_firstname' in priv_cust_df.columns:
                    priv_cust_df = priv_cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])

                priv_cust_df['display'] = priv_cust_df['customer_lastname'].astype(str) + " " + priv_cust_df['customer_firstname'].astype(str) + " (ID: " + priv_cust_df['id_customer'].astype(str) + ")"
                sel = st.selectbox("🔍 Επιλογή Πελάτη (Ιδιωτικά)", options=["-- Επιλέξτε --"] + priv_cust_df['display'].tolist())
                
                if sel != "-- Επιλέξτε --":
                    row = priv_cust_df[priv_cust_df['display'] == sel].iloc[0]
                    c_id = str(row['id_customer']).strip()
                    
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
                                                    sh_pay.update_cell(pr_idx, 6, edit_rec)
                                                    sh_pay.update_cell(pr_idx, 7, edit_amt)
                                                    sh_pay.update_cell(pr_idx, 8, edit_yr)
                                                    sh_pay.update_cell(pr_idx, 9, edit_method)
                                                    break
                                            
                                            update_customer_xreos(c_id, edit_yr, "")
                                            
                                            st.success("✅ Η πληρωμή ενημερώθηκε επιτυχώς!")
                                            time.sleep(1.2)
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Σφάλμα ενημέρωσης πληρωμής: {e}")
                    
                    else:
                        user_regs = reg_df[(reg_df['id_customer'].astype(str).str.strip() == c_id) & (reg_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά')]

                        if user_regs.empty:
                            st.warning("⚠️ Ο συγκεκριμένος πελάτης δεν έχει ενεργή εγγραφή στο τμήμα Ιδιωτικά.")
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

                            st.info(f"💰 **Συμφωνηθέν Ποσό:** {adjusted_val:.2f}€ | 📉 **Υπόλοιπο:** {xreos_val:.2f}€")

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
                                                nid,                            # 1: id_payment
                                                int(c_id),                      # 2: id_customer
                                                str(row['customer_lastname']),    # 3: customer_lastname
                                                str(row['customer_firstname']),   # 4: customer_firstname
                                                datetime.today().strftime('%Y-%m-%d'), # 5: payment_date
                                                rec,                            # 6: receipt_number
                                                amt,                            # 7: amount_paid
                                                str(yr),                        # 8 (Στήλη H): school_year
                                                pay_method                      # 9: pay_method
                                            ])
                                            
                                            update_customer_xreos(c_id, yr, prog)
                                            
                                            st.success("✅ Η πληρωμή καταχωρήθηκε επιτυχώς!")
                                            time.sleep(1.2)
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"Σφάλμα κατά την αποθήκευση: {e}")

    # --- ΣΕΛΙΔΑ: ΑΝΑΦΟΡΑ & ΙΣΤΟΡΙΚΟ ΠΛΗΡΩΜΩΝ ΜΑΘΗΤΗ ---
    elif st.session_state.page == "payment_history_report":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun()

        st.subheader("📊 Αναφορά Πληρωμών")
        
        pays_df = load_specific_db("db_payments")
        cust_df = load_specific_db("db_customers")
        regs_df = load_specific_db("db_registrations")

        st.markdown("### 📋 Συγκεντρωτικός Πίνακας Πληρωμών")
        
        all_years_opt = ["ΟΛΑ"] + (sorted(regs_df['school_year'].dropna().astype(str).unique().tolist(), reverse=True) if not regs_df.empty else ["2026-2027"])
        sel_sum_year = st.selectbox("📅 Επιλογή Σχολικού Έτους", options=all_years_opt, key="sum_year_filter")

        if regs_df.empty or cust_df.empty:
            st.info("Δεν βρέθηκαν δεδομένα εγγραφών ή πελατών.")
        else:
            regs_df.columns = regs_df.columns.astype(str).str.strip()
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            if not pays_df.empty:
                pays_df.columns = pays_df.columns.astype(str).str.strip()

            filtered_regs = regs_df[regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά'].copy()
            if sel_sum_year != "ΟΛΑ":
                filtered_regs = filtered_regs[filtered_regs['school_year'].astype(str).str.strip() == sel_sum_year]

            summary_rows = []
            for _, reg in filtered_regs.iterrows():
                c_id = str(reg.get('id_customer', '')).strip()
                s_year = str(reg.get('school_year', '')).strip()
                u_class = str(reg.get('class', '')).strip()
                
                c_match = cust_df[cust_df['id_customer'].astype(str).str.strip() == c_id]
                if not c_match.empty:
                    c_row = c_match.iloc[0]
                    full_name = f"{c_row.get('customer_lastname', '')} {c_row.get('customer_firstname', '')}"
                else:
                    full_name = f"Άγνωστος (ID: {c_id})"

                paid_sum = 0.0
                if not pays_df.empty and 'id_customer' in pays_df.columns and 'amount_paid' in pays_df.columns:
                    p_match = pays_df[
                        (pays_df['id_customer'].astype(str).str.strip() == c_id) & 
                        (pays_df['school_year'].astype(str).str.strip() == s_year)
                    ] if 'school_year' in pays_df.columns else pays_df[pays_df['id_customer'].astype(str).str.strip() == c_id]
                    
                    paid_sum = pd.to_numeric(p_match['amount_paid'], errors='coerce').fillna(0).sum()

                xreos_raw = reg.iloc[11] if len(reg) > 11 else reg.get('xreos', 0)
                try:
                    xreos_val = float(str(xreos_raw).replace(',', '.') or 0.0)
                except:
                    xreos_val = 0.0

                summary_rows.append({
                    'id_customer': c_id,
                    'Ονοματεπώνυμο': full_name,
                    'Class': u_class,
                    'Έτος': s_year,
                    'Πληρωμένα (€)': paid_sum,
                    'Υπόλοιπο (€)': xreos_val
                })

            summary_df = pd.DataFrame(summary_rows)
            if summary_df.empty:
                st.info("Δεν βρέθηκαν δεδομένα μαθητών 'Ιδιωτικά' με τα επιλεγμένα φίλτρα.")
            else:
                summary_df = summary_df.sort_values(by='Ονοματεπώνυμο', ascending=True)
                st.dataframe(summary_df[['Ονοματεπώνυμο', 'Class', 'Έτος', 'Πληρωμένα (€)', 'Υπόλοιπο (€)']], use_container_width=True, hide_index=True)

                tot_paid = summary_df['Πληρωμένα (€)'].sum()
                tot_xreos = summary_df['Υπόλοιπο (€)'].sum()

                st.markdown("<br>", unsafe_allow_html=True)
                m1, m2 = st.columns(2)
                m1.metric("💳 Συνολικά Πληρωμένα", f"{tot_paid:.2f} €")
                m2.metric("⚠️ Συνολικό Υπόλοιπο", f"{tot_xreos:.2f} €")
                st.markdown("<br>", unsafe_allow_html=True)

                if st.button("📥 Εκτύπωση Συγκεντρωτικού Πίνακα"):
                    pdf_sum_buf = generate_summary_payments_pdf(summary_df, sel_sum_year, "Ιδιωτικά")
                    if pdf_sum_buf:
                        st.download_button(
                            label="📥 Κατεβάστε το Συγκεντρωτικό PDF",
                            data=pdf_sum_buf,
                            file_name="sygentrotikos_oikonomikos_idiotika.pdf",
                            mime="application/pdf"
                        )

        st.markdown("---")
        st.markdown("### 🔍 Ιστορικό Πληρωμών /Μαθητή ")

        if cust_df.empty:
            st.info("Δεν βρέθηκαν δεδομένα πελατών.")
        else:
            priv_customer_ids = regs_df[regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά']['id_customer'].astype(str).unique()
            priv_cust_df = cust_df[cust_df['id_customer'].astype(str).isin(priv_customer_ids)].copy()

            if priv_cust_df.empty:
                st.info("Δεν υπάρχουν μαθητές στο τμήμα 'Ιδιωτικά'.")
            else:
                if 'customer_lastname' in priv_cust_df.columns and 'customer_firstname' in priv_cust_df.columns:
                    priv_cust_df = priv_cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])

                priv_cust_df['display'] = priv_cust_df['customer_lastname'].astype(str) + " " + priv_cust_df['customer_firstname'].astype(str) + " (ID: " + priv_cust_df['id_customer'].astype(str) + ")"
                sel_cust = st.selectbox("🔍 Επιλογή Μαθητή για Ιστορικό", options=["-- Επιλέξτε --"] + priv_cust_df['display'].tolist())

                if sel_cust != "-- Επιλέξτε --":
                    row_c = priv_cust_df[priv_cust_df['display'] == sel_cust].iloc[0]
                    c_id = str(row_c['id_customer']).strip()
                    cust_full_name = f"{row_c.get('customer_lastname', '')} {row_c.get('customer_firstname', '')}"

                    user_regs = pd.DataFrame()
                    if not regs_df.empty:
                        user_regs = regs_df[(regs_df['id_customer'].astype(str).str.strip() == c_id) & (regs_df['class'].astype(str).str.strip().str.lower() == 'ιδιωτικά')]

                    st.markdown(f"#### ⚙️ Πληρωμές: **{cust_full_name}**")
                    ind_years = ["ΟΛΑ"] + (sorted(user_regs['school_year'].dropna().astype(str).unique().tolist(), reverse=True) if not user_regs.empty else ["2026-2027"])
                    sel_ind_year = st.selectbox("📅 Έτος Μαθητή", options=ind_years, key="ind_year")

                    if pays_df.empty:
                        st.info("Δεν υπάρχουν καταχωρημένες πληρωμές στη βάση.")
                    else:
                        cust_pays = pays_df[pays_df['id_customer'].astype(str).str.strip() == c_id].copy()

                        if cust_pays.empty:
                            st.info(f"Ο/Η μαθητής-τρια {cust_full_name} δεν έχει καμία καταχωρημένη πληρωμή.")
                        else:
                            if sel_ind_year != "ΟΛΑ":
                                if 'school_year' in cust_pays.columns:
                                    cust_pays = cust_pays[cust_pays['school_year'].astype(str).str.strip() == sel_ind_year]

                            if cust_pays.empty:
                                st.info("Δεν βρέθηκαν πληρωμές με τα επιλεγμένα φίλτρα.")
                            else:
                                display_pays = pd.DataFrame({
                                    'Ημερομηνία': cust_pays.get('payment_date', ''),
                                    'Αρ. Απόδειξης': cust_pays.get('receipt_number', ''),
                                    'Έτος': cust_pays.get('school_year', ''),
                                    'Τρόπος': cust_pays.get('pay_method', ''),
                                    'Ποσό (€)': pd.to_numeric(cust_pays.get('amount_paid', 0), errors='coerce').fillna(0)
                                })

                                st.markdown(f"#### 📄 Ατομικό Ιστορικό Πληρωμών")
                                st.dataframe(display_pays, use_container_width=True, hide_index=True)

                                total_paid = display_pays['Ποσό (€)'].sum()
                                st.success(f"💰 **Συνολικό Ποσό Πληρωμών Μαθητή:** {total_paid:.2f} €")

                                st.markdown("<br>", unsafe_allow_html=True)
                                if st.button("📥 Εκτύπωση ατομικής αναφοράς"):
                                    pdf_buffer = generate_customer_payment_history_pdf(cust_full_name, sel_ind_year, display_pays)
                                    if pdf_buffer:
                                        st.download_button(
                                            label="📥 Κατεβάστε το PDF",
                                            data=pdf_buffer,
                                            file_name=f"istoriko_pliromon_{c_id}.pdf",
                                            mime="application/pdf"
                                        )

    # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ ΕΞΟΔΩΝ ---
    elif st.session_state.page == "corp_payments":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("💳 Καταχώρηση Εξόδων")
        
        suppliers_list = []
        company_dict = {}
        try:
            gc = get_gspread_client()
            comp_df = pd.DataFrame(gc.open(OLD_SHEET_NAME).worksheet("db_customers_company").get_all_records())
            if not comp_df.empty:
                comp_df.columns = comp_df.columns.astype(str).str.strip()
                for _, r in comp_df.iterrows():
                    c_id = str(r.get('id', '')).strip()
                    c_name = str(r.get('company_name', '')).strip()
                    if c_name:
                        suppliers_list.append(c_name)
                        company_dict[c_name] = c_id
        except:
            pass

        if not suppliers_list:
            st.warning("⚠️ Δεν βρέθηκαν καταχωρημένες εταιρείες. Παρακαλώ καταχωρήστε πρώτα προμηθευτή.")
            if st.button("➕ Μετάβαση σε Καταχώρηση Εταιρειών"):
                st.session_state.page = "corp_suppliers_page"
                st.rerun()
        else:
            with st.form("exp_form"):
                inv = st.text_input("Αρ. Παραστατικού (Invoice Number)")
                sup_name = st.selectbox("Προμηθευτής / Εταιρεία", options=suppliers_list)
                school_yr = st.selectbox("Σχολικό Έτος", options=["2026-2027", "2027-2028", "2028-2029", "2029-2030"])
                inv_date = st.date_input("Ημερομηνία Παραστατικού", value=datetime.today())
                pay_date = st.date_input("Ημερομηνία Πληρωμής", value=datetime.today())
                amt = st.number_input("Ποσό (€)", min_value=0.0, format="%.2f")
                amount_to_save = float(amt)
                pay_method = st.selectbox("Τρόπος Πληρωμής", options=["Μετρητά", "Κατάθεση", "POS"])

                if st.form_submit_button("💾 Αποθήκευση Εξόδου"):
                    try:
                        gc = get_gspread_client()
                        try:
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                        except:
                            sh = gc.open(OLD_SHEET_NAME).add_worksheet(title="db_payments_company", rows="1000", cols="12")
                            sh.append_row([
                                "id", "invoice_number", "type", "customer_id", "company_name", 
                                "school_year", "invoice_date", "payment_date", "entry_date", "amount_to_save", "pay_method"
                            ])

                        recs = sh.get_all_records()
                        nid = int(pd.DataFrame(recs)['id'].max()) + 1 if recs and 'id' in pd.DataFrame(recs).columns and not pd.DataFrame(recs).empty else 1
                        
                        selected_company_id = company_dict.get(sup_name, "")
                        entry_date_str = datetime.today().strftime('%Y-%m-%d')
                        inv_date_str = inv_date.strftime('%Y-%m-%d')
                        pay_date_str = pay_date.strftime('%Y-%m-%d')

                        row_data = [
                            nid, inv, "Έξοδα", selected_company_id, sup_name, 
                            school_yr, inv_date_str, pay_date_str, entry_date_str, amt, pay_method
                        ]
                        
                        sh.append_row(row_data)
                        st.success("✅ Το έξοδο αποθηκεύτηκε επιτυχώς!")
                        time.sleep(1.1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα αποθήκευσης εξόδου: {e}")

    # --- ΣΕΛΙΔΑ: ΠΡΟΒΟΛΗ & ΦΙΛΤΡΑ ΕΞΟΔΩΝ ---
    elif st.session_state.page == "corp_payments_report":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("📊 Πίνακας Εξόδων Επιχείρησης")
        
        try:
            gc = get_gspread_client()
            worksheet = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
            # Χρησιμοποιούμε value_render_option='UNFORMATTED_VALUE' αν υποστηρίζεται ή απλά get_all_records()
            exp_df = pd.DataFrame(worksheet.get_all_records())
        except:
            exp_df = pd.DataFrame()

        if exp_df.empty:
            st.info("Δεν βρέθηκαν καταχωρημένα έξοδα στη βάση `db_payments_company`.")
        else:
            exp_df.columns = exp_df.columns.astype(str).str.strip()

            # Απόλυτα ασφαλής μετατροπή ποσού χωρίς λάθη δεκαδικών
            if 'amount' in exp_df.columns:
                def exact_amount(val):
                    if pd.isna(val) or str(val).strip() == '':
                        return 0.0
                    
                    # Αν έρθε απευθείας σαν αριθμός από το Google Sheets
                    if isinstance(val, (int, float)):
                        return float(val)
                        
                    val_str = str(val).strip().replace('€', '').strip()
                    
                    # Αντικατάσταση κόμματος με τελεία
                    val_str = val_str.replace(',', '.')
                    
                    try:
                        res = float(val_str)
                        return res
                    except:
                        return 0.0

                exp_df['amount_numeric'] = exp_df['amount'].apply(exact_amount)

            col_f1, col_f2 = st.columns(2)
            available_years = ["ΟΛΑ"] + sorted(exp_df['school_year'].dropna().astype(str).unique().tolist(), reverse=True) if 'school_year' in exp_df.columns else ["ΟΛΑ"]
            sel_year_exp = col_f1.selectbox("📅 Φίλτρο ανά Σχολικό Έτος", options=available_years, key="exp_year_filter")

            available_suppliers = ["ΟΛΑ"] + sorted(exp_df['company_name'].dropna().astype(str).unique().tolist()) if 'company_name' in exp_df.columns else ["ΟΛΑ"]
            sel_sup_exp = col_f2.selectbox("🏢 Επιλογή Προμηθευτή", options=available_suppliers, key="exp_sup_filter")

            filtered_exp = exp_df.copy()
            if sel_year_exp != "ΟΛΑ":
                filtered_exp = filtered_exp[filtered_exp['school_year'].astype(str).str.strip() == sel_year_exp]
            if sel_sup_exp != "ΟΛΑ":
                filtered_exp = filtered_exp[filtered_exp['company_name'].astype(str).str.strip() == sel_sup_exp]

            st.markdown(f"**Βρέθηκαν έξοδα:** {len(filtered_exp)}")

            if filtered_exp.empty:
                st.info("Δεν βρέθηκαν έξοδα με τα επιλεγμένα φίλτρα.")
            else:
                display_exp_df = filtered_exp.copy()
                if 'amount_numeric' in display_exp_df.columns:
                    # Εμφάνιση με κόμμα στα δεκαδικά (π.χ. 5,20 €)
                    display_exp_df['amount'] = display_exp_df['amount_numeric'].apply(lambda x: f"{x:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    drop_cols = [c for c in ['amount_numeric'] if c in display_exp_df.columns]
                    display_exp_df = display_exp_df.drop(columns=drop_cols)

                st.dataframe(display_exp_df, use_container_width=True, hide_index=True)

                if 'amount_numeric' in filtered_exp.columns:
                    total_exp_amt = filtered_exp['amount_numeric'].sum()
                    formatted_total = f"{total_exp_amt:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.')
                    st.markdown(f"### 💰 **Συνολικό Ποσό Εξόδων:** `{formatted_total}`")
                    
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
                new_class_name = st.text_input("Όνομα Τμήματος").strip()
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
        st.markdown("#### Ενεργά Τμήματα")
        c_df_refreshed = load_specific_db("db_class")
        if not c_df_refreshed.empty:
            st.dataframe(c_df_refreshed, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν τμήματα.")

        st.markdown("---")
        st.markdown("#### Ενεργά Προγράμματα")
        if not progs_df.empty:
            st.dataframe(progs_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν προγράμματα.")

    # --- ΣΕΛΙΔΑ: ΜΗΤΡΩΟ ΠΕΛΑΤΩΝ ---
    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("📋 Μητρώο Μαθητών")
        reg_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")
        
        if not reg_df.empty:
            reg_df.columns = reg_df.columns.astype(str).str.strip()
            
            bardia_col_found = None
            for col in reg_df.columns:
                if col.lower() in ['m_bardia', 'm bardia', 'bardia', 'βάρδια', 'βαρδια']:
                    bardia_col_found = col
                    break
            if bardia_col_found and bardia_col_found != 'm_bardia':
                reg_df = reg_df.rename(columns={bardia_col_found: 'm_bardia'})

            if not cust_df.empty:
                cust_df.columns = cust_df.columns.astype(str).str.strip()
                if 'id_customer' in reg_df.columns and 'id_customer' in cust_df.columns:
                    reg_df['id_customer_str'] = reg_df['id_customer'].astype(str).str.strip()
                    cust_df['id_customer_str'] = cust_df['id_customer'].astype(str).str.strip()
                    
                    cols_to_use = ['id_customer_str', 'parent_firstname', 'parent_lastname', 'telephone', 'email']
                    existing_cust_cols = [c for c in cols_to_use if c in cust_df.columns]
                    
                    merged_df = pd.merge(
                        reg_df, 
                        cust_df[existing_cust_cols], 
                        on='id_customer_str', 
                        how='left'
                    )
                else:
                    merged_df = reg_df.copy()
            else:
                merged_df = reg_df.copy()

            col_f1, col_f2 = st.columns(2)
            available_years = ["ΟΛΑ"] + sorted(merged_df['school_year'].dropna().astype(str).unique().tolist(), reverse=True)
            sel_year_filter = col_f1.selectbox("📅 Σχολικό Έτος", options=available_years, key="anafora_year_filter")
            
            # Προσθήκη του "ΟΛΑ" ως default επιλογή στα τμήματα
            available_classes = ["ΟΛΑ"] + sorted(merged_df['class'].dropna().astype(str).unique().tolist())
            sel_class_filter = col_f2.selectbox("🏫 Επιλογή Τμήματος", options=available_classes, index=0, key="anafora_class_filter")
            
            filtered_reg = merged_df.copy()
            if sel_year_filter != "ΟΛΑ":
                filtered_reg = filtered_reg[filtered_reg['school_year'].astype(str).str.strip() == sel_year_filter]
            
            if sel_class_filter != "ΟΛΑ":
                filtered_reg = filtered_reg[filtered_reg['class'].astype(str).str.strip().str.lower() == sel_class_filter.strip().lower()]

            st.markdown(f"**Τμήμα:** `{sel_class_filter}` | **Βρέθηκαν εγγραφές:** {len(filtered_reg)}")

            if filtered_reg.empty:
                st.info("Δεν βρέθηκαν εγγραφές με τα επιλεγμένα κριτήρια.")
            else:
                class_lower = sel_class_filter.strip().lower()
                
                # Ορισμός επιθυμητών στηλών ανάλογα με το επιλεγμένο τμήμα
                if class_lower == "ιδιωτικά":
                    desired_cols = [
                        'customer_firstname', 'customer_lastname', 
                        'parent_firstname', 'parent_lastname', 
                        'telephone', 'email', 
                        'adjusted_cost', 'energy', 'xreos'
                    ]
                elif class_lower == "voucher":
                    desired_cols = [
                        'customer_firstname', 'customer_lastname', 
                        'parent_firstname', 'parent_lastname', 
                        'telephone', 'email', 
                        'status', 'stop_date', 'm_bardia'
                    ]
                elif class_lower == "pre-register":
                    desired_cols = [
                        'id_customer', 'customer_firstname', 'customer_lastname', 
                        'parent_firstname', 'parent_lastname', 
                        'class', 'school_year', 'telephone', 'email'
                    ]
                else: # Επιλογή "ΟΛΑ"
                    desired_cols = [
                        'id_customer', 'customer_firstname', 'customer_lastname', 
                        'class', 'school_year', 'telephone', 'email'
                    ]

                existing_cols = [c for c in desired_cols if c in filtered_reg.columns]
                grid_df = filtered_reg[existing_cols].copy()

                renames = {
                    'id_customer': 'ID Πελάτη',
                    'customer_lastname': 'Επώνυμο Μαθητή',
                    'customer_firstname': 'Όνομα Μαθητή',
                    'parent_firstname': 'Όνομα Γονέα',
                    'parent_lastname': 'Επώνυμο Γονέα',
                    'class': 'Τμήμα',
                    'school_year': 'Σχολικό Έτος',
                    'telephone': 'Τηλέφωνο',
                    'email': 'Email',
                    'status': 'Κατάσταση',
                    'stop_date': 'Ημ. Λήξης',
                    'm_bardia': 'Βάρδια',
                    'adjusted_cost': 'Συμφωνηθέν Ποσό',
                    'energy': 'Τύπος Χρέωσης',
                    'xreos': 'Υπόλοιπο (Χρέος)'
                }
                
                grid_df = grid_df.rename(columns=renames)
                st.dataframe(grid_df, use_container_width=True, hide_index=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("📥 Παραγωγή PDF Μητρώου"):
                    rep_pdf = generate_report_pdf_customers(filtered_reg, selected_year=sel_year_filter, selected_class=sel_class_filter)
                    if rep_pdf:
                        st.download_button("📥 Κατεβάστε το αρχείο PDF", data=rep_pdf, file_name=f"mitroo_mathiton_{class_lower}.pdf", mime="application/pdf")
        else:
            st.info("Δεν βρέθηκαν εγγραφές στη βάση `db_registrations`.")

        # --- ΑΝΑΛΥΤΙΚΟ ΙΣΤΟΡΙΚΟ ΣΧΟΛΙΩΝ / ΠΑΡΑΠΟΝΩΝ ΑΝΑ ΜΑΘΗΤΗ ---
        st.markdown("---")
        st.subheader("💬 Αναλυτικό Ιστορικό Σχολίων & Παραπόνων Μαθητή")
        
        if not cust_df.empty:
            if 'customer_lastname' in cust_df.columns and 'customer_firstname' in cust_df.columns:
                cust_sorted = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])
            else:
                cust_sorted = cust_df.copy()
                
            cust_sorted['display_hist'] = cust_sorted['customer_lastname'].astype(str) + " " + cust_sorted['customer_firstname'].astype(str) + " (ID: " + cust_sorted['id_customer'].astype(str) + ")"
            sel_hist_cust = st.selectbox("🔍 Επιλογή Μαθητή για Προβολή Ιστορικού Σχολίων", options=["-- Επιλέξτε --"] + cust_sorted['display_hist'].tolist(), key="hist_cust_sel")
            
            if sel_hist_cust != "-- Επιλέξτε --":
                row_h = cust_sorted[cust_sorted['display_hist'] == sel_hist_cust].iloc[0]
                h_id = str(row_h['id_customer']).strip()
                
                comp_all_df = load_specific_db("db_complaint")
                if comp_all_df.empty:
                    st.info("Δεν υπάρχουν καταχωρημένα σχόλια ή παράπονα στη βάση.")
                else:
                    comp_all_df.columns = comp_all_df.columns.astype(str).str.strip()
                    student_comps = comp_all_df[comp_all_df['id_customer'].astype(str).str.strip() == h_id].copy()
                    
                    if student_comps.empty:
                        st.info("Ο συγκεκριμένος μαθητής δεν έχει καμία καταχώρηση σε σχόλια ή παράπονα.")
                    else:
                        st.markdown(f"**Βρέθηκαν συνολικά {len(student_comps)} εγγραφές (Ανοιχτά & Κλειστά):**")
                        
                        renames_h = {
                            'complaint_date': 'Ημερομηνία',
                            'choice': 'Τύπος',
                            'complaint_subject': 'Θέμα',
                            'complaint_text': 'Κείμενο',
                            'complaint_staff': 'Υπεύθυνος',
                            'status': 'Κατάσταση'
                        }
                        display_comp_df = student_comps.drop(columns=[c for c in ['id_complaint', 'id_customer', 'customer_firstname', 'customer_lastname'] if c in student_comps.columns]).rename(columns=renames_h)
                        st.dataframe(display_comp_df, use_container_width=True, hide_index=True)

    # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ ΕΤΑΙΡΕΙΩΝ / ΠΡΟΜΗΘΕΥΤΩΝ ---
    elif st.session_state.page == "corp_suppliers_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("🏢 Καταχώρηση Νέας Εταιρείας / Προμηθευτή")
        
        with st.form("supplier_form"):
            company_name = st.text_input("Επωνυμία Εταιρείας")
            vat_number = st.text_input("ΑΦΜ")
            tax_office = st.text_input("ΔΟΥ")
            address = st.text_input("Διεύθυνση")
            email = st.text_input("Email")
            phone = st.text_input("Τηλέφωνο")
            
            if st.form_submit_button("💾 Αποθήκευση Εταιρείας"):
                if not company_name.strip():
                    st.warning("⚠️ Παρακαλώ συμπληρώστε την επωνυμία της εταιρείας.")
                else:
                    try:
                        gc = get_gspread_client()
                        try:
                            sh_comp = gc.open(OLD_SHEET_NAME).worksheet("db_customers_company")
                        except:
                            sh_comp = gc.open(OLD_SHEET_NAME).add_worksheet(title="db_customers_company", rows="1000", cols="10")
                            sh_comp.append_row(["id", "company_name", "vat_number", "tax_office", "address", "email", "phone"])

                        recs = sh_comp.get_all_records()
                        nid = int(pd.DataFrame(recs)['id'].max()) + 1 if recs and 'id' in pd.DataFrame(recs).columns and not pd.DataFrame(recs).empty else 1
                        
                        sh_comp.append_row([nid, company_name, vat_number, tax_office, address, email, phone])
                        st.success(f"✅ Η εταιρεία '{company_name}' αποθηκεύτηκε επιτυχώς!")
                        time.sleep(1.1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Σφάλμα αποθήκευσης εταιρείας: {e}")

        st.markdown("---")
        st.subheader("📋 Λίστα Εταιρειών")
        try:
            gc = get_gspread_client()
            comp_df = pd.DataFrame(gc.open(OLD_SHEET_NAME).worksheet("db_customers_company").get_all_records())
            if not comp_df.empty:
                comp_df.columns = comp_df.columns.astype(str).str.strip()
                st.dataframe(comp_df, use_container_width=True, hide_index=True)
            else:
                st.info("Δεν υπάρχουν καταχωρημένες εταιρείες.")
        except:
            st.info("Δεν βρέθηκε το tab db_customers_company ή είναι κενό.")

    # --- ΣΕΛΙΔΑ: PRE-REGISTER ---
    elif st.session_state.page == "pre_register_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("📝 Μαθητές σε Κατάσταση Pre-Register")
        
        reg_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")
        
        if reg_df.empty or cust_df.empty:
            st.info("Δεν βρέθηκαν δεδομένα εγγραφών ή πελατών.")
        else:
            reg_df.columns = reg_df.columns.astype(str).str.strip()
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            
            # Φίλτρο Σχολικού Έτους
            available_years = ["ΟΛΑ"] + sorted(reg_df['school_year'].dropna().astype(str).unique().tolist(), reverse=True)
            sel_pr_year = st.selectbox("📅 Φίλτρο ανά Σχολικό Έτος", options=available_years, key="pr_year_filter")
            
            # Φιλτράρισμα των εγγραφών που έχουν class "Pre-Register" (ανεξαρτήτως πεζών/κεφαλαίων)
            pr_regs = reg_df[reg_df['class'].astype(str).str.strip().str.lower() == 'pre-register'].copy()
            
            if sel_pr_year != "ΟΛΑ":
                pr_regs = pr_regs[pr_regs['school_year'].astype(str).str.strip() == sel_pr_year]
                
            if pr_regs.empty:
                st.info("Δεν βρέθηκαν μαθητές με κατάσταση Pre-Register για τα επιλεγμένα κριτήρια.")
            else:
                pr_regs['id_customer_str'] = pr_regs['id_customer'].astype(str).str.strip()
                cust_df['id_customer_str'] = cust_df['id_customer'].astype(str).str.strip()
                
                # Συνδυάζουμε τις εγγραφές pre-register με ΟΛΑ τα στοιχεία του πελάτη από το db_customers
                merged_pr = pd.merge(pr_regs, cust_df, on='id_customer_str', how='inner', suffixes=('', '_cust'))
                
                if not merged_pr.empty:
                    customer_cols = [c for c in cust_df.columns if c != 'id_customer_str']
                    display_cols = ['school_year'] + customer_cols
                    existing_display_cols = [c for c in display_cols if c in merged_pr.columns]
                    
                    final_pr_df = merged_pr[existing_display_cols].copy()
                    
                    # Πλήρης μετονομασία όλων των στηλών στα ελληνικά
                    renames_pr = {
                        'school_year': 'Σχολικό Έτος',
                        'id_customer': 'ID Πελάτη',
                        'customer_firstname': 'Όνομα Μαθητή',
                        'customer_lastname': 'Επώνυμο Μαθητή',
                        'parent_firstname': 'Όνομα Γονέα',
                        'parent_lastname': 'Επώνυμο Γονέα',
                        'telephone': 'Τηλέφωνο',
                        'email': 'Email',
                        'address': 'Διεύθυνση',
                        'birthday_date': 'Ημερομηνία Γέννησης',
                        'register_date': 'Ημερομηνία Εγγραφής',
                        'Taxis_User': 'Taxis Username',
                        'Taxis_Pass': 'Taxis Password',
                        'ADT': 'ΑΔΤ',
                        'AFM': 'ΑΦΜ'
                    }
                    final_pr_df = final_pr_df.rename(columns=renames_pr)
                    
                    st.markdown(f"**Βρέθηκαν εγγραφές Pre-Register:** {len(final_pr_df)}")
                    st.dataframe(final_pr_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Δεν βρέθηκαν αντιστοιχίσεις πελατών.")

    
    # --- ΣΕΛΙΔΑ: ΕΝΙΑΙΑ ΟΙΚΟΝΟΜΙΚΗ ΕΠΙΣΚΟΠΗΣΗ & ΤΑΜΕΙΟ (DB_FINAL) ---
    elif st.session_state.page == "db_final_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("📊 Συνολική Οικονομική Επισκόπηση")
        
        # 1. Φόρτωση δεδομένων
        pays_df = load_specific_db("db_payments")
        try:
            gc = get_gspread_client()
            sh_exp = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
            exp_df = pd.DataFrame(sh_exp.get_all_records())
        except:
            exp_df = pd.DataFrame()

        # Επεξεργασία Εσόδων
        income_records = []
        if not pays_df.empty:
            pays_df.columns = pays_df.columns.astype(str).str.strip()
            for _, r in pays_df.iterrows():
                p_date = str(r.get('payment_date', '')).strip()
                raw_amt = str(r.get('amount_paid', 0)).replace('€', '').replace(',', '.').strip()
                try: amt = float(raw_amt)
                except: amt = 0.0
                cust_name = f"{r.get('customer_lastname', '')} {r.get('customer_firstname', '')}".strip()
                
                if p_date and amt > 0:
                    income_records.append({
                        'date': p_date,
                        'month_year': p_date[:7] if len(p_date) >= 7 else 'Άγνωστο',
                        'type': 'Έσοδο',
                        'description': f"Πληρωμή συνδρομής: {cust_name}",
                        'income': amt,
                        'expense': 0.0,
                        'is_cash_draw': False
                    })

        # Επεξεργασία Εξόδων & Ταμείου
        expense_records = []
        if not exp_df.empty:
            exp_df.columns = exp_df.columns.astype(str).str.strip()
            for _, r in exp_df.iterrows():
                e_date = str(r.get('payment_date', '')).strip() or str(r.get('entry_date', '')).strip()
                raw_amt = str(r.get('amount', 0)).replace('€', '').replace(',', '.').strip()
                try: amt = float(raw_amt)
                except: amt = 0.0
                comp_name = str(r.get('company_name', '')).strip()
                exp_type_col = str(r.get('type', 'Έξοδα')).strip()
                
                if e_date and amt > 0:
                    is_cash_draw = ("ΤΑΜΕΙΟ" in comp_name.upper())
                    
                    expense_records.append({
                        'date': e_date,
                        'month_year': e_date[:7] if len(e_date) >= 7 else 'Άγνωστο',
                        'type': "Ανάληψη / Κλείσιμο Ταμείου" if is_cash_draw else f"Έξοδος ({comp_name})",
                        'description': f"Ανάληψη Ταμείου: {comp_name}" if is_cash_draw else f"Έξοδος Επιχείρησης: {comp_name}",
                        'income': 0.0,
                        'expense': amt,
                        'is_cash_draw': is_cash_draw
                    })

        all_records = income_records + expense_records
        if not all_records:
            st.info("Δεν βρέθηκαν οικονομικές κινήσεις.")
        else:
            unified_df = pd.DataFrame(all_records)
            unified_df = unified_df.sort_values(by='date', ascending=True) # Ταξινόμηση χρονολογικά για σωστό υπολοιποποιητή

            # Υπολογισμός συνολικού ταμείου χρονολογικά (Running Balance)
            unified_df['net'] = unified_df['income'] - unified_df['expense']
            unified_df['running_drawer'] = unified_df['net'].cumsum()

            # 2. Φίλτρο Αναζήτησης ανά Μήνα
            available_months = ["ΟΛΑ"] + sorted(unified_df['month_year'].unique().tolist(), reverse=True)
            sel_month = st.selectbox("📅 Φίλτρο ανά Μήνα (YYYY-MM)", options=available_months)

            # Υπολογισμός μεταφοράς υπολοίπου από προηγούμενους μήνες (Carry Forward)
            filtered_df = unified_df.copy()
            opening_balance = 0.0
            
            if sel_month != "ΟΛΑ":
                # Υπολογίζουμε τι υπήρχε στο ταμείο πριν από αυτόν τον μήνα
                prior_df = unified_df[unified_df['month_year'] < sel_month]
                if not prior_df.empty:
                    opening_balance = prior_df['net'].sum()
                
                filtered_df = unified_df[unified_df['month_year'] == sel_month]

            # 3. Εμφάνιση Grid
            st.markdown(f"**Συνολικές Κινήσεις περιόδου:** {len(filtered_df)}")
            
            display_df = filtered_df.copy()
            display_df['income'] = display_df['income'].apply(lambda x: f"{x:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
            display_df['expense'] = display_df['expense'].apply(lambda x: f"{x:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            renames = {
                'date': 'Ημερομηνία',
                'month_year': 'Μήνας',
                'type': 'Κατηγορία',
                'description': 'Περιγραφή',
                'income': 'Έσοδα (+)',
                'expense': 'Έξοδος / Ανάληψη (-)'
            }
            display_df = display_df.rename(columns=renames)
            
            st.dataframe(display_df[['Ημερομηνία', 'Κατηγορία', 'Περιγραφή', 'Έσοδα (+)', 'Έξοδος / Ανάληψη (-)']], use_container_width=True, hide_index=True)

            # 4. Σύνολα στο κάτω μέρος ανάλογα με το φίλτρο
            period_inc = filtered_df['income'].sum()
            period_exp = filtered_df['expense'].sum()
            period_net = period_inc - period_exp
            
            # Τελικό ταμείο περιλαμβάνει και το άνοιγμα (opening balance) αν επιλεχθεί συγκεκριμένος μήνας
            total_drawer_balance = opening_balance + period_net if sel_month != "ΟΛΑ" else unified_df['running_drawer'].iloc[-1] if not unified_df.empty else 0.0

            st.divider()
            
            if sel_month != "ΟΛΑ":
                st.info(f"📌 **Άνοιγμα Ταμείου (από προηγούμενους μήνες):** `{opening_balance:,.2f} €`".replace(',', 'X').replace('.', ',').replace('X', '.'))

            col_s1, col_s2, col_s3 = st.columns(3)
            col_s1.metric(label="📈 Έσοδα:", value=f"{period_inc:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
            col_s2.metric(label="📉 Έξοδα - Αναλήψεις", value=f"{period_exp:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
            col_s3.metric(label="💰 Ταμείο", value=f"{total_drawer_balance:,.2f} €".replace(',', 'X').replace('.', ',').replace('X', '.'))
    
    # --- ΣΕΛΙΔΑ: ΠΑΡΑΠΟΝΑ & ΣΧΟΛΙΑ ---
    elif st.session_state.page == "complaints_page":
        if st.button("⬅️ Μενού", use_container_width=True): 
            st.session_state.page = "menu"
            st.rerun() 
        
        st.subheader("🛠️ Διαχείριση Παραπόνων, Σχολίων & Παρατηρήσεων")
        
        cust_df = load_specific_db("db_customers")
        comp_df = load_specific_db("db_complaint")
        
        if cust_df.empty:
            st.warning("⚠️ Δεν βρέθηκαν πελάτες στη βάση.")
        else:
            # Αλφαβητική ταξινόμηση πελατών βάσει επωνύμου
            if 'customer_lastname' in cust_df.columns and 'customer_firstname' in cust_df.columns:
                cust_df = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])

            cust_df['display'] = cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")"
            
            action_mode = st.radio("Ενέργεια", ["➕ Νέα Καταχώρηση", "✏️ Τροποποίηση / Κλείσιμο Υπάρχοντος"], horizontal=True)
            
            if action_mode == "➕ Νέα Καταχώρηση":
                sel_cust = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist(), key="new_comp_cust")
                
                if sel_cust != "-- Επιλέξτε --":
                    row_c = cust_df[cust_df['display'] == sel_cust].iloc[0]
                    c_id = str(row_c['id_customer']).strip()
                    c_fn = str(row_c['customer_firstname']).strip()
                    c_ln = str(row_c['customer_lastname']).strip()
                    
                    with st.form("new_complaint_form", clear_on_submit=True):
                        choice_type = st.selectbox("Επιλογή Τύπου", options=["Σχόλιο", "Παράπονο", "Παρατήρηση"])
                        subject = st.text_input("Θέμα / Τίτλος").strip()
                        text_content = st.text_area("Κείμενο / Περιγραφή")
                        staff_name = st.text_input("Υπεύθυνος / Καθηγητής (Staff)", value=st.session_state.get("username", "")).strip()
                        
                        if st.form_submit_button("💾 Αποθήκευση Καταχώρησης"):
                            if not subject or not text_content:
                                st.warning("⚠️ Συμπληρώστε Θέμα και Κείμενο!")
                            else:
                                try:
                                    gc = get_gspread_client()
                                    try:
                                        sh_comp = gc.open(OLD_SHEET_NAME).worksheet("db_complaint")
                                    except:
                                        sh_comp = gc.open(OLD_SHEET_NAME).add_worksheet(title="db_complaint", rows="1000", cols="10")
                                        sh_comp.append_row(["id_complaint", "id_customer", "customer_firstname", "customer_lastname", "complaint_date", "complaint_subject", "complaint_text", "complaint_staff", "status", "choice"])

                                    recs = sh_comp.get_all_records()
                                    nid = int(pd.DataFrame(recs)['id_complaint'].max()) + 1 if recs and 'id_complaint' in pd.DataFrame(recs).columns and not pd.DataFrame(recs).empty else 1
                                    
                                    today_str = datetime.today().strftime('%Y-%m-%d')
                                    row_data = [
                                        nid, int(c_id), c_fn, c_ln, today_str, subject, text_content, staff_name, "Ανοιχτό", choice_type
                                    ]
                                    
                                    sh_comp.append_row(row_data)
                                    st.success("✅ Η καταχώρηση αποθηκεύτηκε επιτυχώς ως Ανοιχτή!")
                                    time.sleep(1.2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Σφάλμα αποθήκευσης: {e}")

            else: # Τροποποίηση / Κλείσιμο
                st.markdown("#### ✏️ Τροποποίηση ή Κλείσιμο Καταχώρησης")
                if comp_df.empty:
                    st.info("Δεν υπάρχουν καταχωρημένα σχόλια ή παράπονα.")
                else:
                    comp_df.columns = comp_df.columns.astype(str).str.strip()
                    # Φίλτρο για να βλέπει ανοιχτά ή όλα
                    show_only_open = st.checkbox("Εμφάνιση μόνο Ανοιχτών", value=True)
                    
                    working_df = comp_df.copy()
                    if show_only_open and 'status' in working_df.columns:
                        working_df = working_df[working_df['status'].astype(str).str.strip().str.lower() == 'ανοιχτό']
                    
                    if working_df.empty:
                        st.info("Δεν βρέθηκαν εγγραφές με αυτά τα κριτήρια.")
                    else:
                        working_df['display_comp'] = "ID: " + working_df['id_complaint'].astype(str) + " | " + working_df['customer_lastname'].astype(str) + " " + working_df['customer_firstname'].astype(str) + " (" + working_df['choice'].astype(str) + " - " + working_df['complaint_subject'].astype(str) + ")"
                        
                        sel_comp_str = st.selectbox("Επιλογή Εγγραφής προς Επεξεργασία", options=working_df['display_comp'].tolist())
                        comp_row = working_df[working_df['display_comp'] == sel_comp_str].iloc[0]
                        
                        target_id = str(comp_row['id_complaint']).strip()
                        cur_choice = str(comp_row.get('choice', 'Σχόλιο'))
                        cur_subj = str(comp_row.get('complaint_subject', ''))
                        cur_text = str(comp_row.get('complaint_text', ''))
                        cur_staff = str(comp_row.get('complaint_staff', ''))
                        cur_status = str(comp_row.get('status', 'Ανοιχτό'))
                        
                        choice_opts = ["Σχόλιο", "Παράπονο", "Παρατήρηση"]
                        c_idx = choice_opts.index(cur_choice) if cur_choice in choice_opts else 0
                        
                        status_opts = ["Ανοιχτό", "Κλειστό"]
                        s_idx = status_opts.index(cur_status) if cur_status in status_opts else 0

                        with st.form("edit_complaint_form"):
                            edit_choice = st.selectbox("Τύπος", options=choice_opts, index=c_idx)
                            edit_subj = st.text_input("Θέμα / Τίτλος", value=cur_subj)
                            edit_text = st.text_area("Κείμενο / Περιγραφή", value=cur_text)
                            edit_staff = st.text_input("Υπεύθυνος (Staff)", value=cur_staff)
                            edit_status = st.selectbox("Κατάσταση", options=status_opts, index=s_idx)
                            
                            if st.form_submit_button("💾 Αποθήκευση Αλλαγών"):
                                try:
                                    gc = get_gspread_client()
                                    sh_comp = gc.open(OLD_SHEET_NAME).worksheet("db_complaint")
                                    all_rows = sh_comp.get_all_values()
                                    
                                    for r_idx, r_val in enumerate(all_rows[1:], start=2):
                                        if str(r_val[0]).strip() == target_id:
                                            # id_complaint, id_customer, customer_firstname, customer_lastname, complaint_date, complaint_subject, complaint_text, complaint_staff, status, choice
                                            sh_comp.update_cell(r_idx, 6, edit_subj)
                                            sh_comp.update_cell(r_idx, 7, edit_text)
                                            sh_comp.update_cell(r_idx, 8, edit_staff)
                                            sh_comp.update_cell(r_idx, 9, edit_status)
                                            sh_comp.update_cell(r_idx, 10, edit_choice)
                                            break
                                            
                                    st.success("✅ Η εγγραφή ενημερώθηκε επιτυχώς! (Αν έγινε Κλειστό, αφαιρέθηκε από τα ενεργά ανοιχτά).")
                                    time.sleep(1.2)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Σφάλμα ενημέρωσης: {e}")

        # --- ΛΙΣΤΑ ΜΕ ΟΛΑ ΤΑ ΑΝΟΙΧΤΑ ---
        st.markdown("---")
        st.subheader("📋 Λίστα με όλα τα Ανοιχτά Σχόλια / Παράπονα")
        if not comp_df.empty:
            comp_df.columns = comp_df.columns.astype(str).str.strip()
            open_comps = comp_df[comp_df['status'].astype(str).str.strip().str.lower() == 'ανοιχτό'].copy()
            if open_comps.empty:
                st.info("🎉 Δεν υπάρχουν ανοιχτά παράπονα ή σχόλια αυτή τη στιγμή!")
            else:
                renames_c = {
                    'customer_lastname': 'Επώνυμο',
                    'customer_firstname': 'Όνομα',
                    'complaint_date': 'Ημερομηνία',
                    'complaint_subject': 'Θέμα',
                    'complaint_text': 'Κείμενο',
                    'complaint_staff': 'Υπεύθυνος',
                    'status': 'Κατάσταση',
                    'choice': 'Τύπος'
                }
                show_df = open_comps.drop(columns=[c for c in ['id_complaint', 'id_customer'] if c in open_comps.columns]).rename(columns=renames_c)
                st.dataframe(show_df, use_container_width=True, hide_index=True)
        else:
            st.info("Δεν βρέθηκαν δεδομένα στη βάση `db_complaint`.")

    else:
        st.session_state.page = "menu"
        st.rerun()

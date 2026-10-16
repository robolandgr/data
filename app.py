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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 1. ID Παλιού και Νέου Google Sheet
OLD_SHEET_NAME = "Database_Center" # Το παλιό σου sheet στον κύριο λογαριασμό
#NEW_SHEET_ID = "1ZFGvrD-Zr0TEGz-l0PjCWgNIF05Qp96FTScD1eTjy64" #ArgyDB_Second
NEW_SHEET_ID = "1w3PUqUrXbzxYnB0O4JSuz0TX7WQc7lHXAti6GN27dqY" #CNatsik1 

# 1. Αρχικοποίηση όλων των session states
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
# 2. Έλεγχος αδράνειας
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
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(creds)

# --- ΦΟΡΤΩΣΗ ΑΠΟ ΤΟ ΠΑΛΙΟ SHEET (ΓΙΑ ΟΛΑ ΤΑ ΥΠΟΛΟΙΠΑ) ---
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

# --- ΦΟΡΤΩΣΗ ΑΠΟ ΤΟ ΝΕΟ SHEET (ΑΠΟΚΛΕΙΣΤΙΚΑ ΓΙΑ ΤΟ MASTER) ---
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

# --- URL ACCESS GATE ---
query_params = st.query_params
if query_params.get("ref77077") != st.secrets["URL_ACCESS_KEY"]:
    st.error("🚫 Δεν έχετε εξουσιοδότηση πρόσβασης. Παρακαλώ επικοινωνήστε με τον διαχειριστή.")
    st.stop()

st.set_page_config(page_title="Διαχείριση", page_icon="💼", layout="wide")

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
PAYMENT_COLUMNS = ['id_payment', 'id_customer', 'customer_lastname', 'customer_firstname', 'payment_date', 'receipt_number', 'payment_number', 'amount_paid', 'school_year']
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
        
        if not target_indices:
            return

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

#--------------------------------
def parse_db_amount(val):
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).strip().replace(',', '.')
    try:
        # Αν η Google το έχει αποθηκεύσει πολλαπλασιασμένο (π.χ. 75 αντί 7.5), 
        # ή αν είναι καθαρό float, το διαβάζουμε σωστά
        f_val = float(val_str)
        return round(f_val, 2)
    except ValueError:
        cleaned = "".join(c for c in val_str if c.isdigit() or c == '.' or c == '-')
        return round(float(cleaned), 2) if cleaned else 0.0


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
#----------------------------------------------------------------------
def generate_active_customers_pdf(selected_year, df_data):
    """Δημιουργεί PDF για την Αναφορά Πελάτων με Αποδείξεις και Ποσά"""
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        import io

        buffer = io.BytesIO()
        # Χρησιμοποιούμε A4 σε landscape (οριζόντια) προσανατολισμό για να χωράνε άνετα όλες οι στήλες
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        elements = []
        styles = getSampleStyleSheet()

        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Normal'],
            fontSize=10,
            leading=12,
            alignment=1, # Center
            textColor=colors.HexColor('#555555')
        )
        
        title_style = ParagraphStyle(
            'ReportTitle',
            parent=styles['Heading1'],
            fontSize=14,
            leading=16,
            spaceAfter=8,
            alignment=1,
            textColor=colors.HexColor('#2c3e50')
        )

        # --- ΚΕΦΑΛΙΔΑ ΕΠΙΧΕΙΡΗΣΗΣ ---
        elements.append(Paragraph("<b>COMPUTERGR</b>", ParagraphStyle('Company', parent=title_style, fontSize=13)))
        elements.append(Paragraph("ΚΕΝΤΡΟ ΠΛΗΡΟΦΟΡΙΚΗΣ & ΡΟΜΠΟΤΙΚΗΣ & ΔΙΑ ΒΙΟΥ ΜΑΘΗΣΗΣ", header_style))
        elements.append(Spacer(1, 6))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#2c3e50'), spaceAfter=10))

        # --- ΤΙΤΛΟΣ ΑΝΑΦΟΡΑΣ ---
        elements.append(Paragraph("<b>ΑΝΑΦΟΡΑ ΠΕΛΑΤΩΝ ΜΕ ΑΠΟΔΕΙΞΕΙΣ ΑΝΑ ΕΤΟΣ</b>", title_style))
        elements.append(Spacer(1, 4))
        
        year_text = f"ΕΠΙΛΕΓΜΕΝΟ ΕΤΟΣ: {str(selected_year).upper()}"
        elements.append(Paragraph(f"<b>{year_text}</b>", ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=10)))
        elements.append(Spacer(1, 12))

        if not df_data.empty:
            # Κεφαλίδες πίνακα στα ΚΕΦΑΛΑΙΑ
            table_data = [['ΕΠΩΝΥΜΟ', 'ΟΝΟΜΑ', 'ΕΤΟΣ', 'ΠΡΟΓΡΑΜΜΑ', 'ΑΡ. ΑΠΟΔ.', 'ΗΜΕΡ/ΝΙΑ', 'ΠΟΣΟ ΑΠΟΔ.', 'ΣΥΜΦ. ΠΟΣΟ']]
            
            for _, r in df_data.iterrows():
                table_data.append([
                    str(r.get('Επώνυμο', '')).upper(),
                    str(r.get('Όνομα', '')).upper(),
                    str(r.get('Σχολικό Έτος', '')).upper(),
                    str(r.get('Πρόγραμμα', '')).upper(),
                    str(r.get('Αρ. Απόδειξης', '')).upper(),
                    str(r.get('Ημερομηνία', '')).upper(),
                    f"{float(r.get('Ποσό Απόδειξης (€)', 0)):.2f}€",
                    f"{float(r.get('Συμφωνηθέν Ποσό (€)', 0)):.2f}€"
                ])

            # Πλάτη στηλών για A4 Landscape (Συνολικό πλάτος ~800pt)
            t = Table(table_data, colWidths=[110, 110, 75, 170, 75, 75, 80, 80])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(t)
            
            elements.append(Spacer(1, 12))
            
            # Υπολογισμός συνολικών ποσών (με το μοναδικό συμφωνηθέν για αποφυγή διπλασιασμού)
            unique_agreements = df_data.drop_duplicates(subset=['Επώνυμο', 'Όνομα', 'Σχολικό Έτος', 'Πρόγραμμα', 'Συμφωνηθέν Ποσό (€)'])
            total_agreed = unique_agreements['Συμφωνηθέν Ποσό (€)'].sum()
            total_receipts = df_data['Ποσό Απόδειξης (€)'].sum()

            summary_text = f"<b>ΣΥΝΟΛΙΚΟ ΠΟΣΟ ΑΠΟΔΕΙΞΕΩΝ:</b> {total_receipts:.2f}€ &nbsp;&nbsp;|&nbsp;&nbsp; <b>ΣΥΝΟΛΙΚΟ ΣΥΜΦΩΝΗΘΕΝ ΠΟΣΟ:</b> {total_agreed:.2f}€"
            elements.append(Paragraph(summary_text, ParagraphStyle('SumStyle', parent=styles['Normal'], fontSize=10, alignment=1)))

        doc.build(elements)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Σφάλμα δημιουργίας PDF: {e}")
        return None

#----------------------------------------------------------------------
def generate_dypa_reports_pdf(selected_year, selected_class, df_data):
    """Δημιουργεί PDF για το Μαθητολόγιο Χρηματοδοτούμενων Προγραμμάτων - ΔΥΠΑ με έτος και τμήμα στον τίτλο"""
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        import io

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
        elements = []
        styles = getSampleStyleSheet()

        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Normal'],
            fontSize=10,
            leading=12,
            alignment=1,
            textColor=colors.HexColor('#555555')
        )
        
        title_style = ParagraphStyle(
            'ReportTitle',
            parent=styles['Heading1'],
            fontSize=14,
            leading=16,
            spaceAfter=8,
            alignment=1,
            textColor=colors.HexColor('#2c3e50')
        )

        # --- ΚΕΦΑΛΙΔΑ ΕΠΙΧΕΙΡΗΣΗΣ ---
        elements.append(Paragraph("<b>COMPUTERGR</b>", ParagraphStyle('Company', parent=title_style, fontSize=13)))
        elements.append(Paragraph("ΚΕΝΤΡΟ ΠΛΗΡΟΦΟΡΙΚΗΣ & ΡΟΜΠΟΤΙΚΗΣ & ΔΙΑ ΒΙΟΥ ΜΑΘΗΣΗΣ", header_style))
        elements.append(Spacer(1, 6))
        elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#2c3e50'), spaceAfter=10))

        # --- ΤΙΤΛΟΣ ΑΝΑΦΟΡΑΣ ---
        elements.append(Paragraph("<b>ΜΑΘΗΤΟΛΟΓΙΟ ΧΡΗΜΑΤΟΔΟΤΟΥΜΕΝΩΝ ΠΡΟΓΡΑΜΜΑΤΩΝ - ΔΥΠΑ</b>", title_style))
        elements.append(Spacer(1, 4))
        
        subtitle_text = f"ΕΠΙΛΕΓΜΕΝΟ ΕΤΟΣ: {str(selected_year).upper()} &nbsp;&nbsp;|&nbsp;&nbsp; ΤΜΗΜΑ: {str(selected_class).upper()}"
        elements.append(Paragraph(f"<b>{subtitle_text}</b>", ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=10)))
        elements.append(Spacer(1, 12))

        if not df_data.empty:
            table_data = [['ΕΠΩΝΥΜΟ', 'ΟΝΟΜΑ', 'ΕΤΟΣ', 'ΠΡΟΓΡΑΜΜΑ', 'ΤΜΗΜΑ', 'ΑΡ. ΑΠΟΔ.', 'ΗΜΕΡ/ΝΙΑ', 'ΠΟΣΟ (€)']]
            
            for _, r in df_data.iterrows():
                table_data.append([
                    str(r.get('Επώνυμο', '')).upper(),
                    str(r.get('Όνομα', '')).upper(),
                    str(r.get('Σχολικό Έτος', '')).upper(),
                    str(r.get('Πρόγραμμα', '')).upper(),
                    str(r.get('Τμήμα', '')).upper(),
                    str(r.get('Αρ. Απόδειξης', '')).upper(),
                    str(r.get('Ημερομηνία', '')).upper(),
                    f"{float(r.get('Ποσό (€)', 0)):.2f}€"
                ])

            t = Table(table_data, colWidths=[100, 100, 65, 140, 100, 70, 70, 75])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2c3e50')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]))
            elements.append(t)
            
            elements.append(Spacer(1, 12))
            
            total_amount = df_data['Ποσό (€)'].sum() if 'Ποσό (€)' in df_data.columns else 0.0
            summary_text = f"<b>ΣΥΝΟΛΙΚΟ ΠΟΣΟ ΔΥΠΑ:</b> {total_amount:.2f}€"
            elements.append(Paragraph(summary_text, ParagraphStyle('SumStyle', parent=styles['Normal'], fontSize=10, alignment=1)))

        doc.build(elements)
        buffer.seek(0)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"Σφάλμα δημιουργίας PDF ΔΥΠΑ: {e}")
        return None
#----------------------------------------------------------------------
def generate_master_finances_pdf(selected_cust, selected_year, master_df):
    buffer = io.BytesIO()
    from reportlab.lib.pagesizes import landscape
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []
    styles = getSampleStyleSheet()

    # Στυλ
    header_title_style = ParagraphStyle('HeaderTitle', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'))
    header_info_style = ParagraphStyle('HeaderInfo', parent=styles['Normal'], fontSize=8, alignment=1, fontName='Helvetica', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=11, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=9, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=7.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))
    total_style = ParagraphStyle('Total', parent=styles['Normal'], fontSize=9, fontName='Helvetica-Bold', alignment=2, textColor=colors.HexColor('#111827'))

    # Κεφαλίδα Εταιρείας
    story.append(Paragraph("COMPUTERGR | ΚΕΝΤΡΟ ΡΟΜΠΟΤΙΚΗΣ, ΠΛΗΡΟΦΟΡΙΚΗΣ & ΔΙΑ ΒΙΟΥ ΜΑΘΗΣΗΣ", header_title_style))
    company_details = "Μαραγκουδάκη 8, Σέρρες 62121 | Τηλ: 23210 55532 | www.cschool.gr | computergrserres@gmail.com"
    story.append(Paragraph(remove_accents(company_details), header_info_style))
    story.append(Spacer(1, 5))

    # Τίτλος Αναφοράς
    story.append(Paragraph("ΣΥΓΚΕΝΤΡΩΤΙΚΗ ΑΝΑΦΟΡΑ MASTER ΟΙΚΟΝΟΜΙΚΩΝ", title_style))
    filter_info = f"ΠΕΛΑΤΗΣ: {remove_accents(str(selected_cust).upper())} | ΕΤΟΣ: {remove_accents(str(selected_year).upper())}"
    story.append(Paragraph(filter_info, sub_style))
    story.append(Spacer(1, 5))

    table_data = [[
        Paragraph("ΗΜΕΡ.", th_style),
        Paragraph("ΤΥΠΟΣ", th_style),
        Paragraph("ΚΑΤΗΓΟΡΙΑ", th_style),
        Paragraph("ΠΕΛΑΤΗΣ / ΠΡΟΜ.", th_style),
        Paragraph("ΑΡ. ΤΙΜ.", th_style),
        Paragraph("ΠΟΣΟ (€)", th_style),
        Paragraph("ΜΕΘΟΔΟΣ", th_style),
        Paragraph("ΕΤΟΣ", th_style),
        Paragraph("ΠΗΓΗ", th_style)
    ]]

    for _, row in master_df.iterrows():
        try:
            amt = float(str(row.get('Ποσό (€)', 0)).replace(',', '.') or 0)
        except:
            amt = 0.0

        table_data.append([
            Paragraph(str(row.get('Ημερομηνία', '')), td_style),
            Paragraph(remove_accents(str(row.get('Τύπος', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Κατηγορία', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Πελάτης/Προμηθευτής', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Αρ. Τιμολογίου', ''))), td_style),
            Paragraph(f"{amt:.2f}", td_style),
            Paragraph(remove_accents(str(row.get('Μέθοδος', ''))), td_style),
            Paragraph(str(row.get('Σχολικό Έτος', '')), td_style),
            Paragraph(str(row.get('Πηγή', '')), td_style)
        ])

    t = Table(table_data, colWidths=[65, 50, 110, 140, 65, 55, 75, 65, 55], repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')])
    ]))
    
    story.append(t)
    story.append(Spacer(1, 15))

    # Ασφαλης υπολογισμός συνόλων με pandas to_numeric
    amounts_numeric = pd.to_numeric(master_df['Ποσό (€)'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    types_upper = master_df['Τύπος'].astype(str).str.upper()

    total_income = amounts_numeric[types_upper == 'ΕΣΟΔΟ'].sum()
    total_expense = amounts_numeric[types_upper == 'ΕΞΟΔΟ'].sum()
    net_total = total_income - total_expense

    story.append(Paragraph(f"ΣΥΝΟΛΙΚΑ ΕΣΟΔΑ: {total_income:.2f} EUR", total_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph(f"ΣΥΝΟΛΙΚΑ ΕΞΟΔΑ: {total_expense:.2f} EUR", total_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph(f"ΚΑΘΑΡΟ ΥΠΟΛΟΙΠΟ (ΕΣΟΔΑ - ΕΞΟΔΑ): {net_total:.2f} EUR", total_style))

    doc.build(story)
    buffer.seek(0)
    return buffer
#----------------------------------------------------------------------
def generate_customer_payment_history_pdf(customer_name, selected_year, payments_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    # Στυλ
    header_title_style = ParagraphStyle('HeaderTitle', parent=styles['Heading1'], fontSize=13, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'))
    header_info_style = ParagraphStyle('HeaderInfo', parent=styles['Normal'], fontSize=8.5, alignment=1, fontName='Helvetica', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=12, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#111827'), spaceAfter=10)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, alignment=1, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'), spaceAfter=15)
    
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.white)
    td_style = ParagraphStyle('TD', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))
    total_style = ParagraphStyle('Total', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', alignment=2, textColor=colors.HexColor('#111827'))

    # Κεφαλίδα Εταιρείας
    story.append(Paragraph("COMPUTERGR | ΚΕΝΤΡΟ ΡΟΜΠΟΤΙΚΗΣ, ΠΛΗΡΟΦΟΡΙΚΗΣ & ΔΙΑ ΒΙΟΥ ΜΑΘΗΣΗΣ", header_title_style))
    company_details = "Μαραγκουδάκη 8, Σέρρες 62121 | Τηλ: 23210 55532 | www.cschool.gr | computergrserres@gmail.com"
    story.append(Paragraph(remove_accents(company_details), header_info_style))
    story.append(Spacer(1, 5))

    # Τίτλος Αναφοράς
    story.append(Paragraph("ΑΝΑΛΥΤΙΚΟ ΙΣΤΟΡΙΚΟ ΠΛΗΡΩΜΩΝ ΠΕΛΑΤΗ", title_style))
    filter_info = f"ΠΕΛΑΤΗΣ: {remove_accents(str(customer_name).upper())} | ΕΤΟΣ: {remove_accents(str(selected_year).upper())}"
    story.append(Paragraph(filter_info, sub_style))
    story.append(Spacer(1, 5))

    # Δημιουργία Πίνακα Πληρωμών
    table_data = [[
        Paragraph("ΗΜΕΡΟΜΗΝΙΑ", th_style),
        Paragraph("ΑΡ. ΑΠΟΔΕΙΞΗΣ", th_style),
        Paragraph("ΚΑΤΗΓΟΡΙΑ", th_style),
        Paragraph("ΠΡΟΓΡΑΜΜΑ / ΛΕΠΤΟΜΕΡΕΙΕΣ", th_style),
        Paragraph("ΕΤΟΣ", th_style),
        Paragraph("ΤΡΟΠΟΣ", th_style),
        Paragraph("ΠΟΣΟ (€)", th_style)
    ]]

    total_amount = 0.0

    for _, row in payments_df.iterrows():
        amt = float(row.get('Ποσό (€)', 0))
        total_amount += amt
        
        table_data.append([
            Paragraph(str(row.get('Ημερομηνία', '')), td_style),
            Paragraph(remove_accents(str(row.get('Αρ. Απόδειξης', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Κατηγορία', ''))), td_style),
            Paragraph(remove_accents(str(row.get('Πρόγραμμα / Τύπος', ''))), td_style),
            Paragraph(str(row.get('Έτος', '')), td_style),
            Paragraph(remove_accents(str(row.get('Τρόπος', ''))), td_style),
            Paragraph(f"{amt:.2f}", td_style)
        ])

    t = Table(table_data, colWidths=[65, 75, 65, 130, 55, 65, 55], repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F2937')),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D1D5DB')),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F9FAFB')])
    ]))
    
    story.append(t)
    story.append(Spacer(1, 15))
    story.append(Paragraph(f"ΣΥΝΟΛΙΚΟ ΠΟΣΟ ΠΛΗΡΩΜΩΝ: {total_amount:.2f} EUR", total_style))

    doc.build(story)
    buffer.seek(0)
    return buffer
#----------------------------------------------------------------------
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

def generate_report_pdf(report_df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, spaceAfter=15, alignment=1, textColor=colors.HexColor('#1F2937'))
    label_style = ParagraphStyle('LabelStyle', parent=styles['Normal'], fontSize=8.5, fontName='Helvetica-Bold', textColor=colors.HexColor('#4B5563'))
    value_style = ParagraphStyle('ValueStyle', parent=styles['Normal'], fontSize=9.5, fontName='Helvetica', textColor=colors.HexColor('#111827'))
    name_style = ParagraphStyle('NameStyle', parent=styles['Normal'], fontSize=11, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'))
    total_style = ParagraphStyle('TotalStyle', parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=colors.HexColor('#1F2937'), alignment=2)

    story.append(Paragraph("ΟΙΚΟΝΟΜΙΚΗ ΕΚΘΕΣΗ ΚΑΙ ΚΑΡΤΕΛΕΣ ΦΟΙΤΗΣΗΣ", title_style))
    story.append(Spacer(1, 10))

    try:
        total_c = 0
        total_p = 0
        total_b = 0

        for index, row in report_df.iterrows():
            fname = uppercase_clean(row.get('Όνομα', ''))
            lname = uppercase_clean(row.get('Επώνυμο', ''))
            cust_id = row.get('ID Πελάτη', '')
            year = uppercase_clean(row.get('Έτος', ''))
            prog = uppercase_clean(row.get('Πρόγραμμα', ''))
            u_class = uppercase_clean(row.get('Τμήμα', ''))
            cost = row.get('Κόστος', 0)
            paid = row.get('Έχει Δώσει', 0)
            balance = row.get('Υπόλοιπο', 0)
            status_text = "ΔΙΑΚΟΠΗ" if row.get('Κατάσταση', '') == 'STOP' else 'ΕΝΕΡΓΟΣ'

            total_c += cost
            total_p += paid
            total_b += balance

            row1_data = [
                [Paragraph("ΜΑΘΗΤΗΣ:", label_style), Paragraph("ΣΧΟΛΙΚΟ ΕΤΟΣ:", label_style), Paragraph("ΚΑΤΑΣΤΑΣΗ:", label_style)],
                [Paragraph(f"{lname} {fname} (ID: {cust_id})", name_style), Paragraph(year, value_style), Paragraph(status_text, value_style)]
            ]
            t1 = Table(row1_data, colWidths=[280, 140, 120])
            t1.setStyle(TableStyle([('BOTTOMPADDING', (0, 0), (-1, -1), 2)]))

            row2_data = [
                [Paragraph("ΠΡΟΓΡΑΜΜΑ:", label_style), Paragraph("ΤΜΗΜΑ:", label_style)],
                [Paragraph(prog, value_style), Paragraph(u_class, value_style)]
            ]
            t2 = Table(row2_data, colWidths=[270, 270])
            t2.setStyle(TableStyle([('BOTTOMPADDING', (0, 0), (-1, -1), 2)]))

            row3_data = [
                [Paragraph("ΚΟΣΤΟΣ :", label_style), Paragraph("ΣΥΝΟΛΟ ΠΛΗΡΩΜΩΝ:", label_style), Paragraph("ΥΠΟΛΟΙΠΟ:", label_style)],
                [Paragraph(f"{cost} EUR", value_style), Paragraph(f"{paid} EUR", value_style), Paragraph(f"{balance} EUR", value_style)]
            ]
            t3 = Table(row3_data, colWidths=[180, 180, 180])
            t3.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F9FAFB')),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))

            master_box_data = [[t1], [Spacer(1, 4)], [t2], [Spacer(1, 4)], [t3]]
            master_table = Table(master_box_data, colWidths=[540])
            master_table.setStyle(TableStyle([
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#E5E7EB')),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ]))

            story.append(master_table)
            story.append(Spacer(1, 10))

        story.append(Spacer(1, 15))
        story.append(Paragraph(f"ΣΥΝΟΛΙΚΗ ΧΡΕΩΣΗ: {total_c} EUR", total_style))
        story.append(Paragraph(f"ΣΥΝΟΛΙΚΕΣ ΠΛΗΡΩΜΕΣ: {total_p} EUR", total_style))
        story.append(Paragraph(f"ΣΥΝΟΛΙΚΟ ΥΠΟΛΟΙΠΟ: {total_b} EUR", total_style))

        doc.build(story)
        buffer.seek(0)
        return buffer
    except Exception as pdf_err:
        st.error(f"Σφάλμα PDF: {pdf_err}")
        return None

#----------------------------------------------------
# --- ΑΠΟΘΗΚΕΥΣΗ ΑΠΟΚΛΕΙΣΤΙΚΑ ΣΤΟ ΝΕΟ MASTER SHEET ---
def log_to_finance_master(trans_type, category, partner, invoice, amount, method, year, source):
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
        records = sh.get_all_records()
        if records:
            next_id = int(pd.DataFrame(records)['id'].max()) + 1
        else:
            next_id = 1
        
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
    st.markdown("<div style='font-size:20px; font-weight:bold; text-align: center; margin-bottom: 20px;'>Είσοδος στο Σύστημα</div>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        username_input = st.text_input("Username").strip()
        password_input = st.text_input("Password", type="password").strip()
        submit_login = st.form_submit_button("Σύνδεση", disabled=st.session_state.is_submitting)

    if submit_login:
        try:
            gc = get_gspread_client()
            
            # 1. Φόρτωση χρηστών από το 1ο sheet (Database_Center)
            df_main = load_specific_db("db_users")
            
            # 2. Φόρτωση χρηστών από το 2ο sheet (NEW_SHEET_ID)
            sh_second = gc.open_by_key(NEW_SHEET_ID).worksheet("db_users")
            records_second = sh_second.get_all_records()
            df_second = pd.DataFrame(records_second) if records_second else pd.DataFrame()
            
            # 3. Συνένωση των δύο πηγών σε έναν ενιαίο πίνακα χρηστών
            if not df_main.empty and not df_second.empty:
                df_main.columns = df_main.columns.astype(str).str.strip()
                df_second.columns = df_second.columns.astype(str).str.strip()
                users_df = pd.concat([df_main, df_second], ignore_index=True).drop_duplicates()
            elif not df_main.empty:
                df_main.columns = df_main.columns.astype(str).str.strip()
                users_df = df_main
            else:
                df_second.columns = df_second.columns.astype(str).str.strip()
                users_df = df_second

            users_df = users_df.astype(str).apply(lambda x: x.str.strip())
            user_row = users_df[(users_df['username'] == username_input) & 
                                (users_df['password'] == password_input)]
            
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
        
        # Καθαρίζουμε και τυχόν cached dataframes από το session state αν υπάρχουν
        for key in ['cached_reg_df', 'cached_sum_df', 'cached_cust_df']:
            if key in st.session_state:
                del st.session_state[key]
                
        st.success("Τα δεδομένα ενημερώθηκαν επιτυχώς από τη Google!")
        time.sleep(1)
        st.session_state.is_submitting = False
        st.rerun()
    
    # 1. ΚΕΝΤΡΙΚΟ ΜΕΝΟΥ
    if st.session_state.page == "menu":
        st.markdown("<h1 style='text-align: center;'>DB ComputerGr</h1>", unsafe_allow_html=True)
        st.markdown("<h2 style='text-align: center;'>Κεντρικό Μενού Διαχείρισης</h2>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.session_state.role == "master":
            col_m1, col_m2 = st.columns(2)
            
            with col_m1:
                if st.button("📈 Master Οικονομικά", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.page = "finance_master"
                    st.rerun()
                    
            with col_m2:
                if st.button("🛠️ Χειροκίνητες Κινήσεις", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.is_submitting = True
                    st.session_state.page = "edit_master_manual"
                    st.session_state.is_submitting = False
                    st.rerun()
        
        st.subheader("Διαχείριση Πελατών")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            if st.button("➕ Νέος Πελάτης (Πρώτη Εγγραφή)", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "register"; st.rerun()
            if st.button("📋 Επεξεργασία Στοιχείων Πελατών", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "customer_registry"; st.rerun()
        with col_c2:
            if st.button("👥 Διαχείριση Κατάστασης Πελατών (STOP/Επαναφορά)", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "view_customers"; st.rerun()
            if st.button("🔄 Εισαγωγή Πελάτη σε Τμήμα", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "renew_registration"; st.rerun()

        st.markdown("---")
        st.subheader("Πληρωμές")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("💳 Καταχώρηση Πληρωμής (Συνδρομή)", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "add_payment"; st.rerun()
            if st.button("☀️ Πληρωμές Summer Camp", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "add_payment_summer"; st.rerun()
        with col_p2:
            if st.button("💳 Αποδείξεις επί πιστώσει ΔΥΠΑ", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "payment_dypa"; st.rerun() 
            if st.session_state.role != "super":
                if st.button("💳 Καταχώρηση Πληρωμής (Επιπρόσθετες)", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.page = "extra_payments_view";   st.rerun()
                
        st.markdown("---")
        st.subheader("Καθηγητές")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("👨‍🏫 Εγγραφή/Επεξεργασία Καθηγητή", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "register_teacher"; st.rerun()        
        with col_p2:
            if st.button("👩‍🏫 Στοιχεία Καθηγητών", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "teachers_secretariat"; st.rerun()

        
        st.markdown("---")
        st.subheader("Διαχείριση Κέντρου")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            if st.button("📚 Προγράμματα - Τμήματα - Καθηγητές", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "add_structure"; st.rerun()

        st.markdown("---")
        st.subheader("Εταιρικές Συναλλαγές")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            if st.button("📝 Καταχώρηση Στοιχείων Εταιρείας", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "corp_register"; st.rerun()
        with col_e2:
            if st.button("💸 Διαχείριση Συναλλαγών", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "corp_payments"; st.rerun()
                
        st.markdown("---")
        st.subheader("Αναφορές")
        
        col_r1, col_r2 = st.columns(2)
        
        with col_r1:

            if st.session_state.role in ["employee", "admin", "master"]:
                if st.button("📋 Μητρώο Πελατών", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.page = "anafora_pelaton"; st.rerun()
                if st.button("🏫 Αναφορά Καθηγητών & Τμημάτων", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.page = "teacher_report"; st.rerun()
            if st.button("📋 Μαθητολόγιο Αυτοχρηματοδοτούμενων", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "simple_customer_reports"; st.rerun()

        with col_r2:
            if st.button("🚫 Εκδήλωση Ενδιαφέροντος", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.is_submitting = True
                st.session_state.page = "unassigned_customers"
                st.session_state.is_submitting = False
                st.rerun()
            if st.session_state.role in ["admin", "master"]:
                if st.button("📊 Οικονομική Αναφορά", use_container_width=True, disabled=st.session_state.is_submitting):
                    st.session_state.page = "advanced_reports"; st.rerun()
            if st.button("📋 Μαθητολόγιο Χρηματοδοτούμενων - ΔΥΠΑ", use_container_width=True, disabled=st.session_state.is_submitting):
                st.session_state.page = "dypa_customer_reports"; st.rerun()


        st.markdown("---")
        st.markdown("##### 🔗 Διαχείριση Παραπόνων & Σχολίων")
        if st.button("🛠️ Εισαγωγή/ Τροποποίηση Παραπόνων & Σχολίων", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "complaints_page"
            st.rerun()

        
        # --- ΚΑΤΗΓΟΡΙΑ: ΣΥΝΔΕΣΜΟΙ (Κάτω από τις Αναφορές) ---
        st.markdown("---")
        st.markdown("##### 🔗 Χρήσιμοι Σύνδεσμοι")
        
        st.markdown(
            """
            <style>
            .link-grid {
                display: grid;
                grid-template-columns: repeat(2, minmax(0, 1fr));
                gap: 10px;
                margin-bottom: 15px;
            }
            .link-item {
                background-color: #f8f9fa;
                border: 1px solid #e9ecef;
                padding: 8px 12px;
                border-radius: 6px;
                text-align: center;
                transition: all 0.2s ease-in-out;
            }
            .link-item:hover {
                background-color: #e2e8f0;
                border-color: #cbd5e1;
            }
            .link-item a {
                text-decoration: none;
                color: #2c3e50;
                font-weight: 500;
                font-size: 14px;
                display: block;
            }
            </style>

            <div class="link-grid">
                <div class="link-item" style="grid-column: span 2; background-color: #e8f4fd; border-color: #bbe1fa;">📄 <a href="https://drive.google.com/file/d/1YXd41uSKjOTjVrc6Kb8cIWhHIl5IR7KH/view?usp=sharing" target="_blank">Μαθητικά Δίδακτρα 2026-2027</a></div>
                <div class="link-item">🌐 <a href="https://voucher.gov.gr" target="_blank">Voucher Gov</a></div>
                <div class="link-item">🏫 <a href="https://edu4schools.gr" target="_blank">Edu4Schools</a></div>
                <div class="link-item">🏛️ <a href="https://www.dypa.gov.gr" target="_blank">ΔΥΠΑ (DYPA)</a></div>
                <div class="link-item">💻 <a href="https://eliss.peoplecert.org/Login.aspx" target="_blank">PeopleCert (ELISS)</a></div>
                <div class="link-item">📝 <a href="https://i.globalcert.gr/login/" target="_blank">GlobalCert</a></div>
                <div class="link-item">🔒 <a href="https://security.ramcomputer.gr/login.php" target="_blank">Ram Computer Security</a></div>
                <div class="link-item">📋 <a href="https://ops-efet.gr/" target="_blank">OPS EFET</a></div>
                <div class="link-item">🌾 <a href="https://e-services.minagric.gr/" target="_blank">MinAgric e-Services</a></div>
            </div>
            """,
            unsafe_allow_html=True
        )
    # --- ΣΕΛΙΔΑ: ΚΑΤΑΧΩΡΗΣΗ ΕΤΑΙΡΕΙΑΣ ---
    elif st.session_state.page == "corp_register":
        if st.button("⬅️ Μενού", key="menu_corp_reg", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Καταχώρηση & Τροποποίηση Εταιρικών Στοιχείων")
        
        comp_df = load_specific_db("db_customers_company")
        
        if not comp_df.empty:
            comp_df.columns = comp_df.columns.astype(str).str.strip()

        # Επιλογή ενέργειας (Νέα Καταχώρηση ή Τροποποίηση)
        mode_corp = st.radio("Ενέργεια", ["Νέα Καταχώρηση", "Τροποποίηση Υπάρχουσας"], key="corp_mode_sel", horizontal=True)
        
        selected_corp = None
        if mode_corp == "Τροποποίηση Υπάρχουσας" and not comp_df.empty:
            try:
                comp_df['display'] = comp_df['id'].astype(str) + " - " + comp_df['company_name'].astype(str) + " (ΑΦΜ: " + comp_df['vat_number'].astype(str) + ")"
                sel_c_box = st.selectbox("Επιλογή Εταιρείας για Τροποποίηση", options=["-- Επιλέξτε --"] + comp_df['display'].tolist(), key="sel_corp_edit")
                
                if sel_c_box != "-- Επιλέξτε --":
                    selected_corp = comp_df[comp_df['display'] == sel_c_box].iloc[0]
            except Exception:
                st.warning("Υπήρξε πρόβλημα στην ανάγνωση των στηλών της εταιρείας.")

        # Προεπιλεγμένες τιμές
        def_name = str(selected_corp['company_name']) if selected_corp is not None and 'company_name' in selected_corp else ""
        def_vat = str(selected_corp['vat_number']) if selected_corp is not None and 'vat_number' in selected_corp else ""
        def_tax = str(selected_corp['tax_office']) if selected_corp is not None and 'tax_office' in selected_corp else ""
        def_addr = str(selected_corp['address']) if selected_corp is not None and 'address' in selected_corp else ""
        def_email = str(selected_corp['email']) if selected_corp is not None and 'email' in selected_corp else ""
        def_phone = str(selected_corp['phone']) if selected_corp is not None and 'phone' in selected_corp else ""
        def_year = str(selected_corp['school_year']) if selected_corp is not None and 'school_year' in selected_corp else "2026-2027"

        with st.form("corp_reg_form"):
            name = st.text_input("Επωνυμία (company_name)", value=def_name)
            vat = st.text_input("ΑΦΜ (vat_number)", value=def_vat)
            tax_office = st.text_input("ΔΟΥ (tax_office)", value=def_tax)
            addr = st.text_input("Διεύθυνση (address)", value=def_addr)
            email = st.text_input("Email", value=def_email)
            phone = st.text_input("Τηλέφωνο (phone)", value=def_phone)
            
            years_list = [f"{y}-{y+1}" for y in range(2026, 2041)]
            default_year_idx = years_list.index(def_year) if def_year in years_list else 0
            year = st.selectbox("Σχολικό Έτος (school_year)", options=years_list, index=default_year_idx)
            
            btn_label = "💾 Αποθήκευση Τροποποίησης" if mode_corp == "Τροποποίηση Υπάρχουσας" and selected_corp is not None else "Αποθήκευση"
            
            if st.form_submit_button(btn_label, disabled=st.session_state.is_submitting):
                if not name or not vat:
                    st.warning("⚠️ Συμπληρώστε τουλάχιστον την επωνυμία και το ΑΦΜ!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_customers_company")
                        
                        if mode_corp == "Τροποποίηση Υπάρχουσας" and selected_corp is not None:
                            corp_id = str(selected_corp['id']).strip()
                            old_comp_name = str(selected_corp['company_name']).strip()
                            new_comp_name = str(name).strip()
                            
                            all_rows = sh.get_all_values()
                            target_row = None
                            if len(all_rows) > 1:
                                for idx, row_vals in enumerate(all_rows[1:], start=2):
                                    if len(row_vals) > 0 and str(row_vals[0]).strip() == corp_id:
                                        target_row = idx
                                        break
                            
                            if target_row:
                                # Στήλες: id(1), company_name(2), vat_number(3), tax_office(4), address(5), email(6), phone(7), school_year(8)
                                sh.update_cell(target_row, 2, name)
                                sh.update_cell(target_row, 3, vat)
                                sh.update_cell(target_row, 4, tax_office)
                                sh.update_cell(target_row, 5, addr)
                                sh.update_cell(target_row, 6, email)
                                sh.update_cell(target_row, 7, phone)
                                sh.update_cell(target_row, 8, year)
                                
                                # --- ΑΥΤΟΜΑΤΗ ΕΝΗΜΕΡΩΣΗ ΤΟΥ company_name ΣΤΟ db_payments_company ---
                                if old_comp_name != new_comp_name:
                                    try:
                                        sh_trans = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                                        trans_rows = sh_trans.get_all_values()
                                        if len(trans_rows) > 1:
                                            t_header = [h.strip() for h in trans_rows[0]]
                                            # Εντοπισμός στηλών customer_id (ή id) και company_name
                                            t_id_idx = next((i for i, h in enumerate(t_header) if 'customer' in h.lower() or h.lower() == 'customer_id'), 3)
                                            t_name_idx = next((i for i, h in enumerate(t_header) if 'company_name' in h.lower() or 'company' in h.lower()), 4)
                                            
                                            for t_idx, t_row in enumerate(trans_rows[1:], start=2):
                                                # Έλεγχος είτε με βάση το customer_id είτε με βάση το παλιό όνομα εταιρείας
                                                match_by_id = len(t_row) > t_id_idx and str(t_row[t_id_idx]).strip() == corp_id
                                                match_by_name = len(t_row) > t_name_idx and str(t_row[t_name_idx]).strip() == old_comp_name
                                                
                                                if match_by_id or match_by_name:
                                                    sh_trans.update_cell(t_idx, t_name_idx + 1, new_comp_name)
                                    except Exception as trans_err:
                                        st.warning(f"Η εταιρεία ενημερώθηκε αλλά υπήρξε θέμα στις συναλλαγές: {trans_err}")

                                st.success("✅ Τα στοιχεία της εταιρείας τροποποιήθηκαν επιτυχώς και ενημερώθηκαν οι συναλλαγές!")
                            else:
                                st.error("Δεν βρέθηκε η εγγραφή για τροποποίηση.")
                        else:
                            records = sh.get_all_records()
                            next_id = int(pd.DataFrame(records)['id'].max()) + 1 if records and 'id' in pd.DataFrame(records).columns else (len(records) + 1 if records else 1)
                            sh.append_row([next_id, name, vat, tax_office, addr, email, phone, year])
                            st.success("✅ Καταχωρήθηκε!")
                        
                        #load_specific_db.clear()
                        time.sleep(1.5)
                        
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

    # --- ΣΕΛΙΔΑ: ΠΛΗΡΩΜΕΣ ΕΤΑΙΡΕΙΑΣ ---
    elif st.session_state.page == "corp_payments":
        if st.button("⬅️ Μενού", key="menu_corp_pay", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Πληρωμές & Συναλλαγές (Έσοδα/Έξοδα)")
        
        comp_df = load_specific_db("db_customers_company")
        trans_df = load_specific_db("db_payments_company")
        
        if not comp_df.empty:
            comp_df.columns = comp_df.columns.astype(str).str.strip()
        if not trans_df.empty:
            trans_df.columns = trans_df.columns.astype(str).str.strip()

        if 'last_trans_type' not in st.session_state:
            st.session_state.last_trans_type = "Έσοδα"
    
        with st.form("corp_pay_form", clear_on_submit=True):
            st.markdown("###### ➕ Νέα Καταχώρηση Συναλλαγής")
            trans_type = st.selectbox("Τύπος", ["Έσοδα", "Έξοδα"], key="trans_type")
            
            inv_num = st.text_input("Αριθμός Τιμολογίου (Απαιτείται)")
            comp = st.selectbox("Εταιρεία", comp_df['company_name'].tolist() if not comp_df.empty and 'company_name' in comp_df.columns else [])
            year = st.selectbox("Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2041)])
            inv_date = st.date_input("Ημερομηνία Τιμολογίου")
            pay_date = st.date_input("Ημερομηνία Πληρωμής")
            amount = st.number_input("Ποσό", min_value=0.0)
            pay_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "POS", "Επι πιστώση", "Εθνική", "Eurobank"])
            
            submitted = st.form_submit_button("Αποθήκευση Συναλλαγής", disabled=st.session_state.is_submitting)
            
            if submitted:
                if not inv_num:
                    st.warning("Ο Αριθμός Τιμολογίου είναι απαραίτητος!")
                else:
                    try:
                        gc = get_gspread_client()
                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                        
                        c_id = ""
                        if not comp_df.empty and 'company_name' in comp_df.columns and 'id' in comp_df.columns:
                            matched_c = comp_df[comp_df['company_name'] == comp]
                            if not matched_c.empty:
                                c_id = matched_c['id'].iloc[0]
                                
                        records = sh.get_all_records()
                        next_id = 1
                        if records:
                            df_temp_t = pd.DataFrame(records)
                            id_col_t = next((c for c in df_temp_t.columns if 'id' in c.lower()), df_temp_t.columns[0])
                            numeric_ids_t = pd.to_numeric(df_temp_t[id_col_t], errors='coerce').dropna()
                            if not numeric_ids_t.empty:
                                next_id = int(numeric_ids_t.max()) + 1
                            else:
                                next_id = len(df_temp_t) + 1
                        
                        # Νέα σειρά: id(1), invoice_number(2), type(3), customer_id(4), company_name(5), school_year(6), invoice_date(7), payment_date(8), entry_date(9), amount(10), pay_method(11)
                        sh.append_row([
                            next_id, inv_num, trans_type, c_id, comp, year, 
                            inv_date.strftime('%Y-%m-%d'), pay_date.strftime('%Y-%m-%d'), 
                            datetime.today().strftime('%Y-%m-%d'), amount, pay_method
                        ])
                        
                        master_type = "ΕΣΟΔΟ" if trans_type == "Έσοδα" else "ΕΞΟΔΟ"
                        log_to_finance_master(master_type, "Εταιρική Συναλλαγή", comp, inv_num, amount, pay_method, year, "Corporate")
                        
                        st.success("✅ Η συναλλαγή αποθηκεύτηκε!")
                        time.sleep(1.2)
                        
                    except Exception as e:
                        st.error(f"Σφάλμα: {e}")

        # --- ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΑΣ ΣΥΝΑΛΛΑΓΗΣ ---
        st.markdown("---")
        st.markdown("##### 📋 Υπάρχουσες Εταιρικές Συναλλαγές (Προβολή & Τροποποίηση)")
        
        if not trans_df.empty:
            id_trans_col = next((c for c in trans_df.columns if 'id' in c.lower()), trans_df.columns[0])
            
            display_trans = trans_df.copy()
            st.dataframe(display_trans, use_container_width=True, hide_index=True)
            
            trans_ids_list = trans_df[id_trans_col].astype(str).tolist()
            sel_trans_to_edit = st.selectbox("Επιλογή ID Συναλλαγής για Τροποποίηση", ["-- Επιλέξτε ID --"] + trans_ids_list, key="sel_trans_edit_box")
            
            if sel_trans_to_edit != "-- Επιλέξτε ID --":
                matched_t_row = trans_df[trans_df[id_trans_col].astype(str) == str(sel_trans_to_edit)].iloc[0]
                
                old_inv_num = str(matched_t_row.get('invoice_number', matched_t_row.iloc[1] if len(matched_t_row) > 1 else '')).strip()
                old_amount = float(pd.to_numeric(matched_t_row.get('amount', matched_t_row.iloc[9] if len(matched_t_row) > 9 else 0), errors='coerce') or 0.0)
                
                with st.form("edit_corp_trans_form"):
                    st.markdown(f"**Επεξεργασία Συναλλαγής με ID: {sel_trans_to_edit}**")
                    
                    ed_t_type = st.selectbox("Τύπος", ["Έσοδα", "Έξοδα"], index=0 if str(matched_t_row.get('type', 'Έσοδα')) == "Έσοδα" else 1)
                    ed_t_inv = st.text_input("Αριθμός Τιμολογίου", value=old_inv_num).strip()
                    
                    comp_names = comp_df['company_name'].tolist() if not comp_df.empty and 'company_name' in comp_df.columns else []
                    
                    old_comp_name = str(matched_t_row.get('company_name', '')).strip()
                    default_comp_idx = comp_names.index(old_comp_name) if old_comp_name in comp_names else 0
                    
                    ed_t_comp = st.selectbox("Εταιρεία", options=comp_names, index=default_comp_idx)
                    
                    ed_t_year = st.text_input("Σχολικό Έτος", value=str(matched_t_row.get('school_year', '2026-2027'))).strip()
                    ed_t_amount = st.number_input("Ποσό (€)", min_value=0.0, value=old_amount, format="%.2f")
                    ed_t_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "POS", "Επι πιστώση", "Εθνική", "Eurobank"])
                    
                    try:
                        def_inv_dt = datetime.strptime(str(matched_t_row.get('invoice_date', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                    except:
                        def_inv_dt = datetime.today().date()
                    ed_inv_date = st.date_input("Ημερομηνία Τιμολογίου", value=def_inv_dt)
                    
                    try:
                        def_pay_dt = datetime.strptime(str(matched_t_row.get('payment_date', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                    except:
                        def_pay_dt = datetime.today().date()
                    ed_pay_date = st.date_input("Ημερομηνία Πληρωμής", value=def_pay_dt)
                    
                    if st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης Συναλλαγής", disabled=st.session_state.is_submitting):
                        if not ed_t_inv or ed_t_amount <= 0:
                            st.warning("⚠️ Συμπληρώστε αριθμό τιμολογίου και ποσό μεγαλύτερο από 0!")
                        else:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments_company")
                                
                                c_id_ed = ""
                                if not comp_df.empty and 'company_name' in comp_df.columns and 'id' in comp_df.columns:
                                    m_c = comp_df[comp_df['company_name'] == ed_t_comp]
                                    if not m_c.empty:
                                        c_id_ed = m_c['id'].iloc[0]

                                all_raw_t = sh.get_all_values()
                                target_row_idx = None
                                if len(all_raw_t) > 1:
                                    for r_idx, sheet_row in enumerate(all_raw_t[1:], start=2):
                                        if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_trans_to_edit).strip():
                                            target_row_idx = r_idx
                                            break
                                
                                if target_row_idx:
                                    sh.update_cell(target_row_idx, 2, ed_t_inv)
                                    sh.update_cell(target_row_idx, 3, ed_t_type)
                                    sh.update_cell(target_row_idx, 4, str(c_id_ed))
                                    sh.update_cell(target_row_idx, 5, ed_t_comp)
                                    sh.update_cell(target_row_idx, 6, ed_t_year)
                                    sh.update_cell(target_row_idx, 7, ed_inv_date.strftime('%Y-%m-%d'))
                                    sh.update_cell(target_row_idx, 8, ed_pay_date.strftime('%Y-%m-%d'))
                                    sh.update_cell(target_row_idx, 10, float(ed_t_amount))
                                    sh.update_cell(target_row_idx, 11, ed_t_method) # Ενημέρωση στήλης 11 (K)
                                    
                                    try:
                                        sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                        master_rows = sh_master.get_all_values()
                                        if len(master_rows) > 1:
                                            m_header = [h.strip() for h in master_rows[0]]
                                            m_inv_idx = m_header.index('invoice_number') if 'invoice_number' in m_header else 5
                                            
                                            for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                if len(m_row) > m_inv_idx and str(m_row[m_inv_idx]).strip() == old_inv_num:
                                                    master_type = "ΕΣΟΔΟ" if ed_t_type == "Έσοδα" else "ΕΞΟΔΟ"
                                                    sh_master.update_cell(m_idx, 3, master_type)
                                                    sh_master.update_cell(m_idx, 5, ed_t_comp)
                                                    sh_master.update_cell(m_idx, 6, ed_t_inv)
                                                    sh_master.update_cell(m_idx, 7, float(ed_t_amount))
                                                    sh_master.update_cell(m_idx, 8, ed_t_method)
                                                    sh_master.update_cell(m_idx, 9, ed_t_year)
                                                    break
                                    except Exception:
                                        pass

                                st.success("✅ Η συναλλαγή τροποποιήθηκε επιτυχώς και ενημερώθηκε το Master!")
                                time.sleep(1.2)
                                
                            except Exception as edit_err:
                                st.error(f"Σφάλμα τροποποίησης: {edit_err}")
                
                    if 'target_row_idx' in locals() and target_row_idx is None and st.session_state.get('form_submitted', False):
                        st.error("Δεν βρέθηκε η εγγραφή στη βάση.")
        else:
            st.info("Δεν βρέθηκαν καταχωρημένες εταιρικές συναλλαγές.")
    #--------------------------------------------------------------
    elif st.session_state.page == "view_customers":
        if st.button("⬅️ Μενού", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "menu"; st.rerun()

        st.markdown("---")
        st.markdown("### 🛑 Διαχείριση Διακοπών & Καταστάσεων")

        df_cust = load_specific_db("db_customers")
        df_reg = load_specific_db("db_registrations")

        if df_cust.empty or df_reg.empty:
            st.warning("Δεν βρέθηκαν δεδομένα πελατών ή εγγραφών.")
        else:
            df_cust['display_name'] = (df_cust['customer_lastname'].astype(str) + " " + 
                                       df_cust['customer_firstname'].astype(str) + " (ID: " + 
                                       df_cust['id_customer'].astype(str) + ")")
            
            options = ["-- Επιλέξτε Πελάτη --"] + df_cust['display_name'].tolist()
            selected_student = st.selectbox("🔍 Επιλογή πελάτη για αλλαγή κατάστασης", options=options)

            if selected_student != "-- Επιλέξτε Πελάτη --" and "ID: " in selected_student:
                try:
                    stud_id_str = selected_student.split("ID: ")[1].replace(")", "").strip()
                    
                    df_reg['id_customer_str'] = df_reg['id_customer'].astype(str).str.strip()
                    stud_regs = df_reg[df_reg['id_customer_str'] == stud_id_str]

                    if stud_regs.empty:
                        st.info(f"Ο πελάτης (ID: {stud_id_str}) δεν έχει εγγραφές.")
                    else:
                        for idx, r in stud_regs.iterrows():
                            col_r1, col_r2, col_r3 = st.columns([2, 2, 1])
                            with col_r1:
                                st.write(f"📅 **Έτος:** {r['school_year']} | 📚 **Πρόγραμμα:** {r['program']}")
                            with col_r2:
                                current_status = str(r['status']).strip().upper()
                                if current_status == 'FINISH':
                                    status_label = "🏁 FINISH"
                                elif current_status == 'ACTIVE':
                                    status_label = "🟢 ΕΝΕΡΓΟΣ"
                                else:
                                    status_label = f"🛑 ΔΙΑΚΟΠΗ ({r.get('stop_date', '')})"
                                st.write(f"**Κατάσταση:** {status_label}")
                            with col_r3:
                                if st.button("⚙️ Αλλαγή Κατάστασης", key=f"status_btn_{r['id_registration']}", disabled=st.session_state.is_submitting):
                                    st.session_state[f"show_status_form_{r['id_registration']}"] = True

                            # Ενιαία Φόρμα χειροκίνητης επιλογής κατάστασης (ACTIVE, FINISH, STOP)
                            if st.session_state.get(f"show_status_form_{r['id_registration']}", False):
                                with st.form(key=f"status_form_{r['id_registration']}"):
                                    st.write(f"🔄 **Ενημέρωση Κατάστασης Εγγραφής (ID: {r['id_registration']})**")
                                    
                                    # Επιλογή κατάστασης
                                    status_options = ['ACTIVE', 'FINISH', 'STOP']
                                    curr_idx = status_options.index(current_status) if current_status in status_options else 0
                                    chosen_status = st.selectbox("Επιλογή Κατάστασης", options=status_options, index=curr_idx)
                                    
                                    stop_d = st.date_input("Ημερομηνία Διακοπής (αν είναι STOP)", value=datetime.today())
                                    
                                    try:
                                        curr_adj = float(str(r.get('adjusted_cost', 0)).replace(',', '.')) if str(r.get('adjusted_cost', '')).strip() != '' else 0.0
                                    except:
                                        curr_adj = 0.0
                                        
                                    new_cost = st.number_input("Τελικό αναπροσαρμοσμένο ποσό (€)", value=curr_adj, min_value=0.0, step=10.0)
                                    
                                    c1, c2 = st.columns(2)
                                    if c1.form_submit_button("💾 Οριστικοποίηση"):
                                        try:
                                            gc = get_gspread_client()
                                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                            
                                            # Ακριβής ανεύρεση γραμμής βάσει ID Πελάτη, Έτους και Προγράμματος (αποφυγή λαθών του sh.find)
                                            all_rows = sh.get_all_values()
                                            target_row_idx = None
                                            
                                            if len(all_rows) > 1:
                                                header = [h.strip().lower() for h in all_rows[0]]
                                                id_col_idx = next((i for i, h in enumerate(header) if 'customer' in h), 1)
                                                year_col_idx = next((i for i, h in enumerate(header) if 'year' in h or 'έτος' in h), 4)
                                                prog_col_idx = next((i for i, h in enumerate(header) if 'prog' in h or 'προγρ' in h), 5)
                                                
                                                for r_idx, sheet_row in enumerate(all_rows[1:], start=2):
                                                    if len(sheet_row) > max(id_col_idx, year_col_idx, prog_col_idx):
                                                        sheet_c_id = str(sheet_row[id_col_idx]).strip()
                                                        sheet_yr = str(sheet_row[year_col_idx]).strip()
                                                        sheet_prg = str(sheet_row[prog_col_idx]).strip()
                                                        
                                                        if (sheet_c_id == str(stud_id_str) and 
                                                            sheet_yr == str(r['school_year']).strip() and 
                                                            sheet_prg == str(r['program']).strip()):
                                                            target_row_idx = r_idx
                                                            break
                                            
                                            if target_row_idx:
                                                sh.update_cell(target_row_idx, 8, chosen_status) # Status
                                                if chosen_status == 'STOP':
                                                    sh.update_cell(target_row_idx, 9, stop_d.strftime('%Y-%m-%d'))
                                                else:
                                                    sh.update_cell(target_row_idx, 9, '') # Καθαρισμός ημερομηνίας
                                                    
                                                sh.update_cell(target_row_idx, 10, int(new_cost)) # Ποσό
                                                
                                                st.session_state[f"show_status_form_{r['id_registration']}"] = False
                                                st.success(f"✅ Η κατάσταση ενημερώθηκε σε {chosen_status}!")
                                                time.sleep(1.0)
                                                st.rerun()
                                            else:
                                                st.error("Δεν βρέθηκε η ακριβής εγγραφή στη βάση δεδομένων.")
                                        except Exception as e:
                                            st.error(f"Σφάλμα: {e}")
                                            
                                    if c2.form_submit_button("Ακύρωση"):
                                        st.session_state[f"show_status_form_{r['id_registration']}"] = False
                                        st.rerun()
                                        
                except Exception as e:
                    st.error(f"Σφάλμα ανάγνωσης: {e}")
#------------------------------------------------------------------------------------------------------------------    
    # 3. ΣΕΛΙΔΑ: ΠΡΩΤΗ ΚΑΤΑΧΩΡΗΣΗ ΠΕΛΑΤΗ
    elif st.session_state.page == "register":
        if st.button("⬅️ Μενού", key="menu_reg", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 

        st.subheader("Νέα Εγγραφή Πελάτη")
        cat = st.radio("Κατηγορία Πελάτη", ["Μαθητής", "Ενήλικας"], key="reg_cat")

        with st.form("reg_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                cf = st.text_input("Όνομα", key="reg_cf").strip()
                cl = st.text_input("Επώνυμο", key="reg_cl").strip()
                bd = st.date_input("Ημ. Γέννησης", value=datetime(2050,1,1), min_value=datetime(1940,1,1), key="reg_bd")
                tel = st.text_input("Τηλέφωνο", key="reg_tel").strip()
                pf = st.text_input("Όνομα Πατρός", key="reg_pf")
                pl = st.text_input("Επώνυμο Πατρός", key="reg_pl")
                mf = st.text_input("Όνομα Μητρός", key="reg_mf")
                passp = st.text_input("ΑΔΤ / Διαβατήριο", key="reg_passp")
                
            with c2:
                em = st.text_input("Email", key="reg_em").strip()
                adr = st.text_input("Διεύθυνση", key="reg_adr").strip()
                
                if cat == "Ενήλικας":
                    afm = st.text_input("ΑΦΜ", key="reg_afm")
                    town = st.text_input("Πόλη", key="reg_town")
                    zip_code = st.text_input("ΤΚ", key="reg_zip")
                    tax_office = st.text_input("ΔΟΥ (Tax Office)", key="reg_tax")
                    amka = st.text_input("ΑΜΚΑ", key="reg_amka")
                    iban = st.text_input("IBAN", key="reg_iban")
                    bank = st.text_input("Τράπεζα", key="reg_bank")
                else:
                    afm = ""
                    town = ""
                    zip_code = ""
                    tax_office = ""
                    amka = ""
                    iban = ""
                    bank = ""
                vouch = ""

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

                    cf_upper = uppercase_clean(cf)
                    cl_upper = uppercase_clean(cl)
                    pf_upper = uppercase_clean(pf)
                    pl_upper = uppercase_clean(pl)
                    mf_upper = uppercase_clean(mf)
                    passp_upper = uppercase_clean(passp)
                    adr_upper = uppercase_clean(adr)
                    town_upper = uppercase_clean(town)
                    bank_upper = uppercase_clean(bank)
                    em_clean = str(em).strip().lower()

                    row_data = [
                        next_id,                    # 1. id_customer
                        cf_upper,                   # 2. customer_firstname
                        cl_upper,                   # 3. customer_lastname
                        pf_upper,                   # 4. parent_firstname
                        pl_upper,                   # 5. parent_lastname
                        tel,                        # 6. telephone
                        em_clean,                   # 7. email
                        adr_upper,                  # 8. address
                        bd.strftime('%Y-%m-%d'),    # 9. birthday_date
                        datetime.today().strftime('%Y-%m-%d'), # 10. register_date
                        cat,                        # 11. customer_category
                        afm,                        # 12. customer_afm
                        mf_upper,                   # 13. mother_name
                        passp_upper,                # 14. passport
                        town_upper,                 # 15. town
                        zip_code,                   # 16. zipCode
                        tax_office,                 # 17. Tax Office
                        amka,                       # 18. amka
                        iban,                       # 19. iban
                        bank_upper                  # 20. bank
                    ]
                    sh.append_row(row_data)

                    #load_specific_db.clear()
                    st.success("✅ Επιτυχής καταχώρηση!")
                    time.sleep(1.5)
                    st.session_state.is_submitting = False
                    st.rerun()
                    
                except Exception as e:
                    st.session_state.is_submitting = False
                    st.error(f"Σφάλμα κατά την αποθήκευση: {e}")     
     #------------------------------------------------------------------------------------------       
    elif st.session_state.page == "renew_registration":
        if st.button("⬅️ Μενού", key="menu_renew", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Εισαγωγή ή Τροποποίηση Πελάτη σε Τμήμα")
    
        cust_df = load_specific_db("db_customers")
        progs_df = load_specific_db("db_programs")
        classes_df = load_specific_db("db_class")
    
        if not cust_df.empty:
            cust_df['display'] = (cust_df['customer_lastname'].astype(str) + " " + 
                                   cust_df['customer_firstname'].astype(str) + " (ID: " + 
                                   cust_df['id_customer'].astype(str) + ")")
        
            sel_cust = st.selectbox("🔍 Επιλογή πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist())
        
            if sel_cust != "-- Επιλέξτε --":
                row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = str(row['id_customer']).strip()
            
                tab_annual, tab_summer = st.tabs(["📅 Ετήσιο Πρόγραμμα", "☀️ SummerCamp/PreSchool"])
            
                with tab_annual:
                    regs_df = load_specific_db("db_registrations")
                    user_regs = regs_df[regs_df['id_customer'].astype(str).str.strip() == c_id]
                    
                    modes = ["Νέα Εγγραφή"]
                    if not user_regs.empty:
                        modes.append("Τροποποίηση")
                    
                    mode = st.radio("Ενέργεια", modes, key="mode_ann")
                    reg_to_edit = None
                    
                    if mode == "Τροποποίηση":
                        user_regs['display_reg'] = user_regs['school_year'].astype(str) + " - " + user_regs['program'].astype(str) + " (" + user_regs['class'].astype(str) + ") - [ID: " + user_regs['id_registration'].astype(str) + "]"
                        sel_reg = st.selectbox("Επιλογή εγγραφής", user_regs['display_reg'].tolist())
                        reg_to_edit = user_regs[user_regs['display_reg'] == sel_reg].iloc[0]
                    
                    # 1. Κωδικός Καύας
                    kayas_val = reg_to_edit['kayas'] if reg_to_edit is not None else ""
                    kayas = st.text_input("Κωδικός ΚΑΥΑΣ", value=kayas_val).strip()
                    
                    # 2. Σχολικό Έτος
                    default_year_idx = 0
                    if reg_to_edit is not None:
                        target_yr = str(reg_to_edit.get('school_year', '')).strip()
                        years_list = [f"{y}-{y+1}" for y in range(2025, 2030)]
                        if target_yr in years_list:
                            default_year_idx = years_list.index(target_yr)

                    s_year = st.selectbox("📅 Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2025, 2030)], index=default_year_idx, key="ann_s_year")
                    
                    # 3. ΕΠΙΛΟΓΗ ΠΡΟΓΡΑΜΜΑΤΟΣ
                    prog_list = [p for p in progs_df['program_name'].tolist() if p not in ["SummerCamp", "PreSchool"]] if not progs_df.empty else []
                    
                    default_prog_idx = 0
                    if reg_to_edit is not None:
                        target_pr = str(reg_to_edit.get('program', '')).strip()
                        if target_pr in prog_list:
                            default_prog_idx = prog_list.index(target_pr)

                    s_prog = st.selectbox("📚 Πρόγραμμα", options=prog_list, index=default_prog_idx, key="ann_s_prog")
                    
                   # --- ΑΣΦΑΛΗΣ ΑΝΑΓΝΩΣΗ ΚΟΣΤΟΥΣ ---
                    default_cost = 0.0
                    if reg_to_edit is not None:
                        raw_cost = reg_to_edit.get('adjusted_cost', 0)
                        try:
                            clean_cost = str(raw_cost).strip().replace(',', '.')
                            default_cost = float(clean_cost) if clean_cost != '' else 0.0
                        except ValueError:
                            default_cost = 0.0
                    
                    adj_cost = st.number_input("💰 Συμφωνηθέν Ποσό (€)", value=default_cost, step=10.0, key="ann_adj_cost")
                    
                    # 5. ΔΥΝΑΜΙΚΟ ΦΙΛΤΡΑΡΙΣΜΑ ΤΜΗΜΑΤΩΝ
                    class_options = ["Xωρίς Τμήμα"]
                    if not classes_df.empty:
                        classes_df['clean_prog'] = classes_df['program_name'].apply(lambda x: remove_accents(str(x)).strip().upper())
                        clean_selected_prog = remove_accents(str(s_prog)).strip().upper()
                        
                        matched = classes_df[classes_df['clean_prog'] == clean_selected_prog]
                        if not matched.empty:
                            class_options += matched['class_name'].unique().tolist()
                    
                    default_class_idx = 0
                    if reg_to_edit is not None:
                        target_cl = str(reg_to_edit.get('class', '')).strip()
                        if target_cl in class_options:
                            default_class_idx = class_options.index(target_cl)

                    sel_class = st.selectbox("🏫 Τμήμα", options=class_options, index=default_class_idx, key="ann_sel_class")
                    
                    # 6. Κουμπί Υποβολής
                    if st.button("🚀 Οριστικοποίηση Εγγραφής", key="btn_submit_annual", disabled=st.session_state.is_submitting):
                        try:
                            gc = get_gspread_client()
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                            
                            if mode == "Νέα Εγγραφή":
                                records = sh.get_all_records()
                                df_reg = pd.DataFrame(records)
                                
                                if not df_reg.empty and 'id_registration' in df_reg.columns:
                                    numeric_ids = pd.to_numeric(df_reg['id_registration'], errors='coerce')
                                    max_id = numeric_ids.max()
                                    next_reg = int(max_id) + 1 if pd.notna(max_id) else 1
                                else:
                                    next_reg = 1
                                
                                sh.append_row([next_reg, int(c_id), str(row['customer_firstname']), str(row['customer_lastname']), str(s_year), str(s_prog), str(sel_class), "ACTIVE", "", adj_cost, kayas])
                            else:
                                # ΧΡΗΣΗ ΑΥΣΤΗΡΑ ΤΟΥ id_registration ΓΙΑ ΑΠΟΛΥΤΗ ΑΚΡΙΒΕΙΑ (Λύνει το πρόβλημα με τα αδέλφια)
                                target_reg_id = str(reg_to_edit['id_registration']).strip()
                                all_reg_values = sh.get_all_values()
                                target_row_idx = None
                                
                                if len(all_reg_values) > 1:
                                    for r_idx, sheet_row in enumerate(all_reg_values[1:], start=2):
                                        if len(sheet_row) > 0 and str(sheet_row[0]).strip() == target_reg_id:
                                            target_row_idx = r_idx
                                            break
                                            
                                if target_row_idx:
                                    # Ενημέρωση γραμμής βάσει της ακριβούς θέσης της (row)
                                    sh.update_cell(target_row_idx, 5, str(s_year))      # school_year
                                    sh.update_cell(target_row_idx, 6, str(s_prog))      # program
                                    sh.update_cell(target_row_idx, 7, str(sel_class))   # class
                                    sh.update_cell(target_row_idx, 10, adj_cost)        # adjusted_cost
                                    sh.update_cell(target_row_idx, 11, str(kayas))      # kayas
                                else:
                                    st.error("Δεν βρέθηκε η εγγραφή στη βάση δεδομένων.")
                                
                                # --- ΑΥΤΟΜΑΤΗ ΕΝΗΜΕΡΩΣΗ ΤΟΥ db_payments ΜΕ ΒΑΣΗ ΤΟ ID ΚΑΙ ΤΟ ΠΡΟΓΡΑΜΜΑ ---
                                try:
                                    old_prog = str(reg_to_edit.get('program', '')).strip()
                                    new_prog = str(s_prog).strip()
                                    if old_prog != new_prog:
                                        sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                        pays_vals = sh_pays.get_all_values()
                                        if len(pays_vals) > 1:
                                            p_header = [h.strip() for h in pays_vals[0]]
                                            p_cust_idx = next((i for i, h in enumerate(p_header) if 'customer' in h.lower() or 'id_customer' in h.lower()), 1)
                                            p_prog_idx = next((i for i, h in enumerate(p_header) if 'prog' in h.lower() or 'προγρ' in h.lower()), None)
                                            
                                            if p_prog_idx is not None:
                                                for p_i, p_row in enumerate(pays_vals[1:], start=2):
                                                    match_cust = len(p_row) > p_cust_idx and str(p_row[p_cust_idx]).strip() == str(c_id).strip()
                                                    # Ελέγχουμε και το παλιό πρόγραμμα για να αλλάξει μόνο η σωστή πληρωμή αυτού του παιδιού
                                                    match_prog = len(p_row) > p_prog_idx and str(p_row[p_prog_idx]).strip() == old_prog
                                                    if match_cust and match_prog:
                                                        sh_pays.update_cell(p_i, p_prog_idx + 1, new_prog)
                                except Exception as p_err:
                                    st.warning(f"Η εγγραφή ενημερώθηκε αλλά υπήρξε θέμα στις πληρωμές: {p_err}")
                            
                            #load_specific_db.clear()
                            st.success(f"✅ Η εγγραφή για τον πελάτη {row['customer_lastname']} {row['customer_firstname']} ενημερώθηκε επιτυχώς!")
                            time.sleep(1.5)
                            st.rerun()
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")
                
                with tab_summer:
                    mode_sum = st.radio("Ενέργεια", ["Νέα Εγγραφή", "Τροποποίηση"], key="mode_sum")
                    sum_to_edit = None
                    
                    if mode_sum == "Τροποποίηση":
                        sum_df = load_specific_db("db_summer")
                        if sum_df.empty:
                            st.warning("Το φύλλο db_summer είναι άδειο.")
                        else:
                            sum_df.columns = sum_df.columns.astype(str).str.strip()
                            if 'summer_email' in sum_df.columns:
                                target_email = str(row['email']).strip().lower()
                                user_sum = sum_df[sum_df['summer_email'].astype(str).str.strip().str.lower() == target_email]
                                
                                if not user_sum.empty:
                                    user_sum['display_sum'] = user_sum['summer_year'].astype(str) + " - " + user_sum['summer_program'].astype(str)
                                    sel_sum = st.selectbox("Επιλογή εγγραφής Summer", user_sum['display_sum'].tolist())
                                    sum_to_edit = user_sum[user_sum['display_sum'] == sel_sum].iloc[0]
                                else:
                                    st.info("Δεν βρέθηκαν εγγραφές Summer για τον συγκεκριμένο πελάτη.")
        
                    with st.form("summer_form"):
                        s_prog_sum = st.selectbox("📚 Επιλογή Προγράμματος", ["SummerCamp", "PreSchool"], key="prog_sum")
                        sel_y = st.selectbox("📅 Επιλογή Σχολικού Έτους", [f"{y}-{y+1}" for y in range(2025, 2030)], key="year_sum")
                        
                        default_weeks = int(sum_to_edit['summer_weeks']) if sum_to_edit is not None else 1
                        num_weeks = st.number_input("Εβδομάδες", min_value=1, max_value=10, value=default_weeks, step=1, key="weeks_sum")
                        
                        # Υπολογισμός αρχικής τιμής ανά εβδομάδα (αν τροποποιούμε, διαιρούμε το αποθηκευμένο κόστος με τις εβδομάδες)
                        default_weekly_cost = 0.0
                        if sum_to_edit is not None:
                            try:
                                total_saved = float(str(sum_to_edit.get('summer_cost', 0)).replace(',', '.') or 0)
                                saved_weeks = int(sum_to_edit.get('summer_weeks', 1))
                                default_weekly_cost = total_saved / saved_weeks if saved_weeks > 0 else total_saved
                            except:
                                default_weekly_cost = 0.0

                        # Το πεδίο για το ποσό ανά εβδομάδα
                        sum_agreed_cost = st.number_input("💰 Συμφωνηθέν Ποσό ανά Εβδομάδα (€)", value=default_weekly_cost, min_value=0.0, step=10.0, key="sum_agreed_cost_input")
                        
                        submitted_sum = st.form_submit_button("🚀 Αποθήκευση Summer", disabled=st.session_state.is_submitting)
                        
                        if submitted_sum:
                            st.session_state.is_submitting = True
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_summer")
                                
                                # ΑΚΡΙΒΗΣ ΠΟΛΛΑΠΛΑΣΙΑΣΜΟΣ: Εβδομάδες × Συμφωνηθέν ποσό (ανεξαρτήτως προγράμματος)
                                total_cost = float(num_weeks) * float(sum_agreed_cost)
                                
                                if mode_sum == "Νέα Εγγραφή":
                                    sh.append_row([int(len(sh.get_all_values())), str(row['customer_firstname']), str(row['customer_lastname']), str(row['parent_firstname']), str(row['parent_lastname']), str(row['telephone']), str(row['email']), num_weeks, total_cost, str(sel_y), str(s_prog_sum), "ACTIVE"])
                                else:
                                    old_year_val = str(sum_to_edit['summer_year']).strip()
                                    all_rows = sh.get_all_values()
                                    for i, r in enumerate(all_rows[1:]):
                                        if str(r[6]).strip().lower() == str(row['email']).strip().lower() and str(r[9]).strip() == old_year_val:
                                            sh.update(f"H{i+2}:K{i+2}", [[num_weeks, total_cost, str(sel_y), str(s_prog_sum)]])
                                            break
                                    
                                    if old_year_val != str(sel_y):
                                        target_email_str = str(row['email']).strip().lower()
                                        
                                        try:
                                            sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payment_summer")
                                            pays_values = sh_pays.get_all_values()
                                            if len(pays_values) > 1:
                                                p_header = [h.strip() for h in pays_values[0]]
                                                email_col_idx = p_header.index('summer_email_customer') if 'summer_email_customer' in p_header else 4
                                                year_col_idx = p_header.index('school_yearSummer') if 'school_yearSummer' in p_header else 9
                                                rec_col_idx = p_header.index('receipt_numberSummer') if 'receipt_numberSummer' in p_header else 6
                                                
                                                updated_receipts = []
                                                for p_idx, p_row in enumerate(pays_values[1:], start=2):
                                                    if len(p_row) > email_col_idx and str(p_row[email_col_idx]).strip().lower() == target_email_str:
                                                        if len(p_row) > year_col_idx and str(p_row[year_col_idx]).strip() == old_year_val:
                                                            sh_pays.update_cell(p_idx, year_col_idx + 1, str(sel_y))
                                                            if len(p_row) > rec_col_idx:
                                                                updated_receipts.append(str(p_row[rec_col_idx]).strip())
                                                
                                                if updated_receipts:
                                                    try:
                                                        sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                                        master_rows = sh_master.get_all_values()
                                                        if len(master_rows) > 1:
                                                            m_header = [h.strip() for h in master_rows[0]]
                                                            m_inv_idx = m_header.index('invoice_number') if 'invoice_number' in m_header else 5
                                                            m_year_idx = m_header.index('school_year') if 'school_year' in m_header else 8
                                                            
                                                            for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                                if len(m_row) > m_inv_idx and str(m_row[m_inv_idx]).strip() in updated_receipts:
                                                                    sh_master.update_cell(m_idx, m_year_idx + 1, str(sel_y))
                                                    except Exception:
                                                        pass
                                        except Exception as p_err:
                                            st.warning(f"Σφάλμα κατά την ενημέρωση πληρωμών: {p_err}")
                                    
                                #load_master_db.clear()
                                #load_specific_db.clear()
                                st.success("✅ Η εγγραφή Summer αποθηκεύτηκε επιτυχώς!")
                                time.sleep(1.5)
                                st.session_state.is_submitting = False
                                st.rerun()
                                
                            except Exception as e:
                                st.session_state.is_submitting = False
                                st.error(f"Σφάλμα: {e}")
#--------------------------------------------------------------------------------
    elif st.session_state.page == "add_payment":
        if st.button("⬅️ Μενού", key="menu_pay", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 
            
        st.subheader("Διαχείριση Πληρωμών")
        
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        progs_df = load_specific_db("db_programs")
        pays_df = load_specific_db("db_payments")
        
        for df in [cust_df, reg_df, pays_df, progs_df]:
            if not df.empty: df.columns = df.columns.astype(str).str.strip()

        if not cust_df.empty:
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce').fillna(0).astype(int)
            cust_df['customer_lastname'] = cust_df['customer_lastname'].astype(str).str.strip()
            cust_df['customer_firstname'] = cust_df['customer_firstname'].astype(str).str.strip()
            
            # --- ΑΛΦΑΒΗΤΙΚΗ ΤΑΞΙΝΟΜΗΣΗ ΒΑΣΕΙ ΕΠΩΝΥΜΟΥ ΚΑΙ ΟΝΟΜΑΤΟΣ ---
            cust_df = cust_df.sort_values(by=['customer_lastname', 'customer_firstname'], ascending=[True, True])
            
            cust_df['display'] = (cust_df['customer_lastname'] + " " + cust_df['customer_firstname'] + " (ID: " + cust_df['id_customer'].astype(str) + ")")
            
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη για Πληρωμή", options=["-- Επιλέξτε --"] + cust_df['display'].tolist(), index=0)
            
            if sel_cust != "-- Επιλέξτε --":
                c_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = int(c_row['id_customer'])
                
                # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ ΠΛΗΡΩΜΩΝ (ΜΟΝΟ ΜΕ ΑΠΥ/ΤΠΥ ΚΑΙ ΟΧΙ ΑΡΝΗΤΙΚΕΣ/ΜΗΔΕΝΙΚΕΣ) ---
                if not pays_df.empty:
                    pays_df['id_cust_num'] = pd.to_numeric(pays_df['id_customer'], errors='coerce')
                    cust_pays = pays_df[pays_df['id_cust_num'] == float(c_id)].copy()
                    
                    if not cust_pays.empty:
                        # Φιλτράρισμα: Αριθμός απόδειξης να ξεκινάει από ΑΠΥ ή ΤΠΥ ΚΑΙ το ποσό να μην είναι <= 0
                        def is_valid_receipt(row):
                            rec_str = str(row.get('receipt_number', '')).strip().upper()
                            amt_val = float(pd.to_numeric(row.get('amount_paid', 0), errors='coerce') or 0.0)
                            has_prefix = rec_str.startswith('ΑΠΥ') or rec_str.startswith('ΤΠΥ')
                            return has_prefix and amt_val > 0

                        valid_cust_pays = cust_pays[cust_pays.apply(is_valid_receipt, axis=1)].copy()

                        if not valid_cust_pays.empty:
                            st.markdown("---")
                            st.markdown(f"##### 📋 Υπάρχουσες Πληρωμές Πελάτη (ΑΠΥ/ΤΠΥ): {c_row['customer_lastname']} {c_row['customer_firstname']}")
                            
                            display_cols = [c for c in ['id_payment', 'payment_date', 'receipt_number', 'amount_paid', 'school_year', 'pay_method', 'program'] if c in valid_cust_pays.columns]
                            display_pays = valid_cust_pays[display_cols].copy()
                            display_pays.columns = ['ID Πληρωμής', 'Ημερομηνία', 'Αρ. Απόδειξης', 'Ποσό (€)', 'Έτος', 'Τρόπος', 'Πρόγραμμα'][:len(display_cols)]
                            st.dataframe(display_pays, use_container_width=True, hide_index=True)
                            
                            pay_ids_list = valid_cust_pays['id_payment'].astype(str).tolist()
                            sel_pay_to_edit = st.selectbox("Επιλογή ID Πληρωμής για Επεξεργασία / Διαγραφή", ["-- Επιλέξτε ID --"] + pay_ids_list, key="sel_pay_edit_box")
                            
                            if sel_pay_to_edit != "-- Επιλέξτε ID --":
                                matched_p_row = valid_cust_pays[valid_cust_pays['id_payment'].astype(str) == str(sel_pay_to_edit)].iloc[0]
                                old_receipt = str(matched_p_row.get('receipt_number', '')).strip()
                                old_year = str(matched_p_row.get('school_year', '')).strip()
                                old_prog = str(matched_p_row.get('program', '')).strip()
                                
                                try:
                                    def_pay_dt = datetime.strptime(str(matched_p_row.get('payment_date', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                                except:
                                    def_pay_dt = datetime.today().date()
                                
                                with st.form("edit_payment_form"):
                                    st.markdown(f"**Επεξεργασία ή Διαγραφή Πληρωμής με ID: {sel_pay_to_edit}**")
                                    ed_date = st.date_input("Ημερομηνία Πληρωμής", value=def_pay_dt)
                                    ed_rec = st.text_input("Αριθμός Απόδειξης", value=str(matched_p_row.get('receipt_number', ''))).strip()
                                    ed_amt = st.number_input("Ποσό (€)", min_value=0.0, value=float(pd.to_numeric(matched_p_row.get('amount_paid', 0), errors='coerce') or 0.0), format="%.2f")
                                    ed_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "Εθνική", "Eurobank", "POS"], index=["Μετρητά", "Εθνική", "Eurobank", "POS"].index(str(matched_p_row.get('pay_method', 'Μετρητά'))) if str(matched_p_row.get('pay_method', 'Μετρητά')) in ["Μετρητά", "Εθνική", "Eurobank", "POS"] else 0)
                                    
                                    col_btn1, col_btn2 = st.columns(2)
                                    submit_edit = col_btn1.form_submit_button("💾 Οριστικοποίηση Τροποποίησης")
                                    submit_del = col_btn2.form_submit_button("🗑️ Διαγραφή Πληρωμής")

                                    
                                if submit_edit:
                                    if not ed_rec or ed_amt <= 0:
                                        st.warning("⚠️ Συμπλήρωσε αριθμό απόδειξης και ποσό!")
                                    else:
                                        try:
                                            gc = get_gspread_client()
                                            sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                            all_raw_pays = sh_pays.get_all_values()
                                            target_row_idx = None
                                            if len(all_raw_pays) > 1:
                                                for r_idx, sheet_row in enumerate(all_raw_pays[1:], start=2):
                                                    if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_pay_to_edit).strip():
                                                        target_row_idx = r_idx
                                                        break
                                            
                                            if target_row_idx:
                                                sh_pays.update_cell(target_row_idx, 5, ed_date.strftime('%Y-%m-%d'))
                                                sh_pays.update_cell(target_row_idx, 6, ed_rec)
                                                sh_pays.update_cell(target_row_idx, 7, float(ed_amt))  # Ποσό
                                                sh_pays.update_cell(target_row_idx, 9, ed_method)  # Τρόπος (αναλόγως σειράς στη βάση)
                                                
                                                try:
                                                    sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                                    master_rows = sh_master.get_all_values()
                                                    if len(master_rows) > 1:
                                                        for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                            if len(m_row) > 5 and str(m_row[5]).strip() == old_receipt:
                                                                sh_master.update_cell(m_idx, 2, ed_date.strftime('%Y-%m-%d'))
                                                                sh_master.update_cell(m_idx, 6, ed_rec)
                                                                sh_master.update_cell(m_idx, 7, float(ed_amt))
                                                                sh_master.update_cell(m_idx, 8, ed_method)
                                                                break
                                                except Exception:
                                                    pass

                                                # --- ΕΛΕΓΧΟΣ ΚΑΙ ΕΝΗΜΕΡΩΣΗ STATUS (ACTIVE/FINISH) ΜΕΤΑ ΤΗΝ ΤΡΟΠΟΠΟΙΗΣΗ ---
                                                try:
                                                    temp_pays = pays_df.copy()
                                                    temp_pays.loc[temp_pays['id_payment'].astype(str).str.strip() == str(sel_pay_to_edit), 'amount_paid'] = float(ed_amt)
                                                    
                                                    p_matched_edit = temp_pays[
                                                        (temp_pays['id_customer'].astype(str).str.strip() == str(c_id)) & 
                                                        (temp_pays['school_year'].astype(str).str.strip() == old_year) & 
                                                        (temp_pays['program'].astype(str).str.strip() == old_prog)
                                                    ]
                                                    total_paid_after_edit = pd.to_numeric(p_matched_edit['amount_paid'], errors='coerce').fillna(0).sum()

                                                    target_reg_row = reg_df[
                                                        (reg_df['id_customer'].astype(str).str.strip() == str(c_id)) & 
                                                        (reg_df['school_year'].astype(str).str.strip() == old_year) & 
                                                        (reg_df['program'].astype(str).str.strip() == old_prog)
                                                    ]
                                                    
                                                    agreed_cost = 0.0
                                                    if not target_reg_row.empty:
                                                        r_item = target_reg_row.iloc[0]
                                                        adj_c = str(r_item.get('adjusted_cost', '0')).replace(',', '.')
                                                        if adj_c.strip() != '' and float(adj_c) > 0:
                                                            agreed_cost = float(adj_c)
                                                        else:
                                                            p_row_c = progs_df[progs_df['program_name'].astype(str).str.strip().str.lower() == old_prog.lower()]
                                                            if not p_row_c.empty:
                                                                agreed_cost = float(p_row_c.iloc[0].get('program_cost', 0))

                                                    sh_reg_edit = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                                    reg_rows_e = sh_reg_edit.get_all_values()
                                                    if len(reg_rows_e) > 1:
                                                        h_e = [h.strip() for h in reg_rows_e[0]]
                                                        id_c_ie = next((i for i, h in enumerate(h_e) if 'customer' in h.lower()), 1)
                                                        yr_ie = next((i for i, h in enumerate(h_e) if 'year' in h.lower()), 4)
                                                        pr_ie = next((i for i, h in enumerate(h_e) if 'prog' in h.lower()), 5)
                                                        st_ie = next((i for i, h in enumerate(h_e) if 'status' in h.lower()), 7)
                                                        
                                                        for ri_e, rvals_e in enumerate(reg_rows_e[1:], start=2):
                                                            if len(rvals_e) > max(id_c_ie, yr_ie, pr_ie, st_ie):
                                                                if str(rvals_e[id_c_ie]).strip() == str(c_id) and str(rvals_e[yr_ie]).strip() == old_year and str(rvals_e[pr_ie]).strip() == old_prog:
                                                                    new_status = "FINISH" if (agreed_cost > 0 and total_paid_after_edit >= agreed_cost) else "ACTIVE"
                                                                    sh_reg_edit.update_cell(ri_e, st_ie + 1, new_status)
                                                                    break
                                                except Exception:
                                                    pass

                                                st.success("✅ Η πληρωμή τροποποιήθηκε επιτυχώς και το status ενημερώθηκε!")
                                                time.sleep(1.2)
                                                
                                            else:
                                                st.error("Δεν βρέθηκε η εγγραφή στη βάση δεδομένων.")
                                        except Exception as edit_err:
                                            st.error(f"Σφάλμα τροποποίησης: {edit_err}")

                                if submit_del:
                                    try:
                                        gc = get_gspread_client()
                                        sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                        all_raw_pays = sh_pays.get_all_values()
                                        target_row_idx = None
                                        if len(all_raw_pays) > 1:
                                            for r_idx, sheet_row in enumerate(all_raw_pays[1:], start=2):
                                                if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_pay_to_edit).strip():
                                                    target_row_idx = r_idx
                                                    break
                                        
                                        if target_row_idx:
                                            sh_pays.delete_rows(target_row_idx)
                                            
                                            remaining_pays = pays_df[
                                                (pays_df['id_payment'].astype(str).str.strip() != str(sel_pay_to_edit)) &
                                                (pays_df['id_customer'].astype(str).str.strip() == str(c_id)) &
                                                (pays_df['school_year'].astype(str).str.strip() == old_year) &
                                                (pays_df['program'].astype(str).str.strip() == old_prog)
                                            ]
                                            total_paid_rem = pd.to_numeric(remaining_pays['amount_paid'], errors='coerce').fillna(0).sum()

                                            try:
                                                sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                                reg_rows = sh_reg.get_all_values()
                                                if len(reg_rows) > 1:
                                                    h_reg = [h.strip() for h in reg_rows[0]]
                                                    id_c_i = next((i for i, h in enumerate(h_reg) if 'customer' in h.lower()), 1)
                                                    yr_i = next((i for i, h in enumerate(h_reg) if 'year' in h.lower()), 4)
                                                    pr_i = next((i for i, h in enumerate(h_reg) if 'prog' in h.lower()), 5)
                                                    st_i = next((i for i, h in enumerate(h_reg) if 'status' in h.lower()), 7)
                                                    
                                                    for ri, rvals in enumerate(reg_rows[1:], start=2):
                                                        if len(rvals) > max(id_c_i, yr_i, pr_i, st_i):
                                                            if str(rvals[id_c_i]).strip() == str(c_id) and str(rvals[yr_i]).strip() == old_year and str(rvals[pr_i]).strip() == old_prog:
                                                                target_reg_row = reg_df[
                                                                    (reg_df['id_customer'].astype(str).str.strip() == str(c_id)) & 
                                                                    (reg_df['school_year'].astype(str).str.strip() == old_year) & 
                                                                    (reg_df['program'].astype(str).str.strip() == old_prog)
                                                                ]
                                                                agreed_cost = 0.0
                                                                if not target_reg_row.empty:
                                                                    r_item = target_reg_row.iloc[0]
                                                                    adj_c = str(r_item.get('adjusted_cost', '0')).replace(',', '.')
                                                                    if adj_c.strip() != '' and float(adj_c) > 0:
                                                                        agreed_cost = float(adj_c)
                                                                    else:
                                                                        p_row_c = progs_df[progs_df['program_name'].astype(str).str.strip().str.lower() == old_prog.lower()]
                                                                        if not p_row_c.empty:
                                                                            agreed_cost = float(p_row_c.iloc[0].get('program_cost', 0))

                                                                new_status = "FINISH" if (agreed_cost > 0 and total_paid_rem >= agreed_cost) else "ACTIVE"
                                                                sh_reg.update_cell(ri, st_i + 1, new_status)
                                                                break
                                            except Exception:
                                                pass

                                            st.success("🗑️ Η πληρωμή διαγράφηκε επιτυχώς και το status ενημερώθηκε!")
                                            time.sleep(1.2)
                                            
                                        else:
                                            st.error("Δεν βρέθηκε η πληρωμή για διαγραφή.")
                                    except Exception as del_err:
                                        st.error(f"Σφάλμα διαγραφής: {del_err}")
                else:
                    st.info("Ο πελάτης δεν έχει καταχωρημένες πληρωμές.")
                
                st.markdown("---")
                st.markdown("##### ➕ Νέα Καταχώρηση Πληρωμής")

                reg_df['id_customer_numeric'] = pd.to_numeric(reg_df['id_customer'], errors='coerce')
                user_regs = reg_df[reg_df['id_customer_numeric'] == float(c_id)]
                
                if user_regs.empty:
                    st.error("⚠️ Ο πελάτης δεν έχει καταχωρημένη εγγραφή.")
                else:
                    sel_y = st.selectbox("📅 Επιλογή Έτους:", options=sorted(user_regs['school_year'].dropna().astype(str).unique().tolist(), reverse=True))
                    regs_in_year = user_regs[user_regs['school_year'].astype(str) == str(sel_y)]
                    
                    sel_prog = st.selectbox("📚 Επιλογή Προγράμματος:", options=regs_in_year['program'].unique().tolist())
                    
                    if sel_prog:
                        current_regs = regs_in_year[regs_in_year['program'] == sel_prog]
                        
                        total_cost = 0.0
                        for _, row in current_regs.iterrows():
                            adj = str(row.get('adjusted_cost', '0')).replace(',', '.')
                            total_cost += float(adj) if adj.strip() != '' else float(progs_df[progs_df['program_name'] == sel_prog]['program_cost'].iloc[0]) if not progs_df[progs_df['program_name'] == sel_prog].empty else 0.0
                        
                        mask = ((pays_df['id_customer'].astype(int) == c_id) & 
                                (pays_df['school_year'].astype(str) == str(sel_y)) &
                                (pays_df['program'].astype(str) == str(sel_prog)))
                        paid = pd.to_numeric(pays_df.loc[mask, 'amount_paid'], errors='coerce').sum()

                        class_names = ", ".join(current_regs['class'].unique().astype(str))
                        if st.session_state.role not in ["super", "super_admin"]:        
                            st.info(f"""
                            📚 **Πρόγραμμα:** {sel_prog}  
                            🏫 **Τμήματα:** {class_names}  
                            💰 **Συμφωνηθέν Σύνολο:** {total_cost:.2f}€  
                            💸 **Πληρωμένα:** {paid:.2f}€  
                            📉 **Υπόλοιπο:** {total_cost - paid:.2f}€
                            """)                    
                        
                        rec = st.text_input("Αριθμός Απόδειξης").strip()
                        amt = st.number_input("Ποσό (€)", min_value=0.0, value=0.0, format="%.2f")
                        pay_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "Εθνική", "Eurobank", "POS"])
                        
                        if st.button("💳 Αποθήκευση Πληρωμής", disabled=st.session_state.is_submitting):
                            if not rec or amt <= 0:
                                st.warning("⚠️ Συμπλήρωσε αριθμό απόδειξης και ποσό!")
                            else:
                                try:
                                    gc = get_gspread_client()
                                    sh = gc.open(OLD_SHEET_NAME).worksheet("db_payments")
                                    
                                    records = sh.get_all_records()
                                    if records:
                                        df_pays = pd.DataFrame(records)
                                        df_pays.columns = df_pays.columns.astype(str).str.strip()
                                        last_id = pd.to_numeric(df_pays['id_payment'], errors='coerce').max()
                                        next_id = int(last_id) + 1 if pd.notna(last_id) else 1
                                    else:
                                        next_id = 1
                                    
                                    # Καταχώρηση χωρίς το payment_number (αφαιρέθηκε)
                                    sh.append_row([
                                        next_id, 
                                        c_id, 
                                        str(c_row['customer_lastname']), 
                                        str(c_row['customer_firstname']), 
                                        datetime.today().strftime('%Y-%m-%d'), 
                                        rec, 
                                        amt, 
                                        str(sel_y), 
                                        pay_method, 
                                        str(sel_prog)
                                    ])
                                    
                                    check_and_update_status_to_finish(c_id, sel_y, sel_prog)
                                    log_to_finance_master("ΕΣΟΔΟ", sel_prog, f"{c_row['customer_lastname']} {c_row['customer_firstname']}", rec, amt, pay_method, sel_y, "Payments")
                                    
                                    st.success("✅ Η πληρωμή καταχωρήθηκε επιτυχώς!")
                                    time.sleep(1.5)
                                    
                                except Exception as e:
                                    st.error(f"Σφάλμα κατά την καταχώρηση: {e}")
#---------------------------------------------------------------------------------------------------
    elif st.session_state.page == "add_payment_summer":
        if st.button("⬅️ Επιστροφή στο Μενού", key="back_to_menu", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 
            
        st.header("💳 Πληρωμές Summer & Pre-School")
    
        df_summer = load_specific_db("db_summer")
        df_pays = load_specific_db("db_payment_summer")
    
        if df_summer.empty:
            st.warning("Δεν βρέθηκαν εγγραφές στο db_summer.")
        else:
            df_summer.columns = df_summer.columns.astype(str).str.strip()
            df_summer['summer_lastname'] = df_summer['summer_lastname'].astype(str).str.strip()
            df_summer['summer_firstname'] = df_summer['summer_firstname'].astype(str).str.strip()
            
            # --- ΕΓΓΥΗΜΕΝΗ ΕΛΛΗΝΙΚΗ ΑΛΦΑΒΗΤΙΚΗ ΤΑΞΙΝΟΜΗΣΗ ---
            import unicodedata
            def remove_accents(input_str):
                if not isinstance(input_str, str): return ""
                nfkd_form = unicodedata.normalize('NFKD', input_str)
                return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper()

            df_summer['sort_key'] = df_summer['summer_lastname'].apply(remove_accents) + " " + df_summer['summer_firstname'].apply(remove_accents)
            df_summer = df_summer.sort_values(by='sort_key', ascending=True)
            
            # Ενσωμάτωση του summer_id στο label για εύκολη αναγνώριση
            s_id_col = next((c for c in df_summer.columns if 'summer_id' in c.lower() or 'id' in c.lower()), 'summer_id')
            
            df_summer['label'] = (
                df_summer['summer_lastname'] + " " + 
                df_summer['summer_firstname'] + " (ID: " + 
                df_summer[s_id_col].astype(str) + " - " + 
                df_summer['summer_email'].astype(str) + ")"
            )
            
            selected_reg = st.selectbox("🔍 Επιλογή Εγγραφής Πελάτη", options=["-- Επιλέξτε --"] + df_summer['label'].tolist())
            if selected_reg != "-- Επιλέξτε --":
                reg_row = df_summer[df_summer['label'] == selected_reg].iloc[0]
                
                target_summer_id = str(reg_row[s_id_col]).strip()
                email = str(reg_row['summer_email']).strip().lower()
                year = str(reg_row['summer_year']).strip()
                target_lname = str(reg_row['summer_lastname']).strip().lower()
                target_fname = str(reg_row['summer_firstname']).strip().lower()
                
                base_cost = float(str(reg_row['summer_cost']).replace(',', '.')) if reg_row['summer_cost'] else 0.0
                num_weeks = int(reg_row['summer_weeks']) if 'summer_weeks' in reg_row else 1
                total_cost = base_cost  
                
                # --- ΥΠΟΛΟΓΙΣΜΟΣ ΠΛΗΡΩΜΩΝ (ΜΕ ΥΠΟΣΤΗΡΙΞΗ ΚΟΜΜΑΤΟΣ) ---
                total_paid = 0.0
                if not df_pays.empty:
                    df_pays.columns = df_pays.columns.astype(str).str.strip()
                    s_id_pay_col = next((c for c in df_pays.columns if 'summer_id' in c.lower()), None)
                    
                    if s_id_pay_col:
                        mask = df_pays[s_id_pay_col].astype(str).str.strip() == target_summer_id
                    else:
                        df_pays['email_clean'] = df_pays['summer_email_customer'].astype(str).str.strip().str.lower()
                        df_pays['year_clean'] = df_pays['school_yearSummer'].astype(str).str.strip()
                        df_pays['lname_clean'] = df_pays['customer_lastnameSummer'].astype(str).str.strip().str.lower()
                        df_pays['fname_clean'] = df_pays['customer_firstnameSummer'].astype(str).str.strip().str.lower()

                        mask = (
                            (df_pays['email_clean'] == email) & 
                            (df_pays['year_clean'] == year) &
                            (df_pays['lname_clean'] == target_lname) &
                            (df_pays['fname_clean'] == target_fname)
                        )
                    
                    # Διόρθωση εδώ για τα κόμματα
                    total_paid = pd.to_numeric(df_pays.loc[mask, 'amount_paidSummer'].astype(str).str.replace(',', '.'), errors='coerce').sum()
                
                st.info(f"💰 **Συνολικό Κόστος:** {total_cost:.2f}€ | **Ήδη Πληρωμένα:** {total_paid:.2f}€ | **Υπόλοιπο:** {total_cost - total_paid:.2f}€")    
                
                # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ SUMMER ΠΛΗΡΩΜΩΝ ---
                st.markdown("---")
                st.markdown("##### 📋 Υπάρχουσες Πληρωμές Summer")
                
                if not df_pays.empty:
                    if s_id_pay_col:
                        cust_summer_pays = df_pays[df_pays[s_id_pay_col].astype(str).str.strip() == target_summer_id].copy()
                    else:
                        cust_summer_pays = df_pays[mask].copy()
                    
                    if not cust_summer_pays.empty:
                        display_summer_pays = cust_summer_pays[['id_paymentSummer', 'payment_dateSummer', 'receipt_numberSummer', 'amount_paidSummer', 'school_yearSummer', 'pay_methodSummer']].copy()
                        display_summer_pays.columns = ['ID', 'Ημερομηνία', 'Αρ. Απόδειξης', 'Ποσό (€)', 'Έτος', 'Τρόπος']
                        st.dataframe(display_summer_pays, use_container_width=True, hide_index=True)
                        
                        pay_ids_summer = cust_summer_pays['id_paymentSummer'].astype(str).tolist()
                        sel_summer_edit = st.selectbox("Επιλογή ID για Τροποποίηση", ["-- Επιλέξτε ID --"] + pay_ids_summer, key="sel_summer_edit_box")
                        
                        if sel_summer_edit != "-- Επιλέξτε ID --":
                            matched_s_row = cust_summer_pays[cust_summer_pays['id_paymentSummer'].astype(str) == str(sel_summer_edit)].iloc[0]
                            old_rec_summer = str(matched_s_row.get('receipt_numberSummer', '')).strip()
                            
                            try:
                                def_pay_dt_s = datetime.strptime(str(matched_s_row.get('payment_dateSummer', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                            except:
                                def_pay_dt_s = datetime.today().date()
                            
                            with st.form("edit_summer_payment_form"):
                                st.markdown(f"**Επεξεργασία Πληρωμής με ID: {sel_summer_edit}**")
                                ed_date_summer = st.date_input("Ημερομηνία Πληρωμής", value=def_pay_dt_s)
                                ed_rec_summer = st.text_input("Αριθμός Απόδειξης", value=old_rec_summer).strip()
                                
                                # Διόρθωση ανάγνωσης ποσού με κόμμα στην επεξεργασία
                                raw_amt_s = str(matched_s_row.get('amount_paidSummer', 0)).replace(',', '.')
                                default_amt_s = float(pd.to_numeric(raw_amt_s, errors='coerce') or 0.0)
                                ed_amt_summer = st.number_input("Ποσό (€)", min_value=0.0, value=default_amt_s, format="%.2f")
                                
                                ed_method_summer = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "Εθνική", "Eurobank", "POS"], index=["Μετρητά", "Εθνική", "Eurobank", "POS"].index(str(matched_s_row.get('pay_methodSummer', 'Μετρητά'))) if str(matched_s_row.get('pay_methodSummer', 'Μετρητά')) in ["Μετρητά", "Εθνική", "Eurobank", "POS"] else 0)
                                
                                if st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης Summer", disabled=st.session_state.is_submitting):
                                    st.session_state.is_submitting = True
                                    if not ed_rec_summer or ed_amt_summer <= 0:
                                        st.warning("Παρακαλώ συμπληρώστε αριθμό απόδειξης και ποσό.")
                                        st.session_state.is_submitting = False
                                    else:
                                        try:
                                            gc = get_gspread_client()
                                            sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payment_summer")
                                            
                                            all_raw_spays = sh_pays.get_all_values()
                                            target_row_idx = None
                                            if len(all_raw_spays) > 1:
                                                for r_idx, sheet_row in enumerate(all_raw_spays[1:], start=2):
                                                    if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_summer_edit).strip():
                                                        target_row_idx = r_idx
                                                        break
                                            
                                            if target_row_idx:
                                                header_spays = [h.strip() for h in all_raw_spays[0]]
                                                date_idx = header_spays.index('payment_dateSummer') + 1 if 'payment_dateSummer' in header_spays else 6
                                                rec_idx = header_spays.index('receipt_numberSummer') + 1 if 'receipt_numberSummer' in header_spays else 7
                                                amt_idx = header_spays.index('amount_paidSummer') + 1 if 'amount_paidSummer' in header_spays else 9
                                                meth_idx = header_spays.index('pay_methodSummer') + 1 if 'pay_methodSummer' in header_spays else 11

                                                sh_pays.update_cell(target_row_idx, date_idx, ed_date_summer.strftime('%Y-%m-%d'))
                                                sh_pays.update_cell(target_row_idx, rec_idx, ed_rec_summer)
                                                sh_pays.update_cell(target_row_idx, amt_idx, float(ed_amt_summer))
                                                sh_pays.update_cell(target_row_idx, meth_idx, ed_method_summer)
                                                
                                                try:
                                                    sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                                    master_rows = sh_master.get_all_values()
                                                    if len(master_rows) > 1:
                                                        for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                            if len(m_row) > 5 and str(m_row[5]).strip() == old_rec_summer:
                                                                sh_master.update_cell(m_idx, 2, ed_date_summer.strftime('%Y-%m-%d'))
                                                                sh_master.update_cell(m_idx, 6, ed_rec_summer)
                                                                sh_master.update_cell(m_idx, 7, float(ed_amt_summer))
                                                                sh_master.update_cell(m_idx, 8, ed_method_summer)
                                                                break
                                                except Exception:
                                                    pass

                                                st.success("✅ Η πληρωμή Summer τροποποιήθηκε επιτυχώς!")
                                                time.sleep(1.2)
                                                st.session_state.is_submitting = False
                                                st.rerun()
                                            else:
                                                st.error("Δεν βρέθηκε η εγγραφή στη βάση.")
                                                st.session_state.is_submitting = False
                                        except Exception as edit_err:
                                            st.session_state.is_submitting = False
                                            st.error(f"Σφάλμα τροποποίησης: {edit_err}")
                    else:
                        st.info("Δεν υπάρχουν καταχωρημένες πληρωμές για αυτή την εγγραφή.")
                else:
                    st.info("Δεν βρέθηκαν δεδομένα πληρωμών Summer.")          

                # --- ΝΕΑ ΚΑΤΑΧΩΡΗΣΗ SUMMER ΠΛΗΡΩΜΗΣ ---
                st.markdown("---")
                st.markdown("##### ➕ Νέα Καταχώρηση Πληρωμής Summer")
                
                with st.form("payment_form", clear_on_submit=True):
                    col1, col2 = st.columns(2)
                    rec_num = col1.text_input("Αρ. Απόδειξης")
                    amount = col2.number_input("Ποσό (€)", min_value=0.0, format="%.2f")
                    pay_method = st.selectbox("Τρόπος Πληρωμής", ["Μετρητά", "Εθνική", "Eurobank", "POS"])
                    
                    if st.form_submit_button("✅ Αποθήκευση Πληρωμής", disabled=st.session_state.is_submitting):
                        st.session_state.is_submitting = True
                        if not rec_num or amount <= 0:
                            st.warning("Παρακαλώ συμπληρώστε αριθμό απόδειξης και ποσό.")
                            st.session_state.is_submitting = False
                        else:
                            try:
                                gc = get_gspread_client()
                                sh_pays = gc.open(OLD_SHEET_NAME).worksheet("db_payment_summer")
                                sh_sum = gc.open(OLD_SHEET_NAME).worksheet("db_summer")
                                
                                actual_summer_prog = str(reg_row.get('summer_program', 'SummerCamp')).strip()

                                next_id = int(pd.to_numeric(df_pays['id_paymentSummer'], errors='coerce').max()) + 1 if not df_pays.empty else 1
                                
                                row_to_append = [
                                    next_id, target_summer_id, reg_row['summer_lastname'], reg_row['summer_firstname'], 
                                    email, datetime.today().strftime('%Y-%m-%d'), rec_num, "1", 
                                    amount, year, pay_method, actual_summer_prog
                                ]
                                sh_pays.append_row(row_to_append)
                                
                                new_total_paid = total_paid + amount
                                new_balance = total_cost - new_total_paid
                                new_status = "FINISH" if new_balance <= 0 else "ACTIVE"
                                
                                all_summer_rows = sh_sum.get_all_values()
                                if len(all_summer_rows) > 1:
                                    header_s = [h.strip() for h in all_summer_rows[0]]
                                    stat_idx = next((i for i, h in enumerate(header_s) if 'status' in h.lower()), 11)
                                    s_id_idx = next((i for i, h in enumerate(header_s) if 'summer_id' in h.lower() or 'id' in h.lower()), 0)

                                    for i, r in enumerate(all_summer_rows[1:], start=2):
                                        if len(r) > s_id_idx and str(r[s_id_idx]).strip() == target_summer_id:
                                            sh_sum.update_cell(i, stat_idx + 1, new_status)
                                            break

                                log_to_finance_master("ΕΣΟΔΟ", actual_summer_prog, f"{reg_row['summer_lastname']} {reg_row['summer_firstname']}", rec_num, amount, pay_method, year, "Summer")
                                
                                st.success(f"✅ Η πληρωμή καταχωρήθηκε! Status: {new_status}")
                                time.sleep(1.2)
                                st.session_state.is_submitting = False
                                st.rerun()
                                
                            except Exception as e:
                                st.session_state.is_submitting = False
                                st.error(f"Σφάλμα: {e}")
#---------------------------------------------------------------
    elif st.session_state.page == "advanced_reports":
        if st.session_state.role not in ["admin", "master"]:
            st.session_state.page = "menu"; st.rerun() 
    
        if st.button("⬅️ Επιστροφή στο Μενού", key="menu_reports", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("📊 Οικονομικές Αναφορές")
        
        report_type = st.radio("Τύπος Αναφοράς", ["Ιδιώτες / Μαθητές", "Εταιρικοί Πελάτες"], horizontal=True)
        
        if report_type == "Ιδιώτες / Μαθητές":
            reg_df = load_specific_db("db_registrations")
            sum_df = load_specific_db("db_summer")
            pays_df = load_specific_db("db_payments")
            sum_pays_df = load_specific_db("db_payment_summer")
            extra_pays_df = load_specific_db("db_extra_payments")
            
            # --- 1. ΚΑΘΑΡΙΣΜΟΣ ΓΙΑ ΝΑ ΦΥΓΟΥΝ ΟΙ ΚΕΝΕΣ ΓΡΑΜΜΕΣ ---
            if not reg_df.empty:
                reg_df = reg_df.dropna(subset=['customer_lastname', 'customer_firstname'], how='all')
                reg_df = reg_df[reg_df['customer_lastname'].astype(str).str.strip() != '']
            
            if not sum_df.empty:
                sum_df = sum_df.dropna(subset=['summer_lastname', 'summer_firstname'], how='all')
                sum_df = sum_df[sum_df['summer_lastname'].astype(str).str.strip() != '']

            df1 = pd.DataFrame()
            if not reg_df.empty:
                df1 = reg_df[['id_customer', 'customer_lastname', 'customer_firstname', 'school_year', 'program', 'class', 'status', 'adjusted_cost']].copy()
                df1.columns = ['ID', 'Επώνυμο', 'Όνομα', 'Έτος', 'Πρόγραμμα', 'Τμήμα', 'Κατάσταση', 'Συμφωνηθέν']
            else:
                df1 = pd.DataFrame(columns=['ID', 'Επώνυμο', 'Όνομα', 'Έτος', 'Πρόγραμμα', 'Τμήμα', 'Κατάσταση', 'Συμφωνηθέν'])
            
            if not sum_df.empty:
                df2 = sum_df[['summer_lastname', 'summer_firstname', 'summer_year', 'summer_program', 'summer_cost', 'status_summer']].copy()
                df2['ID'] = "SUMMER"
                df2['Τμήμα'] = 'Summer/Pre'
                df2.columns = ['Επώνυμο', 'Όνομα', 'Έτος', 'Πρόγραμμα', 'Συμφωνηθέν', 'Κατάσταση', 'ID', 'Τμήμα']
                df2 = df2[['ID', 'Επώνυμο', 'Όνομα', 'Έτος', 'Πρόγραμμα', 'Τμήμα', 'Κατάσταση', 'Συμφωνηθέν']]
            else:
                df2 = pd.DataFrame(columns=['ID', 'Επώνυμο', 'Όνομα', 'Έτος', 'Πρόγραμμα', 'Τμήμα', 'Κατάσταση', 'Συμφωνηθέν'])
    
            master_report = pd.concat([df1, df2], ignore_index=True)
            master_report['Συμφωνηθέν'] = pd.to_numeric(master_report['Συμφωνηθέν'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
            master_report['Επώνυμο'] = master_report['Επώνυμο'].astype(str).str.strip()
            master_report['Όνομα'] = master_report['Όνομα'].astype(str).str.strip()
            master_report['display'] = master_report['Επώνυμο'] + " " + master_report['Όνομα']

            # --- 2. ΣΩΣΤΟΣ ΥΠΟΛΟΓΙΣΜΟΣ ΠΛΗΡΩΜΩΝ ΑΝΑ ΜΑΘΗΤΗ/ΕΤΟΣ (Χωρίς διπλομετρήσεις) ---
            total_pays_per_cust = {}
            if not pays_df.empty:
                pays_clean = pays_df.copy()
                pays_clean['amt'] = pd.to_numeric(pays_clean['amount_paid'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                grouped_pays = pays_clean.groupby(['id_customer', 'school_year'])['amt'].sum().to_dict()
                for (c_id, yr), val in grouped_pays.items():
                    total_pays_per_cust[(str(c_id).strip(), str(yr).strip())] = val

            total_sum_pays_per_cust = {}
            if not sum_pays_df.empty:
                s_lname_col = 'summer_lastname' if 'summer_lastname' in sum_pays_df.columns else ('customer_lastnameSummer' if 'customer_lastnameSummer' in sum_pays_df.columns else '')
                s_fname_col = 'summer_firstname' if 'summer_firstname' in sum_pays_df.columns else ('customer_firstnameSummer' if 'customer_firstnameSummer' in sum_pays_df.columns else '')
                s_year_col = 'school_yearSummer' if 'school_yearSummer' in sum_pays_df.columns else ('school_year' if 'school_year' in sum_pays_df.columns else '')
                s_amt_col = 'amount_paidSummer' if 'amount_paidSummer' in sum_pays_df.columns else ('amount_paid' if 'amount_paid' in sum_pays_df.columns else 'amount')

                if s_lname_col and s_fname_col and s_year_col:
                    sum_clean = sum_pays_df.copy()
                    sum_clean['amt'] = pd.to_numeric(sum_clean[s_amt_col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                    sum_clean['lname_l'] = sum_clean[s_lname_col].astype(str).str.strip().str.lower()
                    sum_clean['fname_l'] = sum_clean[s_fname_col].astype(str).str.strip().str.lower()
                    sum_clean['yr_l'] = sum_clean[s_year_col].astype(str).str.strip()
                    
                    grouped_sum_pays = sum_clean.groupby(['lname_l', 'fname_l', 'yr_l'])['amt'].sum().to_dict()
                    for (lname, fname, yr), val in grouped_sum_pays.items():
                        total_sum_pays_per_cust[(lname, fname, yr)] = val

            paid_list = []
            assigned_customers_year = set()

            for _, row in master_report.iterrows():
                p_sum = 0.0
                c_id = str(row['ID']).strip()
                yr = str(row['Έτος']).strip()
                lname = str(row['Επώνυμο']).strip().lower()
                fname = str(row['Όνομα']).strip().lower()

                if c_id != "SUMMER":
                    key = (c_id, yr)
                    # Αν ο μαθητής έχει πολλαπλές εγγραφές/προγράμματα στο ίδιο έτος, 
                    # εμφανίζουμε το συνολικό ποσό των πληρωμών του στην πρώτη του εγγραφή για να μην διπλοεμφανίζεται.
                    if key not in assigned_customers_year:
                        p_sum = total_pays_per_cust.get(key, 0.0)
                        assigned_customers_year.add(key)
                    else:
                        p_sum = 0.0  # Στις επόμενες γραμμές του ίδιου μαθητή για το ίδιο έτος η πληρωμή φαίνεται 0 για αποφυγή διπλογραφής
                else:
                    sum_key = (lname, fname, yr)
                    if sum_key not in assigned_customers_year:
                        p_sum = total_sum_pays_per_cust.get(sum_key, 0.0)
                        assigned_customers_year.add(sum_key)
                    else:
                        p_sum = 0.0

                paid_list.append(p_sum)

            master_report['Έχει Δώσει'] = paid_list
            master_report['Υπόλοιπο'] = master_report['Συμφωνηθέν'] - master_report['Έχει Δώσει']

            # --- ΦΙΛΤΡΑ ΑΝΑΦΟΡΑΣ ---
            c1, c2, c3 = st.columns(3)
            sel_f_year = c1.selectbox("📅 Έτος", ["ΟΛΑ"] + sorted(master_report['Έτος'].unique().tolist(), reverse=True))
            sel_f_prog = c2.selectbox("📚 Πρόγραμμα", ["ΟΛΑ"] + sorted(master_report['Πρόγραμμα'].unique().tolist()))
            sel_f_class = c3.selectbox("🏫 Τμήμα", ["ΟΛΑ"] + sorted(master_report['Τμήμα'].unique().tolist()))
            
            c4, c5 = st.columns(2)
            sel_f_status = c4.selectbox("🛑 Κατάσταση Εγγραφής", ["ΟΛΑ", "ACTIVE", "STOP", "FINISH"])
            sel_f_cust = c5.selectbox("👤 Πελάτης", ["ΟΛΟΙ"] + sorted(master_report['display'].unique().tolist()))
    
            f_df = master_report.copy()
            if sel_f_year != "ΟΛΑ": f_df = f_df[f_df['Έτος'] == sel_f_year]
            if sel_f_prog != "ΟΛΑ": f_df = f_df[f_df['Πρόγραμμα'] == sel_f_prog]
            if sel_f_class != "ΟΛΑ": f_df = f_df[f_df['Τμήμα'] == sel_f_class]
            if sel_f_status != "ΟΛΑ": f_df = f_df[f_df['Κατάσταση'] == sel_f_status]
            if sel_f_cust != "ΟΛΟΙ": f_df = f_df[f_df['display'] == sel_f_cust]
            
            # --- ΠΡΑΣΙΝΟ ΧΡΩΜΑ ΟΤΑΝ ΤΟ ΥΠΟΛΟΙΠΟ ΕΙΝΑΙ 0 & ΣΩΣΤΑ ΔΕΚΑΔΙΚΑ ---
            def highlight_zero_balance(row):
                if row['Υπόλοιπο'] == 0:
                    return ['background-color: #d4edda'] * len(row)
                return [''] * len(row)

            display_f_df = (
                f_df.drop(columns=['display'])
                .style.apply(highlight_zero_balance, axis=1)
                .format("{:.2f}", subset=['Συμφωνηθέν', 'Έχει Δώσει', 'Υπόλοιπο'])
            )
            st.dataframe(display_f_df, use_container_width=True, hide_index=True)
            
            st.markdown(f"**Σύνολα:** Χρέωση: {f_df['Συμφωνηθέν'].sum():.2f}€ | Πληρωμές: {f_df['Έχει Δώσει'].sum():.2f}€ | Υπόλοιπο: {f_df['Υπόλοιπο'].sum():.2f}€")

            # --- 3. EXTRA GRID ΠΛΗΡΩΜΩΝ (Όταν επιλέγεται συγκεκριμένος πελάτης) ---
            if sel_f_cust != "ΟΛΟΙ":
                st.markdown("---")
                st.markdown(f"### 💳 Αναλυτικό Ιστορικό Πληρωμών για: **{sel_f_cust}**")
                
                parts = sel_f_cust.split(" ")
                target_lname = parts[0].strip().lower()
                target_fname = " ".join(parts[1:]).strip().lower()

                all_cust_payments = []

                # Α. db_payments
                if not pays_df.empty:
                    p_reg = pays_df[
                        (pays_df['customer_lastname'].astype(str).str.strip().str.lower() == target_lname) & 
                        (pays_df['customer_firstname'].astype(str).str.strip().str.lower() == target_fname)
                    ].copy()
                    if not p_reg.empty:
                        for _, pr in p_reg.iterrows():
                            all_cust_payments.append({
                                'Ημερομηνία': pr.get('payment_date', ''),
                                'Αρ. Απόδειξης': pr.get('receipt_number', ''),
                                'Ποσό (€)': float(str(pr.get('amount_paid', 0)).replace(',', '.') or 0),
                                'Τρόπος': pr.get('pay_method', ''),
                                'Έτος': pr.get('school_year', ''),
                                'Πρόγραμμα / Τύπος': pr.get('program', ''),
                                'Κατηγορία': 'Συνδρομή'
                            })

                # Β. db_payment_summer
                if not sum_pays_df.empty:
                    s_lname_col = 'summer_lastname' if 'summer_lastname' in sum_pays_df.columns else 'customer_lastnameSummer'
                    s_fname_col = 'summer_firstname' if 'summer_firstname' in sum_pays_df.columns else 'customer_firstnameSummer'
                    
                    p_sum = sum_pays_df[
                        (sum_pays_df[s_lname_col].astype(str).str.strip().str.lower() == target_lname) & 
                        (sum_pays_df[s_fname_col].astype(str).str.strip().str.lower() == target_fname)
                    ].copy()
                    if not p_sum.empty:
                        for _, pr in p_sum.iterrows():
                            all_cust_payments.append({
                                'Ημερομηνία': pr.get('payment_dateSummer', pr.get('payment_date', '')),
                                'Αρ. Απόδειξης': pr.get('receipt_numberSummer', pr.get('receipt_number', '')),
                                'Ποσό (€)': float(str(pr.get('amount_paidSummer', 0)).replace(',', '.') or 0),
                                'Τρόπος': pr.get('pay_methodSummer', pr.get('pay_method', '')),
                                'Έτος': pr.get('school_yearSummer', pr.get('school_year', '')),
                                'Πρόγραμμα / Τύπος': pr.get('programSummer', pr.get('program', 'SummerCamp')),
                                'Κατηγορία': 'Summer Camp'
                            })

                # Γ. db_extra_payments
                if not extra_pays_df.empty:
                    p_ext = extra_pays_df[
                        (extra_pays_df['customer_lastname'].astype(str).str.strip().str.lower() == target_lname) & 
                        (extra_pays_df['customer_firstname'].astype(str).str.strip().str.lower() == target_fname)
                    ].copy()
                    if not p_ext.empty:
                        for _, pr in p_ext.iterrows():
                            all_cust_payments.append({
                                'Ημερομηνία': pr.get('payment_date', ''),
                                'Αρ. Απόδειξης': pr.get('receipt_number', ''),
                                'Ποσό (€)': float(str(pr.get('cost', pr.get('amount', 0))).replace(',', '.') or 0),
                                'Τρόπος': pr.get('extra_pay_method', pr.get('method', '')),
                                'Έτος': sel_f_year if sel_f_year != "ΟΛΑ" else "---",
                                'Πρόγραμμα / Τύπος': pr.get('details', ''),
                                'Κατηγορία': 'Επιπρόσθετη'
                            })
                # --- Δ. db_payment_dypa (Επί Πιστώσει ΔΥΠΑ) ---
                dypa_pays_df = load_specific_db("db_payment_dypa")
                if not dypa_pays_df.empty:
                    p_dypa = dypa_pays_df[
                        (dypa_pays_df['customer_lastname'].astype(str).str.strip().str.lower() == target_lname) & 
                        (dypa_pays_df['customer_firstname'].astype(str).str.strip().str.lower() == target_fname)
                    ].copy()
                    if not p_dypa.empty:
                        for _, pr in p_dypa.iterrows():
                            all_cust_payments.append({
                                'Ημερομηνία': pr.get('payment_date', ''),
                                'Αρ. Απόδειξης': pr.get('receipt_number', ''),
                                'Ποσό (€)': float(str(pr.get('amount_paid', 0)).replace(',', '.') or 0),
                                'Τρόπος': pr.get('pay_method', 'ΕΠΙ ΠΙΣΤΩΣΕΙ'),
                                'Έτος': pr.get('school_year', ''),
                                'Πρόγραμμα / Τύπος': pr.get('program', ''),
                                'Κατηγορία': 'Επί Πιστώσει (ΔΥΠΑ)'
                            })
                if all_cust_payments:
                    df_payments_view = pd.DataFrame(all_cust_payments)
                    if sel_f_year != "ΟΛΑ" and 'Έτος' in df_payments_view.columns:
                        df_payments_view = df_payments_view[(df_payments_view['Έτος'].astype(str) == str(sel_f_year)) | (df_payments_view['Έτος'] == "---")]
                    
                    st.dataframe(df_payments_view, use_container_width=True, hide_index=True)
                    
                    if 'Κατηγορία' in df_payments_view.columns:
                        paid_df = df_payments_view[df_payments_view['Κατηγορία'] != 'Επί Πιστώσει (ΔΥΠΑ)']
                    else:
                        paid_df = df_payments_view
                        
                    total_paid_amount = paid_df['Ποσό (€)'].sum() if not paid_df.empty else 0.0
                    st.markdown(f"**Συνολικό Ποσό Πληρωμών Πελάτη:** {total_paid_amount:.2f}€")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("📥 Παραγωγή PDF Ιστορικού Πληρωμών", disabled=st.session_state.is_submitting):
                        pdf_data = generate_customer_payment_history_pdf(sel_f_cust, sel_f_year, df_payments_view)
                        if pdf_data:
                            st.download_button(
                                label="📥 Κατεβάστε το PDF Πληρωμών",
                                data=pdf_data,
                                file_name=f"histoliko_pliromon_{sel_f_cust.replace(' ', '_')}.pdf",
                                mime="application/pdf"
                            )
                else:
                    st.info("Δεν βρέθηκαν καταχωρημένες πληρωμές για αυτόν τον πελάτη.")
        
        # --- ΕΤΑΙΡΙΚΟΙ ΠΕΛΑΤΕΣ ---
        elif report_type == "Εταιρικοί Πελάτες":
            comp_reg_df = load_specific_db("db_customers_company")
            comp_pays_df = load_specific_db("db_payments_company")
            
            for df in [comp_reg_df, comp_pays_df]:
                if not df.empty:
                    df.columns = df.columns.astype(str).str.strip()
        
            all_companies_set = set()
            
            comp_col_reg = next((c for c in comp_reg_df.columns if 'company' in c.lower() or 'εταιρ' in c.lower() or 'name' in c.lower()), None) if not comp_reg_df.empty else None
            if comp_col_reg:
                for val in comp_reg_df[comp_col_reg].dropna():
                    v_str = str(val).strip()
                    if v_str and v_str.lower() != 'nan':
                        all_companies_set.add(v_str)
        
            comp_col_p = next((c for c in comp_pays_df.columns if 'company' in c.lower() or 'εταιρ' in c.lower() or 'name' in c.lower()), None) if not comp_pays_df.empty else None
            if comp_col_p:
                for val in comp_pays_df[comp_col_p].dropna():
                    v_str = str(val).strip()
                    if v_str and v_str.lower() != 'nan':
                        all_companies_set.add(v_str)
        
            if not comp_pays_df.empty:
                col_f1, col_f2, col_f3, col_f4 = st.columns(4)
                
                type_col = next((c for c in comp_pays_df.columns if 'type' in c.lower() or 'trans' in c.lower() or 'τύπος' in c.lower()), None)
                unique_types = sorted([str(t) for t in comp_pays_df[type_col].unique() if t and str(t).lower() != 'nan']) if type_col else []
                sel_comp_type = col_f1.selectbox("📂 Τύπος", ["ΟΛΑ"] + unique_types, key="comp_type_filter")
        
                method_col = next((c for c in comp_pays_df.columns if c.lower() == 'pay_method'), None)
                if not method_col:
                    method_col = next((c for c in comp_pays_df.columns if 'method' in c.lower() or 'τρόπος' in c.lower() or 'πληρωμ' in c.lower()), None)
                    
                unique_methods = sorted([str(m) for m in comp_pays_df[method_col].unique() if m and str(m).lower() != 'nan']) if method_col else []
                sel_comp_method = col_f2.selectbox("💳 Τρόπος Πληρωμής", ["ΟΛΑ"] + unique_methods, key="comp_method_filter")
        
                unique_comps = sorted(list(all_companies_set))
                sel_comp_name = col_f3.selectbox("🏢 Επιλογή Εταιρείας", ["ΟΛΕΣ"] + unique_comps, key="comp_name_filter")
        
                year_col = next((c for c in comp_pays_df.columns if 'year' in c.lower() or 'έτος' in c.lower() or 'ετος' in c.lower()), None)
                unique_years = sorted([str(y) for y in comp_pays_df[year_col].unique() if y and str(y).lower() != 'nan']) if year_col else []
                sel_comp_year = col_f4.selectbox("📅 Σχολικό Έτος", ["ΟΛΑ"] + unique_years, key="comp_year_filter")
        
                date_col = next((c for c in comp_pays_df.columns if 'date' in c.lower() or 'invoice' in c.lower()), None)
                col_d1, col_d2 = st.columns(2)
                start_inv_date = col_d1.date_input("📅 Από Ημερομηνία Τιμολογίου", value=None, key="comp_start_date")
                end_inv_date = col_d2.date_input("📅 Έως Ημερομηνία Τιμολογίου", value=None, key="comp_end_date")
        
                st.markdown("---")
        
                filtered_comps = comp_pays_df.copy()
        
                if sel_comp_type != "ΟΛΑ" and type_col:
                    filtered_comps = filtered_comps[filtered_comps[type_col].astype(str).str.strip().str.lower() == sel_comp_type.lower()]
                    
                if sel_comp_method != "ΟΛΑ" and method_col:
                    filtered_comps = filtered_comps[filtered_comps[method_col].astype(str).str.strip().str.lower() == sel_comp_method.lower()]
                    
                if sel_comp_name != "ΟΛΕΣ" and comp_col_p:
                    filtered_comps = filtered_comps[filtered_comps[comp_col_p].astype(str).str.strip() == sel_comp_name]
        
                if sel_comp_year != "ΟΛΑ" and year_col:
                    filtered_comps = filtered_comps[filtered_comps[year_col].astype(str).str.strip() == sel_comp_year]
        
                if date_col and (start_inv_date or end_inv_date):
                    parsed_inv_dates = pd.to_datetime(filtered_comps[date_col], errors='coerce').dt.date
                    if start_inv_date:
                        filtered_comps = filtered_comps[parsed_inv_dates >= start_inv_date]
                    if end_inv_date:
                        filtered_comps = filtered_comps[parsed_inv_dates <= end_inv_date]
        
                if not filtered_comps.empty:
                    st.dataframe(filtered_comps, use_container_width=True, hide_index=True)
        
                    amt_col = next((c for c in filtered_comps.columns if 'amount' in c.lower() or 'cost' in c.lower() or 'ποσο' in c.lower()), filtered_comps.columns[0])
                    
                    def parse_comp_amount(val):
                        if pd.isna(val):
                            return 0.0
                        val_str = str(val).strip().replace(',', '.')
                        try:
                            return float(val_str)
                        except ValueError:
                            cleaned = "".join(c for c in val_str if c.isdigit() or c == '.' or c == '-')
                            return float(cleaned) if cleaned else 0.0
        
                    filtered_comps['Clean_Amount'] = filtered_comps[amt_col].apply(parse_comp_amount)
                    
                    total_income = 0.0
                    total_expense = 0.0
        
                    if type_col:
                        for _, r in filtered_comps.iterrows():
                            t_val = str(r[type_col]).strip().lower()
                            amt_val = r['Clean_Amount']
                            if 'εσόδ' in t_val or 'εσοδ' in t_val or 'income' in t_val or 'revenue' in t_val:
                                total_income += amt_val
                            elif 'έξοδ' in t_val or 'εξοδ' in t_val or 'expense' in t_val:
                                total_expense += amt_val
                    else:
                        total_income = filtered_comps['Clean_Amount'].sum()
        
                    net_company_balance = total_income - total_expense
        
                    st.markdown(
                        f"📊 **Σύνολα Εταιρικών Κινήσεων:** "
                        f"🟢 Έσοδα: **{total_income:.2f}€** | "
                        f"🔴 Έξοδα: **{total_expense:.2f}€** | "
                        f"⚖️ Καθαρό Υπόλοιπο: **{net_company_balance:.2f}€**"
                    )
                else:
                    st.info("Δεν βρέθηκαν εταιρικές συναλλαγές με βάση τα επιλεγμένα φίλτρα.")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένες εταιρικές πληρωμές/συναλλαγές (`db_payments_company`).")
                
    #-----------------------------------------------------------------------------------------------------------
    elif st.session_state.page == "add_structure":
        if st.button("⬅️ Μενού", key="menu_struct", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("🛠️ Διαχείριση Προγραμμάτων, Τμημάτων & Καθηγητών")

        tab_prog, tab_class, tab_assign, tab_pay = st.tabs(["📚 Προγράμματα", "🏫 Τμήματα", "👨‍🏫 Ανάθεση Καθηγητών", "💳 Αποπληρωμή"])

        # --- TAB 1: ΠΡΟΓΡΑΜΜΑΤΑ ---
        with tab_prog:
            st.markdown("##### ➕ Προσθήκη Προγράμματος")
            prog_df = load_specific_db("db_programs")
            
            with st.form("form_prog_new", clear_on_submit=True):
                p_name = st.text_input("Όνομα Προγράμματος")
                p_cost = st.number_input("Κόστος / Ποσό (€)", min_value=0.0)
                if st.form_submit_button("💾 Αποθήκευση Προγράμματος", disabled=st.session_state.is_submitting):
                    if not p_name:
                        st.warning("Συμπληρώστε όνομα προγράμματος!")
                    else:
                        try:
                            gc = get_gspread_client()
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_programs")
                            recs = sh.get_all_records()
                            
                            # Υπολογισμός νέου id_program με βάση τη στήλη id_program ή τα δεδομένα
                            next_id = 1
                            if recs:
                                df_temp = pd.DataFrame(recs)
                                id_col_name = 'id_program' if 'id_program' in df_temp.columns else ('id' if 'id' in df_temp.columns else df_temp.columns[0])
                                numeric_ids = pd.to_numeric(df_temp[id_col_name], errors='coerce').dropna()
                                if not numeric_ids.empty:
                                    next_id = int(numeric_ids.max()) + 1
                                else:
                                    next_id = len(df_temp) + 1
                            
                            sh.append_row([next_id, p_name, float(p_cost)])
                            #load_specific_db.clear()
                            st.success("✅ Το πρόγραμμα αποθηκεύτηκε επιτυχώς!")
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")

            st.markdown("---")
            st.markdown("##### 📋 Υπάρχοντα Προγράμματα (Τροποποίηση / Διαγραφή)")
            prog_df_current = load_specific_db("db_programs")
            
            if not prog_df_current.empty:
                for _, r in prog_df_current.iterrows():
                    # Εντοπισμός σωστού ID από τη στήλη id_program
                    r_id = r.get('id_program', r.get('id', r.name))
                    cur_pname = r.get('program_name', r.get('Πρόγραμμα', ''))
                    cur_pcost = r.get('program_cost', r.get('Ποσό', 0))

                    with st.expander(f"📚 ID: {r_id} | {cur_pname} | Κόστος: {cur_pcost}€"):
                        e_c1, e_c2 = st.columns(2)
                        do_edit = e_c1.button("✏️ Τροποποίηση", key=f"edit_p_{r_id}")
                        do_del = e_c2.button("🗑️ Διαγραφή", key=f"del_p_{r_id}")

                        if do_edit:
                            st.session_state[f"edit_p_mode_{r_id}"] = True

                        if st.session_state.get(f"edit_p_mode_{r_id}", False):
                            with st.form(key=f"form_edit_prog_{r_id}"):
                                new_pname = st.text_input("Όνομα Προγράμματος", value=str(cur_pname))
                                new_pcost = st.number_input("Κόστος (€)", value=float(cur_pcost))
                                if st.form_submit_button("💾 Αποθήκευση Αλλαγών", disabled=st.session_state.is_submitting):
                                    try:
                                        gc = get_gspread_client()
                                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_programs")
                                        
                                        # Εύρεση της σωστής γραμμής βάσει του id_program (στήλη 1)
                                        cell = sh.find(str(r_id))
                                        if cell:
                                            sh.update_cell(cell.row, 2, new_pname) # Στήλη 2: program_name
                                            sh.update_cell(cell.row, 3, float(new_pcost)) # Στήλη 3: program_cost
                                            st.session_state[f"edit_p_mode_{r_id}"] = False
                                            #load_specific_db.clear()
                                            st.success("✅ Ενημερώθηκε!")
                                    except Exception as err:
                                        st.error(f"Σφάλμα: {err}")

                        if do_del:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_programs")
                                cell = sh.find(str(r_id))
                                if cell:
                                    sh.delete_rows(cell.row)
                                    #load_specific_db.clear()
                                    st.success("🗑️ Διαγράφηκε επιτυχώς!")
                            except Exception as err:
                                st.error(f"Σφάλμα: {err}")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένα προγράμματα.")
        # --- TAB 2: ΤΜΗΜΑΤΑ ---
        with tab_class:
            st.markdown("##### ➕ Δημιουργία Τμήματος")
            class_df = load_specific_db("db_class")
            prog_df_for_class = load_specific_db("db_programs")
            
            prog_name_col = 'program_name' if not prog_df_for_class.empty and 'program_name' in prog_df_for_class.columns else (prog_df_for_class.columns[1] if not prog_df_for_class.empty and len(prog_df_for_class.columns) > 1 else None)
            prog_names_list = prog_df_for_class[prog_name_col].tolist() if not prog_df_for_class.empty and prog_name_col else []

            with st.form("form_class_new", clear_on_submit=True):
                c_name = st.text_input("Όνομα Τμήματος")
                c_prog = st.selectbox("Επιλογή Προγράμματος", prog_names_list)
                c_year = st.selectbox("Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2041)])
                c_start = st.date_input("Ημερομηνία Έναρξης", value=datetime.today())
                c_end = st.date_input("Ημερομηνία Λήξης", value=datetime.today())

                if st.form_submit_button("💾 Δημιουργία Τμήματος", disabled=st.session_state.is_submitting):
                    if not c_name:
                        st.warning("Συμπληρώστε όνομα τμήματος!")
                    else:
                        try:
                            gc = get_gspread_client()
                            sh = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                            recs = sh.get_all_records()
                            
                            next_c_id = 1
                            if recs:
                                df_temp = pd.DataFrame(recs)
                                possible_id_cols = [c for c in df_temp.columns if 'id' in str(c).lower()]
                                id_col_c = possible_id_cols[0] if possible_id_cols else df_temp.columns[0]
                                numeric_ids_c = pd.to_numeric(df_temp[id_col_c], errors='coerce').dropna()
                                if not numeric_ids_c.empty:
                                    next_c_id = int(numeric_ids_c.max()) + 1
                                else:
                                    next_c_id = len(df_temp) + 1
                            if next_c_id < 1: next_c_id = 1
                            
                            sh.append_row([
                                next_c_id, c_name, c_prog, "", c_year, 
                                c_start.strftime('%Y-%m-%d'), c_end.strftime('%Y-%m-%d')
                            ])
                            #load_specific_db.clear()
                            st.success("✅ Το τμήμα δημιουργήθηκε επιτυχώς!")
                        except Exception as e:
                            st.error(f"Σφάλμα: {e}")

            st.markdown("---")
            st.markdown("##### 📋 Υπάρχοντα Τμήματα (Τροποποίηση)")
            class_df_current = load_specific_db("db_class")
            
            if not class_df_current.empty:
                for idx, r in class_df_current.iterrows():
                    possible_id_keys = [k for k in r.keys() if 'id' in str(k).lower()]
                    r_id = r.get(possible_id_keys[0], r.iloc[0]) if possible_id_keys else r.iloc[0]
                    if pd.isna(r_id) or str(r_id).strip() == "":
                        r_id = idx + 1

                    c_name_val = str(r.get('class_name', r.get('Τμήμα', r.iloc[1] if len(r) > 1 else ''))).strip()
                    c_prog_val = str(r.get('program_name', r.get('Πρόγραμμα', r.iloc[2] if len(r) > 2 else ''))).strip()
                    c_teach_val = r.get('teachers', r.get('Καθηγητής', r.iloc[3] if len(r) > 3 else ''))
                    c_year_val = r.get('year', r.get('Έτος', r.iloc[4] if len(r) > 4 else ''))
                    c_start_val = r.get('start_date', r.iloc[5] if len(r) > 5 else '')
                    c_end_val = r.get('end_date', r.iloc[6] if len(r) > 6 else '')

                    unique_key_suffix = f"{r_id}_{idx}"

                    with st.expander(f"🏫 ID: {r_id} | Τμήμα: {c_name_val} | Πρόγραμμα: {c_prog_val} | Καθηγητής: {c_teach_val} | Έτος: {c_year_val}"):
                        edit_c = st.button("✏️ Τροποποίηση Τμήματος", key=f"edit_c_{unique_key_suffix}", disabled=st.session_state.is_submitting)

                        if edit_c:
                            st.session_state[f"edit_c_mode_{unique_key_suffix}"] = True

                        if st.session_state.get(f"edit_c_mode_{unique_key_suffix}", False):
                            with st.form(key=f"form_edit_class_{unique_key_suffix}"):
                                ec_name = st.text_input("Όνομα Τμήματος", value=str(c_name_val))
                                ec_prog = st.selectbox("Πρόγραμμα", prog_names_list, index=prog_names_list.index(c_prog_val) if c_prog_val in prog_names_list else 0)
                                ec_teach = st.text_input("Καθηγητής", value=str(c_teach_val))
                                ec_year = st.text_input("Σχολικό Έτος", value=str(c_year_val))
                                
                                try:
                                    dt_s = datetime.strptime(str(c_start_val), '%Y-%m-%d').date()
                                except:
                                    dt_s = datetime.today().date()
                                try:
                                    dt_e = datetime.strptime(str(c_end_val), '%Y-%m-%d').date()
                                except:
                                    dt_e = datetime.today().date()

                                ec_start = st.date_input("Έναρξη", value=dt_s)
                                ec_end = st.date_input("Λήξη", value=dt_e)

                                if st.form_submit_button("💾 Αποθήκευση", disabled=st.session_state.is_submitting):
                                    try:
                                        gc = get_gspread_client()
                                        sh = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                                        cell = sh.find(str(r_id))
                                        if cell:
                                            sh.update_cell(cell.row, 2, ec_name)
                                            sh.update_cell(cell.row, 3, ec_prog)
                                            sh.update_cell(cell.row, 4, ec_teach)
                                            sh.update_cell(cell.row, 5, ec_year)
                                            sh.update_cell(cell.row, 6, ec_start.strftime('%Y-%m-%d'))
                                            sh.update_cell(cell.row, 7, ec_end.strftime('%Y-%m-%d'))
                                            
                                            # --- ΑΥΤΟΜΑΤΗ ΕΝΗΜΕΡΩΣΗ ΣΕ ΣΥΝΔΕΔΕΜΕΝΑ ΦΥΛΛΑ ---
                                            old_class_name = c_name_val
                                            new_class_name = str(ec_name).strip()
                                            new_prog_name = str(ec_prog).strip()
                                            
                                            # 1. Ενημέρωση db_registrations
                                            try:
                                                sh_reg = gc.open(OLD_SHEET_NAME).worksheet("db_registrations")
                                                reg_vals = sh_reg.get_all_values()
                                                if len(reg_vals) > 1:
                                                    r_header = [h.strip() for h in reg_vals[0]]
                                                    class_col_idx = next((i for i, h in enumerate(r_header) if 'class' in h.lower() or 'τμήμ' in h.lower()), None)
                                                    prog_col_idx = next((i for i, h in enumerate(r_header) if 'program' in h.lower() or 'προγρ' in h.lower()), None)
                                                    
                                                    for r_i, r_row in enumerate(reg_vals[1:], start=2):
                                                        if class_col_idx is not None and len(r_row) > class_col_idx and str(r_row[class_col_idx]).strip() == old_class_name:
                                                            sh_reg.update_cell(r_i, class_col_idx + 1, new_class_name)
                                                        if prog_col_idx is not None and len(r_row) > prog_col_idx and str(r_row[prog_col_idx]).strip() == c_prog_val:
                                                            sh_reg.update_cell(r_i, prog_col_idx + 1, new_prog_name)
                                            except Exception:
                                                pass

                                            # 2. Ενημέρωση db_payment_dypa
                                            try:
                                                sh_dypa = gc.open(OLD_SHEET_NAME).worksheet("db_payment_dypa")
                                                dypa_vals = sh_dypa.get_all_values()
                                                if len(dypa_vals) > 1:
                                                    d_header = [h.strip() for h in dypa_vals[0]]
                                                    d_class_idx = next((i for i, h in enumerate(d_header) if 'class' in h.lower() or 'τμήμ' in h.lower()), None)
                                                    d_prog_idx = next((i for i, h in enumerate(d_header) if 'program' in h.lower() or 'προγρ' in h.lower()), None)
                                                    
                                                    for d_i, d_row in enumerate(dypa_vals[1:], start=2):
                                                        if d_class_idx is not None and len(d_row) > d_class_idx and str(d_row[d_class_idx]).strip() == old_class_name:
                                                            sh_dypa.update_cell(d_i, d_class_idx + 1, new_class_name)
                                                        if d_prog_idx is not None and len(d_row) > d_prog_idx and str(d_row[d_prog_idx]).strip() == c_prog_val:
                                                            sh_dypa.update_cell(d_i, d_prog_idx + 1, new_prog_name)
                                            except Exception:
                                                pass

                                            # 3. Ενημέρωση db_class_payments
                                            try:
                                                sh_cp = gc.open(OLD_SHEET_NAME).worksheet("db_class_payments")
                                                cp_vals = sh_cp.get_all_values()
                                                if len(cp_vals) > 1:
                                                    cp_header = [h.strip() for h in cp_vals[0]]
                                                    cp_class_idx = next((i for i, h in enumerate(cp_header) if 'class' in h.lower() or 'τμήμ' in h.lower()), None)
                                                    
                                                    if cp_class_idx is not None:
                                                        for cp_i, cp_row in enumerate(cp_vals[1:], start=2):
                                                            if len(cp_row) > cp_class_idx and str(cp_row[cp_class_idx]).strip() == old_class_name:
                                                                sh_cp.update_cell(cp_i, cp_class_idx + 1, new_class_name)
                                            except Exception:
                                                pass

                                            st.session_state[f"edit_c_mode_{unique_key_suffix}"] = False
                                            #load_specific_db.clear()
                                            st.success("✅ Ενημερώθηκε επιτυχώς και συγχρονίστηκαν τα σχετικά αρχεία!")
                                    except Exception as err:
                                        st.error(f"Σφάλμα: {err}")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένα τμήματα.")         
        # --- TAB 3: ΑΝΑΘΕΣΗ ΚΑΘΗΓΗΤΩΝ ---
        with tab_assign:
            st.markdown("##### 👨‍🏫 Ανάθεση Καθηγητών σε Τμήματα")
            class_df_assign = load_specific_db("db_class")
            teachers_df = load_specific_db("db_teachers")
            
            # Δημιουργία καθαρής λίστας καθηγητών από τα teacher_firstname & teacher_lastname
            teachers_list = []
            if not teachers_df.empty:
                for _, r in teachers_df.iterrows():
                    fname = str(r.get('teacher_firstname', '')).strip()
                    lname = str(r.get('teacher_lastname', '')).strip()
                    
                    parts = []
                    if fname and fname.lower() != 'nan':
                        parts.append(fname)
                    if lname and lname.lower() != 'nan':
                        parts.append(lname)
                        
                    if parts:
                        teachers_list.append(" ".join(parts))

            teachers_list = sorted(list(set([t for t in teachers_list if t])))
            
            if not teachers_list:
                teachers_list = ["--- Δεν βρέθηκαν καθηγητές στο db_teachers ---"]

            if not class_df_assign.empty:
                with st.form("form_assign_teacher", clear_on_submit=True):
                    c_names_list = class_df_assign['class_name'].tolist() if 'class_name' in class_df_assign.columns else []
                    sel_class_to_assign = st.selectbox("Επιλογή Τμήματος", c_names_list)
                    
                    sel_teacher_input = st.selectbox("Επιλογή Καθηγητή", teachers_list)
                    
                    if st.form_submit_button("💾 Ενημέρωση Ανάθεσης Καθηγητή", disabled=st.session_state.is_submitting):
                        if sel_teacher_input.startswith("---"):
                            st.warning("Παρακαλώ επιλέξτε έναν έγκυρο καθηγητή από τη λίστα!")
                        else:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_class")
                                row_match = class_df_assign[class_df_assign['class_name'] == sel_class_to_assign]
                                if not row_match.empty:
                                    r_id_val = row_match.iloc[0].iloc[0]
                                    cell = sh.find(str(r_id_val))
                                    if cell:
                                        sh.update_cell(cell.row, 4, sel_teacher_input)
                                        #load_specific_db.clear()
                                        st.success("✅ Ο καθηγητής ανατέθηκε επιτυχώς!")
                            except Exception as err:
                                st.error(f"Σφάλμα ανάθεσης: {err}")

                st.markdown("---")
                available_cols = [c for c in ['id', 'class_name', 'program_name', 'teachers', 'year'] if c in class_df_assign.columns]
                if not available_cols:
                    available_cols = class_df_assign.columns.tolist()
                st.dataframe(class_df_assign[available_cols], use_container_width=True, hide_index=True)
            else:
                st.info("Δεν υπάρχουν διαθέσιμα τμήματα για ανάθεση.")
                
        # --- TAB 4: ΑΠΟΠΛΗΡΩΜΗ ---
        with tab_pay:
            st.markdown("##### 💳 Καταχώρηση & Έλεγχος Αποπληρωμών ανά Τμήμα")
            
            class_df_pay = load_specific_db("db_class")
            
            with st.form("form_payment_completion", clear_on_submit=True):
                st.markdown("###### ➕ Ολοκλήρωση / Εξόφληση Τμήματος")
                
                if not class_df_pay.empty:
                    c_names_for_pay = class_df_pay['class_name'].tolist() if 'class_name' in class_df_pay.columns else []
                    sel_class_pay = st.selectbox("Επιλογή Τμήματος", c_names_for_pay)
                    
                    pay_completion_date = st.date_input("Ημερομηνία Αποπληρωμής", value=datetime.today())
                    pay_amount_val = st.number_input("Συνολικό Πσό Αποπληρωμής (€)", min_value=0.0)
                    
                    submit_pay = st.form_submit_button("💾 Οριστικοποίηση Αποπληρωμής & Καταχώρηση", disabled=st.session_state.is_submitting)
                    
                    if submit_pay:
                        try:
                            gc = get_gspread_client()
                            
                            # Εύρεση του ακριβούς σχολικού έτους (year) για το επιλεγμένο τμήμα από το db_class
                            matched_class_row = class_df_pay[class_df_pay['class_name'] == sel_class_pay]
                            target_school_year = "2026-2027"
                            if not matched_class_row.empty:
                                for col_y in ['year', 'Έτος', 'school_year']:
                                    if col_y in matched_class_row.columns:
                                        val_y = str(matched_class_row.iloc[0].get(col_y, '')).strip()
                                        if val_y and val_y.lower() != 'nan':
                                            target_school_year = val_y
                                            break

                            sh_cp = gc.open(OLD_SHEET_NAME).worksheet("db_class_payments")
                            recs_cp = sh_cp.get_all_records()
                            
                            # Υπολογισμός ID
                            next_cp_id = 1
                            if recs_cp:
                                df_temp_cp = pd.DataFrame(recs_cp)
                                id_col_cp = next((c for c in df_temp_cp.columns if 'id' in c.lower()), df_temp_cp.columns[0])
                                numeric_ids_cp = pd.to_numeric(df_temp_cp[id_col_cp], errors='coerce').dropna()
                                if not numeric_ids_cp.empty:
                                    next_cp_id = int(numeric_ids_cp.max()) + 1
                                else:
                                    next_cp_id = len(df_temp_cp) + 1
                            else:
                                next_cp_id = len(recs_cp) + 1
                            
                            # ΑΚΡΙΒΗΣ ΣΕΙΡΑ: id_payment, class_name, school_year, payment_date, amount_paid
                            sh_cp.append_row([
                                next_cp_id,
                                sel_class_pay,
                                target_school_year,
                                pay_completion_date.strftime('%Y-%m-%d'),
                                float(pay_amount_val)
                            ])

                            # Καταχώρηση στο Master Οικονομικά
                            log_to_finance_master(
                                "ΕΣΟΔΟ",
                                f"Αποπληρωμή Τμήματος: {sel_class_pay}",
                                "ΚΡΑΤΟΣ",
                                "",
                                float(pay_amount_val),
                                "Τραπεζική/POS",
                                target_school_year,
                                "Αποπληρωμή"
                            )
                            
                            #load_specific_db.clear()
                            st.success(f"✅ Η αποπληρωμή για το τμήμα '{sel_class_pay}' (Έτος: {target_school_year}) καταχωρήθηκε επιτυχώς!")
                        except Exception as err:
                            st.error(f"Σφάλμα αποθήκευσης αποπληρωμής: {err}")
                else:
                    st.info("Δεν βρέθηκαν διαθέσιμα τμήματα.")

            st.markdown("---")
            st.markdown("##### 📋 Υπάρχουσες Αποπληρωμές Τμημάτων (Προβολή & Τροποποίηση)")
            class_pays_df = load_specific_db("db_class_payments")
            
            if not class_pays_df.empty:
                class_pays_df.columns = class_pays_df.columns.astype(str).str.strip()
                id_cp_col = next((c for c in class_pays_df.columns if 'id' in c.lower()), class_pays_df.columns[0])
                
                display_cp_df = class_pays_df.copy()
                
                # Μετονομασία στηλών στα Ελληνικά
                rename_map = {
                    id_cp_col: 'ID',
                    'class_name': 'Τμήμα',
                    'school_year': 'Σχολικό Έτος',
                    'payment_date': 'Ημερομηνία',
                    'amount_paid': 'Ποσό (€)'
                }
                display_cp_df = display_cp_df.rename(columns=rename_map)
                display_cp_df['Πελάτης/Προμηθευτής'] = "ΚΡΑΤΟΣ"
                
                desired_columns = ['ID', 'Τμήμα', 'Σχολικό Έτος', 'Ημερομηνία', 'Ποσό (€)', 'Πελάτης/Προμηθευτής']
                existing_cols = [col for col in desired_columns if col in display_cp_df.columns]
                st.dataframe(display_cp_df[existing_cols], use_container_width=True, hide_index=True)
                
                # --- ΕΠΙΛΟΓΗ ΓΙΑ ΤΡΟΠΟΠΟΙΗΣΗ ---
                st.markdown("###### ✏️ Τροποποίηση Υπάρχουσας Αποπληρωμής")
                cp_ids_list = class_pays_df[id_cp_col].astype(str).tolist()
                sel_cp_to_edit = st.selectbox("Επιλογή ID Αποπληρωμής για Τροποποίηση", ["-- Επιλέξτε ID --"] + cp_ids_list, key="sel_cp_edit_box")
                
                if sel_cp_to_edit != "-- Επιλέξτε ID --":
                    matched_cp_row = class_pays_df[class_pays_df[id_cp_col].astype(str) == str(sel_cp_to_edit)].iloc[0]
                    
                    old_cp_class = str(matched_cp_row.get('class_name', matched_cp_row.iloc[1] if len(matched_cp_row) > 1 else '')).strip()
                    old_cp_amount = float(pd.to_numeric(matched_cp_row.get('amount_paid', matched_cp_row.iloc[4] if len(matched_cp_row) > 4 else 0), errors='coerce') or 0.0)
                    
                    with st.form("edit_class_payment_form"):
                        st.markdown(f"**Επεξεργασία Αποπληρωμής με ID: {sel_cp_to_edit}**")
                        
                        c_names_list = class_df_pay['class_name'].tolist() if not class_df_pay.empty and 'class_name' in class_df_pay.columns else [old_cp_class]
                        default_c_idx = c_names_list.index(old_cp_class) if old_cp_class in c_names_list else 0
                        
                        ed_cp_class = st.selectbox("Τμήμα", options=c_names_list, index=default_c_idx)
                        ed_cp_amt = st.number_input("Ποσό (€)", min_value=0.0, value=old_cp_amount, format="%.2f")
                        ed_cp_year = st.text_input("Σχολικό Έτος", value=str(matched_cp_row.get('school_year', '2026-2027'))).strip()
                        
                        try:
                            default_dt = datetime.strptime(str(matched_cp_row.get('payment_date', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                        except:
                            default_dt = datetime.today().date()
                        ed_cp_date = st.date_input("Ημερομηνία", value=default_dt)
                        
                        if st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης Αποπληρωμής", disabled=st.session_state.is_submitting):
                            if ed_cp_amt <= 0:
                                st.warning("⚠️ Το ποσό πρέπει να είναι μεγαλύτερο από 0.")
                            else:
                                try:
                                    gc = get_gspread_client()
                                    sh_cp = gc.open(OLD_SHEET_NAME).worksheet("db_class_payments")
                                    
                                    # Εύρεση ακριβούς γραμμής στο Google Sheet βάσει του ID
                                    all_raw_cp = sh_cp.get_all_values()
                                    target_row_idx = None
                                    if len(all_raw_cp) > 1:
                                        for r_idx, sheet_row in enumerate(all_raw_cp[1:], start=2):
                                            if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_cp_to_edit).strip():
                                                target_row_idx = r_idx
                                                break
                                    
                                    if target_row_idx:
                                        # Στήλες db_class_payments: id(1), class_name(2), school_year(3), payment_date(4), amount_paid(5)
                                        sh_cp.update_cell(target_row_idx, 2, ed_cp_class)
                                        sh_cp.update_cell(target_row_idx, 3, ed_cp_year)
                                        sh_cp.update_cell(target_row_idx, 4, ed_cp_date.strftime('%Y-%m-%d'))
                                        sh_cp.update_cell(target_row_idx, 5, float(ed_cp_amt))
                                        
                                        # Ενημέρωση και στο Master (db_finance_master) ψάχνοντας την κατηγορία/ποσό
                                        try:
                                            sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                            master_rows = sh_master.get_all_values()
                                            if len(master_rows) > 1:
                                                m_header = [h.strip() for h in master_rows[0]]
                                                m_cat_idx = m_header.index('category') if 'category' in m_header else 3
                                                m_amt_idx = m_header.index('amount') if 'amount' in m_header else 6
                                                
                                                for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                    if len(m_row) > m_cat_idx and old_cp_class in str(m_row[m_cat_idx]):
                                                        sh_master.update_cell(m_idx, m_cat_idx + 1, f"Αποπληρωμή Τμήματος: {ed_cp_class}")
                                                        sh_master.update_cell(m_idx, m_amt_idx + 1, float(ed_cp_amt))
                                                        break
                                        except Exception:
                                            pass

                                        #load_specific_db.clear()
                                        #load_master_db.clear()
                                        st.success("✅ Η αποπληρωμή τροποποιήθηκε επιτυχώς και ενημερώθηκε το Master!")
                                        time.sleep(1.2)
                                      
                                    else:
                                        st.error("Δεν βρέθηκε η εγγραφή στη βάση.")
                                except Exception as edit_err:
                                    st.error(f"Σφάλμα τροποποίησης: {edit_err}")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένες αποπληρωμές τμημάτων.")
    #-----------------------------------------------------------------------------------                 
    #anafora pelaton
    elif st.session_state.page == "anafora_pelaton":
        if st.button("⬅️ Μενού", key="menu_anafora", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "menu"; st.rerun() 
        
        st.subheader("Αναφορά & Μητρώο Πελατών")
        
        # 1. Φόρτωση δεδομένων
        reg_df = load_specific_db("db_registrations")
        sum_df = load_specific_db("db_summer")
        cust_df = load_specific_db("db_customers")
        
        # --- ΚΑΘΑΡΙΣΜΟΣ ΚΕΝΩΝ ΓΡΑΜΜΩΝ ΑΠΟ ΟΛΑ ΤΑ DATASETS ---
        if not reg_df.empty:
            reg_df = reg_df.dropna(subset=['customer_lastname', 'customer_firstname'], how='all')
            reg_df = reg_df[reg_df['customer_lastname'].astype(str).str.strip() != '']
            reg_df = reg_df[reg_df['customer_lastname'].astype(str).str.upper() != 'NAN']
    
        if not sum_df.empty:
            sum_df = sum_df.dropna(subset=['summer_lastname', 'summer_firstname'], how='all')
            sum_df = sum_df[sum_df['summer_lastname'].astype(str).str.strip() != '']
            sum_df = sum_df[sum_df['summer_lastname'].astype(str).str.upper() != 'NAN']
    
        if not cust_df.empty:
            cust_df = cust_df.dropna(subset=['customer_lastname', 'customer_firstname'], how='all')
            cust_df = cust_df[cust_df['customer_lastname'].astype(str).str.strip() != '']
            cust_df = cust_df[cust_df['customer_lastname'].astype(str).str.upper() != 'NAN']
    
        # 2. Merge για προσθήκη id_customer / πατρωνύμου στο reg_df
        if not cust_df.empty:
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce')
            reg_df['id_customer'] = pd.to_numeric(reg_df['id_customer'], errors='coerce')
            reg_df = pd.merge(reg_df, cust_df[['id_customer', 'parent_firstname']], on='id_customer', how='left')
        else:
            reg_df['parent_firstname'] = ""
    
        # 3. Προετοιμασία Summer DataFrame
        if not sum_df.empty:
            sum_prep = sum_df.rename(columns={
                'summer_lastname': 'customer_lastname',
                'summer_firstname': 'customer_firstname',
                'summer_year': 'school_year',
                'summer_program': 'program',
                'status_summer': 'status'
            })
            sum_prep['class'] = 'SummerCamp'
            sum_prep['parent_firstname'] = sum_df['summer_parent_first'] if 'summer_parent_first' in sum_df.columns else ""
            sum_prep = sum_prep[['id_customer', 'customer_lastname', 'customer_firstname', 'parent_firstname', 'school_year', 'program', 'class', 'status']] if 'id_customer' in sum_prep.columns else sum_prep[['customer_lastname', 'customer_firstname', 'parent_firstname', 'school_year', 'program', 'class', 'status']]
        else:
            sum_prep = pd.DataFrame()
        
        # Ενοποίηση (Φροντίζουμε να διατηρείται το id_customer αν υπάρχει)
        cols_to_use = ['customer_lastname', 'customer_firstname', 'parent_firstname', 'school_year', 'program', 'class', 'status']
        if 'id_customer' in reg_df.columns:
            cols_to_use = ['id_customer'] + cols_to_use
        if not sum_prep.empty and 'id_customer' not in sum_prep.columns:
            sum_prep['id_customer'] = None
    
        full_df = pd.concat([
            reg_df[[c for c in cols_to_use if c in reg_df.columns]], 
            sum_prep[[c for c in cols_to_use if c in sum_prep.columns]] if not sum_prep.empty else pd.DataFrame(columns=cols_to_use)
        ], ignore_index=True)
        
        # Τελικός καθαρισμός τυχόν κενών γραμμών μετά το concat
        full_df = full_df.dropna(subset=['customer_lastname', 'customer_firstname'], how='all')
        full_df = full_df[full_df['customer_lastname'].astype(str).str.strip() != '']
        full_df = full_df[full_df['customer_lastname'].astype(str).str.upper() != 'NAN']
    
        # Τυποποίηση πεζών/κεφαλαίων
        full_df['customer_lastname'] = full_df['customer_lastname'].astype(str).str.strip().str.upper()
        full_df['customer_firstname'] = full_df['customer_firstname'].astype(str).str.strip().str.upper()
        
        # Δημιουργία καθαρού display_name για το grid
        full_df['display_name'] = full_df['customer_lastname'] + " " + full_df['customer_firstname']
        
        if full_df.empty:
            st.info("Δεν βρέθηκαν δεδομένα.")
        else:
            # 100% ΜΟΝΑΔΙΚΗ ΛΙΣΤΑ ΠΕΛΑΤΩΝ ΓΙΑ ΤΟ DROPDOWN
            unique_customers_df = full_df[['customer_lastname', 'customer_firstname', 'display_name']].drop_duplicates(subset=['customer_lastname', 'customer_firstname'])
            unique_customers_list = sorted(unique_customers_df['display_name'].tolist())
    
            # Φίλτρα
            f_col1, f_col2 = st.columns(2)
            sel_f_year = f_col1.selectbox("📅 Έτος", ["ΟΛΑ"] + sorted(full_df['school_year'].astype(str).unique().tolist(), reverse=True))
            sel_f_cust = f_col2.selectbox("👤 Πελάτης", ["ΟΛΟΙ"] + unique_customers_list, key="filter_cust_select")
            
            f_col3, f_col4, f_col5 = st.columns(3)
            sel_f_prog = f_col3.selectbox("📚 Πρόγραμμα", ["ΟΛΑ"] + sorted(full_df['program'].astype(str).unique().tolist()))
            sel_f_class = f_col4.selectbox("🏫 Τμήμα", ["ΟΛΑ"] + sorted(full_df['class'].astype(str).unique().tolist()))
            sel_f_status = f_col5.selectbox("🛑 Κατάσταση", ["ΟΛΑ", "ACTIVE", "STOP", "FINISH"])
        
            # Εφαρμογή Φίλτρων
            filtered_df = full_df.copy()
            if sel_f_year != "ΟΛΑ": filtered_df = filtered_df[filtered_df['school_year'].astype(str) == str(sel_f_year)]
            if sel_f_prog != "ΟΛΑ": filtered_df = filtered_df[filtered_df['program'].astype(str) == str(sel_f_prog)]
            if sel_f_class != "ΟΛΑ": filtered_df = filtered_df[filtered_df['class'].astype(str) == str(sel_f_class)]
            if sel_f_status != "ΟΛΑ": filtered_df = filtered_df[filtered_df['status'].astype(str).str.upper() == str(sel_f_status).upper()]
            if sel_f_cust != "ΟΛΟΙ": filtered_df = filtered_df[filtered_df['display_name'] == sel_f_cust]
        
            # 1ο Grid: Προετοιμασία για εμφάνιση
            display_df = filtered_df.drop(columns=[c for c in ['display_name', 'parent_firstname', 'id_customer'] if c in filtered_df.columns]).rename(columns={
                'customer_lastname': 'Επώνυμο',
                'customer_firstname': 'Όνομα',
                'school_year': 'Σχολικό Έτος',
                'program': 'Πρόγραμμα',
                'class': 'Τμήμα',
                'status': 'Κατάσταση'
            })
            
            display_df.insert(0, 'Α/Α', range(1, len(display_df) + 1))
            st.dataframe(display_df, use_container_width=True, hide_index=True)
        
            if st.button("📥 Παραγωγή PDF", disabled=st.session_state.is_submitting):
                rep_pdf = generate_report_pdf_customers(filtered_df, selected_year=sel_f_year, selected_student=sel_f_cust)
                if rep_pdf:
                    st.download_button("📥 Κατεβάστε το αρχείο", data=rep_pdf, file_name="mitroo.pdf", mime="application/pdf")
            
            # --- 2Ο GRID: ΦΙΛΤΡΑΡΙΣΜΕΝΟ ΜΗΤΡΩΟ ΠΕΛΑΤΩΝ ΑΠΟ DB_CUSTOMERS ---
            st.markdown("---")
            st.subheader("📋 Αναλυτικά Στοιχεία Πελατών")
            
            if not cust_df.empty:
                valid_cust_ids = filtered_df['id_customer'].dropna().unique() if 'id_customer' in filtered_df.columns else []
                
                matched_cust_df = pd.DataFrame()
                if len(valid_cust_ids) > 0 and 'id_customer' in cust_df.columns:
                    matched_cust_df = cust_df[cust_df['id_customer'].isin(valid_cust_ids)]
                
                filtered_names = set(zip(filtered_df['customer_lastname'], filtered_df['customer_firstname']))
                if len(filtered_names) > 0:
                    name_matched = cust_df[
                        cust_df.apply(lambda r: (str(r.get('customer_lastname', '')).strip().upper(), str(r.get('customer_firstname', '')).strip().upper()) in filtered_names, axis=1)
                    ]
                    if not matched_cust_df.empty:
                        matched_cust_df = pd.concat([matched_cust_df, name_matched]).drop_duplicates()
                    else:
                        matched_cust_df = name_matched
    
                if not matched_cust_df.empty:
                    grid_cust_df = matched_cust_df.copy()
                    if 'customer_afm' in grid_cust_df.columns:
                        grid_cust_df['customer_afm'] = grid_cust_df['customer_afm'].astype(str).str.replace("'", "")
                    
                    rename_map = {
                        'id_customer': 'ID Πελάτη',
                        'customer_category': 'Κατηγορία',
                        'customer_firstname': 'Όνομα',
                        'customer_lastname': 'Επώνυμο',
                        'parent_firstname': 'Όνομα Πατρός',
                        'parent_lastname': 'Επώνυμο Πατρός',
                        'mother_name': 'Όνομα Μητρός',
                        'passport': 'ΑΔΤ / Διαβατήριο',
                        'telephone': 'Τηλέφωνο',
                        'email': 'Email',
                        'address': 'Διεύθυνση',
                        'town': 'Πόλη',
                        'zipCode': 'ΤΚ',
                        'tax_office': 'ΔΟΥ',
                        'Tax Office': 'ΔΟΥ',
                        'customer_afm': 'ΑΦΜ',
                        'amka': 'ΑΜΚΑ',
                        'iban': 'IBAN',
                        'bank': 'Τράπεζα',
                        'birthday_date': 'Ημ. Γέννησης',
                        'register_date': 'Ημ. Εγγραφής'
                    }
                    grid_cust_df = grid_cust_df.rename(columns={k: v for k, v in rename_map.items() if k in grid_cust_df.columns})
                    st.dataframe(grid_cust_df, use_container_width=True, hide_index=True)
                    
                    # --- ΚΟΥΜΠΙ ΛΗΨΗΣ ΣΕ EXCEL ΜΕ ΔΙΑΧΩΡΙΣΜΕΝΑ ΠΕΔΙΑ ---
                    import io
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        grid_cust_df.to_excel(writer, index=False, sheet_name='Mitroo_Pelaton')
                    excel_data = output.getvalue()
    
                    st.download_button(
                        label="📥 Λήψη Μητρώου σε Excel",
                        data=excel_data,
                        file_name="mitroo_pelaton_analytiko.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel_grid_cust"
                    )
                else:
                    st.info("Δεν βρέθηκαν αντίστοιχα στοιχεία πελατών για τα επιλεγμένα φίλτρα.")
            else:
                st.info("Δεν βρέθηκαν δεδομένα στη βάση πελατών (`db_customers`).")
    
            # --- ΕΜΦΑΝΙΣΗ ΣΤΟΙΧΕΙΩΝ ΠΕΛΑΤΗ & ΠΑΡΑΠΟΝΩΝ (ΜΟΝΟ ΟΤΑΝ ΕΠΙΛΕΓΕΤΑΙ ΣΥΓΚΕΚΡΙΜΕΝΟΣ ΠΕΛΑΤΗΣ) ---
            if sel_f_cust != "ΟΛΟΙ":
                st.markdown("---")
                
                cust_matched_rows = filtered_df[filtered_df['display_name'] == sel_f_cust]
                if not cust_matched_rows.empty:
                    target_lname = cust_matched_rows.iloc[0]['customer_lastname']
                    target_fname = cust_matched_rows.iloc[0]['customer_firstname']
                    
                    if not cust_df.empty:
                        cust_db_match = cust_df[
                            (cust_df['customer_lastname'].astype(str).str.strip().str.upper() == target_lname) & 
                            (cust_df['customer_firstname'].astype(str).str.strip().str.upper() == target_fname)
                        ]
                        
                        if not cust_db_match.empty:
                            row_c = cust_db_match.iloc[0]
                            target_c_id = int(pd.to_numeric(row_c.get('id_customer', 0), errors='coerce'))
                            
                            # --- ΟΜΟΡΦΗ ΕΜΦΑΝΙΣΗ ΣΤΟΙΧΕΙΩΝ ΠΕΛΑΤΗ (ΕΚΤΟΣ GRID) ---
                            st.markdown(f"### 👤 Στοιχεία Πελάτη: **{sel_f_cust}**")
                            
                            st.markdown("""
                            <style>
                            .cust-card {
                                background-color: #f8f9fa;
                                border: 1px solid #e9ecef;
                                border-radius: 10px;
                                padding: 20px;
                                margin-bottom: 20px;
                            }
                            .cust-row {
                                display: flex;
                                margin-bottom: 8px;
                            }
                            .cust-label {
                                font-weight: bold;
                                color: #495057;
                                width: 170px;
                                flex-shrink: 0;
                            }
                            .cust-val {
                                color: #212529;
                            }
                            </style>
                            """, unsafe_allow_html=True)
    
                            c_id = row_c.get('id_customer', '')
                            c_cat = row_c.get('customer_category', '')
                            c_fname = row_c.get('customer_firstname', '')
                            c_lname = row_c.get('customer_lastname', '')
                            c_pfname = row_c.get('parent_firstname', '')
                            c_plname = row_c.get('parent_lastname', '')
                            c_mfname = row_c.get('mother_name', '')
                            c_pass = row_c.get('passport', '')
                            c_tel = row_c.get('telephone', '')
                            c_email = row_c.get('email', '')
                            c_addr = row_c.get('address', '')
                            c_town = row_c.get('town', '')
                            c_zip = row_c.get('zipCode', '')
                            c_taxoff = row_c.get('Tax Office', row_c.get('tax_office', ''))
                            c_afm = str(row_c.get('customer_afm', '')).replace("'", "")
                            c_amka = row_c.get('amka', '')
                            c_iban = row_c.get('iban', '')
                            c_bank = row_c.get('bank', '')
                            c_bday = row_c.get('birthday_date', '')
                            c_regdate = row_c.get('register_date', '')
    
                            st.markdown(f"""
                            <div class="cust-card">
                                <div style="display: flex; flex-wrap: wrap;">
                                    <div style="flex: 1; min-width: 300px; padding-right: 15px;">
                                        <div class="cust-row"><span class="cust-label">ID Πελάτη:</span><span class="cust-val">{c_id}</span></div>
                                        <div class="cust-row"><span class="cust-label">Κατηγορία:</span><span class="cust-val">{c_cat}</span></div>
                                        <div class="cust-row"><span class="cust-label">Όνομα:</span><span class="cust-val">{c_fname}</span></div>
                                        <div class="cust-row"><span class="cust-label">Επώνυμο:</span><span class="cust-val">{c_lname}</span></div>
                                        <div class="cust-row"><span class="cust-label">Όνομα Πατρός:</span><span class="cust-val">{c_pfname}</span></div>
                                        <div class="cust-row"><span class="cust-label">Επώνυμο Πατρός:</span><span class="cust-val">{c_plname}</span></div>
                                        <div class="cust-row"><span class="cust-label">Όνομα Μητρός:</span><span class="cust-val">{c_mfname}</span></div>
                                        <div class="cust-row"><span class="cust-label">ΑΔΤ / Διαβατήριο:</span><span class="cust-val">{c_pass}</span></div>
                                        <div class="cust-row"><span class="cust-label">Τηλέφωνο:</span><span class="cust-val">{c_tel}</span></div>
                                        <div class="cust-row"><span class="cust-label">Email:</span><span class="cust-val">{c_email}</span></div>
                                    </div>
                                    <div style="flex: 1; min-width: 300px;">
                                        <div class="cust-row"><span class="cust-label">Διεύθυνση:</span><span class="cust-val">{c_addr}</span></div>
                                        <div class="cust-row"><span class="cust-label">Πόλη:</span><span class="cust-val">{c_town}</span></div>
                                        <div class="cust-row"><span class="cust-label">ΤΚ:</span><span class="cust-val">{c_zip}</span></div>
                                        <div class="cust-row"><span class="cust-label">ΔΟΥ:</span><span class="cust-val">{c_taxoff}</span></div>
                                        <div class="cust-row"><span class="cust-label">ΑΦΜ:</span><span class="cust-val">{c_afm}</span></div>
                                        <div class="cust-row"><span class="cust-label">ΑΜΚΑ:</span><span class="cust-val">{c_amka}</span></div>
                                        <div class="cust-row"><span class="cust-label">IBAN:</span><span class="cust-val">{c_iban}</span></div>
                                        <div class="cust-row"><span class="cust-label">Τράπεζα:</span><span class="cust-val">{c_bank}</span></div>
                                        <div class="cust-row"><span class="cust-label">Ημ. Γέννησης:</span><span class="cust-val">{c_bday}</span></div>
                                        <div class="cust-row"><span class="cust-label">Ημ. Εγγραφής:</span><span class="cust-val">{c_regdate}</span></div>
                                    </div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
    
                            st.markdown("---")
                            
                            # --- ΕΜΦΑΝΙΣΗ ΠΑΡΑΠΟΝΩΝ & ΣΧΟΛΙΩΝ ---
                            st.markdown(f"### 🛠️ Παράπονα & Σχόλια για: **{sel_f_cust}**")
                            
                            complaints_df = load_specific_db("db_complaint")
                            
                            if not complaints_df.empty:
                                complaints_df.columns = complaints_df.columns.astype(str).str.strip()
                                complaints_df['id_cust_num'] = pd.to_numeric(complaints_df['id_customer'], errors='coerce')
                                cust_comps = complaints_df[complaints_df['id_cust_num'] == float(target_c_id)].copy()
                                
                                if not cust_comps.empty:
                                    if 'status' not in cust_comps.columns:
                                        cust_comps['status'] = 'ΕΝΕΡΓΟ'
                                    if 'choice' not in cust_comps.columns:
                                        cust_comps['choice'] = 'Παράπονο'
    
                                    view_cust_comps = cust_comps[['id_complaint', 'choice', 'complaint_date', 'complaint_subject', 'complaint_text', 'complaint_staff', 'status']].copy()
                                    view_cust_comps.columns = ['ID', 'Τύπος', 'Ημερομηνία', 'Θέμα', 'Κείμενο', 'Υπεύθυνος', 'Κατάσταση']
                                    
                                    st.dataframe(view_cust_comps, use_container_width=True, hide_index=True, height=180)
                                else:
                                    st.info("Δεν βρέθηκαν καταχωρημένα παράπονα ή σχόλια για αυτόν τον πελάτη.")
                            else:
                                st.info("Η βάση παραπόνων είναι κενή.")
                        else:
                            st.info("Δεν βρέθηκαν τα αναλυτικά στοιχεία του πελάτη στη βάση `db_customers`.")
                    else:
                        st.info("Η βάση πελατών είναι κενή.")
                else:
                    st.info("Δεν ήταν δυνατός ο εντοπισμός του πελάτη.")
    #----------------------------------------------------------------------------------------------------------------------------------
    elif st.session_state.page == "customer_registry":
        if st.button("⬅️ Μενού", key="menu_cust_reg", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 

        st.subheader("Επεξεργασία Στοιχείων Πελάτη")
        
        cust_df = load_specific_db("db_customers")
        
        if not cust_df.empty:
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            cust_df['customer_lastname'] = cust_df['customer_lastname'].astype(str).str.strip()
            cust_df['customer_firstname'] = cust_df['customer_firstname'].astype(str).str.strip()
            cust_df['parent_firstname'] = cust_df['parent_firstname'].astype(str).str.strip()
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce').fillna(0).astype(int)

            import unicodedata
            def remove_accents(input_str):
                if not isinstance(input_str, str): return ""
                nfkd_form = unicodedata.normalize('NFKD', input_str)
                return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).upper()

            cust_df['sort_key'] = cust_df['customer_lastname'].apply(remove_accents) + " " + cust_df['customer_firstname'].apply(remove_accents)
            cust_df = cust_df.sort_values(by='sort_key', ascending=True)

            cust_df['dropdown'] = (cust_df['customer_lastname'] + " " + 
                                   cust_df['customer_firstname'] + " (ΠΑΤΡΟΣ: " + 
                                   cust_df['parent_firstname'] + " - ID: " + cust_df['id_customer'].astype(str) + ")")
            
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη για Επεξεργασία", ["-- Επιλέξτε --"] + cust_df['dropdown'].tolist(), key="sel_edit_reg")

            if sel_cust != "-- Επιλέξτε --":
                selected_row = cust_df[cust_df['dropdown'] == sel_cust].iloc[0]
                cat = selected_row['customer_category']
                cust_id = str(selected_row['id_customer']).strip()
                old_lname = str(selected_row['customer_lastname']).strip()
                old_fname = str(selected_row['customer_firstname']).strip()

                with st.form("edit_customer_form"):
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        new_fname = st.text_input("Όνομα", value=selected_row['customer_firstname'])
                        new_lname = st.text_input("Επώνυμο", value=selected_row['customer_lastname'])
                        new_tel = st.text_input("Τηλέφωνο", value=selected_row['telephone'])
                        new_mf = st.text_input("Όνομα Μητρός", value=selected_row.get('mother_name', ''))
                        new_passp = st.text_input("ΑΔΤ / Διαβατήριο", value=selected_row.get('passport', ''))
                        
                        new_pfname, new_plname, new_afm = "", "", ""
                        new_town, new_zip, new_tax, new_amka, new_iban, new_bank = "", "", "", "", "", ""
                        
                        if cat == "Μαθητής":
                            new_pfname = st.text_input("Όνομα Πατρός", value=selected_row['parent_firstname'])
                            new_plname = st.text_input("Επώνυμο Πατρός", value=selected_row['parent_lastname'])
                        else:
                            new_afm = st.text_input("ΑΦΜ", value=str(selected_row.get('customer_afm', '')).replace("'", ""))
                            new_town = st.text_input("Πόλη", value=selected_row.get('town', ''))
                            new_zip = st.text_input("ΤΚ", value=str(selected_row.get('zipCode', '')))
                            new_tax = st.text_input("ΔΟΥ (Tax Office)", value=selected_row.get('Tax Office', selected_row.get('tax_office', '')))
                            new_amka = st.text_input("ΑΜΚΑ", value=str(selected_row.get('amka', '')))
                            new_iban = st.text_input("IBAN", value=str(selected_row.get('iban', '')))
                            new_bank = st.text_input("Τράπεζα", value=selected_row.get('bank', ''))

                    with col_e2:
                        new_email = st.text_input("Email", value=selected_row['email'])
                        new_adr = st.text_input("Διεύθυνση", value=selected_row['address'])
                        
                        raw_date = selected_row.get('birthday_date')
                        try:
                            dt_val = pd.to_datetime(raw_date, errors='coerce')
                            current_date = dt_val.date() if pd.notna(dt_val) else datetime(2000, 1, 1).date()
                        except:
                            current_date = datetime(2000, 1, 1).date()
                        
                        new_bd = st.date_input(
                            "Ημ. Γέννησης", 
                            value=current_date,
                            min_value=datetime(1950, 1, 1).date(),
                            max_value=datetime(2050, 12, 31).date()
                        )

                    if st.form_submit_button("💾 Αποθήκευση αλλαγών", disabled=st.session_state.is_submitting):
                        st.session_state.is_submitting = True
                        try:
                            gc = get_gspread_client()
                            
                            new_fname_upper = uppercase_clean(new_fname)
                            new_lname_upper = uppercase_clean(new_lname)
                            new_pfname_upper = uppercase_clean(new_pfname)
                            new_plname_upper = uppercase_clean(new_plname)
                            new_mf_upper = uppercase_clean(new_mf)
                            new_passp_upper = uppercase_clean(new_passp)
                            new_adr_upper = uppercase_clean(new_adr)
                            new_town_upper = uppercase_clean(new_town)
                            new_bank_upper = uppercase_clean(new_bank)
                            new_email_clean = str(new_email).strip().lower()

                            sh_cust = gc.open(OLD_SHEET_NAME).worksheet("db_customers")
                            cell = sh_cust.find(cust_id)
                            row_idx = cell.row
                            
                            sh_cust.update_cell(row_idx, 2, new_fname_upper)
                            sh_cust.update_cell(row_idx, 3, new_lname_upper)
                            sh_cust.update_cell(row_idx, 4, new_pfname_upper)
                            sh_cust.update_cell(row_idx, 5, new_plname_upper)
                            sh_cust.update_cell(row_idx, 6, new_tel)
                            sh_cust.update_cell(row_idx, 7, new_email_clean)
                            sh_cust.update_cell(row_idx, 8, new_adr_upper)
                            sh_cust.update_cell(row_idx, 9, new_bd.strftime('%Y-%m-%d'))
                            
                            if cat != "Μαθητής":
                                sh_cust.update_cell(row_idx, 12, f"'{new_afm}")
                                
                            sh_cust.update_cell(row_idx, 13, new_mf_upper)
                            sh_cust.update_cell(row_idx, 14, new_passp_upper)
                            
                            if cat != "Μαθητής":
                                sh_cust.update_cell(row_idx, 15, new_town_upper)
                                sh_cust.update_cell(row_idx, 16, str(new_zip))
                                sh_cust.update_cell(row_idx, 17, new_tax)
                                sh_cust.update_cell(row_idx, 18, str(new_amka))
                                sh_cust.update_cell(row_idx, 19, str(new_iban))
                                sh_cust.update_cell(row_idx, 20, new_bank_upper)

                            def update_sheet_names(sheet_name, id_col_name, fname_col_name, lname_col_name):
                                try:
                                    ws = gc.open(OLD_SHEET_NAME).worksheet(sheet_name)
                                    recs = ws.get_all_records()
                                    if recs:
                                        df_tmp = pd.DataFrame(recs)
                                        df_tmp.columns = df_tmp.columns.astype(str).str.strip()
                                        for i, r in df_tmp.iterrows():
                                            match_id = False
                                            if id_col_name in df_tmp.columns and str(r.get(id_col_name, '')).strip() == cust_id:
                                                match_id = True
                                            elif str(r.get('customer_lastname', '')).strip().lower() == old_lname.lower() and str(r.get('customer_firstname', '')).strip().lower() == old_fname.lower():
                                                match_id = True
                                            
                                            if match_id:
                                                r_idx = i + 2
                                                if fname_col_name in df_tmp.columns:
                                                    f_col_idx = df_tmp.columns.get_loc(fname_col_name) + 1
                                                    ws.update_cell(r_idx, f_col_idx, new_fname_upper)
                                                if lname_col_name in df_tmp.columns:
                                                    l_col_idx = df_tmp.columns.get_loc(lname_col_name) + 1
                                                    ws.update_cell(r_idx, l_col_idx, new_lname_upper)
                                except Exception:
                                    pass

                            update_sheet_names("db_registrations", "id_customer", "customer_firstname", "customer_lastname")
                            
                            try:
                                ws_sum = gc.open(OLD_SHEET_NAME).worksheet("db_summer")
                                recs_sum = ws_sum.get_all_records()
                                if recs_sum:
                                    df_sum = pd.DataFrame(recs_sum)
                                    df_sum.columns = df_sum.columns.astype(str).str.strip()
                                    for i, r in df_sum.iterrows():
                                        if str(r.get('summer_lastname', r.get('customer_lastnameSummer', ''))).strip().lower() == old_lname.lower() and str(r.get('summer_firstname', r.get('customer_firstnameSummer', ''))).strip().lower() == old_fname.lower():
                                            r_idx = i + 2
                                            if 'summer_firstname' in df_sum.columns:
                                                ws_sum.update_cell(r_idx, df_sum.columns.get_loc('summer_firstname') + 1, new_fname_upper)
                                            if 'summer_lastname' in df_sum.columns:
                                                ws_sum.update_cell(r_idx, df_sum.columns.get_loc('summer_lastname') + 1, new_lname_upper)
                            except Exception:
                                pass

                            update_sheet_names("db_payment_dypa", "id_customer", "customer_firstname", "customer_lastname")
                            update_sheet_names("db_payments", "id_customer", "customer_firstname", "customer_lastname")
                            update_sheet_names("db_extra_payments", "id_customer", "customer_firstname", "customer_lastname")

                            try:
                                ws_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                recs_m = ws_master.get_all_records()
                                if recs_m:
                                    df_m = pd.DataFrame(recs_m)
                                    df_m.columns = df_m.columns.astype(str).str.strip()
                                    old_full = f"{old_lname} {old_fname}".strip().lower()
                                    new_full = f"{new_lname_upper} {new_fname_upper}".strip()
                                    for i, r in df_m.iterrows():
                                        partner_val = str(r.get('partner', '')).strip().lower()
                                        if old_full in partner_val:
                                            r_idx = i + 2
                                            p_col_idx = df_m.columns.get_loc('partner') + 1
                                            ws_master.update_cell(r_idx, p_col_idx, new_full)
                            except Exception:
                                pass

                            #load_specific_db.clear()
                            #load_master_db.clear()
                            st.success("✅ Τα στοιχεία ενημερώθηκαν επιτυχώς σε όλες τις βάσεις και στο Master!")
                            time.sleep(1.5)
                            st.session_state.is_submitting = False
                            st.rerun()
                            
                        except Exception as e: 
                            st.session_state.is_submitting = False
                            st.error(f"Σφάλμα: {e}")
    #-----------------------------------------------------------------------------                        
    # 9. ΣΕΛΙΔΑ: ΚΑΘΗΓΗΤΕΣ
    elif st.session_state.page == "register_teacher":
        if st.button("⬅️ Μενού", key="menu_teach", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("Διαχείριση Καθηγητών")
        
        df_teachers = load_specific_db("db_teachers")

        if not df_teachers.empty and 'teacher_lastname' in df_teachers.columns:
            if 'full_name' not in df_teachers.columns:
                df_teachers['full_name'] = df_teachers['teacher_lastname'].astype(str) + " " + df_teachers['teacher_firstname'].astype(str)
            options = ["-- Νέα Εγγραφή --"] + df_teachers['full_name'].tolist()
        else:
            options = ["-- Νέα Εγγραφή --"]

        sel_teacher = st.selectbox("🔍 Επιλογή Καθηγητή", options=options, key="TEACHER_SELECT_UNIQUE_ID")

        selected_t = None
        if sel_teacher != "-- Νέα Εγγραφή --" and not df_teachers.empty:
            temp_df = df_teachers[df_teachers['full_name'] == sel_teacher]
            if not temp_df.empty:
                selected_t = temp_df.iloc[0]

        with st.form(f"form_{sel_teacher}"):
            col1, col2 = st.columns(2)
            with col1:
                t_fname = st.text_input("Όνομα", value=selected_t.get('teacher_firstname', '') if selected_t is not None else "")
                t_lname = st.text_input("Επώνυμο", value=selected_t.get('teacher_lastname', '') if selected_t is not None else "")
                t_email = st.text_input("Email", value=selected_t.get('teacher_email', '') if selected_t is not None else "")
            with col2:
                t_tel = st.text_input("Τηλέφωνο", value=selected_t.get('teacher_telephone', '') if selected_t is not None else "")
                banks = ["Εθνική", "Πειραιώς", "Alpha Bank", "Eurobank", "Attica Bank"]
                stored_bank = selected_t.get('teacher_bank', 'Εθνική') if selected_t is not None else "Εθνική"
                idx = banks.index(stored_bank) if stored_bank in banks else 0
                t_bank = st.selectbox("Τράπεζα", banks, index=idx)
                t_iban = st.text_input("IBAN", value=selected_t.get('teacher_iban', '') if selected_t is not None else "")

            sub = st.form_submit_button("💾 Αποθήκευση", disabled=st.session_state.is_submitting)

        if sub:
            try:
                gc = get_gspread_client()
                sh = gc.open(OLD_SHEET_NAME).worksheet("db_teachers")
                
                if selected_t is None:
                    next_id = int(pd.to_numeric(df_teachers['teacher_id'], errors='coerce').max()) + 1 if not df_teachers.empty else 1
                    sh.append_row([next_id, t_fname, t_lname, t_email, t_tel, t_bank, t_iban])
                    st.success("✅ Νέος καθηγητής καταχωρήθηκε!")
                else:
                    row_idx = selected_t.name + 2
                    sh.update(f"B{row_idx}:G{row_idx}", [[t_fname, t_lname, t_email, t_tel, t_bank, t_iban]])
                    st.success("✅ Τα στοιχεία ενημερώθηκαν!")
                
                #load_specific_db.clear()
            except Exception as e:
                st.error(f"Σφάλμα αποθήκευση: {e}")

    # 11. ΣΕΛΙΔΑ: EXTRA ΠΛΗΡΩΜΕΣ
    elif st.session_state.page == "extra_payments_view":
        if st.button("⬅️ Μενού", key="menu_extra_view_final", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Διαχείριση Επιπρόσθετων Πληρωμών")
        
        cust_df = load_specific_db("db_customers")
        extra_df = load_specific_db("db_extra_payments")
        
        for df in [cust_df, extra_df]:
            if not df.empty: df.columns = df.columns.astype(str).str.strip()

        if not cust_df.empty:
            cust_df = cust_df.copy()
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce').fillna(0).astype(int)
            cust_df['display'] = (cust_df['customer_lastname'].astype(str) + " " + 
                                cust_df['customer_firstname'].astype(str) + " (ID: " + 
                                cust_df['id_customer'].astype(str) + ")")
            
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη", options=["-- Επιλέξτε --"] + cust_df['display'].tolist(), key="ex_sel")
            
            if sel_cust != "-- Επιλέξτε --":
                c_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = int(c_row['id_customer'])
                
                # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ EXTRA ΠΛΗΡΩΜΩΝ ---
                st.markdown("---")
                st.markdown(f"##### 📋 Υπάρχουσες Επιπρόσθετες Πληρωμές Πελάτη: {c_row['customer_lastname']} {c_row['customer_firstname']}")
                
                if not extra_df.empty:
                    extra_df['id_cust_num'] = pd.to_numeric(extra_df['id_customer'], errors='coerce')
                    cust_extras = extra_df[extra_df['id_cust_num'] == float(c_id)].copy()
                    
                    if not cust_extras.empty:
                        # Εμφάνιση grid
                        display_cols = [c for c in cust_extras.columns if c != 'id_cust_num']
                        st.dataframe(cust_extras[display_cols], use_container_width=True, hide_index=True)
                        
                        # Επιλογή για Τροποποίηση
                        record_ids = cust_extras['id_payment'].astype(str).tolist()
                        sel_extra_to_edit = st.selectbox("Επιλογή ID Πληρωμής για Τροποποίηση", ["-- Επιλέξτε ID --"] + record_ids, key="sel_extra_edit_box")
                        
                        if sel_extra_to_edit != "-- Επιλέξτε ID --":
                            matched_e_row = cust_extras[cust_extras['id_payment'].astype(str) == str(sel_extra_to_edit)].iloc[0]
                            old_receipt = str(matched_e_row.get('receipt_number', '')).strip()
                            
                            with st.form("edit_extra_form_view"):
                                st.markdown(f"**Επεξεργασία Εγγραφής με ID: {sel_extra_to_edit}**")
                                ed_rec = st.text_input("Αριθμός Απόδειξης", value=old_receipt).strip()
                                ed_details = st.text_input("Αιτιολογία Πληρωμής", value=str(matched_e_row.get('details', ''))).strip()
                                ed_amt = st.number_input("Ποσό (€)", min_value=0.0, value=float(pd.to_numeric(matched_e_row.get('amount', matched_e_row.get('cost', 0)), errors='coerce') or 0.0), format="%.2f")
                                ed_method = st.selectbox("Τρόπος Πληρωμής", options=["Μετρητά", "Εθνική", "Eurobank", "POS"], index=["Μετρητά", "Εθνική", "Eurobank", "POS"].index(str(matched_e_row.get('method', 'Μετρητά'))) if str(matched_e_row.get('method', 'Μετρητά')) in ["Μετρητά", "Εθνική", "Eurobank", "POS"] else 0)
                                
                                if st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης", disabled=st.session_state.is_submitting):
                                    if ed_amt <= 0:
                                        st.warning("⚠️ Το ποσό πρέπει να είναι μεγαλύτερο από 0.")
                                    else:
                                        try:
                                            gc = get_gspread_client()
                                            sh_extra = gc.open(OLD_SHEET_NAME).worksheet("db_extra_payments")
                                            
                                            # Εύρεση ακριβούς γραμμής στο Google Sheet βάσει του id_payment
                                            all_raw_extras = sh_extra.get_all_values()
                                            target_row_idx = None
                                            if len(all_raw_extras) > 1:
                                                for r_idx, sheet_row in enumerate(all_raw_extras[1:], start=2):
                                                    if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_extra_to_edit).strip():
                                                        target_row_idx = r_idx
                                                        break
                                            
                                            if target_row_idx:
                                                # Ενημέρωση γραμμής στο db_extra_payments
                                                # Στήλες: id_payment(1), id_customer(2), lastname(3), firstname(4), date(5), receipt(6), details(7), amount(8), method(9)
                                                sh_extra.update_cell(target_row_idx, 6, ed_rec)
                                                sh_extra.update_cell(target_row_idx, 7, ed_details)
                                                sh_extra.update_cell(target_row_idx, 8, float(ed_amt))
                                                sh_extra.update_cell(target_row_idx, 9, ed_method)
                                                
                                                # Ενημέρωση και στο Master (db_finance_master) στο νέο sheet
                                                try:
                                                    sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                                    master_rows = sh_master.get_all_values()
                                                    if len(master_rows) > 1:
                                                        for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                            if len(m_row) > 5 and str(m_row[5]).strip() == old_receipt:
                                                                sh_master.update_cell(m_idx, 4, ed_details) # category
                                                                sh_master.update_cell(m_idx, 6, ed_rec) # invoice_number
                                                                sh_master.update_cell(m_idx, 7, float(ed_amt)) # amount
                                                                sh_master.update_cell(m_idx, 8, ed_method) # payment_method
                                                                break
                                                except Exception:
                                                    pass

                                                #load_specific_db.clear()
                                                #load_master_db.clear()
                                                st.success("✅ Η επιπρόσθετη πληρωμή τροποποιήθηκε επιτυχώς!")
                                                time.sleep(1.2)
                                                
                                            else:
                                                st.error("Δεν βρέθηκε η εγγραφή στη βάση.")
                                        except Exception as edit_err:
                                            st.error(f"Σφάλμα τροποποίησης: {edit_err}")
                    else:
                        st.info("Ο πελάτης δεν έχει επιπρόσθετες πληρωμές.")
                else:
                    st.info("Δεν βρέθηκαν δεδομένα.")

                # --- ΝΕΑ ΚΑΤΑΧΩΡΗΣΗ EXTRA ΠΛΗΡΩΜΗΣ ---
                st.markdown("---")
                st.markdown("##### ➕ Νέα Επιπρόσθετη Πληρωμή")
                
                with st.form("extra_pay_form", clear_on_submit=True):
                    rec = st.text_input("Αριθμός Απόδειξης (Προαιρετικό)", key="in_rec")
                    details = st.text_input("Αιτιολογία Πληρωμής", key="in_details")
                    amt = st.number_input("Ποσό (€)", min_value=0.0, format="%.2f", key="in_amt")
                    method = st.selectbox("Τρόπος Πληρωμής", options=["Μετρητά", "Εθνική", "Eurobank", "POS"], key="in_method")
                    sub_ep = st.form_submit_button("💳 Αποθήκευση", disabled=st.session_state.is_submitting)
                
                    if sub_ep:
                        if amt <= 0:
                            st.error("⚠️ Το ποσό πρέπει να είναι μεγαλύτερο από 0.")
                        else:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_extra_payments")
                                records = sh.get_all_records()
                                next_id = int(pd.to_numeric(pd.DataFrame(records)['id_payment'], errors='coerce').max()) + 1 if records else 1
                                
                                sh.append_row([
                                    next_id, c_id, str(c_row['customer_lastname']), str(c_row['customer_firstname']), 
                                    datetime.today().strftime('%Y-%m-%d'), rec if rec else "", details, amt, method
                                ])
                                
                                log_to_finance_master("ΕΣΟΔΟ", details, sel_cust, rec, amt, method, "2026-2027", "Extra")
                                
                                #load_specific_db.clear()
                                #load_master_db.clear()
                                st.success(f"✅ Η πληρωμή ({method}) καταχωρήθηκε!")
                                time.sleep(1.5)
                                
                            except Exception as e:
                                st.error(f"Σφάλμα αποθήκευσης: {e}")
        else:
            st.warning("Δεν βρέθηκαν πελάτες.")
#----------------------------------------------------------------------------------------------------------
    #καθηγητες - αναφορα
    elif st.session_state.page == "teacher_report":
        if st.button("⬅️ Μενού", key="menu_teach_rep", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 
            
        st.subheader("Αναφορά Καθηγητών & Τμημάτων")
        
        reg_df = load_specific_db("db_registrations")
        classes_df = load_specific_db("db_class")
        teachers_df = load_specific_db("db_teachers")
        customers_df = load_specific_db("db_customers")
        cp_df = load_specific_db("db_class_payments")
        sum_df = load_specific_db("db_summer")
        
        for df in [reg_df, classes_df, teachers_df, customers_df, cp_df, sum_df]:
            if not df.empty:
                df.columns = df.columns.astype(str).str.strip()
                # Καθαρισμός κενών από όλα τα κείμενα για να αποφεύγονται διπλότυπα λόγω spaces
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].astype(str).str.strip()
    
        # Διόρθωση σειράς ονόματος: Επώνυμο + Όνομα
        if not teachers_df.empty:
            if 'teacher_firstname' in teachers_df.columns and 'teacher_lastname' in teachers_df.columns:
                teachers_df['full_name'] = teachers_df['teacher_firstname'] + " " + teachers_df['teacher_lastname']
            else:
                teachers_df['full_name'] = teachers_df.iloc[:, 1].astype(str) + " " + teachers_df.iloc[:, 2].astype(str)
    
        tab1, tab2 = st.tabs(["👨‍🏫 Ανά Καθηγητή", "🏫 Ανά Τμήμα & Μαθητές"])
    
        with tab1:
            if not classes_df.empty:
                year_col_c = next((c for c in classes_df.columns if 'year' in c.lower() or 'έτος' in c.lower() or 'ετος' in c.lower()), classes_df.columns[4] if len(classes_df.columns) > 4 else None)
                years_list = ["ΟΛΑ"] + sorted(classes_df[year_col_c].astype(str).unique().tolist(), reverse=True) if year_col_c else ["ΟΛΑ"]
                sel_year = st.selectbox("📅 Επιλογή Σχολικού Έτους", options=years_list, key="year_t")
                
                teach_opts = teachers_df['full_name'].tolist() if not teachers_df.empty and 'full_name' in teachers_df.columns else []
                sel_teacher = st.selectbox("👨‍🏫 Επιλογή Καθηγητή", options=["ΟΛΟΙ"] + teach_opts, key="teach_t")
                
                # Υπολογισμός μαθητών ανά τμήμα (όλοι εκτός από STOP)
                if not reg_df.empty:
                    class_col_r = next((c for c in reg_df.columns if c.lower() in ['class', 'class_name', 'τμήμα', 'τμημα']), 'class')
                    year_col_r = next((c for c in reg_df.columns if 'year' in c.lower() or 'έτος' in c.lower()), 'school_year')
                    status_col_r = next((c for c in reg_df.columns if 'status' in c.lower() or 'κατάσταση' in c.lower()), 'status')
                    
                    active_regs = reg_df[reg_df[status_col_r].str.upper() != 'STOP'] if status_col_r in reg_df.columns else reg_df
                    counts = active_regs.groupby([class_col_r, year_col_r]).size().reset_index(name='Μαθητές')
                else:
                    counts = pd.DataFrame(columns=['class', 'school_year', 'Μαθητές'])
    
                class_name_col = next((c for c in classes_df.columns if 'class_name' in c.lower() or c.lower() in ['class', 'τμήμα', 'τμημα']), classes_df.columns[1] if len(classes_df.columns) > 1 else classes_df.columns[0])
                
                # Καθαρισμός και αποφυγή διπλότυπων στα τμήματα
                classes_df = classes_df.drop_duplicates(subset=[class_name_col])
    
                if not counts.empty and year_col_c:
                    report = pd.merge(classes_df, counts, left_on=[class_name_col, year_col_c], right_on=[counts.columns[0], counts.columns[1]], how='left')
                else:
                    report = classes_df.copy()
                    report['Μαθητές'] = 0
                
                if 'Μαθητές' in report.columns:
                    report['Μαθητές'] = report['Μαθητές'].fillna(0).astype(int).replace(0, "")
                else:
                    report['Μαθητές'] = ""
                
                teach_col_in_class = next((c for c in report.columns if 'teach' in c.lower() or 'καθηγ' in c.lower() or 'teacher' in c.lower()), report.columns[3] if len(report.columns) > 3 else None)
                if teach_col_in_class:
                    report['resolved_teacher'] = report[teach_col_in_class]
                else:
                    report['resolved_teacher'] = ""
    
                # Φιλτράρισμα
                if sel_year != "ΟΛΑ" and year_col_c: 
                    report = report[report[year_col_c].astype(str) == sel_year]
                
                if sel_teacher != "ΟΛΟΙ": 
                    report = report[report['resolved_teacher'].str.contains(sel_teacher, na=False, case=False)]
                
                prog_col_c = next((c for c in report.columns if 'prog' in c.lower() or 'προγρ' in c.lower()), report.columns[2] if len(report.columns) > 2 else '')
                
                display_data = pd.DataFrame({
                    'Τμήμα': report[class_name_col],
                    'Πρόγραμμα': report[prog_col_c] if prog_col_c else "",
                    'Καθηγητής': report['resolved_teacher'],
                    'Έτος': report[year_col_c] if year_col_c else "",
                    'Μαθητές': report['Μαθητές']
                })
                
                st.dataframe(display_data, use_container_width=True, hide_index=True)
                
                if st.button("📥 PDF Καθηγητών", disabled=st.session_state.is_submitting):
                    rep = generate_teacher_pdf(report, sel_year, sel_teacher)
                    st.download_button("Κατέβασμα PDF", rep, "report_teachers.pdf", "application/pdf")
            else:
                st.info("Δεν βρέθηκαν καταχωρημένα τμήματα.")
    
        with tab2:
            if not reg_df.empty or not sum_df.empty:
                years_set = set()
                if not reg_df.empty and 'school_year' in reg_df.columns:
                    years_set.update(reg_df['school_year'].unique().tolist())
                if not sum_df.empty and 'summer_year' in sum_df.columns:
                    years_set.update(sum_df['summer_year'].unique().tolist())
                
                years_list_t2 = sorted(list(years_set), reverse=True)
                sel_year_t2 = st.selectbox("📅 Επιλογή Έτους", options=years_list_t2, key="year_t2")
                
                # Φιλτράρισμα και καθαρισμός τμημάτων (απόρριψη 'Χωρίς Τμήμα', 'Προσθήκη ΑΦΜ' κλπ.)
                ignored_values = ['χωρίς τμήμα', 'προσθήκη αφμ', 'nan', 'none', '']
                class_opts_t2 = []
                
                if not reg_df.empty and 'class' in reg_df.columns:
                    year_matched_regs = reg_df[reg_df['school_year'] == str(sel_year_t2)]
                    for c in year_matched_regs['class'].unique():
                        c_str = str(c).strip()
                        if c_str.lower() not in ignored_values:
                            class_opts_t2.append(c_str)
                
                if not sum_df.empty:
                    class_opts_t2.extend(['SummerCamp', 'PreSchool'])
                
                class_opts_t2 = sorted(list(set(class_opts_t2)))
                
                if not class_opts_t2:
                    class_opts_t2 = ["Δεν βρέθηκαν τμήματα"]
    
                sel_class = st.selectbox("🏫 Επιλογή Τμήματος", options=class_opts_t2, key="class_t")
                
                # --- ΕΛΕΓΧΟΣ ΑΠΟΠΛΗΡΩΜΗΣ ---
                is_paid = False
                if not cp_df.empty:
                    c_col = next((c for c in cp_df.columns if 'class' in c.lower() or 'τμήμ' in c.lower()), cp_df.columns[1] if len(cp_df.columns) > 1 else '')
                    y_col = next((c for c in cp_df.columns if 'year' in c.lower() or 'έτος' in c.lower() or 'ετος' in c.lower()), cp_df.columns[2] if len(cp_df.columns) > 2 else '')
                    
                    if c_col and y_col:
                        matched_cp = cp_df[
                            (cp_df[c_col] == str(sel_class)) & 
                            (cp_df[y_col] == str(sel_year_t2))
                        ]
                        is_paid = not matched_cp.empty
    
                status_pay = "✅ Ναι" if is_paid else "❌ Όχι"
    
                prog_val = "---"
                teach_val = "---"
                if sel_class.lower() in ['summercamp', 'preschool']:
                    prog_val = sel_class
                elif not classes_df.empty:
                    c_name_col = next((c for c in classes_df.columns if 'class_name' in c.lower() or c.lower() in ['class', 'τμήμα', 'τμημα']), 'class_name')
                    matched_cls = classes_df[classes_df[c_name_col] == str(sel_class)]
                    if not matched_cls.empty:
                        cls_row = matched_cls.iloc[0]
                        prog_val = str(cls_row.get('program_name', cls_row.get('Πρόγραμμα', '---')))
                        teach_val = str(cls_row.get('teachers', cls_row.get('Καθηγητής', '---')))
    
                active_students_list = []
    
                if not reg_df.empty and sel_class.lower() not in ['summercamp', 'preschool']:
                    class_col_r = next((c for c in reg_df.columns if c.lower() in ['class', 'class_name', 'τμήμα', 'τμημα']), 'class')
                    year_col_r = next((c for c in reg_df.columns if 'year' in c.lower() or 'έτος' in c.lower()), 'school_year')
                    status_col_r = next((c for c in reg_df.columns if 'status' in c.lower() or 'κατάσταση' in c.lower()), 'status')
                    
                    mask_reg = (
                        (reg_df[class_col_r] == str(sel_class)) & 
                        (reg_df[year_col_r] == str(sel_year_t2)) & 
                        (reg_df[status_col_r].str.upper() != 'STOP')
                    ) if status_col_r in reg_df.columns else (
                        (reg_df[class_col_r] == str(sel_class)) & 
                        (reg_df[year_col_r] == str(sel_year_t2))
                    )
                    
                    matched_regs = reg_df[mask_reg]
                    for _, r in matched_regs.iterrows():
                        c_id = str(r.get('id_customer', ''))
                        fname = str(r.get('customer_firstname', ''))
                        lname = str(r.get('customer_lastname', ''))
                        if prog_val == "---" and 'program' in r:
                            prog_val = str(r.get('program', ''))
                        
                        tel = ""
                        email = ""
                        
                        if not customers_df.empty and c_id:
                            cust_id_col = next((c for c in customers_df.columns if 'id' in c.lower() or 'customer' in c.lower()), 'id_customer')
                            c_match = customers_df[customers_df[cust_id_col] == c_id]
                            if not c_match.empty:
                                tel = str(c_match.iloc[0].get('telephone', ''))
                                email = str(c_match.iloc[0].get('email', ''))
                                if not fname: fname = str(c_match.iloc[0].get('customer_firstname', ''))
                                if not lname: lname = str(c_match.iloc[0].get('customer_lastname', ''))
    
                        if fname or lname:
                            active_students_list.append({
                                'Όνομα': fname.upper(),
                                'Επώνυμο': lname.upper(),
                                'Τηλέφωνο': tel,
                                'Email': email
                            })
    
                if not sum_df.empty and sel_class.lower() in ['summercamp', 'preschool']:
                    for _, s in sum_df.iterrows():
                        s_year = str(s.get('summer_year', ''))
                        s_prog = str(s.get('summer_program', '')).lower()
                        s_stat = str(s.get('status_summer', 'ACTIVE')).upper()
                        
                        if s_year == str(sel_year_t2) and s_stat != 'STOP' and s_prog == sel_class.lower():
                            s_fname = str(s.get('summer_firstname', s.get('customer_firstnameSummer', '')))
                            s_lname = str(s.get('summer_lastname', s.get('customer_lastnameSummer', '')))
                            s_tel = str(s.get('summer_tel', s.get('telephone', '')))
                            s_email = str(s.get('summer_email', ''))
                            
                            if s_fname or s_lname:
                                active_students_list.append({
                                    'Όνομα': s_fname.upper(),
                                    'Επώνυμο': s_lname.upper(),
                                    'Τηλέφωνο': s_tel,
                                    'Email': s_email
                                })
    
                if active_students_list:
                    final_df = pd.DataFrame(active_students_list).drop_duplicates(subset=['Επώνυμο', 'Όνομα'])
                    
                    st.info(f"📚 Πρόγραμμα: {prog_val} | 👨‍🏫 Καθηγητής: {teach_val}\n\n**Αποπληρωμή Τμήματος:** {status_pay}")
                    st.write(f"**Σύνολο Μαθητών:** {len(final_df)}")
                    st.dataframe(final_df, use_container_width=True, hide_index=True)
                    
                    if st.button("📥 PDF Τμήματος", disabled=st.session_state.is_submitting):
                        rep = generate_class_pdf(final_df, sel_class, teach_val, prog_val, is_paid)
                        st.download_button("Κατέβασμα PDF", rep, f"class_{sel_class}.pdf", "application/pdf")
                else:
                    st.warning("Δεν βρέθηκαν μαθητές για αυτό το τμήμα και έτος.")
            else:
                st.warning("Δεν βρέθηκαν δεδομένα εγγραφών.")
    #-----------------------------------------------------------------------                    
    elif st.session_state.page == "finance_master":
        if st.session_state.role not in ["admin", "master"]:
            st.session_state.page = "menu"; st.rerun() 
    
        if st.button("⬅️ Επιστροφή στο Μενού", key="menu_master_fin", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 
    
        st.subheader("💰 Master Οικονομικά (Συγκεντρωτικές Κινήσεις)")
    
        st.markdown("---")
    
        # Απευθείας φόρτωση του db_finance_master (χωρίς να διαβάζει από άλλες βάσεις)
        final_master_df = load_master_db("db_finance_master")
    
        if not final_master_df.empty:
            # Καθαρισμός στηλών
            final_master_df.columns = final_master_df.columns.astype(str).str.strip()
            
            display_m_df = final_master_df.rename(columns={
                'date_entry': 'Ημερομηνία',
                'type': 'Τύπος',
                'category': 'Κατηγορία',
                'partner': 'Πελάτης/Προμηθευτής',
                'invoice_number': 'Αρ. Τιμολογίου',
                'amount': 'Ποσό (€)',
                'payment_method': 'Μέθοδος',
                'school_year': 'Σχολικό Έτος',
                'source': 'Πηγή'
            })
            
            # --- ΦΙΛΤΡΑ ΑΝΑΖΗΤΗΣΗΣ (ΜΕ ΕΥΡΟΣ ΗΜΕΡΟΜΗΝΙΩΝ) ---
            col_date1, col_date2 = st.columns(2)
            start_date = col_date1.date_input("📅 Από Ημερομηνία", value=None)
            end_date = col_date2.date_input("📅 Έως Ημερομηνία", value=None)
    
            st.markdown("---")
    
            col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
            
            unique_types = sorted([str(t) for t in display_m_df['Τύπος'].unique() if t and t != ""])
            sel_type = col_f1.selectbox("📂 Τύπος", ["ΟΛΑ"] + unique_types, key="master_type_filter")
            
            unique_custs = sorted([c for c in display_m_df['Πελάτης/Προμηθευτής'].unique() if c and c != ""])
            sel_cust = col_f2.selectbox("👤 Αναζήτηση Πελάτη", ["ΟΛΟΙ"] + unique_custs, key="master_cust_filter")
            
            unique_years = sorted([str(y) for y in display_m_df['Σχολικό Έτος'].unique() if y and y != ""])
            sel_year = col_f3.selectbox("📅 Σχολικό Έτος", ["ΟΛΑ"] + unique_years, key="master_year_filter")
            
            unique_methods = sorted([str(m) for m in display_m_df['Μέθοδος'].unique() if m and m != ""])
            sel_method = col_f4.selectbox("💳 Μέθοδος", ["ΟΛΑ"] + unique_methods, key="master_method_filter")
            
            unique_sources = sorted([str(s) for s in display_m_df['Πηγή'].unique() if s and s != ""])
            sel_source = col_f5.selectbox("📂 Πηγή", ["ΟΛΑ"] + unique_sources, key="master_source_filter")
    
            # Εφαρμογή φίλτρων
            filtered_master = display_m_df.copy()
            
            # Φίλτρο εύρους ημερομηνιών
            if start_date or end_date:
                parsed_dates = pd.to_datetime(filtered_master['Ημερομηνία'], errors='coerce').dt.date
                if start_date:
                    filtered_master = filtered_master[parsed_dates >= start_date]
                if end_date:
                    filtered_master = filtered_master[parsed_dates <= end_date]
    
            if sel_type != "ΟΛΑ":
                filtered_master = filtered_master[filtered_master['Τύπος'] == sel_type]
            if sel_cust != "ΟΛΟΙ":
                filtered_master = filtered_master[filtered_master['Πελάτης/Προμηθευτής'] == sel_cust]
            if sel_year != "ΟΛΑ":
                filtered_master = filtered_master[filtered_master['Σχολικό Έτος'] == sel_year]
            if sel_method != "ΟΛΑ":
                filtered_master = filtered_master[filtered_master['Μέθοδος'] == sel_method]
            if sel_source != "ΟΛΑ":
                filtered_master = filtered_master[filtered_master['Πηγή'] == sel_source]
    
            st.dataframe(filtered_master, use_container_width=True, hide_index=True)
            
            # --- ΑΣΦΑΛΗΣ ΥΠΟΛΟΓΙΣΜΟΣ ΣΥΝΟΛΩΝ (ΕΣΟΔΑ / ΕΞΟΔΑ) ---
            def parse_safe_amount(val):
                if pd.isna(val):
                    return 0.0
                val_str = str(val).strip().replace(',', '.')
                try:
                    return float(val_str)
                except ValueError:
                    cleaned = "".join(c for c in val_str if c.isdigit() or c == '.' or c == '-')
                    return float(cleaned) if cleaned else 0.0
    
            filtered_master['Amount_Clean'] = filtered_master['Ποσό (€)'].apply(parse_safe_amount)
            filtered_master['Type_Clean'] = filtered_master['Τύπος'].astype(str).str.strip().str.upper()
    
            inc_vals = filtered_master[filtered_master['Type_Clean'] == 'ΕΣΟΔΟ']['Amount_Clean'].sum()
            exp_vals = filtered_master[filtered_master['Type_Clean'] == 'ΕΞΟΔΟ']['Amount_Clean'].sum()
            net_vals = inc_vals - exp_vals
    
            st.markdown(
                f"📊 **Σύνολα Κινήσεων:** "
                f"🟢 Έσοδα: **{inc_vals:.2f}€** | "
                f"🔴 Έξοδα: **{exp_vals:.2f}€** | "
                f"⚖️ Καθαρό Υπόλοιπο (Έσοδα - Έξοδα): **{net_vals:.2f}€**"
            )
    
            # --- ΠΑΡΑΓΩΓΗ PDF MASTER ΟΙΚΟΝΟΜΙΚΩΝ ---
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Παραγωγή PDF Master Κινήσεων", key="btn_pdf_master_k", disabled=st.session_state.is_submitting):
                pdf_data = generate_master_finances_pdf(
                    selected_cust=sel_cust, 
                    selected_year=sel_year, 
                    master_df=filtered_master
                )
                if pdf_data:
                    st.download_button(
                        label="📥 Κατεβάστε το PDF Master Οικονομικών",
                        data=pdf_data,
                        file_name="master_finances_report.pdf",
                        mime="application/pdf",
                        key="dl_pdf_master_k"
                    )
        else:
            st.info("Το Master sheet είναι κενό.")
#---------------------------------------------------------------------------------------------------------------
    elif st.session_state.page == "simple_customer_reports":
        if "role" not in st.session_state:
            st.session_state.page = "menu"

        if st.button("⬅️️ Επιστροφή στο Μενού", key="menu_simple_rep", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun()

        st.subheader("Μαθητολόγιο Αυτοχρηματοδοτούμενων Προγραμμάτων")

        pays_df = load_specific_db("db_payments")
        sum_pays_df = load_specific_db("db_payment_summer")
        reg_df = load_specific_db("db_registrations")
        summer_reg_df = load_specific_db("db_summer")
        progs_df = load_specific_db("db_programs")

        for df in [pays_df, sum_pays_df, reg_df, summer_reg_df, progs_df]:
            if not df.empty:
                df.columns = df.columns.astype(str).str.strip()

        valid_records = []

        # 1. ΒΑΣΗ ΟΛΩΝ ΤΩΝ ΕΓΓΡΑΦΩΝ
        if not reg_df.empty:
            grouped_regs = reg_df.groupby(['id_customer', 'customer_lastname', 'customer_firstname', 'school_year'])
            
            for (cust_id, lname_raw, fname_raw, year_raw), group in grouped_regs:
                lname = str(lname_raw).strip()
                fname = str(fname_raw).strip()
                year = str(year_raw).strip()
                
                if not lname or not fname or not year:
                    continue

                # Υπολογισμός συνολικού κόστους και εύρεσης κύριου προγράμματος
                best_prog = "---"
                total_student_cost = 0.0
                max_single_cost = -1.0

                for _, r in group.iterrows():
                    p_name = str(r.get('program', '')).strip()
                    adj_c = str(r.get('adjusted_cost', '0')).replace(',', '.')
                    p_cost = float(adj_c) if adj_c.strip() != '' and float(adj_c) > 0 else 0.0
                    
                    # Αν το adjusted_cost είναι 0, ΔΕΝ τραβάμε αναγκαστικά το βασικό κόστος προγράμματος 
                    # αν ο μαθητής είναι σε πολλαπλά προγράμματα και το ποσό έχει μπει αλλού.
                    # Ελέγχουμε αν η γραμμή έχει ρητά adjusted_cost.
                    if p_cost > max_single_cost:
                        max_single_cost = p_cost
                        if p_name:
                            best_prog = p_name

                    total_student_cost += p_cost

                # Αν το total_student_cost από τα adjusted_costs είναι 0 αλλά υπάρχουν προγράμματα, 
                # κοιτάμε μήπως το κόστος είναι καταχωρημένο στον τιμοκατάλογο (για μοναδικό πρόγραμμα).
                if total_student_cost == 0.0 and len(group) == 1 and not progs_df.empty:
                    p_name = str(group.iloc[0].get('program', '')).strip()
                    p_match = progs_df[progs_df['program_name'].astype(str).str.strip().str.lower() == p_name.lower()]
                    if not p_match.empty:
                        total_student_cost = float(p_match.iloc[0].get('program_cost', 0) or 0.0)

                # Συλλογή ΟΛΩΝ των πληρωμών για διαχωρισμό ΑΠΥ/ΤΠΥ και :-)
                matched_pays = pd.DataFrame()
                if not pays_df.empty:
                    matched_pays = pays_df[
                        (pays_df['id_customer'].astype(str).str.strip() == str(cust_id)) &
                        (pays_df['school_year'].astype(str).str.strip() == year)
                    ]

                receipts_list = []
                dates_list = []
                valid_receipts_sum = 0.0
                non_official_payments_sum = 0.0  # Για τις πληρωμές :-)

                if not matched_pays.empty:
                    for _, p_row in matched_pays.iterrows():
                        rec_num = str(p_row.get('receipt_number', '')).strip().upper()
                        p_amt = float(str(p_row.get('amount_paid', 0)).replace(',', '.') or 0)
                        p_date = str(p_row.get('payment_date', '')).strip()
                        
                        if (rec_num.startswith("ΑΠΥ") or rec_num.startswith("ΤΠΥ")) and p_amt > 0:
                            receipts_list.append(rec_num)
                            if p_date and p_date.lower() != 'nan':
                                dates_list.append(p_date)
                            valid_receipts_sum += p_amt
                        else:
                            if p_amt > 0:
                                non_official_payments_sum += p_amt

                receipts_str = ", ".join(receipts_list) if receipts_list else "---"
                dates_str = ", ".join(dates_list) if dates_list else "---"

                # Μείωση του συμφωνηθέντος ποσού σύμφωνα με τις πληρωμές :-)
                adjusted_student_cost = max(0.0, total_student_cost - non_official_payments_sum)

                valid_records.append({
                    'Επώνυμο': lname.upper(),
                    'Όνομα': fname.upper(),
                    'Σχολικό Έτος': year,
                    'Πρόγραμμα': best_prog,
                    'Αρ. Απόδειξης': receipts_str,
                    'Ημερομηνία': dates_str,
                    'Ποσό Απόδειξης (€)': float(valid_receipts_sum),
                    'Συμφωνηθέν Ποσό (€)': float(adjusted_student_cost),
                    'Πηγή': 'Registrations'
                })

        # 2. ΠΡΟΣΘΗΚΗ SUMMER ΕΓΓΡΑΦΩΝ
        if not summer_reg_df.empty:
            for _, s in summer_reg_df.iterrows():
                lname = str(s.get('summer_lastname', s.get('customer_lastnameSummer', ''))).strip()
                fname = str(s.get('summer_firstname', s.get('customer_firstnameSummer', ''))).strip()
                year = str(s.get('summer_year', '')).strip()
                prog = str(s.get('summer_program', 'SummerCamp')).strip()
                s_cost = float(str(s.get('summer_cost', 0)).replace(',', '.') or 0)

                if lname and fname and year:
                    matched_s_pays = pd.DataFrame()
                    if not sum_pays_df.empty:
                        s_lname_col = 'summer_lastname' if 'summer_lastname' in sum_pays_df.columns else 'customer_lastnameSummer'
                        s_fname_col = 'summer_firstname' if 'summer_firstname' in sum_pays_df.columns else 'customer_firstnameSummer'
                        s_year_col = 'school_yearSummer' if 'school_yearSummer' in sum_pays_df.columns else 'school_year'
                        
                        matched_s_pays = sum_pays_df[
                            (sum_pays_df[s_lname_col].astype(str).str.strip().str.lower() == lname.lower()) &
                            (sum_pays_df[s_fname_col].astype(str).str.strip().str.lower() == fname.lower()) &
                            (sum_pays_df[s_year_col].astype(str).str.strip() == year)
                        ]

                    receipts_list_s = []
                    dates_list_s = []
                    valid_receipts_sum_s = 0.0
                    non_official_sum_s = 0.0

                    if not matched_s_pays.empty:
                        for _, sp_row in matched_s_pays.iterrows():
                            rec_num = str(sp_row.get('receipt_numberSummer', sp_row.get('receipt_number', ''))).strip().upper()
                            sp_amt = float(str(sp_row.get('amount_paidSummer', sp_row.get('amount_paid', 0))).replace(',', '.') or 0)
                            sp_date = str(sp_row.get('payment_dateSummer', sp_row.get('payment_date', ''))).strip()
                            
                            if (rec_num.startswith("ΑΠΥ") or rec_num.startswith("ΤΠΥ")) and sp_amt > 0:
                                receipts_list_s.append(rec_num)
                                if sp_date and sp_date.lower() != 'nan':
                                    dates_list_s.append(sp_date)
                                valid_receipts_sum_s += sp_amt
                            else:
                                if sp_amt > 0:
                                    non_official_sum_s += sp_amt

                    receipts_str_s = ", ".join(receipts_list_s) if receipts_list_s else "---"
                    dates_str_s = ", ".join(dates_list_s) if dates_list_s else "---"
                    adjusted_summer_cost = max(0.0, s_cost - non_official_sum_s)

                    valid_records.append({
                        'Επώνυμο': lname.upper(),
                        'Όνομα': fname.upper(),
                        'Σχολικό Έτος': year,
                        'Πρόγραμμα': prog,
                        'Αρ. Απόδειξης': receipts_str_s,
                        'Ημερομηνία': dates_str_s,
                        'Ποσό Απόδειξης (€)': float(valid_receipts_sum_s),
                        'Συμφωνηθέν Ποσό (€)': float(adjusted_summer_cost),
                        'Πηγή': 'Summer'
                    })

        if valid_records:
            df_active_cust = pd.DataFrame(valid_records)

            # --- ΦΙΛΤΡΑ ΑΝΑΖΗΤΗΣΗΣ ---
            col_f1, col_f2 = st.columns(2)
            unique_years = sorted(df_active_cust['Σχολικό Έτος'].unique().tolist(), reverse=True)
            sel_year_filter = col_f1.selectbox("📅 Φίλτρο ανά Σχολικό Έτος", ["ΟΛΑ"] + unique_years, key="simple_rep_year_filter")

            unique_customers = sorted(list(set([
                f"{r['Επώνυμο']} {r['Όνομα']}" for r in valid_records
            ])))

            sel_cust_filter = col_f2.selectbox("🔍 Αναζήτηση Πελάτη", ["ΟΛΟΙ"] + unique_customers, key="simple_rep_cust_filter")
        
            # Εφαρμογή φίλτρων
            filtered_df = df_active_cust.copy()
            if sel_year_filter != "ΟΛΑ":
                filtered_df = filtered_df[filtered_df['Σχολικό Έτος'] == sel_year_filter]
            if sel_cust_filter != "ΟΛΟΙ":
                filtered_df = filtered_df[(filtered_df['Επώνυμο'] + " " + filtered_df['Όνομα']) == sel_cust_filter]

            st.dataframe(filtered_df[['Επώνυμο', 'Όνομα', 'Σχολικό Έτος', 'Πρόγραμμα', 'Αρ. Απόδειξης', 'Ημερομηνία', 'Ποσό Απόδειξης (€)', 'Συμφωνηθέν Ποσό (€)']], use_container_width=True, hide_index=True)
            
            # Υπολογισμός συνολικών ποσών (με μοναδικά συμφωνηθέντα ανά μαθητή/έτος για αποφυγή διπλασιασμού)
            unique_agreements = filtered_df.drop_duplicates(subset=['Επώνυμο', 'Όνομα', 'Σχολικό Έτος'])
            total_agreed = unique_agreements['Συμφωνηθέν Ποσό (€)'].sum()
            total_receipts = filtered_df['Ποσό Απόδειξης (€)'].sum()

            st.markdown(f"**Συνολικό Ποσό Αποδείξεων:** {total_receipts:.2f}€ | **Συνολικό Συμφωνηθέν Ποσό:** {total_agreed:.2f}€")

            # --- ΠΑΡΑΓΩΓΗ PDF ---
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📥 Παραγωγή PDF Αναφοράς", key="btn_pdf_simple_rep", disabled=st.session_state.is_submitting):
                pdf_data = generate_active_customers_pdf(sel_year_filter, filtered_df)
                if pdf_data:
                    st.download_button(
                        label="📥 Κατεβάστε το PDF Αναφοράς",
                        data=pdf_data,
                        file_name="active_customers_receipts_report.pdf",
                        mime="application/pdf",
                        key="dl_pdf_simple_rep"
                    )
        else:
            st.info("Δεν βρέθηκαν καταχωρημένες εγγραφές μαθητών.")
            
    #------------------------------------------------------------------
    elif st.session_state.page == "payment_dypa":
        if "role" not in st.session_state:
            st.session_state.page = "menu"
            

        if st.button("⬅️ Επιστροφή στο Μενού", key="menu_dypa_back", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.page = "menu"
            

        st.markdown("### Διαχείριση Αποδείξεων επί Πιστώσει ΔΥΠΑ")
        st.markdown("---")

        dypa_df = load_specific_db("db_payment_dypa")
        reg_df = load_specific_db("db_registrations")
        cust_df = load_specific_db("db_customers")

        # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ ΑΠΟΔΕΙΞΕΩΝ ΔΥΠΑ ---
        st.markdown("#### 📋 Υπάρχουσες Αποδείξεις ΔΥΠΑ")
        if not dypa_df.empty:
            dypa_df.columns = dypa_df.columns.astype(str).str.strip()
            st.dataframe(dypa_df, use_container_width=True, hide_index=True)
            
            # Εντοπισμός στήλης ID (συνήθως id_payment ή id)
            id_col_dypa = next((c for c in dypa_df.columns if 'id' in c.lower()), dypa_df.columns[0])
            dypa_ids = dypa_df[id_col_dypa].astype(str).tolist()
            
            sel_dypa_edit = st.selectbox("Επιλογή ID για Τροποποίηση", ["-- Επιλέξτε ID --"] + dypa_ids, key="sel_dypa_edit_box")
            
            if sel_dypa_edit != "-- Επιλέξτε ID --":
                matched_d_row = dypa_df[dypa_df[id_col_dypa].astype(str) == str(sel_dypa_edit)].iloc[0]
                old_dypa_receipt = str(matched_d_row.get('receipt_number', matched_d_row.iloc[5] if len(matched_d_row) > 5 else '')).strip()
                
                with st.form("edit_dypa_form"):
                    st.markdown(f"**Επεξεργασία Απόδειξης ΔΥΠΑ με ID: {sel_dypa_edit}**")
                    ed_rec_dypa = st.text_input("Αριθμός Απόδειξης (πρέπει να ξεκινάει με ΑΠΥ)", value=old_dypa_receipt).strip()
                    ed_amt_dypa = st.number_input("Ποσό (€)", min_value=0.0, value=float(pd.to_numeric(matched_d_row.get('amount_paid', matched_d_row.get('amount', 0)), errors='coerce') or 0.0), format="%.2f")
                    ed_year_dypa = st.text_input("Σχολικό Έτος", value=str(matched_d_row.get('school_year', '2026-2027'))).strip()
                    
                    if st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης ΔΥΠΑ", disabled=st.session_state.is_submitting):
                        if not ed_rec_dypa.upper().startswith("ΑΠΥ"):
                            st.error("Σφάλμα: Ο αριθμός απόδειξης πρέπει να ξεκινάει υποχρεωτικά με τη λέξη 'ΑΠΥ'!")
                        elif ed_amt_dypa <= 0:
                            st.error("Σφάλμα: Το ποσό πρέπει να είναι μεγαλύτερο από 0.")
                        else:
                            try:
                                gc = get_gspread_client()
                                sh_old = gc.open(OLD_SHEET_NAME).worksheet("db_payment_dypa")
                                
                                # Εύρεση γραμμής στο Google Sheet
                                all_raw_dypa = sh_old.get_all_values()
                                target_row_idx = None
                                if len(all_raw_dypa) > 1:
                                    for r_idx, sheet_row in enumerate(all_raw_dypa[1:], start=2):
                                        if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_dypa_edit).strip():
                                            target_row_idx = r_idx
                                            break
                                
                                if target_row_idx:
                                    # Στήλες db_payment_dypa: id(1), customer_id(2), lastname(3), firstname(4), date(5), receipt_number(6), amount(7), method(8), program(9), class(10), school_year(11)
                                    sh_old.update_cell(target_row_idx, 6, ed_rec_dypa)
                                    sh_old.update_cell(target_row_idx, 7, float(ed_amt_dypa))
                                    sh_old.update_cell(target_row_idx, 11, ed_year_dypa)
                                    
                                    # Ενημέρωση και στο Master (db_finance_master) στο νέο sheet
                                    try:
                                        sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                        master_rows = sh_master.get_all_values()
                                        if len(master_rows) > 1:
                                            for m_idx, m_row in enumerate(master_rows[1:], start=2):
                                                if len(m_row) > 5 and str(m_row[5]).strip() == old_dypa_receipt:
                                                    sh_master.update_cell(m_idx, 6, ed_rec_dypa) # invoice_number
                                                    sh_master.update_cell(m_idx, 7, float(ed_amt_dypa)) # amount
                                                    sh_master.update_cell(m_idx, 9, ed_year_dypa) # school_year
                                                    break
                                    except Exception:
                                        pass

                                    #load_specific_db.clear()
                                    #load_master_db.clear()
                                    st.success("✅ Η απόδειξη ΔΥΠΑ τροποποιήθηκε επιτυχώς και ενημερώθηκε το Master!")
                                    time.sleep(1.2)
                                    
                                else:
                                    st.error("Δεν βρέθηκε η εγγραφή στη βάση.")
                            except Exception as edit_err:
                                st.error(f"Σφάλμα τροποποίησης: {edit_err}")
        else:
            st.info("Δεν βρέθηκαν καταχωρημένες αποδείξεις ΔΥΠΑ.")

        st.markdown("---")
        st.markdown("#### Νέα Καταχώρηση επί Πιστώσει")
        
        cust_options = []
        cust_records_map = {}
        
        # Καθαρισμός στηλών customers για αποφυγή λαθών
        if not cust_df.empty:
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            
            for _, c_row in cust_df.iterrows():
                # Εύρεση κατηγορίας (ψάχνουμε τη στήλη customer_category ή παρόμοια)
                cat_val = ""
                for col in cust_df.columns:
                    if 'category' in col.lower() or 'κατηγορ' in col.lower():
                        cat_val = str(c_row.get(col, '')).strip()
                        break
                
                # Αν είναι ενήλικας
                if cat_val.lower() in ["ενήλικας", "ενηλικας", "adult", "ενήλικες"]:
                    # Εύρεση ID, Επωνύμου, Ονόματος
                    cid = str(c_row.get('id_customer', c_row.get('id', c_row.iloc[0] if len(c_row) > 0 else ''))).strip()
                    lname = str(c_row.get('customer_lastname', c_row.get('lastname', c_row.iloc[2] if len(c_row) > 2 else ''))).strip()
                    fname = str(c_row.get('customer_firstname', c_row.get('firstname', c_row.iloc[1] if len(c_row) > 1 else ''))).strip()
                    
                    if lname or fname:
                        display_name = f"{lname} {fname} (ID: {cid})"
                        if display_name not in cust_options:
                            cust_options.append(display_name)
                            cust_records_map[display_name] = []

                        # Ψάχναμε αν έχει εγγραφές στο reg_df
                        has_reg = False
                        if not reg_df.empty:
                            reg_df.columns = reg_df.columns.astype(str).str.strip()
                            matched_regs = reg_df[reg_df['id_customer'].astype(str).str.strip() == cid]
                            if not matched_regs.empty:
                                has_reg = True
                                for _, r in matched_regs.iterrows():
                                    cust_records_map[display_name].append({
                                        'id_customer': cid,
                                        'firstname': fname,
                                        'lastname': lname,
                                        'school_year': str(r.get('school_year', '2026-2027')).strip(),
                                        'program': str(r.get('program', 'Ενήλικες')).strip(),
                                        'class': str(r.get('class', '---')).strip()
                                    })
                        
                        # Αν δεν έχει καμία εγγραφή στο reg_df, του βάζουμε προεπιλεγμένα
                        if not has_reg:
                            cust_records_map[display_name].append({
                                'id_customer': cid,
                                'firstname': fname,
                                'lastname': lname,
                                'school_year': '2026-2027',
                                'program': 'Ενήλικες',
                                'class': '---'
                            })

        # Επιλογή πελάτη εκτός φόρμας για να ενημερώνονται άμεσα τα πεδία
        selected_cust = st.selectbox("Επιλογή Πελάτη / Μαθητή (Ενήλικες)", ["--- Επιλέξτε ---"] + sorted(cust_options), key="dypa_sel_cust")
        
        available_years = []
        available_programs = []
        available_classes = []
        def_cid = ""
        def_lname = ""
        def_fname = ""

        if selected_cust != "--- Επιλέξτε ---" and selected_cust in cust_records_map:
            records = cust_records_map[selected_cust]
            if records:
                def_cid = records[0]['id_customer']
                def_lname = records[0]['lastname']
                def_fname = records[0]['firstname']
                available_years = sorted(list(set([x['school_year'] for x in records if x['school_year'] and x['school_year'] != '---'])), reverse=True)
                if not available_years: available_years = ['2026-2027']
                
                available_programs = sorted(list(set([x['program'] for x in records if x['program'] and x['program'] != '---'])))
                if not available_programs: available_programs = ['Ενήλικες']
                
                available_classes = sorted(list(set([x['class'] for x in records if x['class'] and x['class'] != '---'])))
                if not available_classes: available_classes = ['---']

        c1, c2 = st.columns(2)
        pay_date_input = c1.date_input("Ημερομηνία Πληρωμής", value=datetime.today(), key="dypa_date")
        receipt_input = c2.text_input("Αριθμός Απόδειξης (πρέπει να ξεκινάει με ΑΠΥ)", key="dypa_receipt")

        c3, c4 = st.columns(2)
        amount_input = c3.text_input("Ποσό (€)", value="0.0", key="dypa_amount")
        year_input = c4.selectbox("Σχολικό Έτος", available_years if available_years else ["2026-2027"], key="dypa_year")

        c5, c6 = st.columns(2)
        prog_input = c5.selectbox("Πρόγραμμα", available_programs if available_programs else ["Ενήλικες"], key="dypa_prog")
        class_input = c6.selectbox("Τμήμα", available_classes if available_classes else ["---"], key="dypa_class")

        st.markdown("**Τρόπος Πληρωμής:** ΕΠΙ ΠΙΣΤΩΣΕΙ (Μόνιμο)")

        if st.button("Αποθήκευση Απόδειξης ΔΥΠΑ", key="btn_save_dypa", disabled=st.session_state.is_submitting):
            if selected_cust == "--- Επιλέξτε ---":
                st.error("Σφάλμα: Παρακαλώ επιλέξτε πελάτη!")
            elif not receipt_input.upper().startswith("ΑΠΥ"):
                st.error("Σφάλμα: Ο αριθμός απόδειξης πρέπει να ξεκινάει υποχρεωτικά με τη λέξη 'ΑΠΥ'!")
            else:
                try:
                    gc = get_gspread_client()
                    
                    # 1. Αποθήκευση κανονικά στο παλιό sheet (db_payment_dypa)
                    sh_old = gc.open(OLD_SHEET_NAME).worksheet("db_payment_dypa")
                    recs_old = sh_old.get_all_records()
                    next_id_old = 1
                    if recs_old:
                        df_t_old = pd.DataFrame(recs_old)
                        num_ids_old = pd.to_numeric(df_t_old.iloc[:, 0], errors='coerce').dropna()
                        next_id_old = int(num_ids_old.max()) + 1 if not num_ids_old.empty else len(df_t_old) + 1
                    
                    parsed_amount = float(str(amount_input).replace(',', '.') or 0)
                    
                    sh_old.append_row([
                        next_id_old,
                        def_cid,
                        def_lname,
                        def_fname,
                        pay_date_input.strftime('%Y-%m-%d'),
                        receipt_input,
                        parsed_amount,
                        "ΕΠΙ ΠΙΣΤΩΣΕΙ",
                        prog_input,
                        class_input,
                        year_input
                    ])

                    # 2. Αποθήκευση απευθείας στο db_finance_master του νέου sheet (NEW_SHEET_ID)
                    try:
                        sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                        master_recs = sh_master.get_all_records()
                        next_id_master = 1
                        if master_recs:
                            df_m = pd.DataFrame(master_recs)
                            num_m = pd.to_numeric(df_m.iloc[:, 0], errors='coerce').dropna()
                            next_id_master = int(num_m.max()) + 1 if not num_m.empty else len(df_m) + 1
                        
                        cust_full_name = f"{def_lname} {def_fname}".strip()
                        sh_master.append_row([
                            next_id_master,
                            pay_date_input.strftime('%Y-%m-%d'),
                            "ΕΣΟΔΟ",
                            f"Επί Πιστώσει ΔΥΠΑ ({prog_input})",
                            cust_full_name,
                            receipt_input,
                            parsed_amount,
                            "ΕΠΙ ΠΙΣΤΩΣΕΙ",
                            year_input,
                            "DYPA"
                        ])
                    except Exception as err_m:
                        st.warning(f"Η εγγραφή αποθηκεύτηκε τοπικά αλλά υπήρξε θέμα στο master sheet: {err_m}")

                    #load_specific_db.clear()
                    #load_master_db.clear()
                    st.success("Η απόδειξη ΔΥΠΑ αποθηκεύτηκε επιτυχώς και ενημερώθηκε το Master!")
                    time.sleep(1.2)
                    
                except Exception as e:
                    st.error(f"Σφάλμα αποθήκευσης: {e}")

    elif st.session_state.page == "dypa_customer_reports":
        if "role" not in st.session_state:
            st.session_state.page = "menu"; st.rerun()

        if st.button("⬅️ Επιστροφή στο Μενού", key="menu_dypa_rep", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun()

        st.subheader("Μαθητολόγιο Χρηματοδοτούμενων Προγραμμάτων - ΔΥΠΑ")

        dypa_df = load_specific_db("db_payment_dypa")
        cust_df = load_specific_db("db_customers")

        for df in [dypa_df, cust_df]:
            if not df.empty:
                df.columns = df.columns.astype(str).str.strip()

        if dypa_df.empty or cust_df.empty:
            st.warning("Δεν βρέθηκαν δεδομένα στις βάσεις `db_payment_dypa` ή `db_customers`.")
        else:
            # Φιλτράρισμα πελατών αυστηρά με κατηγορία 'Ενήλικας'
            adult_custs = cust_df[cust_df['customer_category'].astype(str).str.strip().str.lower().isin(['ενήλικας', 'ενηλικας', 'adult'])]
            adult_cust_ids = set(adult_custs['id_customer'].astype(str).str.strip().tolist())

            # Επεξεργασία δεδομένων dypa_df (ΜΟΝΟ επί πιστώσει και ΜΟΝΟ Ενήλικες)
            valid_dypa_records = []
            for _, r in dypa_df.iterrows():
                pay_method = str(r.get('pay_method', '')).strip().upper()
                cust_id = str(r.get('id_customer', '')).strip()

                # Έλεγχος αν είναι επί πιστώσει και αν ο πελάτης είναι Ενήλικας
                if ("ΠΙΣΤΩΣΕΙ" in pay_method or pay_method == "ΕΠΙ ΠΙΣΤΩΣΕΙ") and cust_id in adult_cust_ids:
                    lname = str(r.get('customer_lastname', '')).strip()
                    fname = str(r.get('customer_firstname', '')).strip()
                    year = str(r.get('school_year', '')).strip()
                    prog = str(r.get('program', '')).strip()
                    clazz = str(r.get('class', '')).strip()
                    receipt = str(r.get('receipt_number', '')).strip()
                    date_val = str(r.get('payment_date', '')).strip()
                    amt = float(str(r.get('amount_paid', 0)).replace(',', '.') or 0)

                    valid_dypa_records.append({
                        'ID': cust_id,
                        'Επώνυμο': lname,
                        'Όνομα': fname,
                        'Σχολικό Έτος': year,
                        'Πρόγραμμα': prog,
                        'Τμήμα': clazz if clazz and clazz.lower() != 'nan' else '---',
                        'Αρ. Απόδειξης': receipt,
                        'Ημερομηνία': date_val,
                        'Ποσό (€)': amt,
                        'display': f"{lname} {fname}"
                    })

            if valid_dypa_records:
                df_dypa_cust = pd.DataFrame(valid_dypa_records)

                # --- ΦΙΛΤΡΑ ΑΝΑΖΗΤΗΣΗΣ (Ανά Πελάτη & ανά Τμήμα) ---
                col_f1, col_f2, col_f3 = st.columns(3)
                unique_years = sorted(df_dypa_cust['Σχολικό Έτος'].unique().tolist(), reverse=True)
                sel_year_filter = col_f1.selectbox("📅 Έτος", ["ΟΛΑ"] + unique_years, key="dypa_rep_year_filter")

                unique_customers = sorted(df_dypa_cust['display'].unique().tolist())
                sel_cust_filter = col_f2.selectbox("👤 Πελάτης", ["ΟΛΟΙ"] + unique_customers, key="dypa_rep_cust_filter")

                unique_classes = sorted(df_dypa_cust['Τμήμα'].unique().tolist())
                sel_class_filter = col_f3.selectbox("🏫 Τμήμα", ["ΟΛΑ"] + unique_classes, key="dypa_rep_class_filter")

                # Εφαρμογή φίλτρων
                filtered_df = df_dypa_cust.copy()
                if sel_year_filter != "ΟΛΑ":
                    filtered_df = filtered_df[filtered_df['Σχολικό Έτος'] == sel_year_filter]
                if sel_cust_filter != "ΟΛΟΙ":
                    filtered_df = filtered_df[filtered_df['display'] == sel_cust_filter]
                if sel_class_filter != "ΟΛΑ":
                    filtered_df = filtered_df[filtered_df['Τμήμα'] == sel_class_filter]

                st.dataframe(filtered_df[['Επώνυμο', 'Όνομα', 'Σχολικό Έτος', 'Πρόγραμμα', 'Τμήμα', 'Αρ. Απόδειξης', 'Ημερομηνία', 'Ποσό (€)']], use_container_width=True, hide_index=True)
                
                total_dypa_amt = filtered_df['Ποσό (€)'].sum() if not filtered_df.empty else 0.0
                st.markdown(f"**Συνολικό Πσό ΔΥΠΑ (Επί Πιστώσει):** {total_dypa_amt:.2f}€")

                # --- ΠΑΡΑΓΩΓΗ PDF ---
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("📥 Παραγωγή PDF Αναφοράς ΔΥΠΑ", key="btn_pdf_dypa_rep", disabled=st.session_state.is_submitting):
                    pdf_data = generate_dypa_reports_pdf(sel_year_filter, sel_class_filter, filtered_df)
                    if pdf_data:
                        st.download_button(
                            label="📥 Κατεβάστε το PDF Αναφοράς ΔΥΠΑ",
                            data=pdf_data,
                            file_name="dypa_credit_reports_report.pdf",
                            mime="application/pdf",
                            key="dl_pdf_dypa_rep"
                        )
            else:
                st.info("Δεν βρέθηκαν καταχωρήσεις επί πιστώσει για ενήλικες πελάτες στο `db_payment_dypa`.")
    #---------------------------------------------------------------------------------
    elif st.session_state.page == "complaints_page":
        #load_specific_db.clear()

        if st.button("⬅️ Μενού", key="menu_complaints_back", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 
            
        st.subheader("🛠️ Διαχείριση & Παράπονα / Σχόλια Πελατών")
        
        cust_df = load_specific_db("db_customers")
        complaints_df = load_specific_db("db_complaint")
        teachers_df = load_specific_db("db_teachers")
        
        for df in [cust_df, complaints_df, teachers_df]:
            if not df.empty: df.columns = df.columns.astype(str).str.strip()

        if not cust_df.empty:
            cust_df['id_customer'] = pd.to_numeric(cust_df['id_customer'], errors='coerce').fillna(0).astype(int)
            cust_df['display'] = (cust_df['customer_lastname'].astype(str) + " " + cust_df['customer_firstname'].astype(str) + " (ID: " + cust_df['id_customer'].astype(str) + ")")
            
            sel_cust = st.selectbox("🔍 Επιλογή Πελάτη για Καταχώρηση / Προβολή", options=["-- Επιλέξτε --"] + cust_df['display'].tolist(), index=0)
            
            if sel_cust != "-- Επιλέξτε --":
                c_row = cust_df[cust_df['display'] == sel_cust].iloc[0]
                c_id = int(c_row['id_customer'])
                
                # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ ΕΓΓΡΑΦΩΝ ΠΕΛΑΤΗ ---
                st.markdown("---")
                st.markdown(f"##### 📋 Υπάρχοντα Παράπονα / Σχόλια Πελάτη: {c_row['customer_lastname']} {c_row['customer_firstname']}")
                
                if not complaints_df.empty:
                    complaints_df['id_cust_num'] = pd.to_numeric(complaints_df['id_customer'], errors='coerce')
                    cust_complaints = complaints_df[complaints_df['id_cust_num'] == float(c_id)].copy()
                    
                    if not cust_complaints.empty:
                        if 'status' not in cust_complaints.columns:
                            cust_complaints['status'] = 'ΕΝΕΡΓΟ'
                        if 'choice' not in cust_complaints.columns:
                            cust_complaints['choice'] = 'Παράπονο'

                        display_comps = cust_complaints[['id_complaint', 'choice', 'complaint_date', 'complaint_subject', 'complaint_text', 'complaint_staff', 'status']].copy()
                        display_comps.columns = ['ID', 'Τύπος', 'Ημερομηνία', 'Θέμα', 'Κείμενο', 'Υπεύθυνος', 'Κατάσταση']
                        
                        # Εμφάνιση με ρύθμιση ύψους grid για καλύτερη ανάγνωση μεγάλων κειμένων
                        st.dataframe(display_comps, use_container_width=True, hide_index=True, height=200)
                        
                        pay_ids_list = cust_complaints['id_complaint'].astype(str).tolist()
                        sel_comp_to_edit = st.selectbox("Επιλογή ID για Επεξεργασία", ["-- Επιλέξτε ID --"] + pay_ids_list, key="sel_comp_edit_box")
                        
                        if sel_comp_to_edit != "-- Επιλέξτε ID --":
                            matched_c_row = cust_complaints[cust_complaints['id_complaint'].astype(str) == str(sel_comp_to_edit)].iloc[0]
                            
                            teacher_options_edit = []
                            if not teachers_df.empty:
                                if 'full_name' in teachers_df.columns:
                                    teacher_options_edit = sorted(teachers_df['full_name'].dropna().astype(str).unique().tolist())
                                elif 'teacher_firstname' in teachers_df.columns and 'teacher_lastname' in teachers_df.columns:
                                    teacher_options_edit = sorted((teachers_df['teacher_lastname'].astype(str).str.strip() + " " + teachers_df['teacher_firstname'].astype(str).str.strip()).unique().tolist())
                                else:
                                    teacher_options_edit = sorted(teachers_df.iloc[:, 1].dropna().astype(str).unique().tolist())

                            current_staff = str(matched_c_row.get('complaint_staff', ''))
                            default_staff_idx = teacher_options_edit.index(current_staff) if current_staff in teacher_options_edit else 0

                            current_status = str(matched_c_row.get('status', 'ΕΝΕΡΓΟ')).strip().upper()
                            status_idx = 0 if current_status == 'ΕΝΕΡΓΟ' else 1

                            current_choice = str(matched_c_row.get('choice', 'Παράπονο')).strip()
                            choice_idx = 0 if current_choice == 'Παράπονο' else 1

                            try:
                                def_comp_dt = datetime.strptime(str(matched_c_row.get('complaint_date', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                            except:
                                def_comp_dt = datetime.today().date()

                            with st.form("edit_complaint_form"):
                                st.markdown(f"**Επεξεργασία Εγγραφής με ID: {sel_comp_to_edit}**")
                                ed_choice = st.selectbox("Τύπος Καταχώρησης", options=["Παράπονο", "Σχόλιο"], index=choice_idx)
                                ed_date = st.date_input("Ημερομηνία Καταχώρησης", value=def_comp_dt)
                                ed_subject = st.text_input("Θέμα", value=str(matched_c_row.get('complaint_subject', ''))).strip()
                                ed_text = st.text_area("Κείμενο", value=str(matched_c_row.get('complaint_text', ''))).strip()
                                ed_staff = st.selectbox("Υπεύθυνος / Staff", options=teacher_options_edit if teacher_options_edit else [st.session_state.get('username', 'Διαχειριστής')], index=default_staff_idx)
                                ed_status = st.selectbox("Κατάσταση", ["ΕΝΕΡΓΟ", "ΔΙΕΚΠΕΡΑΙΩΜΕΝΟ"], index=status_idx)
                                
                                submit_edit = st.form_submit_button("💾 Οριστικοποίηση Τροποποίησης", disabled=st.session_state.is_submitting)

                                if submit_edit:
                                    st.session_state.is_submitting = True
                                    if not ed_subject or not ed_text:
                                        st.warning("⚠️ Συμπλήρωσε θέμα και κείμενο!")
                                        st.session_state.is_submitting = False
                                    else:
                                        try:
                                            gc = get_gspread_client()
                                            sh_comps = gc.open(OLD_SHEET_NAME).worksheet("db_complaint")
                                            all_raw_comps = sh_comps.get_all_values()
                                            target_row_idx = None
                                            if len(all_raw_comps) > 1:
                                                header_cols = [h.strip().lower() for h in all_raw_comps[0]]
                                                for r_idx, sheet_row in enumerate(all_raw_comps[1:], start=2):
                                                    if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_comp_to_edit).strip():
                                                        target_row_idx = r_idx
                                                        break
                                    
                                            if target_row_idx:
                                                date_idx = header_cols.index('complaint_date') + 1 if 'complaint_date' in header_cols else 5
                                                sub_idx = header_cols.index('complaint_subject') + 1 if 'complaint_subject' in header_cols else 6
                                                txt_idx = header_cols.index('complaint_text') + 1 if 'complaint_text' in header_cols else 7
                                                staff_idx = header_cols.index('complaint_staff') + 1 if 'complaint_staff' in header_cols else 8
                                                stat_idx = header_cols.index('status') + 1 if 'status' in header_cols else 9
                                                choice_idx_col = header_cols.index('choice') + 1 if 'choice' in header_cols else 10

                                                sh_comps.update_cell(target_row_idx, date_idx, ed_date.strftime('%Y-%m-%d'))
                                                sh_comps.update_cell(target_row_idx, sub_idx, ed_subject)
                                                sh_comps.update_cell(target_row_idx, txt_idx, ed_text)
                                                sh_comps.update_cell(target_row_idx, staff_idx, ed_staff)
                                                sh_comps.update_cell(target_row_idx, stat_idx, ed_status)
                                                
                                                try:
                                                    sh_comps.update_cell(target_row_idx, choice_idx_col, ed_choice)
                                                except:
                                                    pass
                                                
                                                #load_specific_db.clear()
                                                st.success("✅ Η εγγραφή τροποποιήθηκε επιτυχώς!")
                                                time.sleep(1.2)
                                                st.session_state.is_submitting = False
                                                st.rerun()
                                            else:
                                                st.error("Δεν βρέθηκε η εγγραφή στη βάση δεδομένων.")
                                                st.session_state.is_submitting = False
                                        except Exception as edit_err:
                                            st.session_state.is_submitting = False
                                            st.error(f"Σφάλμα τροποποίησης: {edit_err}")
                    else:
                        st.info("Ο πελάτης δεν έχει καταχωρημένα παράπονα ή σχόλια.")
                else:
                    st.info("Δεν βρέθηκαν δεδομένα στον πίνακα παραπόνων.")
                
                # --- ΝΕΑ ΚΑΤΑΧΩΡΗΣΗ ΠΑΡΑΠΟΝΟΥ / ΣΧΟΛΙΟΥ ΜΕ CLEAR_ON_SUBMIT ---
                st.markdown("---")
                st.markdown("##### ➕ Νέα Καταχώρηση (Παράπονο / Σχόλιο)")

                teacher_options_new = []
                if not teachers_df.empty:
                    if 'full_name' in teachers_df.columns:
                        teacher_options_new = sorted(teachers_df['full_name'].dropna().astype(str).unique().tolist())
                    elif 'teacher_firstname' in teachers_df.columns and 'teacher_lastname' in teachers_df.columns:
                        teacher_options_new = sorted((teachers_df['teacher_lastname'].astype(str).str.strip() + " " + teachers_df['teacher_firstname'].astype(str).str.strip()).unique().tolist())
                    else:
                        teacher_options_new = sorted(teachers_df.iloc[:, 1].dropna().astype(str).unique().tolist())

                with st.form("new_complaint_form_submit", clear_on_submit=True):
                    new_choice = st.selectbox("Τύπος Καταχώρησης", options=["Παράπονο", "Σχόλιο"], key="comp_choice_new")
                    c_date = st.date_input("Ημερομηνία Καταχώρησης", value=datetime.today(), key="comp_date_new")
                    c_subject = st.text_input("Θέμα", key="comp_subj_new").strip()
                    c_text = st.text_area("Κείμενο", key="comp_text_new").strip()
                    sel_staff = st.selectbox("Υπεύθυνος / Staff", options=teacher_options_new if teacher_options_new else [st.session_state.get('username', 'Διαχειριστής')], key="comp_staff_new")
                    
                    submit_new_comp = st.form_submit_button("💾 Αποθήκευση Καταχώρησης", disabled=st.session_state.is_submitting)

                    if submit_new_comp:
                        st.session_state.is_submitting = True
                        if not c_subject or not c_text:
                            st.warning("⚠️ Συμπλήρωσε θέμα και κείμενο!")
                            st.session_state.is_submitting = False
                        else:
                            try:
                                gc = get_gspread_client()
                                sh = gc.open(OLD_SHEET_NAME).worksheet("db_complaint")
                                
                                records = sh.get_all_records()
                                if records:
                                    df_comps = pd.DataFrame(records)
                                    df_comps.columns = df_comps.columns.astype(str).str.strip()
                                    numeric_ids = pd.to_numeric(df_comps['id_complaint'].astype(str).str.replace('COMP-', ''), errors='coerce')
                                    last_id = numeric_ids.max()
                                    next_id = int(last_id) + 1 if pd.notna(last_id) else len(df_comps) + 1
                                else:
                                    next_id = 1
                                
                                new_id = f"COMP-{next_id}"

                                sh.append_row([
                                    new_id, 
                                    c_id, 
                                    str(c_row['customer_firstname']), 
                                    str(c_row['customer_lastname']), 
                                    c_date.strftime('%Y-%m-%d'), 
                                    c_subject, 
                                    c_text, 
                                    sel_staff, 
                                    'ΕΝΕΡΓΟ',
                                    new_choice
                                ])
                                
                                #load_specific_db.clear()
                                st.success("✅ Η καταχώρηση αποθηκεύτηκε επιτυχώς!")
                                time.sleep(1.2)
                                st.session_state.is_submitting = False
                                st.rerun()
                                
                            except Exception as e:
                                st.session_state.is_submitting = False
                                st.error(f"Σφάλμα κατά την καταχώρηση: {e}")

        # --- ΓΕΝΙΚΗ ΛΙΣΤΑ ΜΕ ΟΛΑ ΤΑ ΕΝΕΡΓΑ ΠΑΡΑΠΟΝΑ / ΣΧΟΛΙΑ ΣΤΟ ΚΑΤΩ ΜΕΡΟΣ ---
        st.markdown("---")
        st.subheader("🔥 Λίστα Όλων των Ενεργών Παραπόνων & Σχολίων (Κέντρου)")
        
        all_comps_fresh = load_specific_db("db_complaint")
        if not all_comps_fresh.empty:
            all_comps_fresh.columns = all_comps_fresh.columns.astype(str).str.strip()
            if 'status' not in all_comps_fresh.columns:
                all_comps_fresh['status'] = 'ΕΝΕΡΓΟ'
            if 'choice' not in all_comps_fresh.columns:
                all_comps_fresh['choice'] = 'Παράπονο'
                
            global_active = all_comps_fresh[all_comps_fresh['status'].astype(str).str.upper() == 'ΕΝΕΡΓΟ'].copy()
            
            if not global_active.empty:
                global_active['Πελάτης'] = global_active['customer_lastname'].astype(str) + " " + global_active['customer_firstname'].astype(str)
                
                view_global_active = global_active[['id_complaint', 'choice', 'Πελάτης', 'complaint_date', 'complaint_subject', 'complaint_text', 'complaint_staff']].copy()
                view_global_active.columns = ['ID', 'Τύπος', 'Πελάτης', 'Ημερομηνία', 'Θέμα', 'Κείμενο', 'Υπεύθυνος']
                st.dataframe(view_global_active, use_container_width=True, hide_index=True, height=220)
            else:
                st.info("Δεν υπάρχουν ενεργά παράπονα ή σχόλια καταχωρημένα στο σύστημα.")
        else:
            st.info("Η βάση παραπόνων είναι κενή.")
#-------------------------------------------------------
    elif st.session_state.page == "unassigned_customers":
        if st.button("⬅️ Μενού", key="menu_unassigned", use_container_width=True, disabled=st.session_state.is_submitting):
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun()
    
        st.subheader("🚫 Εκδήλωση Ενδιαφέροντος")
    
        cust_df = load_specific_db("db_customers")
        reg_df = load_specific_db("db_registrations")
        sum_df = load_specific_db("db_summer")
    
        if cust_df.empty:
            st.warning("Δεν βρέθηκαν πελάτες στη βάση.")
        else:
            cust_df.columns = cust_df.columns.astype(str).str.strip()
            
            # 1. Συλλογή από db_registrations (id_customer, customer_firstname, customer_lastname)
            reg_keys = set()
            if not reg_df.empty:
                reg_df.columns = reg_df.columns.astype(str).str.strip()
                for _, r in reg_df.iterrows():
                    r_id = str(r.get('id_customer', '')).strip()
                    r_fn = str(r.get('customer_firstname', '')).strip().lower()
                    r_ln = str(r.get('customer_lastname', '')).strip().lower()
                    if r_id and r_id != 'nan':
                        reg_keys.add(r_id)
                    if r_fn and r_ln:
                        reg_keys.add((r_ln, r_fn))
    
            # 2. Συλλογή από db_summer (summer_firstname, summer_lastname)
            sum_keys = set()
            if not sum_df.empty:
                sum_df.columns = sum_df.columns.astype(str).str.strip()
                for _, r in sum_df.iterrows():
                    s_fn = str(r.get('summer_firstname', '')).strip().lower()
                    s_ln = str(r.get('summer_lastname', '')).strip().lower()
                    if s_fn and s_ln:
                        sum_keys.add((s_ln, s_fn))
    
            # 3. Έλεγχος ποιός πελάτης από το db_customers ΔΕΝ υπάρχει ούτε στο reg_keys ούτε στο sum_keys
            unassigned_rows = []
            for _, c_row in cust_df.iterrows():
                c_id = str(c_row.get('id_customer', '')).strip()
                c_fn = str(c_row.get('customer_firstname', '')).strip().lower()
                c_ln = str(c_row.get('customer_lastname', '')).strip().lower()
    
                exists_in_regs = c_id in reg_keys or (c_ln, c_fn) in reg_keys
                exists_in_sum = (c_ln, c_fn) in sum_keys
    
                if not exists_in_regs and not exists_in_sum:
                    unassigned_rows.append(c_row)
    
            if not unassigned_rows:
                st.success("🎉 Όλοι οι πελάτες έχουν εγγραφή σε τμήμα!")
            else:
                unassigned_df = pd.DataFrame(unassigned_rows)
                
                # Δημιουργία καθαρού display_name για το φίλτρο
                unassigned_df['display_name'] = (
                    unassigned_df['customer_lastname'].astype(str).str.strip().str.upper() + " " + 
                    unassigned_df['customer_firstname'].astype(str).str.strip().str.upper()
                )
                
                # --- ΦΙΛΤΡΟ ΑΝΑΖΗΤΗΣΗΣ ΑΝΑ ΠΕΛΑΤΗ (ΜΟΝΟ ΟΣΟΙ ΕΙΝΑΙ ΣΤΟ GRID) ---
                unique_unassigned_list = sorted(unassigned_df['display_name'].tolist())
                sel_unassigned_cust = st.selectbox("👤 Αναζήτηση / Επιλογή Ανένταχτου Πελάτη", ["ΟΛΟΙ"] + unique_unassigned_list, key="filter_unassigned_cust")
    
                # Εφαρμογή φίλτρου πελάτη αν έχει επιλεγεί κάποιος συγκεκριμένος
                filtered_unassigned_df = unassigned_df.copy()
                if sel_unassigned_cust != "ΟΛΟΙ":
                    filtered_unassigned_df = filtered_unassigned_df[filtered_unassigned_df['display_name'] == sel_unassigned_cust]
    
                display_cols = [c for c in ['id_customer', 'customer_lastname', 'customer_firstname', 'telephone', 'email', 'register_date'] if c in filtered_unassigned_df.columns]
                
                rename_dict = {
                    'id_customer': 'ID',
                    'customer_lastname': 'Επώνυμο',
                    'customer_firstname': 'Όνομα',
                    'telephone': 'Τηλέφωνο',
                    'email': 'Email',
                    'register_date': 'Ημ. Εγγραφής'
                }
                final_unassigned = filtered_unassigned_df[display_cols].rename(columns=rename_dict)
                
                st.info(f"Βρέθηκαν **{len(unassigned_df)}** συνολικά πελάτες που δεν έχουν ενταχθεί σε κάποιο πρόγραμμα.")
                st.dataframe(final_unassigned, use_container_width=True, hide_index=True)
    
                # --- ΕΜΦΑΝΙΣΗ ΣΤΟΙΧΕΙΩΝ & ΠΑΡΑΠΟΝΩΝ (ΟΤΑΝ ΕΠΙΛΕΓΕΤΑΙ ΣΥΓΚΕΚΡΙΜΕΝΟΣ ΠΕΛΑΤΗΣ) ---
                if sel_unassigned_cust != "ΟΛΟΙ":
                    st.markdown("---")
                    match_row_data = filtered_unassigned_df.iloc[0]
                    target_c_id = int(pd.to_numeric(match_row_data.get('id_customer', 0), errors='coerce'))
    
                    st.markdown(f"### 👤 Στοιχεία Πελάτη: **{sel_unassigned_cust}**")
                    
                    st.markdown("""
                    <style>
                    .cust-card {
                        background-color: #f8f9fa;
                        border: 1px solid #e9ecef;
                        border-radius: 10px;
                        padding: 20px;
                        margin-bottom: 20px;
                    }
                    .cust-row {
                        display: flex;
                        margin-bottom: 8px;
                    }
                    .cust-label {
                        font-weight: bold;
                        color: #495057;
                        width: 170px;
                        flex-shrink: 0;
                    }
                    .cust-val {
                        color: #212529;
                    }
                    </style>
                    """, unsafe_allow_html=True)
    
                    c_id = match_row_data.get('id_customer', '')
                    c_cat = match_row_data.get('customer_category', '')
                    c_fname = match_row_data.get('customer_firstname', '')
                    c_lname = match_row_data.get('customer_lastname', '')
                    c_pfname = match_row_data.get('parent_firstname', '')
                    c_plname = match_row_data.get('parent_lastname', '')
                    c_mfname = match_row_data.get('mother_name', '')
                    c_pass = match_row_data.get('passport', '')
                    c_tel = match_row_data.get('telephone', '')
                    c_email = match_row_data.get('email', '')
                    c_addr = match_row_data.get('address', '')
                    c_town = match_row_data.get('town', '')
                    c_zip = match_row_data.get('zipCode', '')
                    c_taxoff = match_row_data.get('Tax Office', match_row_data.get('tax_office', ''))
                    c_afm = str(match_row_data.get('customer_afm', '')).replace("'", "")
                    c_amka = match_row_data.get('amka', '')
                    c_iban = match_row_data.get('iban', '')
                    c_bank = match_row_data.get('bank', '')
                    c_bday = match_row_data.get('birthday_date', '')
                    c_regdate = match_row_data.get('register_date', '')
    
                    st.markdown(f"""
                    <div class="cust-card">
                        <div style="display: flex; flex-wrap: wrap;">
                            <div style="flex: 1; min-width: 300px; padding-right: 15px;">
                                <div class="cust-row"><span class="cust-label">ID Πελάτη:</span><span class="cust-val">{c_id}</span></div>
                                <div class="cust-row"><span class="cust-label">Κατηγορία:</span><span class="cust-val">{c_cat}</span></div>
                                <div class="cust-row"><span class="cust-label">Όνομα:</span><span class="cust-val">{c_fname}</span></div>
                                <div class="cust-row"><span class="cust-label">Επώνυμο:</span><span class="cust-val">{c_lname}</span></div>
                                <div class="cust-row"><span class="cust-label">Όνομα Πατρός:</span><span class="cust-val">{c_pfname}</span></div>
                                <div class="cust-row"><span class="cust-label">Επώνυμο Πατρός:</span><span class="cust-val">{c_plname}</span></div>
                                <div class="cust-row"><span class="cust-label">Όνομα Μητρός:</span><span class="cust-val">{c_mfname}</span></div>
                                <div class="cust-row"><span class="cust-label">ΑΔΤ / Διαβατήριο:</span><span class="cust-val">{c_pass}</span></div>
                                <div class="cust-row"><span class="cust-label">Τηλέφωνο:</span><span class="cust-val">{c_tel}</span></div>
                                <div class="cust-row"><span class="cust-label">Email:</span><span class="cust-val">{c_email}</span></div>
                            </div>
                            <div style="flex: 1; min-width: 300px;">
                                <div class="cust-row"><span class="cust-label">Διεύθυνση:</span><span class="cust-val">{c_addr}</span></div>
                                <div class="cust-row"><span class="cust-label">Πόλη:</span><span class="cust-val">{c_town}</span></div>
                                <div class="cust-row"><span class="cust-label">ΤΚ:</span><span class="cust-val">{c_zip}</span></div>
                                <div class="cust-row"><span class="cust-label">ΔΟΥ:</span><span class="cust-val">{c_taxoff}</span></div>
                                <div class="cust-row"><span class="cust-label">ΑΦΜ:</span><span class="cust-val">{c_afm}</span></div>
                                <div class="cust-row"><span class="cust-label">ΑΜΚΑ:</span><span class="cust-val">{c_amka}</span></div>
                                <div class="cust-row"><span class="cust-label">IBAN:</span><span class="cust-val">{c_iban}</span></div>
                                <div class="cust-row"><span class="cust-label">Τράπεζα:</span><span class="cust-val">{c_bank}</span></div>
                                <div class="cust-row"><span class="cust-label">Ημ. Γέννησης:</span><span class="cust-val">{c_bday}</span></div>
                                <div class="cust-row"><span class="cust-label">Ημ. Εγγραφής:</span><span class="cust-val">{c_regdate}</span></div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
    
                    st.markdown("---")
                    
                    # --- ΕΜΦΑΝΙΣΗ ΠΑΡΑΠΟΝΩΝ & ΣΧΟΛΙΩΝ ---
                    st.markdown(f"### 🛠️ Παράπονα & Σχόλια για: **{sel_unassigned_cust}**")
                    
                    complaints_df = load_specific_db("db_complaint")
                    
                    if not complaints_df.empty:
                        complaints_df.columns = complaints_df.columns.astype(str).str.strip()
                        complaints_df['id_cust_num'] = pd.to_numeric(complaints_df['id_customer'], errors='coerce')
                        cust_comps = complaints_df[complaints_df['id_cust_num'] == float(target_c_id)].copy()
                        
                        if not cust_comps.empty:
                            if 'status' not in cust_comps.columns:
                                cust_comps['status'] = 'ΕΝΕΡΓΟ'
                            if 'choice' not in cust_comps.columns:
                                cust_comps['choice'] = 'Παράπονο'
    
                            view_cust_comps = cust_comps[['id_complaint', 'choice', 'complaint_date', 'complaint_subject', 'complaint_text', 'complaint_staff', 'status']].copy()
                            view_cust_comps.columns = ['ID', 'Τύπος', 'Ημερομηνία', 'Θέμα', 'Κείμενο', 'Υπεύθυνος', 'Κατάσταση']
                            
                            st.dataframe(view_cust_comps, use_container_width=True, hide_index=True, height=180)
                        else:
                            st.info("Δεν βρέθηκαν καταχωρημένα παράπονα ή σχόλια για αυτόν τον πελάτη.")
                    else:
                        st.info("Η βάση παραπόνων είναι κενή.")
#-------------------------------------------------------------------------------
    elif st.session_state.page == "teachers_secretariat":
        if st.session_state.role not in ["admin", "master"]:
            st.session_state.page = "menu"; st.rerun() 

        if st.button("⬅️ Επιστροφή στο Μενού", key="menu_teachers_back", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.page = "menu"; st.rerun() 

        st.subheader("👩‍🏫 Στοιχεία Καθηγητών - Γραμματείας")
        st.markdown("---")

        try:
            # Φόρτωση δεδομένων από το tab db_teachers
            teachers_df = load_specific_db("db_teachers")

            if not teachers_df.empty:
                teachers_df.columns = teachers_df.columns.astype(str).str.strip()

                # Επιθυμητές στήλες
                desired_cols = [
                    'teacher_id', 
                    'teacher_firstname', 
                    'teacher_lastname', 
                    'teacher_email', 
                    'teacher_telephone', 
                    'teacher_bank', 
                    'teacher_iban'
                ]

                # Κράτησε μόνο όσες από τις επιθυμητές στήλες υπάρχουν πραγματικά στο DataFrame
                available_cols = [col for col in desired_cols if col in teachers_df.columns]

                if available_cols:
                    display_teachers = teachers_df[available_cols].copy()

                    # Μετονομασία στηλών για πιο επαγγελματική εμφάνιση στα ελληνικά
                    rename_map = {
                        'teacher_id': 'ID',
                        'teacher_firstname': 'Όνομα',
                        'teacher_lastname': 'Επώνυμο',
                        'teacher_email': 'Email',
                        'teacher_telephone': 'Τηλέφωνο',
                        'teacher_bank': 'Τράπεζα',
                        'teacher_iban': 'IBAN'
                    }
                    display_teachers = display_teachers.rename(columns=rename_map)

                    # Εμφάνιση σε grid (dataframe)
                    st.dataframe(display_teachers, use_container_width=True, hide_index=True)
                else:
                    st.warning("⚠️ Δεν βρέθηκαν οι αναμενόμενες στήλες στο tab `db_teachers`.")
            else:
                st.info("Το tab `db_teachers` είναι προσωρινά κενό.")

        except Exception as t_err:
            st.error(f"Σφάλμα φόρτωσης στοιχείων καθηγητών: {t_err}")

#------------------------
    elif st.session_state.page == "edit_master_manual":
        if st.session_state.role not in ["master"]:
            st.session_state.page = "menu"; st.rerun()
    
        if st.button("⬅️ Μενού", key="menu_master_manual_back", use_container_width=True, disabled=st.session_state.is_submitting): 
            st.session_state.is_submitting = True
            st.session_state.page = "menu"
            st.session_state.is_submitting = False
            st.rerun() 
    
        st.subheader("🛠️ Διαχείριση & Επεξεργασία Χειροκίνητων Κινήσεων Master")
    
        master_df = load_master_db("db_finance_master")
        if not master_df.empty:
            master_df.columns = master_df.columns.astype(str).str.strip()
            
            # Φιλτράρουμε μόνο τις χειροκίνητες (Manual) κινήσεις
            if 'source' in master_df.columns:
                manual_df = master_df[master_df['source'].astype(str).str.strip() == 'Manual'].copy()
            else:
                manual_df = pd.DataFrame()
    
            # --- GRID & ΤΡΟΠΟΠΟΙΗΣΗ ΥΠΑΡΧΟΥΣΩΝ ΚΙΝΗΣΕΩΝ ---
            st.markdown("---")
            st.markdown("##### 📋 Υπάρχουσες Χειροκίνητες Κινήσεις Master")
    
            if not manual_df.empty:
                display_cols = [c for c in ['id', 'date_entry', 'type', 'category', 'partner', 'invoice_number', 'amount', 'payment_method', 'school_year'] if c in manual_df.columns]
                display_m_table = manual_df[display_cols].copy()
                display_m_table.columns = ['ID', 'Ημερομηνία', 'Τύπος', 'Κατηγορία', 'Πελάτης/Προμηθευτής', 'Αρ. Παραστατικού', 'Ποσό (€)', 'Τρόπος', 'Έτος'][:len(display_cols)]
                st.dataframe(display_m_table, use_container_width=True, hide_index=True)
    
                man_ids_list = manual_df['id'].astype(str).tolist()
                sel_man_to_edit = st.selectbox("Επιλογή ID Κίνησης για Επεξεργασία / Διαγραφή", ["-- Επιλέξτε ID --"] + man_ids_list, key="sel_master_man_edit_box")
    
                if sel_man_to_edit != "-- Επιλέξτε ID --":
                    matched_m_row = manual_df[manual_df['id'].astype(str) == str(sel_man_to_edit)].iloc[0]
                    
                    try:
                        def_m_dt = datetime.strptime(str(matched_m_row.get('date_entry', datetime.today().strftime('%Y-%m-%d'))), '%Y-%m-%d').date()
                    except:
                        def_m_dt = datetime.today().date()
    
                    with st.form("edit_master_manual_form"):
                        st.markdown(f"**Επεξεργασία ή Διαγραφή Κίνησης Master με ID: {sel_man_to_edit}**")
                        
                        ed_m_type = st.selectbox("Τύπος Κίνησης", ["ΕΣΟΔΟ", "ΕΞΟΔΟ"], index=0 if str(matched_m_row.get('type', 'ΕΣΟΔΟ')).upper() == "ΕΣΟΔΟ" else 1)
                        ed_m_cat = st.text_input("Κατηγορία / Πρόγραμμα / Αιτιολογία", value=str(matched_m_row.get('category', ''))).strip()
                        ed_m_partner = st.text_input("Πελάτης / Προμηθευτής", value=str(matched_m_row.get('partner', ''))).strip()
                        ed_m_inv = st.text_input("Αρ. Τιμολογίου / Απόδειξης", value=str(matched_m_row.get('invoice_number', ''))).strip()
                        
                        try:
                            def_amt_val = float(str(matched_m_row.get('amount', 0)).replace(',', '.'))
                        except:
                            def_amt_val = 0.0
                            
                        ed_m_amt = st.number_input("Ποσό (€)", min_value=0.0, value=def_amt_val, format="%.2f")
                        
                        methods_opts = ["Μετρητά", "POS", "Επι πιστώση", "Εθνική", "Eurobank"]
                        curr_meth = str(matched_m_row.get('payment_method', 'Μετρητά'))
                        m_idx = methods_opts.index(curr_meth) if curr_meth in methods_opts else 0
                        ed_m_method = st.selectbox("Μέθοδος Πληρωμής", methods_opts, index=m_idx)
                        
                        years_opts = [f"{y}-{y+1}" for y in range(2026, 2041)]
                        curr_yr = str(matched_m_row.get('school_year', '2026-2027'))
                        y_idx = years_opts.index(curr_yr) if curr_yr in years_opts else 0
                        ed_m_year = st.selectbox("Σχολικό Έτος", years_opts, index=y_idx)
                        
                        ed_m_date = st.date_input("Ημερομηνία", value=def_m_dt)
                        
                        col_btn1, col_btn2 = st.columns(2)
                        submit_m_edit = col_btn1.form_submit_button("💾 Οριστικοποίηση Τροποποίησης")
                        submit_m_del = col_btn2.form_submit_button("🗑️ Διαγραφή Κίνησης")
    
                        if submit_m_edit:
                            if not ed_m_cat or ed_m_amt <= 0:
                                st.warning("⚠️ Συμπληρώστε τουλάχιστον την κατηγορία και ποσό μεγαλύτερο από 0!")
                            else:
                                try:
                                    gc = get_gspread_client()
                                    sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                    all_master_rows = sh_master.get_all_values()
                                    
                                    target_row_idx = None
                                    if len(all_master_rows) > 1:
                                        for r_idx, sheet_row in enumerate(all_master_rows[1:], start=2):
                                            if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_man_to_edit).strip():
                                                target_row_idx = r_idx
                                                break
                                    
                                    if target_row_idx:
                                        sh_master.update_cell(target_row_idx, 2, ed_m_date.strftime('%Y-%m-%d'))
                                        sh_master.update_cell(target_row_idx, 3, ed_m_type)
                                        sh_master.update_cell(target_row_idx, 4, ed_m_cat)
                                        sh_master.update_cell(target_row_idx, 5, ed_m_partner if ed_m_partner else "---")
                                        sh_master.update_cell(target_row_idx, 6, ed_m_inv if ed_m_inv else "---")
                                        sh_master.update_cell(target_row_idx, 7, float(ed_m_amt))
                                        sh_master.update_cell(target_row_idx, 8, ed_m_method)
                                        sh_master.update_cell(target_row_idx, 9, ed_m_year)
                                        
                                        st.success("✅ Η κίνηση Master τροποποιήθηκε επιτυχώς!")
                                        time.sleep(1.2)
                                        st.rerun()
                                    else:
                                        st.error("Δεν βρέθηκε η εγγραφή στο Master sheet.")
                                except Exception as ed_err:
                                    st.error(f"Σφάλμα τροποποίησης: {ed_err}")
    
                        if submit_m_del:
                            try:
                                gc = get_gspread_client()
                                sh_master = gc.open_by_key(NEW_SHEET_ID).worksheet("db_finance_master")
                                all_master_rows = sh_master.get_all_values()
                                
                                target_row_idx = None
                                if len(all_master_rows) > 1:
                                    for r_idx, sheet_row in enumerate(all_master_rows[1:], start=2):
                                        if len(sheet_row) > 0 and str(sheet_row[0]).strip() == str(sel_man_to_edit).strip():
                                            target_row_idx = r_idx
                                            break
                                            
                                if target_row_idx:
                                    sh_master.delete_rows(target_row_idx)
                                    st.success("🗑️ Η κίνηση διαγράφηκε επιτυχώς από το Master!")
                                    time.sleep(1.2)
                                    st.rerun()
                                else:
                                    st.error("Δεν βρέθηκε η κίνηση για διαγραφή.")
                            except Exception as del_err:
                                st.error(f"Σφάλμα διαγραφής: {del_err}")
            else:
                st.info("Δεν υπάρχουν καταχωρημένες χειροκίνητες κινήσεις στο Master.")
    
            # --- ΝΕΑ ΚΑΤΑΧΩΡΗΣΗ ΧΕΙΡΟΚΙΝΗΤΗΣ ΚΙΝΗΣΗΣ ---
            st.markdown("---")
            st.markdown("##### ➕ Νέα Χειροκίνητη Καταχώρηση Κίνησης στο Master")
    
            with st.form("new_master_manual_clone_form", clear_on_submit=True):
                m_type = st.selectbox("Τύπος Κίνησης", ["ΕΣΟΔΟ", "ΕΞΟΔΟ"])
                m_cat = st.text_input("Κατηγορία / Πρόγραμμα / Αιτιολογία").strip()
                m_partner = st.text_input("Πελάτης / Προμηθευτής").strip()
                m_inv = st.text_input("Αρ. Τιμολογίου / Απόδειξης").strip()
                m_amt = st.number_input("Ποσό (€)", min_value=0.0, format="%.2f")
                m_method = st.selectbox("Μέθοδος Πληρωμής", ["Μετρητά", "POS", "Επι πιστώση", "Εθνική", "Eurobank"])
                m_year = st.selectbox("Σχολικό Έτος", [f"{y}-{y+1}" for y in range(2026, 2041)])
                m_date = st.date_input("Ημερομηνία", value=datetime.today())
    
                if st.form_submit_button("💾 Αποθήκευση Χειροκίνητης Κίνησης", disabled=st.session_state.is_submitting):
                    if not m_cat or m_amt <= 0:
                        st.warning("⚠️ Συμπληρώστε τουλάχιστον την κατηγορία και ποσό μεγαλύτερο από 0!")
                    else:
                        try:
                            log_to_finance_master(
                                trans_type=m_type,
                                category=m_cat,
                                partner=m_partner if m_partner else "---",
                                invoice=m_inv if m_inv else "---",
                                amount=m_amt,
                                method=m_method,
                                year=m_year,
                                source="Manual"
                            )
                            st.success("✅ Η χειροκίνητη κίνηση καταχωρήθηκε επιτυχώς στο Master!")
                            time.sleep(1.2)
                            st.rerun()
                        except Exception as manual_err:
                            st.error(f"Σφάλμα αποθήκευσης: {manual_err}")
        else:
            st.warning("Το Master sheet είναι κενό.")

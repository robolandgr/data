import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

st.title("🧪 Κλώνος Βάσης Roboland - Google Sheets")

# Ορισμός δικαιωμάτων για τα Google Sheets
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Φόρτωση των στοιχείων σύνδεσης απευθείας από τα Streamlit Secrets
credentials_dict = dict(st.secrets["connections"]["gsheets"])
# Αφαίρεση του πεδίου spreadsheet αν υπάρχει στο λεξικό των connections για να μην μπερδέψει το Credentials
credentials_dict.pop("spreadsheet", None)

credentials = Credentials.from_service_account_info(
    credentials_dict,
    scopes=scopes
)

# Σύνδεση με το gspread
gc = gspread.authorize(credentials)

# Άνοιγμα του Google Sheet με βάση το URL που έχεις στα secrets
sheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
sh = gc.open_by_url(sheet_url)

# Επιλογή του πρώτου φύλλου (worksheet)
worksheet = sh.get_worksheet(0)

# Μετατροπή των δεδομένων σε Pandas DataFrame
data = worksheet.get_all_records()
df = pd.DataFrame(data)

st.success("Η σύνδεση και η ανάγνωση πέτυχαν!")
st.dataframe(df)

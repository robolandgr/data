import streamlit as st
from streamlit_gsheets import GSheetsConnection

st.title("🧪 Test Σύνδεσης με Google Sheets")

# Δημιουργία σύνδεσης με τη βάση
conn = st.connection("gsheets", type=GSheetsConnection)

# Ανάγνωση δεδομένων (διαβάζει ολόκληρο το φύλλο)
df = conn.read()

# Εμφάνιση μηνύματος επιτυχίας και των δεδομένων
st.success("Η σύνδεση πέτυχε!")
st.subheader("Δεδομένα από το Google Sheet:")
st.dataframe(df)

import streamlit as st
import firebase_admin
from firebase_admin import credentials, firestore, storage
import fitz  # PyMuPDF
import pandas as pd
import io
import re
from datetime import datetime

# --- 1. FIREBASE SETUP ---
if not firebase_admin._apps:
    cred = credentials.Certificate("serviceAccountKey.json")
    firebase_admin.initialize_app(cred, {
        'storageBucket': 'gso-database.firebasestorage.app' # Replace with your bucket link
    })

db = firestore.client()
bucket = storage.bucket()

# --- 2. HELPER FUNCTIONS ---
def format_date_to_string(date_str):
    months = {'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
              'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'}
    try:
        parts = date_str.split()
        return f"{parts[0].zfill(2)}{months.get(parts[1].upper(), '00')}{parts[2][-2:]}"
    except: return "000000"

def is_expired(expiry_ddmmyy):
    try:
        exp_date = datetime.strptime(expiry_ddmmyy, "%d%m%y")
        return exp_date < datetime.today()
    except: return True

def add_signature_to_pdf(page):
    text = "MADE BY ABDULLAH ALHAKIM"
    page_rect = page.rect
    point = fitz.Point(page_rect.width - 200, page_rect.height - 20)
    page.insert_text(point, text, fontsize=10, color=(0.4, 0.4, 0.4))

def create_template(temp_type):
    output = io.BytesIO()
    df = pd.DataFrame(columns=["Ref Number", "Country"]) if temp_type == "MICHELIN" else pd.DataFrame(columns=["Brand", "Size", "Pattern"])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. UI DESIGN (MATT PURPLE) ---
st.set_page_config(page_title="GSO Expert Pro", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #F3F0F7; }
    [data-testid="stSidebar"] { background-color: #4B3F72 !important; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    h1, h2, h3 { color: #2E2841; font-family: 'Segoe UI', sans-serif; }
    div[data-testid="stMetric"] { background: #FFFFFF; border: 1px solid #D1C4E9; padding: 15px; border-radius: 12px; }
    .stButton>button { background: #7A61BA; color: white; border-radius: 8px; font-weight: bold; border: none; height: 3em; }
    .footer { position: fixed; left: 0; bottom: 0; width: 100%; background-color: #4B3F72; color: #FFFFFF; text-align: center; padding: 8px; z-index: 100; font-weight: bold; }
    </style>
    <div class="footer">MADE BY ABDULLAH ALHAKIM</div>
    """, unsafe_allow_html=True)

with st.sidebar:
    st.title("GSO Finder")
    menu = st.radio("WORKFLOW", ["Dashboard", "Add Certificates", "Search & Merge"])

# --- PAGE: DASHBOARD ---
if menu == "Dashboard":
    st.title("📊 Control Center")
    today_display = datetime.now().strftime("%d %B %Y")
    c1, c2 = st.columns(2)
    with c1: st.metric("System Date", today_display)
    with c2: st.metric("Database", "Online")
    st.markdown("### 📥 Templates")
    st.download_button("Download Michelin Template", create_template("MICHELIN"), "Michelin_Template.xlsx")
    st.download_button("Download Others Template", create_template("OTHERS"), "Others_Template.xlsx")

# --- PAGE: ADD NEW ---
elif menu == "Add Certificates":
    st.title("📥 Batch Upload")
    uploaded_pdfs = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)
    if st.button("Sync to Cloud"):
        for uploaded_file in uploaded_pdfs:
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            for page_num in range(0, len(doc), 2):
                text = doc[page_num].get_text()
                if "GSO Conformity Certificate" in text:
                    try:
                        brand = re.search(r"Brand:\s*(.*)", text).group(1).strip().upper()
                        expiry_raw = re.search(r"Date of Expiry:\s*(\d{1,2}\s*[A-Z]{3}\s*\d{4})", text).group(1).strip()
                        exp = format_date_to_string(expiry_raw)
                        if is_expired(exp): continue
                        
                        ref = re.search(r"Manufacturer Ref No:\s*(.*)", text).group(1).strip().zfill(6)
                        size = re.search(r"Type:\s*(.*)", text).group(1).strip()
                        pattern = re.search(r"Pattern:\s*(.*)", text).group(1).strip().upper()
                        country = re.search(r"Country of Production:\s*(.*)", text).group(1).strip().upper()

                        clean_size = size.replace('/', '-')
                        # We use a unique ID for the document to make searching 100% accurate
                        doc_id = f"{brand}_{ref}_{country}_{exp}" if brand in ["MICHELIN", "BFGOODRICH"] else f"{brand}_{clean_size}_{pattern}_{exp}"

                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        blob = bucket.blob(f"certificates/{doc_id}.pdf")
                        blob.upload_from_string(new_doc.tobytes(), content_type='application/pdf')
                        blob.make_public()
                        
                        db.collection("gso_database").document(doc_id).set({
                            "brand": brand, "expiry": exp, "url": blob.public_url,
                            "ref_no": ref, "country": country, "size": clean_size, "pattern": pattern
                        })
                        st.success(f"Uploaded: {doc_id}")
                    except: continue
        st.success("Sync Complete!")

# --- PAGE: SEARCH ---
elif menu == "Search & Merge":
    st.title("🔍 Report Generation")
    mode = st.radio("Category", ["MICHELIN / BFG", "OTHER BRANDS"], horizontal=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if excel_file and st.button("Generate Report"):
        df = pd.read_excel(excel_file).astype(str).apply(lambda x: x.str.replace(r'\.0$', '', regex=True))
        combined_pdf = fitz.open()
        missing = []
        
        # Optimized Search: We fetch only relevant documents
        progress_bar = st.progress(0)
        total = len(df)

        for index, row in df.iterrows():
            found_doc = None
            if mode == "MICHELIN / BFG":
                ref = row.iloc[0].strip().zfill(6)
                country = row.iloc[1].strip().upper()
                # Use a specific query on documented fields
                results = db.collection("gso_database").where("ref_no", "==", ref).where("country", "==", country).limit(1).get()
            else:
                brand = row.iloc[0].strip().upper()
                size = row.iloc[1].strip().replace('/', '-')
                pattern = row.iloc[2].strip().upper()
                results = db.collection("gso_database").where("brand", "==", brand).where("size", "==", size).where("pattern", "==", pattern).limit(1).get()

            if results:
                data = results[0].to_dict()
                if is_expired(data['expiry']):
                    missing.append(f"Row {index+2}: Expired")
                else:
                    pdf_bytes = bucket.blob(f"certificates/{results[0].id}.pdf").download_as_bytes()
                    match_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                    for page in match_doc: add_signature_to_pdf(page)
                    combined_pdf.insert_pdf(match_doc)
            else:
                missing.append(f"Row {index+2}: Not Found")
            
            progress_bar.progress((index + 1) / total)

        if len(combined_pdf) > 0:
            out = io.BytesIO()
            combined_pdf.save(out)
            st.success(f"Generated {len(combined_pdf)} Page PDF")
            st.download_button("📥 DOWNLOAD REPORT", out.getvalue(), "GSO_Final_Report.pdf", "application/pdf")
        if missing:
            with st.expander("Errors/Missing"):
                for m in missing: st.error(m)

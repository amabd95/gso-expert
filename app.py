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
    try:
        # Fix private key formatting issue
        creds_dict = dict(st.secrets["firebase_credentials"])
        
        # Ensure private_key has proper newlines
        if "private_key" in creds_dict:
            # Replace literal \n with actual newlines
            creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')
        
        cred = credentials.Certificate(creds_dict)
        firebase_admin.initialize_app(cred, {
            'storageBucket': 'gso-database.firebasestorage.app'
        })
    except Exception as e:
        st.error(f"Database Connection Failed: {e}")
        st.stop()

db = firestore.client()
bucket = storage.bucket()

# --- 2. HELPER FUNCTIONS ---
def format_date_to_string(date_str):
    # Converts dates like '12 JAN 2025' to '120125'
    months = {'JAN': '01', 'FEB': '02', 'MAR': '03', 'APR': '04', 'MAY': '05', 'JUN': '06',
              'JUL': '07', 'AUG': '08', 'SEP': '09', 'OCT': '10', 'NOV': '11', 'DEC': '12'}
    try:
        parts = date_str.split()
        return f"{parts[0].zfill(2)}{months.get(parts[1].upper(), '00')}{parts[2][-2:]}"
    except: return "000000"

def is_expired(expiry_ddmmyy):
    try:
        exp_date = datetime.strptime(expiry_ddmmyy, "%d%m%y")
        return exp_date.date() < datetime.today().date()
    except: return True

def add_signature_to_pdf(page):
    text = "MADE BY ABDULLAH ALHAKIM"
    page_rect = page.rect
    point = fitz.Point(page_rect.width - 200, page_rect.height - 20)
    page.insert_text(point, text, fontsize=10, color=(0.4, 0.4, 0.4))

def create_template(temp_type):
    output = io.BytesIO()
    if temp_type == "MICHELIN":
        df = pd.DataFrame(columns=["Ref Number", "Country"])
    else:
        df = pd.DataFrame(columns=["Brand", "Size", "Pattern"])
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# --- 3. FIXED LOADER WITH BATCH FETCHING ---
@st.cache_data(ttl=600)
def load_database_index():
    """
    Load database with proper pagination to avoid timeouts
    """
    data = []
    
    try:
        # Method 1: Use get() instead of stream() - it's more reliable
        docs_ref = db.collection("gso_database")
        
        # Fetch in batches to avoid timeout
        batch_size = 500
        last_doc = None
        
        while True:
            query = docs_ref.limit(batch_size)
            
            if last_doc:
                query = query.start_after(last_doc)
            
            # Use get() which is more reliable than stream()
            docs = query.get()
            
            if not docs:
                break
            
            for doc in docs:
                d = doc.to_dict()
                d['id'] = doc.id
                data.append(d)
            
            if len(docs) < batch_size:
                break
            
            last_doc = docs[-1]
        
    except Exception as e:
        st.error(f"Error loading database: {e}")
        return pd.DataFrame(columns=["brand", "size", "pattern", "ref_no", "country", "expiry", "id"])
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Handle empty database case
    if df.empty:
        return pd.DataFrame(columns=["brand", "size", "pattern", "ref_no", "country", "expiry", "id"])

    # Ensure all columns exist and are clean strings
    cols = ["brand", "size", "pattern", "ref_no", "country", "expiry"]
    for col in cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip().str.upper()

    # Normalization: Force all sizes to use hyphens internally
    df['size'] = df['size'].str.replace('/', '-')
    
    return df

# --- 4. UI DESIGN ---
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
    
    if st.button("🔄 Refresh Database"):
        load_database_index.clear()
        st.success("Database cache refreshed!")

    c1, c2 = st.columns(2)
    with c1: st.metric("System Date", today_display)
    with c2: st.metric("Database", "Online")
    
    st.markdown("### 📥 Templates")
    tc1, tc2 = st.columns(2)
    with tc1: st.download_button("Download Michelin Template", create_template("MICHELIN"), "Michelin_Template.xlsx")
    with tc2: st.download_button("Download Others Template", create_template("OTHERS"), "Others_Template.xlsx")

# --- PAGE: ADD NEW ---
elif menu == "Add Certificates":
    st.title("📥 Batch Upload")
    uploaded_pdfs = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True)
    
    if st.button("Sync to Cloud"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for i, uploaded_file in enumerate(uploaded_pdfs):
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
                        
                        if brand in ["MICHELIN", "BFGOODRICH"]:
                            doc_id = f"{brand}_{ref}_{country}_{exp}"
                        else:
                            doc_id = f"{brand}_{clean_size}_{pattern}_{exp}"

                        new_doc = fitz.open()
                        new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                        blob = bucket.blob(f"certificates/{doc_id}.pdf")
                        blob.upload_from_string(new_doc.tobytes(), content_type='application/pdf')
                        blob.make_public()
                        
                        db.collection("gso_database").document(doc_id).set({
                            "brand": brand, "expiry": exp, "url": blob.public_url,
                            "ref_no": ref, "country": country, "size": clean_size, "pattern": pattern
                        })
                    except: continue
            
            progress_bar.progress((i + 1) / len(uploaded_pdfs))
            status_text.text(f"Processed file {i+1} of {len(uploaded_pdfs)}")
        
        load_database_index.clear()
        st.success("Sync Complete!")

# --- PAGE: SEARCH ---
elif menu == "Search & Merge":
    st.title("🔍 Report Generation")
    mode = st.radio("Category", ["MICHELIN / BFG", "OTHER BRANDS"], horizontal=True)
    excel_file = st.file_uploader("Upload Excel", type=["xlsx"])

    if excel_file and st.button("Generate Report"):
        # 1. LOAD INDEX: Download the "map" of the database
        with st.spinner("Loading Database Index..."):
            db_df = load_database_index()
        
        if db_df.empty:
            st.error("Database is empty or failed to load")
            st.stop()
        
        # 2. LOAD EXCEL
        df = pd.read_excel(excel_file).astype(str).apply(lambda x: x.str.replace(r'\.0$', '', regex=True))
        combined_pdf = fitz.open()
        missing = []
        progress_bar = st.progress(0)
        
        # 3. INSTANT SEARCH (0 Latency because db_df is in memory)
        for index, row in df.iterrows():
            matches = pd.DataFrame()
            
            if mode == "MICHELIN / BFG":
                t_ref = row.iloc[0].strip().zfill(6)
                t_country = row.iloc[1].strip().upper()
                if not db_df.empty:
                    matches = db_df[
                        (db_df['ref_no'] == t_ref) & 
                        (db_df['country'] == t_country)
                    ]
            else:
                t_brand = row.iloc[0].strip().upper()
                t_size = row.iloc[1].strip().replace('/', '-').upper()
                t_pattern = row.iloc[2].strip().upper()
                
                if not db_df.empty:
                    matches = db_df[
                        (db_df['brand'] == t_brand) & 
                        (db_df['pattern'] == t_pattern) &
                        (db_df['size'] == t_size)
                    ]

            # 4. DOWNLOAD PDF
            if not matches.empty:
                found_item = matches.iloc[0]
                if is_expired(found_item['expiry']):
                    missing.append(f"Row {index+2}: Certificate Expired")
                else:
                    try:
                        pdf_bytes = bucket.blob(f"certificates/{found_item['id']}.pdf").download_as_bytes()
                        match_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                        for page in match_doc: add_signature_to_pdf(page)
                        combined_pdf.insert_pdf(match_doc)
                    except Exception as e:
                        missing.append(f"Row {index+2}: Found in DB but PDF missing in Storage")
            else:
                missing.append(f"Row {index+2}: Not Found")
            
            progress_bar.progress((index + 1) / len(df))

        if len(combined_pdf) > 0:
            out = io.BytesIO()
            combined_pdf.save(out)
            st.success(f"Generated {len(combined_pdf)} Page PDF")
            st.download_button("📥 DOWNLOAD REPORT", out.getvalue(), "GSO_Final_Report.pdf", "application/pdf")
        if missing:
            with st.expander("Errors/Missing"):
                for m in missing: st.error(m)

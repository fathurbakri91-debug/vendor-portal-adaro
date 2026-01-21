import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import xlsxwriter
import os

# ==========================================
# KONFIGURASI
# ==========================================
SHEET_NAME = "DB_VENDOR_ADARO"

st.set_page_config(page_title="Internal Monitoring - ADARO", layout="wide", page_icon="👀")

# ==========================================
# FUNGSI KONEKSI (HYBRID: CLOUD & LOKAL)
# ==========================================
def connect_gsheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    try:
        # CARA 1: Cek apakah jalan di Streamlit Cloud (Pakai st.secrets)
        if "gcp_service_account" in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scope)
        
        # CARA 2: Cek apakah jalan di Laptop Lokal (Pakai secrets.json)
        elif os.path.exists("secrets.json"):
            creds = ServiceAccountCredentials.from_json_keyfile_name("secrets.json", scope)
        
        else:
            st.error("❌ File 'secrets.json' tidak ditemukan dan 'st.secrets' belum disetting!")
            return None

        client = gspread.authorize(creds)
        return client.open(SHEET_NAME).sheet1
        
    except Exception as e:
        st.error(f"Gagal Login Google Sheet: {e}")
        return None

def load_data_online():
    try:
        sheet = connect_gsheet()
        if not sheet: return pd.DataFrame()

        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        df = df.astype(str)
        
        # Helper: Bersihkan data kosong
        if 'Estimasi Kirim' not in df.columns: df['Estimasi Kirim'] = ""
        if 'Keterangan Vendor' not in df.columns: df['Keterangan Vendor'] = ""
        
        # --- FORMATTING TANGGAL (DD/MM/YYYY) ---
        if 'Document Date' in df.columns:
             df['Tahun_PO'] = pd.to_datetime(df['Document Date'], errors='coerce').dt.year.astype(str).str.replace(r'\.0', '', regex=True)
             df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
             
        if 'Delivery Date' in df.columns:
             df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("-")

        # [UPDATE] Format ETA Vendor jadi dd/mm/yyyy
        if 'Estimasi Kirim' in df.columns:
             # Ubah dulu ke datetime, lalu format ulang ke string dd/mm/yyyy
             # Jika kosong atau error, biarkan kosong ("")
             df['Estimasi Kirim'] = pd.to_datetime(df['Estimasi Kirim'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
             
        return df
    except Exception as e:
        st.error(f"Gagal koneksi ke Google Sheet: {e}")
        return pd.DataFrame()

def format_rupiah_idr(nilai):
    try:
        clean_val = str(nilai).replace('IDR','').strip()
        if ',' in clean_val and '.' in clean_val:
             clean_val = clean_val.replace('.','').replace(',','.')
        elif ',' in clean_val:
             clean_val = clean_val.replace(',','.')
        angka = float(clean_val)
        return "{:,.0f}".format(angka).replace(',', '.') + " IDR"
    except: return nilai

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Monitoring')
        workbook = writer.book
        worksheet = writer.sheets['Monitoring']
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#D7E4BC', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column(0, len(df.columns) - 1, 15)
    return output.getvalue()

# ==========================================
# UI DASHBOARD
# ==========================================
st.title("👀 Dashboard Monitoring Supply (Live Cloud)")
st.markdown("Data terhubung langsung ke Google Sheet. Tekan tombol di bawah untuk refresh.")

if st.button("🔄 Refresh Data Sekarang"):
    st.cache_data.clear()
    st.rerun()

df = load_data_online()

if not df.empty:
    st.sidebar.header("🔍 Filter Data")
    status_filter = st.sidebar.radio("Status Respon Vendor:", ["Semua Data", "✅ Sudah Direspon", "❌ Belum Direspon"], index=0)
    st.sidebar.divider()

    col_vendor = [c for c in df.columns if 'Supplier' in c or 'Vendor' in c]
    if col_vendor:
        col_vendor = col_vendor[0]
        all_vendors = sorted(df[col_vendor].unique().tolist())
        selected_vendors = st.sidebar.multiselect("Pilih Vendor (Opsional)", all_vendors)
    else:
        col_vendor = None
        selected_vendors = []
    
    if 'Tahun_PO' in df.columns:
        all_years = sorted(df['Tahun_PO'].dropna().unique().tolist())
        selected_year = st.sidebar.selectbox("Tahun PO", ['All'] + all_years)
    
    df_show = df.copy()
    
    # === FILTER LOGIC ===
    if status_filter == "✅ Sudah Direspon":
        mask = (df_show['Estimasi Kirim'].str.len() > 2) | (df_show['Keterangan Vendor'].str.len() > 2)
        df_show = df_show[mask]
    elif status_filter == "❌ Belum Direspon":
        mask = (df_show['Estimasi Kirim'].str.len() <= 2) & (df_show['Keterangan Vendor'].str.len() <= 2)
        df_show = df_show[mask]
    
    if col_vendor and selected_vendors: 
        df_show = df_show[df_show[col_vendor].isin(selected_vendors)]
    
    if 'Tahun_PO' in df.columns and selected_year != 'All': 
        df_show = df_show[df_show['Tahun_PO'] == selected_year]

    # === SCORECARD METRICS ===
    total_item = len(df_show)
    def clean_qty(x):
        try: return float(str(x).replace('.','').replace(',','.'))
        except: return 0.0
    total_qty = df_show['Still to be delivered (qty)'].apply(clean_qty).sum()
    
    def clean_money_raw(x):
        try: return float(str(x).replace('IDR','').replace('.','').replace(',','.').strip())
        except: return 0.0
    total_value_raw = 0
    if 'Net Order Value' in df_show.columns:
        total_value_raw = df_show['Net Order Value'].apply(clean_money_raw).sum()
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Item (Filtered)", f"{total_item} Baris")
    m2.metric("Total Qty Sisa", f"{total_qty:,.0f} Unit")
    m3.metric("Total Nilai PO", "{:,.0f} IDR".format(total_value_raw).replace(',', '.'))
    
    st.divider()

    # Format Harga untuk Tabel
    if 'Net Order Value' in df_show.columns:
        df_show['Net Order Value'] = df_show['Net Order Value'].apply(format_rupiah_idr)

    st.subheader(f"Detail Data")
    
    # === TABEL ===
    cols_display = [
        col_vendor, 
        'Kategori_Item',
        'Estimasi Kirim', 'Keterangan Vendor', 
        'Delivery Date', 
        'Purchasing Document', 'Item', 
        'Net Order Value', 'Material', 'Short Text', 
        'Order Quantity', 'Still to be delivered (qty)', 
        'Document Date'
    ]
    cols_final = [c for c in cols_display if c in df_show.columns]
    
    st.dataframe(
        df_show[cols_final],
        column_config={
            "Estimasi Kirim": st.column_config.TextColumn("ETA Vendor"),
            "Delivery Date": st.column_config.TextColumn("Target (Plan)"),
            "Net Order Value": st.column_config.TextColumn("Nilai PO (IDR)"),
            "Keterangan Vendor": st.column_config.TextColumn("Keterangan", width="large"),
        },
        use_container_width=True,
        hide_index=True
    )
    
    excel_data = convert_df_to_excel(df_show[cols_final])
    st.download_button(label="Download Laporan.xlsx", data=excel_data, file_name='Laporan_Monitoring.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

else: st.warning("Data Google Sheet Kosong / Belum Terhubung.")
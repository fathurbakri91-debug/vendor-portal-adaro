import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

# ==========================================
# KONFIGURASI
# ==========================================
SHEET_NAME = "DB_VENDOR_ADARO"
# File Excel Akun (Nanti diupload ke Github)
FILE_AKUN = "VENDOR_ACCOUNTS.xlsx" 

st.set_page_config(page_title="Vendor Portal - ADARO", layout="wide")

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

def load_data_cloud():
    try:
        sheet = connect_gsheet()
        if not sheet: return pd.DataFrame()
        
        data = sheet.get_all_records()
        df = pd.DataFrame(data)
        df = df.astype(str)
        
        # --- FORMATTING (Sama seperti sebelumnya) ---
        if 'Document Date' in df.columns:
            df['Tahun_PO'] = pd.to_datetime(df['Document Date'], errors='coerce').dt.year.astype(str).str.replace(r'\.0', '', regex=True)
            df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        if 'Delivery Date' in df.columns:
             df['Delivery Date'] = pd.to_datetime(df['Delivery Date'], errors='coerce').dt.strftime('%d/%m/%Y').fillna("-")

        if 'Estimasi Kirim' in df.columns:
            df['Estimasi Kirim'] = pd.to_datetime(df['Estimasi Kirim'], errors='coerce')

        if 'Net Order Value' in df.columns:
            def fmt_money(x):
                try: 
                    clean = float(str(x).replace('IDR','').replace('.','').replace(',','.').strip())
                    return "{:,.0f} IDR".format(clean).replace(',', '.')
                except: return x
            df['Net Order Value'] = df['Net Order Value'].apply(fmt_money)
            
        return df
    except Exception as e:
        st.error(f"Gagal koneksi server: {e}")
        return pd.DataFrame()

def save_data_cloud(df_edited):
    try:
        sheet = connect_gsheet()
        df_upload = df_edited.copy()
        
        if 'Estimasi Kirim' in df_upload.columns:
            df_upload['Estimasi Kirim'] = df_upload['Estimasi Kirim'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notnull(x) else "")
            
        if 'Tahun_PO' in df_upload.columns: df_upload = df_upload.drop(columns=['Tahun_PO'])
        
        sheet.clear()
        data_to_upload = [df_upload.columns.values.tolist()] + df_upload.values.tolist()
        sheet.update(data_to_upload)
        
        st.success("✅ Data tersimpan di Cloud!")
        st.cache_data.clear()
    except Exception as e:
        st.error(f"Gagal simpan: {e}")

def load_akun():
    try: return pd.read_excel(FILE_AKUN, dtype=str)
    except: return pd.DataFrame()

# ==========================================
# UI APLIKASI
# ==========================================
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.title("🔒 Login Vendor Portal")
    df_akun = load_akun()
    if not df_akun.empty:
        list_vendor = sorted(df_akun['Username'].astype(str).tolist())
        with st.form("login_form"):
            st.write("Silakan pilih Nama Perusahaan Anda:")
            user_input = st.selectbox("Nama Vendor", options=list_vendor)
            pass_input = st.text_input("Password", type="password")
            if st.form_submit_button("Masuk"):
                user_data = df_akun[df_akun['Username'] == user_input]
                if not user_data.empty and str(pass_input) == str(user_data.iloc[0]['Password']):
                    st.session_state['logged_in'] = True
                    st.session_state['user'] = user_input
                    st.rerun()
                else: st.error("Password Salah!")
    else: st.error("Database Akun Kosong/Gagal Dimuat.")
else:
    vendor_name = st.session_state['user']
    st.sidebar.title("Menu")
    st.sidebar.info(f"Login: **{vendor_name}**")
    if st.sidebar.button("Logout"):
        st.session_state['logged_in'] = False
        st.rerun()
    st.sidebar.divider()
    
    df = load_data_cloud()
    
    if not df.empty:
        col_vendor = [c for c in df.columns if 'Supplier' in c or 'Vendor' in c]
        if col_vendor:
            df_vendor = df[df[col_vendor[0]] == vendor_name].copy()
            
            years = sorted(df_vendor['Tahun_PO'].dropna().unique().tolist())
            pilih_tahun = st.sidebar.selectbox("Pilih Tahun:", ['All'] + years)
            
            if pilih_tahun != 'All':
                df_vendor = df_vendor[df_vendor['Tahun_PO'] == pilih_tahun]

            st.title("☁️ Dashboard Pengiriman")
            
            # === SCORECARD ===
            total_item = len(df_vendor)
            def clean_qty(x):
                try: return float(str(x).replace('.','').replace(',','.'))
                except: return 0.0
            total_qty_sisa = df_vendor['Still to be delivered (qty)'].apply(clean_qty).sum()
            def clean_idr(x):
                try: return float(str(x).replace(' IDR','').replace('.','').replace(',','.').strip())
                except: return 0.0
            total_nilai = df_vendor['Net Order Value'].apply(clean_idr).sum()
            total_nilai_str = "{:,.0f} IDR".format(total_nilai).replace(',', '.')

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Item", f"{total_item} Baris")
            col2.metric("Total Qty Sisa", f"{total_qty_sisa:,.0f} Unit")
            col3.metric("Total Nilai PO", total_nilai_str)
            st.divider()

            # === TABEL ===
            column_cfg = {
                "Kategori_Item": st.column_config.TextColumn("Kategori", width="small"),
                "Document Date": st.column_config.TextColumn("Tgl PO", width="small"),
                "Purchasing Document": st.column_config.TextColumn("No. PO", width="medium"),
                "Item": st.column_config.TextColumn("Item", width="small"),
                "Short Text": st.column_config.TextColumn("Deskripsi", width="large"),
                "Net Order Value": st.column_config.TextColumn("Nilai PO (IDR)", width="medium"),
                "Order Quantity": st.column_config.NumberColumn("Qty Order", format="%.2f"),
                "Still to be delivered (qty)": st.column_config.NumberColumn("Sisa Qty", format="%.2f"),
                "Delivery Date": st.column_config.TextColumn("Target (Plan)", width="medium"),
                "Estimasi Kirim": st.column_config.DateColumn(
                    "Janji Kirim (ETA)", 
                    min_value=pd.to_datetime("2020-01-01"),
                    max_value=pd.to_datetime("2030-12-31"),
                    format="DD/MM/YYYY", step=1
                ),
                "Keterangan Vendor": st.column_config.TextColumn("Keterangan", width="large")
            }
            cols_wanted = ['Kategori_Item', 'Document Date', 'Delivery Date', 'Purchasing Document', 'Item', 'Material', 'Short Text', 'Net Order Value', 'Order Quantity', 'Still to be delivered (qty)', 'Estimasi Kirim', 'Keterangan Vendor']
            cols_final = [c for c in cols_wanted if c in df_vendor.columns]
            kolom_edit = ['Estimasi Kirim', 'Keterangan Vendor']
            kolom_kunci = [c for c in cols_final if c not in kolom_edit]

            edited_df = st.data_editor(
                df_vendor[cols_final],
                column_config=column_cfg,
                num_rows="fixed",
                key="editor",
                use_container_width=True,
                disabled=kolom_kunci
            )

            if st.button("💾 SIMPAN UPDATE KE CLOUD"):
                for i in edited_df.index:
                    if 'Estimasi Kirim' in edited_df.columns:
                        df.loc[i, 'Estimasi Kirim'] = edited_df.loc[i, 'Estimasi Kirim']
                    if 'Keterangan Vendor' in edited_df.columns:
                        df.loc[i, 'Keterangan Vendor'] = edited_df.loc[i, 'Keterangan Vendor']
                save_data_cloud(df)
        else: st.error("Kolom Vendor tidak ditemukan.")
    else: st.warning("Sedang mengambil data dari Cloud...")
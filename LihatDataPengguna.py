import streamlit as st
import pandas as pd
import os

# === JUDUL HALAMAN ===
st.set_page_config(page_title="Data Pengguna Tersimpan", page_icon="📂", layout="wide")
st.title("📂 Data Pengguna Weton")
st.markdown("---")

# === NAMA FILE EXCEL ===
FILE_EXCEL = "data_weton_jodoh.xlsx"

# === CEK DAN TAMPILKAN DATA ===
if os.path.exists(FILE_EXCEL):
    try:
        # Baca file Excel
        df = pd.read_excel(FILE_EXCEL, engine="openpyxl")
        
        # Tampilkan jumlah data
        st.success(f"✅ File ditemukan — Total {len(df)} catatan")
        
        # Tampilkan tabel
        st.dataframe(df, use_container_width=True)
        
        # Tombol unduh
        with open(FILE_EXCEL, "rb") as f:
            st.download_button(
                label="📥 Unduh File Excel",
                data=f,
                file_name="data_weton_jodoh.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except PermissionError:
        st.error("❌ File sedang terbuka di Excel! Tutup terlebih dahulu lalu coba lagi.")
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan: {str(e)}")
else:
    st.warning("⚠️ File data_weton_jodoh.xlsx belum ditemukan.")
    st.info("Pastikan file Excel berada di **folder yang sama** dengan program ini.")

st.markdown("---")
st.caption("Program penampil data — Kalkulator Weton Jawa")
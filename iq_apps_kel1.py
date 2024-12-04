import pickle
import streamlit as st
import pandas as pd
import os
from io import BytesIO
import xlwt  # Library untuk membuat file Excel dalam format .xls
from datetime import datetime  # Untuk menangani input tanggal

# *Validasi file model*
if not os.path.exists('scaler_iq.sav') or not os.path.exists('model_iq.sav'):
    st.error("❌ File model tidak ditemukan. Pastikan file 'scaler_iq.sav' dan 'model_iq.sav' tersedia.")
    st.stop()

# Membaca model scaler dan model IQ yang telah disimpan
scaler_iq = pickle.load(open('scaler_iq.sav', 'rb'))
model_iq = pickle.load(open('model_iq.sav', 'rb'))

# *Judul dan header aplikasi*

st.markdown(
    """
    <h1 style="text-align: center;">🌟 Prediksi Nilai IQ 🌟</h1>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <h3 style="text-align: center;">Masukkan data berikut untuk memulai prediksi IQ Anda</h3>
    """,
    unsafe_allow_html=True
)


# *Input data pengguna*
col1, col2 = st.columns(2)

with col1:
    nama = st.text_input("Masukkan Nama Anda", value="")
    jenis_kelamin = st.selectbox("Pilih Jenis Kelamin", ["👨 Laki-laki", "👩 Perempuan"])

with col2:
    tanggal = st.date_input("Masukkan Tanggal", value=datetime.today())
    skor_mentah = st.text_input("Masukkan Skor Mentah", value="0")

# *Simpan histori data dalam session_state*
if "histori_prediksi" not in st.session_state:
    st.session_state["histori_prediksi"] = []  # Menyimpan semua data prediksi sebagai list of dict

# *Tombol untuk memproses data*
if st.button("🔍 Hitung Nilai IQ"):
    try:
        # Konversi skor mentah ke tipe float
        skor_mentah_float = float(skor_mentah)

        # Masukkan data sebagai DataFrame
        df_input = pd.DataFrame([[skor_mentah_float]], columns=['Skor Mentah'])

        # Pastikan kolom input cocok dengan model
        df_input.columns = model_iq.feature_names_in_

        # Standarisasi skor mentah
        skor_mentah_standar = scaler_iq.transform(df_input)[0][0]

        # Hitung nilai IQ
        nilai_iq = (skor_mentah_standar * 15) + 100

        # Tentukan kategori IQ
        if nilai_iq >= 110:
            keterangan_iq = "🧠 Di Atas Rata-Rata"
        elif nilai_iq >= 92:
            keterangan_iq = "🌟 Rata-Rata"
        elif nilai_iq >= 56:
            keterangan_iq = "🔻 Di Bawah Rata-Rata"
        else:
            keterangan_iq = "❌ Defisiensi"

        # Prediksi outcome menggunakan model
        prediksi_outcome = model_iq.predict(df_input)[0]
        outcome_label = "✔ Lulus" if prediksi_outcome else "❌ Tidak Lulus"

        # Tampilkan hasil analisis
        st.success(f"💡 Nilai IQ Anda: {nilai_iq:.2f}")
        st.info(f"📋 Keterangan: {keterangan_iq} (Prediksi Outcome: {outcome_label})")

        # Simpan data prediksi ke histori
        st.session_state["histori_prediksi"].append({
            "Nama": nama,
            "Jenis Kelamin": jenis_kelamin,
            "Tanggal": tanggal.strftime("%Y-%m-%d"),
            "Skor Mentah": skor_mentah_float,
            "Nilai IQ": round(nilai_iq, 2),
            "Keterangan": keterangan_iq,
            "Outcome": outcome_label
        })

    except ValueError:
        st.error("⚠ Pastikan input Skor Mentah berupa angka yang valid.")
    except Exception as e:
        st.error(f"⚠ Terjadi kesalahan: {str(e)}")

# *Fungsi untuk menyimpan data ke file Excel (.xls) dalam memori*
def save_to_excel_xls_in_memory(data):
    output = BytesIO()
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Hasil Prediksi")

    # Menuliskan header
    for col_num, column_title in enumerate(data.columns):
        sheet.write(0, col_num, column_title)

    # Menuliskan data baris demi baris
    for row_num, row in enumerate(data.values, start=1):
        for col_num, cell_value in enumerate(row):
            sheet.write(row_num, col_num, cell_value)

    workbook.save(output)
    output.seek(0)
    return output

# *Tombol untuk mengunduh hasil prediksi dalam bentuk Excel (.xls)*
if st.session_state["histori_prediksi"]:
    try:
        # Konversi histori prediksi ke DataFrame
        df_hasil = pd.DataFrame(st.session_state["histori_prediksi"])

        # Simpan ke file Excel (.xls) dalam memori
        excel_file = save_to_excel_xls_in_memory(df_hasil)

        # Unduh file dalam format .xls
        st.download_button(
            label="💾 Download Hasil Prediksi (.xls)",
            data=excel_file,
            file_name="hasil_prediksi.xls",
            mime="application/vnd.ms-excel"
        )

    except Exception as e:
        st.error(f"⚠ Terjadi kesalahan saat mengunduh file: {str(e)}")
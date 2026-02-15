import streamlit as st
import requests
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os
import urllib.parse

st.title("Dashboard Laporan Rekap Publikasi Media Online")

API_KEY = os.getenv("APIFLASH_KEY")

if not API_KEY:
    st.error("API Key belum diisi di Settings → Secrets.")
    st.stop()

links_input = st.text_area("Masukkan link (1 link per baris)")

if st.button("Buat Laporan"):

    if not links_input.strip():
        st.warning("Masukkan link terlebih dahulu.")
        st.stop()

    links = [l.strip() for l in links_input.split("\n") if l.strip()]

    document = Document()

    # =========================
    # JUDUL DOKUMEN
    # =========================
    document.add_heading("LAPORAN REKAP PUBLIKASI MEDIA ONLINE", level=1)
    document.add_paragraph("")

    # =========================
    # BUAT TABEL
    # =========================
    table = document.add_table(rows=1, cols=6)
    table.style = "Table Grid"

    headers = ["NO", "TANGGAL", "JUDUL KEGIATAN", "LINK PUBLIKASI", "DOKUMENTASI", "KET"]

    for i, header in enumerate(headers):
        table.rows[0].cells[i].text = header

    # =========================
    # ISI DATA PER LINK
    # =========================
    for idx, link in enumerate(links, start=1):

        row = table.add_row().cells

        # Nomor
        row[0].text = str(idx)

        # Tanggal
        row[1].text = datetime.now().strftime("%d/%m/%Y")

        # =========================
        # Judul dari slug URL
        # =========================
        try:
            slug = link.rstrip("/").split("/")[-1]
            slug = slug.replace("-", " ")
            title = slug.title()
        except:
            title = "Judul tidak ditemukan"

        row[2].text = title

        # Link Publikasi
        row[3].text = link

        # =========================
        # Screenshot
        # =========================
        try:
            encoded_url = urllib.parse.quote(link, safe="")

            screenshot_url = (
                f"https://api.apiflash.com/v1/urltoimage?"
                f"access_key={API_KEY}"
                f"&url={encoded_url}"
                f"&width=1280"
                f"&height=800"
                f"&format=png"
                f"&fresh=true"
            )

            response = requests.get(screenshot_url, timeout=30)

            if response.status_code == 200 and "image" in response.headers.get("Content-Type", ""):

                image_path = f"screenshot_{idx}.png"

                with open(image_path, "wb") as f:
                    f.write(response.content)

                row[4].paragraphs[0].add_run().add_picture(
                    image_path, width=Inches(1.2)
                )

            else:
                row[4].text = "Screenshot gagal"

        except:
            row[4].text = "Error screenshot"

        # Keterangan
        row[5].text = "-"

    # =========================
    # SIMPAN FILE
    # =========================
    file_name = "Laporan_Rekap_Publikasi.docx"
    document.save(file_name)

    with open(file_name, "rb") as f:
        st.download_button(
            "Download Laporan Word",
            f,
            file_name=file_name
        )

import streamlit as st
import requests
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os
import urllib.parse

st.title("Dashboard Laporan Rekap Publikasi Media Online")

# Ambil API key dari Streamlit Secrets
SCREENSHOT_API_KEY = os.getenv("SCREENSHOT_API_KEY")

if not SCREENSHOT_API_KEY:
    st.error("API KEY tidak ditemukan. Pastikan sudah diisi di Settings → Secrets.")
    st.stop()

links_input = st.text_area("Masukkan link (1 link per baris)")

if st.button("Buat Laporan"):

    if links_input.strip() == "":
        st.warning("Masukkan link terlebih dahulu.")
    else:

        # Bersihkan link kosong
        links = [l.strip() for l in links_input.split("\n") if l.strip() != ""]

        document = Document()
        document.add_heading("LAPORAN REKAP PUBLIKASI MEDIA ONLINE", level=1)

        table = document.add_table(rows=2, cols=6)
        table.style = "Table Grid"

        headers = ["NO", "TANGGAL", "JUDUL KEGIATAN", "LINK PUBLIKASI", "DOKUMENTASI", "KET"]
        for i in range(6):
            table.rows[0].cells[i].text = headers[i]

        tanggal = datetime.now().strftime("%d/%m/%Y")
        row = table.rows[1]

        row.cells[0].text = "1"
        row.cells[1].text = tanggal
        row.cells[2].text = "Rekap Publikasi Media Online"

        # Isi kolom link
        link_text = ""
        for i, link in enumerate(links, start=1):
            link_text += f"{i}. {link}\n"
        row.cells[3].text = link_text

        # ===== SCREENSHOT SECTION =====
        for i, link in enumerate(links, start=1):

            try:
                encoded_url = urllib.parse.quote(link, safe="")

                screenshot_url = (
                    f"https://api.screenshotone.com/take?"
                    f"access_key={SCREENSHOT_API_KEY}"
                    f"&url={encoded_url}"
                    f"&viewport_width=1280"
                    f"&viewport_height=800"
                    f"&format=png"
                    f"&block_ads=true"
                    f"&block_cookie_banners=true"
                    f"&wait_until=network_idle"
                )

                image_response = requests.get(screenshot_url, timeout=30)

                # Pastikan benar-benar gambar
                if image_response.status_code == 200 and "image" in image_response.headers.get("Content-Type", ""):

                    image_path = f"screenshot_{i}.png"

                    with open(image_path, "wb") as f:
                        f.write(image_response.content)

                    row.cells[4].paragraphs[0].add_run().add_picture(image_path, width=Inches(1.2))

                else:
                    row.cells[4].paragraphs[0].add_run(f"\nScreenshot gagal untuk link {i}\n")

            except Exception as e:
                row.cells[4].paragraphs[0].add_run(f"\nError pada link {i}\n")

        row.cells[5].text = "-"

        file_name = "Laporan_Rekap_Publikasi.docx"
        document.save(file_name)

        with open(file_name, "rb") as file:
            st.download_button(
                label="Download Laporan Word",
                data=file,
                file_name=file_name
            )

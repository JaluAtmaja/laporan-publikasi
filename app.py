import streamlit as st
import requests
from docx import Document
from docx.shared import Inches
from datetime import datetime
import os

st.title("Dashboard Laporan Rekap Publikasi Media Online")

SCREENSHOT_API_KEY = os.getenv("SCREENSHOT_API_KEY")

links_input = st.text_area("Masukkan link (1 link per baris)")

if st.button("Buat Laporan"):

    if links_input.strip() == "":
        st.warning("Masukkan link terlebih dahulu.")
    else:
        links = links_input.strip().split("\n")

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

        link_text = ""
        for i, link in enumerate(links, start=1):
            link_text += f"{i}. {link}\n"
        row.cells[3].text = link_text

        # Screenshot otomatis
        for i, link in enumerate(links, start=1):
            screenshot_url = f"https://api.screenshotone.com/take?access_key={SCREENSHOT_API_KEY}&url={link}&viewport_width=1280&viewport_height=800&format=png"
            image_response = requests.get(screenshot_url)

            image_path = f"screenshot_{i}.png"
            with open(image_path, "wb") as f:
                f.write(image_response.content)

            row.cells[4].paragraphs[0].add_run().add_picture(image_path, width=Inches(1.2))

        row.cells[5].text = "-"

        file_name = "Laporan_Rekap_Publikasi.docx"
        document.save(file_name)

        with open(file_name, "rb") as file:
            st.download_button(
                label="Download Laporan Word",
                data=file,
                file_name=file_name
            )

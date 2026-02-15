import streamlit as st
from docx import Document
from datetime import datetime

st.title("Dashboard Laporan Rekap Publikasi Media Online")

links_input = st.text_area("Masukkan link (1 link per baris)")

if st.button("Buat Laporan"):

    if links_input.strip() == "":
        st.warning("Masukkan link terlebih dahulu.")
    else:
        links = links_input.strip().split("\n")

        document = Document()
        document.add_heading("LAPORAN REKAP PUBLIKASI MEDIA ONLINE", level=1)

        table = document.add_table(rows=2, cols=5)
        table.style = "Table Grid"

        headers = ["NO", "TANGGAL", "JUDUL KEGIATAN", "LINK PUBLIKASI", "KET"]
        for i in range(5):
            table.rows[0].cells[i].text = headers[i]

        tanggal = datetime.now().strftime("%d/%m/%Y")
        row = table.rows[1]

        row.cells[0].text = "1"
        row.cells[1].text = tanggal
        row.cells[2].text = "Satgas TMMD Reguler Ke-127 mempererat silaturahmi melalui aktivitas bersama warga"
        row.cells[3].text = "\n".join(links)
        row.cells[4].text = "-"

        file_name = "Laporan_Rekap_Publikasi.docx"
        document.save(file_name)

        with open(file_name, "rb") as file:
            st.download_button(
                label="Download Laporan Word",
                data=file,
                file_name=file_name
            )

import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Cm
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import os
from datetime import date
from docx2pdf import convert

st.set_page_config(layout="wide")
st.title("📸 Relatório Oleões - Município de Lisboa")

# Upload inputs
company_logo = st.file_uploader("📤 Upload company logo", type=["png", "jpg", "jpeg"])
ghg_logo = st.file_uploader("📤 Upload GHG logo", type=["png", "jpg", "jpeg"])
excel_file = st.file_uploader("📥 Upload Excel file (.xlsx)", type=["xlsx"])
photos = st.file_uploader("📸 Upload photo files", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if excel_file and photos and company_logo and ghg_logo:
    df = pd.read_excel(excel_file)
    st.write("📋 Detected columns:", df.columns.tolist())

    doc = Document()

    # Set page layout to A4 landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # -- COVER PAGE --
    doc.add_paragraph().add_run().add_break()

    table = doc.add_table(rows=1, cols=2)
    row = table.rows[0]
    cell_logo = row.cells[0]
    cell_text = row.cells[1]

    # Add logo to left cell
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
        tmp_logo.write(company_logo.read())
        cell_logo.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cell_logo.paragraphs[0].add_run()
        run.add_picture(tmp_logo.name, width=Cm(6))

    # Add title to right cell
    title_para = cell_text.paragraphs[0]
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_para.add_run("Município de Lisboa\nRelatório Final de Instalações\nAlargamento Rede de Oleões 2025")
    run.bold = True
    run.font.size = doc.styles['Normal'].font.size

    doc.add_paragraph().add_run().add_break()

    # Add GHG logo bottom-right on cover
    table_ghg = doc.add_table(rows=1, cols=2)
    row_ghg = table_ghg.rows[0]
    row_ghg.cells[0].text = "GHG savings certified by:"
    row_ghg.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ghg:
        tmp_ghg.write(ghg_logo.read())
        row_ghg.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        row_ghg.cells[1].paragraphs[0].add_run().add_picture(tmp_ghg.name, width=Cm(3.5))

    # Add page break after cover
    doc.add_page_break()

    # -- PHOTO MAPPING --
    photo_map = {os.path.splitext(photo.name)[0]: photo for photo in photos}

    # -- CONTENT PAGES --
    for index, row in df.iterrows():
        internal = str(row["ID"])  # Make sure this column exists
        doc.add_paragraph(f"📌 CÓDIGO DO OLEÃO: {internal}", style="Normal")

        if internal in photo_map:
    try:
        img_file = photo_map[internal]
        img = Image.open(img_file).convert("RGB")  # Ensure RGB
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as img_tmp:
            resized = img.resize((int(472.44), int(629.92)))  # ≈ 12cm x 16cm at 96 DPI
            resized.save(img_tmp.name, format="JPEG")
            doc.add_picture(img_tmp.name, width=Cm(12), height=Cm(16))
    except Exception as e:
        st.error(f"⚠️ Failed to add image for '{internal}': {e}")
        doc.add_paragraph("⚠️ Error loading image.")


        else:
            doc.add_paragraph("❌ Photo not found.")

        doc.add_page_break()

    # -- LAST PAGE --
    doc.add_paragraph().add_run().add_break()
    table_final = doc.add_table(rows=1, cols=2)
    final_row = table_final.rows[0]
    left_cell = final_row.cells[0]
    right_cell = final_row.cells[1]

    # Company logo centered
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo2:
        tmp_logo2.write(company_logo.read())
        left_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = left_cell.paragraphs[0].add_run()
        run.add_picture(tmp_logo2.name, width=Cm(6))

    # Right: date and GHG logo
    right_para = right_cell.paragraphs[0]
    right_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = right_para.add_run(f"Lisboa, {date.today().strftime('%d/%m/%Y')}\nGHG savings certified by:\n")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ghg2:
        tmp_ghg2.write(ghg_logo.read())
        right_para.add_run().add_picture(tmp_ghg2.name, width=Cm(3.5))

    # Finalize
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as docx_file:
        doc.save(docx_file.name)
        docx_file.seek(0)
        st.success("✅ DOCX report ready!")
        st.download_button("⬇️ Download Word Report", docx_file, file_name="relatorio_oleoes.docx")

        # Export to PDF
        try:
            pdf_path = docx_file.name.replace(".docx", ".pdf")
            convert(docx_file.name, pdf_path)
            with open(pdf_path, "rb") as pdf_file:
                st.download_button("⬇️ Download PDF Report", pdf_file, file_name="relatorio_oleoes.pdf")
        except Exception as e:
            st.warning(f"⚠️ PDF export failed: {e}")

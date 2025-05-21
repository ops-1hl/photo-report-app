import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import io
import os

st.title("📸 Photo Report Generator")

# Upload files
logo_file = st.file_uploader("Upload company logo (left side of cover)", type=["jpg", "jpeg", "png"])
cert_logo_file = st.file_uploader("Upload GHG certifier logo (bottom right of cover)", type=["jpg", "jpeg", "png"])
excel_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
images = st.file_uploader("Upload photo files (.jpg or .png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if logo_file and cert_logo_file and excel_file and images:
    try:
        df = pd.read_excel(excel_file)
        st.write("📄 Columns in Excel:", df.columns.tolist())

        if "ID" not in df.columns:
            st.error("❌ Excel must contain an 'ID' column.")
        else:
            document = Document()

            # Landscape orientation
            section = document.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)

            # === COVER PAGE ===
            table = document.add_table(rows=1, cols=2)
            table.autofit = False
            table.allow_autofit = False
            table.columns[0].width = Inches(4)
            table.columns[1].width = Inches(6)

            row_cells = table.rows[0].cells

            # Left: Main company logo
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_logo:
                img = Image.open(logo_file)
                img.save(tmp_logo.name)
                row_cells[0].paragraphs[0].add_run().add_picture(tmp_logo.name, width=Inches(3))

            # Right: Title text, vertically centered
            paragraph = row_cells[1].paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run("\n\nMunicípio de Lisboa\n")
            run.bold = True
            run.font.size = Pt(24)

            run = paragraph.add_run("Relatório Final de Instalações\n")
            run.font.size = Pt(18)

            run = paragraph.add_run("\nAlargamento Rede de Oleões 2025")
            run.font.size = Pt(16)

            # Add some space before bottom-right corner
            document.add_paragraph("\n" * 10)

            # Bottom-right corner: certification text + logo
            cert_paragraph = document.add_paragraph()
            cert_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = cert_paragraph.add_run("GHG savings certified by:  ")
            run.font.size = Pt(10)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_cert_logo:
                cert_img = Image.open(cert_logo_file)
                cert_img.save(tmp_cert_logo.name)
                run.add_picture(tmp_cert_logo.name, width=Inches(1))

            # Page break
            document.add_page_break()

            # === IMAGE PAGES ===
            image_map = {
                os.path.splitext(img.name)[0]: img
                for img in images
            }

            for index, row in df.iterrows():
                internal = str(row["ID"])
                document.add_paragraph(f"📌 Internal Number: {internal}", style='Heading2')

                if internal in image_map:
                    image_map[internal].seek(0)
                    image = Image.open(image_map[internal])
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                        image.save(tmp.name)
                        document.add_picture(tmp.name, width=Inches(6))
                else:
                    document.add_paragraph("❌ Photo not found.")

                document.add_page_break()

            # Save and provide download
            docx_buffer = io.BytesIO()
            document.save(docx_buffer)
            docx_buffer.seek(0)

            st.success("✅ Report generated successfully!")
            st.download_button("⬇️ Download Report", docx_buffer, file_name="photo_report.docx")

    except Exception as e:
        st.error(f"⚠️ An error occurred: {str(e)}")

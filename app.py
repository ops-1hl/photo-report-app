import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
import tempfile
import io
import os

st.set_page_config(layout="wide")
st.title("📸 Photo Report Generator")

# Upload assets
logo_file = st.file_uploader("Upload company logo (left side of cover)", type=["png", "jpg", "jpeg"])
cert_logo_file = st.file_uploader("Upload GHG certifier logo (bottom right of cover)", type=["png", "jpg", "jpeg"])
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

            # Set A4 landscape for the first section
            section = document.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Cm(29.7)
            section.page_height = Cm(21.0)
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)

            # === COVER PAGE ===
            table = document.add_table(rows=1, cols=2)
            table.columns[0].width = Cm(10)
            table.columns[1].width = Cm(17)
            row_cells = table.rows[0].cells

            # Left side logo
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
                img = Image.open(logo_file).convert("RGBA")
                img.save(tmp_logo.name, format="PNG")
                row_cells[0].paragraphs[0].add_run().add_picture(tmp_logo.name, width=Cm(6))

            # Right side centered text
            paragraph = row_cells[1].paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run("\nMunicípio de Lisboa\n").bold = True
            paragraph.runs[-1].font.size = Pt(24)
            paragraph.add_run("Relatório Final de Instalações\n").font.size = Pt(18)
            paragraph.add_run("\nAlargamento Rede de Oleões 2025").font.size = Pt(16)

            # Add GHG logo in bottom-right
            document.add_paragraph("\n" * 10)
            cert_paragraph = document.add_paragraph()
            cert_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = cert_paragraph.add_run("GHG savings certified by:  ")
            run.font.size = Pt(10)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_cert:
                cert_img = Image.open(cert_logo_file).convert("RGBA")
                cert_img.save(tmp_cert.name, format="PNG")
                run.add_picture(tmp_cert.name, width=Cm(2.5))

            document.add_page_break()

            # === CONTENT PAGES ===
            image_map = {os.path.splitext(img.name)[0]: img for img in images}

            for index, row in df.iterrows():
                codigo = str(row["ID"])

                # Start new section for each entry with A4 landscape
                section = document.add_section(WD_SECTION.NEW_PAGE)
                section.orientation = WD_ORIENT.LANDSCAPE
                section.page_width = Cm(29.7)
                section.page_height = Cm(21.0)
                section.top_margin = Cm(2)
                section.bottom_margin = Cm(2)
                section.left_margin = Cm(2)
                section.right_margin = Cm(2)

                # Title
                p = document.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                run = p.add_run(f"📌 CÓDIGO DO OLEÃO: {codigo}")
                run.bold = True
                run.font.size = Pt(16)

                # Add photo (resized to 12 x 16 cm)
                if codigo in image_map:
                    img = Image.open(image_map[codigo])
                    img = img.convert("RGB")  # fix RGBA issue

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_img:
                        img.save(tmp_img.name, format="JPEG")
                        document.add_picture(tmp_img.name, width=Cm(12), height=Cm(16))
                else:
                    document.add_paragraph("❌ Photo not found.")

                # Page number footer (auto-field, NOT static text)
                footer = section.footer.paragraphs[0]
                footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                footer_run = footer.add_run("Page ")
                footer_run.font.size = Pt(9)
                footer_run.add_field('PAGE')  # this will auto-number in Word

            # Export to buffer
            buffer = io.BytesIO()
            document.save(buffer)
            buffer.seek(0)

            st.success("✅ Report generated successfully!")
            st.download_button("⬇️ Download Report", buffer, file_name="photo_report.docx")

    except Exception as e:
        st.error(f"⚠️ An error occurred: {str(e)}")

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

# Uploads
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

            # Set A4 Landscape orientation
            section = document.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width = Cm(29.7)
            section.page_height = Cm(21.0)
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)

            # === COVER PAGE ===
            table = document.add_table(rows=1, cols=2)
            table.autofit = False
            table.columns[0].width = Inches(4)
            table.columns[1].width = Inches(6)
            row_cells = table.rows[0].cells

            # Left: logo (PNG preserved)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
                img = Image.open(logo_file)
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGBA")
                img.save(tmp_logo.name, format="PNG")
                row_cells[0].paragraphs[0].add_run().add_picture(tmp_logo.name, width=Inches(3))

            # Right: title text
            paragraph = row_cells[1].paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.add_run("\n\nMunicípio de Lisboa\n").bold = True
            paragraph.runs[-1].font.size = Pt(24)
            paragraph.add_run("Relatório Final de Instalações\n").font.size = Pt(18)
            paragraph.add_run("\nAlargamento Rede de Oleões 2025").font.size = Pt(16)

            # Spacer and bottom-right logo
            document.add_paragraph("\n" * 10)

            cert_paragraph = document.add_paragraph()
            cert_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = cert_paragraph.add_run("GHG savings certified by:  ")
            run.font.size = Pt(10)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_cert_logo:
                cert_img = Image.open(cert_logo_file)
                if cert_img.mode in ("RGBA", "P"):
                    cert_img = cert_img.convert("RGBA")
                cert_img.save(tmp_cert_logo.name, format="PNG")
                run.add_picture(tmp_cert_logo.name, width=Inches(1))

            document.add_page_break()

            # === REPORT PAGES ===
            image_map = {os.path.splitext(img.name)[0]: img for img in images}

            for index, row in df.iterrows():
                internal = str(row["ID"])

                # Use table layout for side-by-side or vertically structured content
                doc_paragraph = document.add_paragraph()
                doc_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                doc_paragraph.add_run(f"📌 Internal Number: {internal}").bold = True

                if internal in image_map:
                    image = Image.open(image_map[internal])
                    if image.mode in ("RGBA", "P"):
                        image = image.convert("RGB")

                    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                        image.save(tmp.name, format="JPEG")
                        document.add_picture(tmp.name, width=Inches(6))
                else:
                    document.add_paragraph("❌ Photo not found.")

                document.add_page_break()

            # Save document to buffer and provide download
            docx_buffer = io.BytesIO()
            document.save(docx_buffer)
            docx_buffer.seek(0)

            st.success("✅ Report generated successfully!")
            st.download_button("⬇️ Download Report", docx_buffer, file_name="photo_report.docx")

    except Exception as e:
        st.error(f"⚠️ An error occurred: {str(e)}")

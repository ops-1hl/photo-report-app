import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.section import WD_ORIENT
import tempfile
import io
import os

st.title("📸 Photo Location Report Generator")

# Upload Excel file
excel_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

# Upload photo files
images = st.file_uploader("Upload photo files (.jpg or .png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if excel_file and images:
    try:
        df = pd.read_excel(excel_file)

        # Show columns to help debugging
        st.write("📄 Columns found in Excel:", df.columns.tolist())

        # Check for required columns
        if "ID" not in df.columns or "LOCAL" not in df.columns:
            st.error("❌ Excel must contain 'ID' and 'LOCAL' columns.")
        else:
            # Create Word document
            document = Document()

            # Set to landscape mode
            section = document.sections[0]
            section.orientation = WD_ORIENT.LANDSCAPE
            section.page_width, section.page_height = section.page_height, section.page_width
            section.left_margin = Cm(2)
            section.right_margin = Cm(2)
            section.top_margin = Cm(2)
            section.bottom_margin = Cm(2)

            # Map image filenames (without extension) to uploaded files
            image_map = {
                os.path.splitext(img.name)[0]: img
                for img in images
            }

            for index, row in df.iterrows():
                internal = str(row["ID"])
                location = str(row["LOCAL"])

                document.add_paragraph(f"📌 Internal Number: {internal}", style='Heading2')

                if internal in image_map:
                    image_map[internal].seek(0)
                    image = Image.open(image_map[internal])
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                        image.save(tmp.name)
                        document.add_picture(tmp.name, width=Inches(6))
                else:
                    document.add_paragraph("❌ Photo not found.")

                document.add_paragraph(f"📍 Location: {location}")
                document.add_page_break()

            # Save to in-memory buffer for download
            docx_buffer = io.BytesIO()
            document.save(docx_buffer)
            docx_buffer.seek(0)

            st.success("✅ Report generated successfully!")
            st.download_button("⬇️ Download Report", docx_buffer, file_name="photo_report.docx")

    except Exception as e:
        st.error(f"⚠️ An error occurred: {str(e)}")

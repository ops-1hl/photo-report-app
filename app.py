import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches
import tempfile
import os

st.title("📸 Photo Location Report Generator")

# Upload Excel
excel_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])

# Upload multiple images
images = st.file_uploader("Upload photo files (.jpg or .png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

if excel_file and images:
    df = pd.read_excel(excel_file)

    # Create a temp Word document
    document = Document()

    # Build a mapping of image names without extension
    image_map = {}
    for img in images:
        img_id = os.path.splitext(img.name)[0]
        image_map[img_id] = img

    for index, row in df.iterrows():
    internal = str(row["ID"])           # Fixed
    location = str(row["LOCAL"])        # Fixed

    document.add_paragraph(f"📌 Internal Number: {internal}")

    if internal in image_map:
        image = Image.open(image_map[internal])
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
            image.save(tmp.name)
            document.add_picture(tmp.name, width=Inches(4))
    else:
        document.add_paragraph("❌ Photo not found.")

    document.add_paragraph(f"📍 Location: {location}")
    document.add_page_break()


    # Save to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        document.save(tmp.name)
        tmp.seek(0)
        st.success("✅ Report generated!")
        st.download_button("⬇️ Download Report", tmp, file_name="photo_report.docx")

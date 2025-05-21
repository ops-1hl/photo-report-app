import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches, Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from datetime import datetime
import tempfile
import os
import shutil
import platform

# PDF conversion imports
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

st.set_page_config(layout="centered")
st.title("📸 Relatório Fotográfico - Oleões")

# Uploads
excel_file = st.file_uploader("📄 Upload Excel (.xlsx)", type=["xlsx"])
images = st.file_uploader("🖼️ Upload Fotos (.jpg/.png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
company_logo = st.file_uploader("🏢 Logo da Empresa", type=["png", "jpg"])
ghg_logo = st.file_uploader("♻️ Logo GHG", type=["png", "jpg"])

export_format = st.radio("Formato de Exportação", options=["DOCX", "PDF"])

if excel_file and images and company_logo and ghg_logo:
    df = pd.read_excel(excel_file)
    document = Document()

    # Landscape layout
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    # --- COVER PAGE ---
    cover_table = document.add_table(rows=1, cols=2)
    row_cells = cover_table.rows[0].cells

    # Company Logo (Left)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
        tmp_logo.write(company_logo.read())
        logo_path = tmp_logo.name
    row_cells[0].paragraphs[0].add_run().add_picture(logo_path, width=Cm(6))

    # Title (Right)
    title_para = row_cells[1].paragraphs[0]
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title_run = title_para.add_run("Município de Lisboa\nRelatório Final de Instalações\n\nAlargamento Rede de Oleões 2025")
    title_run.bold = True
    title_run.font.size = Pt(20)

    # GHG Logo (Bottom right)
    footer_para = document.add_paragraph()
    footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    footer_para.add_run("GHG savings certified by:   ")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ghg:
        tmp_ghg.write(ghg_logo.read())
        ghg_path = tmp_ghg.name
    footer_para.add_run().add_picture(ghg_path, width=Cm(2))

    document.add_page_break()

    # --- BODY ---
    image_map = {os.path.splitext(img.name)[0]: img for img in images}

    for _, row in df.iterrows():
        codigo = str(row["ID"])
        document.add_paragraph().add_run(f"📌 CÓDIGO DO OLEÃO: {codigo}").bold = True

        if codigo in image_map:
            try:
                image = Image.open(image_map[codigo]).convert("RGB")
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_img:
                    image = image.resize((int(Cm(12).pt), int(Cm(16).pt)))
                    image.save(tmp_img.name, format="JPEG")
                    document.add_picture(tmp_img.name, width=Cm(12), height=Cm(16))
            except Exception as e:
                document.add_paragraph(f"❌ Erro com a imagem: {e}")
        else:
            document.add_paragraph("❌ Fotografia não encontrada.")

        document.add_page_break()

    # --- FINAL PAGE ---
    final_section = document.add_section(WD_ORIENT.LANDSCAPE)
    final_section.page_width, final_section.page_height = section.page_width, section.page_height

    final_table = document.add_table(rows=1, cols=2)
    final_table.autofit = False
    final_table.columns[0].width = Cm(12)
    final_table.columns[1].width = Cm(21)

    # Left Logo
    left_cell = final_table.cell(0, 0)
    for _ in range(10):
        left_cell.add_paragraph()
    left_cell.paragraphs[-1].add_run().add_picture(logo_path, width=Cm(6))

    # Right Text + GHG Logo
    right_cell = final_table.cell(0, 1)
    right_para = right_cell.paragraphs[0]
    right_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    today = datetime.today().strftime("%d/%m/%Y")
    right_para.add_run(f"Avanca, {today}\n\n").bold = True

    right_cell.add_paragraph().add_run(
        "Escritórios/Centro Logístico e Operacional\n"
        "Rua Padre António Maria Pinho, n.º 71\n"
        "3860-130, Avanca - Estarreja | Portugal\n\n"
        "🌐  www.carbon-foote.com    |    www.hardlevel.pt\n\n"
    )

    last_ghg = right_cell.add_paragraph()
    last_ghg.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    last_ghg.add_run("GHG savings certified by:   ")
    last_ghg.add_run().add_picture(ghg_path, width=Cm(2))

    # --- SAVE & EXPORT ---
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        docx_path = tmp_docx.name
        document.save(docx_path)

    if export_format == "DOCX":
        with open(docx_path, "rb") as f:
            st.success("✅ Relatório DOCX gerado com sucesso!")
            st.download_button("⬇️ Download DOCX", f.read(), file_name="relatorio_oleoes.docx")

    elif export_format == "PDF":
        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path = os.path.join(tmp_dir, "relatorio_oleoes.pdf")

            if DOCX2PDF_AVAILABLE and platform.system() in ["Windows", "Darwin"]:
                docx2pdf_convert(docx_path, pdf_path)
            else:
                st.warning("📌 PDF export requires `docx2pdf` on Windows/macOS or LibreOffice on Linux.")
                st.stop()

            with open(pdf_path, "rb") as f:
                st.success("✅ Relatório PDF gerado com sucesso!")
                st.download_button("⬇️ Download PDF", f.read(), file_name="relatorio_oleoes.pdf")

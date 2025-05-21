import streamlit as st
import pandas as pd
from PIL import Image
from docx import Document
from docx.shared import Inches, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import tempfile
import os
import datetime

st.set_page_config(layout="centered")
st.title("📸 Gerador de Relatório - Oleões 2025")

# File uploaders
excel_file = st.file_uploader("📄 Carregar ficheiro Excel (.xlsx)", type=["xlsx"])
images = st.file_uploader("🖼️ Carregar fotos (.jpg ou .png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

logo_file = st.file_uploader("🏢 Logotipo da empresa", type=["jpg", "jpeg", "png"])
ghg_logo_file = st.file_uploader("🌍 Logotipo de certificação GHG", type=["jpg", "jpeg", "png"])

if excel_file and images and logo_file and ghg_logo_file:
    df = pd.read_excel(excel_file)

    st.success("✔️ Ficheiros carregados com sucesso!")
    document = Document()

    # Configure landscape layout
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)

    # Build photo map
    photo_map = {os.path.splitext(photo.name)[0]: photo for photo in images}

    # Cover page
    document.add_paragraph().add_run()  # spacer
    table = document.add_table(rows=1, cols=2)
    row = table.rows[0]
    row.cells[0].width = Cm(8)
    row.cells[1].width = Cm(20)

    # Left side - logo
    logo_img = Image.open(logo_file).convert("RGB")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_logo:
        logo_img.save(tmp_logo.name, format="JPEG")
        row.cells[0].paragraphs[0].add_run().add_picture(tmp_logo.name, width=Cm(6))

    # Right side - title
    p = row.cells[1].paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("Município de Lisboa\nRelatório Final de Instalações\n\nAlargamento Rede de Oleões 2025")
    run.bold = True
    run.font.size = Pt(20)

    # Bottom right - GHG certification
    table = document.add_table(rows=1, cols=2)
    row = table.rows[0]
    row.cells[0].width = Cm(18)
    row.cells[1].width = Cm(10)
    p = row.cells[0].paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p.add_run("GHG savings certified by:")
    ghg_img = Image.open(ghg_logo_file).convert("RGB")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_ghg:
        ghg_img.save(tmp_ghg.name, format="JPEG")
        row.cells[1].paragraphs[0].add_run().add_picture(tmp_ghg.name, width=Cm(3))

    document.add_page_break()

    # Report content
    for _, row in df.iterrows():
        codigo = str(row["ID"])  # "Código do Oleão"

        # Header
        p = document.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        p.add_run(f"📌 Código do Oleão: {codigo}").bold = True

        # Image
        if codigo in photo_map:
            try:
                img = Image.open(photo_map[codigo]).convert("RGB")
                resized = img.resize((int(12 * 37.8), int(16 * 37.8)))  # 12cm x 16cm at 96 DPI
                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as img_tmp:
                    resized.save(img_tmp.name, format="JPEG")
                    document.add_picture(img_tmp.name, width=Cm(12), height=Cm(16))
            except Exception as e:
                st.error(f"Erro com a imagem '{codigo}': {e}")
                document.add_paragraph("❌ Imagem inválida.")
        else:
            document.add_paragraph("❌ Foto não encontrada.")

        document.add_page_break()

    # Final page
    final_section = document.add_section(WD_ORIENT.LANDSCAPE)
    final_section.page_width = new_width
    final_section.page_height = new_height

    # Logo centered left
    table = document.add_table(rows=1, cols=2)
    row = table.rows[0]
    row.cells[0].width = Cm(8)
    row.cells[1].width = Cm(20)
    row.cells[0].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    row.cells[0].paragraphs[0].add_run().add_picture(tmp_logo.name, width=Cm(6))

    # Right side branding
    p = row.cells[1].paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("\nRelatório Final\nRede de Oleões 2025\n")
    run.bold = True
    run.font.size = Pt(20)
    p.add_run(f"\nEmitido em: {datetime.date.today().strftime('%d/%m/%Y')}")

    # Bottom right GHG logo
    table = document.add_table(rows=1, cols=2)
    row = table.rows[0]
    row.cells[0].width = Cm(18)
    row.cells[1].width = Cm(10)
    p = row.cells[0].paragraphs[0]
    p.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    p.add_run("GHG savings certified by:")
    row.cells[1].paragraphs[0].add_run().add_picture(tmp_ghg.name, width=Cm(3))

    # Export DOCX
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
        document.save(tmp_docx.name)
        tmp_docx.seek(0)
        st.download_button("⬇️ Baixar relatório DOCX", tmp_docx, file_name="relatorio_oleoes.docx")

    # Option to export as PDF
    if st.button("💾 Exportar como PDF (apenas local)"):
        try:
            from docx2pdf import convert
            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, "report.docx")
                pdf_path = os.path.join(tmpdir, "report.pdf")
                document.save(docx_path)
                convert(docx_path, pdf_path)
                with open(pdf_path, "rb") as pdf_file:
                    st.download_button("📥 Baixar PDF", pdf_file, file_name="relatorio_oleoes.pdf")
        except Exception as e:
            st.error(f"Erro ao exportar como PDF: {e}")

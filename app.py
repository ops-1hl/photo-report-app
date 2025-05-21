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

st.set_page_config(layout="centered")
st.title("📸 Relatório Fotográfico - Oleões")

# Upload files
excel_file = st.file_uploader("📄 Upload Excel file (.xlsx)", type=["xlsx"])
images = st.file_uploader("🖼️ Upload photo files (.jpg/.png)", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
company_logo = st.file_uploader("🏢 Upload company logo (cover + last page)", type=["png", "jpg"])
ghg_logo = st.file_uploader("♻️ Upload GHG logo", type=["png", "jpg"])

if excel_file and images and company_logo and ghg_logo:
    df = pd.read_excel(excel_file)
    document = Document()

    # Set orientation to landscape
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height

    # --- COVER PAGE ---
    table = document.add_table(rows=1, cols=2)
    row_cells = table.rows[0].cells

    # Left: Company logo
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_logo:
        tmp_logo.write(company_logo.read())
        tmp_path_logo = tmp_logo.name
    row_cells[0].paragraphs[0].add_run().add_picture(tmp_path_logo, width=Cm(6))

    # Right: Title text
    title_para = row_cells[1].paragraphs[0]
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = title_para.add_run("Município de Lisboa\nRelatório Final de Instalações\n\nAlargamento Rede de Oleões 2025")
    run.bold = True
    run.font.size = Pt(20)

    # Bottom right: GHG logo + text
    cover_footer = document.add_paragraph()
    cover_footer.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    cover_footer.add_run("GHG savings certified by:   ")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_ghg:
        tmp_ghg.write(ghg_logo.read())
        tmp_path_ghg = tmp_ghg.name
    cover_footer.add_run().add_picture(tmp_path_ghg, width=Cm(2))

    document.add_page_break()

    # --- BODY PAGES ---
    image_map = {os.path.splitext(img.name)[0]: img for img in images}

    for _, row in df.iterrows():
        code = str(row["ID"])
        document.add_paragraph(f"📌 CÓDIGO DO OLEÃO: {code}", style='Normal')

        if code in image_map:
            try:
                image = Image.open(image_map[code]).convert("RGB")
                image.thumbnail((Cm(12).pt * 0.75, Cm(16).pt * 0.75))

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_img:
                    image.save(tmp_img.name, format="JPEG")
                    document.add_picture(tmp_img.name, width=Cm(12), height=Cm(16))
            except Exception as e:
                document.add_paragraph(f"❌ Erro ao processar imagem: {e}")
        else:
            document.add_paragraph("❌ Fotografia não encontrada.")

        document.add_page_break()

    # --- FINAL PAGE ---
    section = document.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = new_width, new_height

    final_table = document.add_table(rows=1, cols=2)
    final_table.autofit = False
    final_table.columns[0].width = Cm(12)
    final_table.columns[1].width = Cm(21)

    # Left: Centered logo
    logo_cell = final_table.cell(0, 0)
    for _ in range(10):
        logo_cell.add_paragraph()
    logo_cell.paragraphs[-1].add_run().add_picture(tmp_path_logo, width=Cm(6))

    # Right: Address and GHG
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

    # GHG logo on last page
    last_ghg_para = right_cell.add_paragraph()
    last_ghg_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    last_ghg_para.add_run("GHG savings certified by:   ")
    last_ghg_para.add_run().add_picture(tmp_path_ghg, width=Cm(2))

    # --- OUTPUT DOCX ---
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_doc:
        document.save(tmp_doc.name)
        tmp_doc.seek(0)
        st.success("✅ Relatório gerado com sucesso!")
        st.download_button("⬇️ Download do Relatório", tmp_doc.read(), file_name="relatorio_oleoes.docx")

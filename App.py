import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from jinja2 import Environment, BaseLoader
import pypandoc
import tempfile
from pathlib import Path
import os

st.set_page_config(page_title="Offer Letter Generator", layout="centered")
st.title("📄 Offer Letter PDF Generator (Streamlit)")

# Upload files
template_file = st.file_uploader("📝 Upload Word Template (.docx)", type=["docx"])
excel_file = st.file_uploader("📊 Upload Excel File (.xlsx)", type=["xlsx"])

# Process button
if st.button("Generate Offer Letters"):
    if not template_file or not excel_file:
        st.warning("Please upload both the template and Excel file.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_dir = Path(tmpdir) / "pdf_output"
            output_dir.mkdir(exist_ok=True)

            # Load Excel
            df = pd.read_excel(excel_file)

            # Prepare template
            jinja_env = Environment(loader=BaseLoader(), autoescape=False)

            for _, row in df.iterrows():
                name = str(row['Name']).strip()
                domain = str(row['Domain']).strip().replace("&", "&amp;")
                start_date = pd.to_datetime(row['StartDate']).strftime('%d-%m-%Y')
                end_date = pd.to_datetime(row['EndDate']).strftime('%d-%m-%Y')

                context = {
                    'name': name,
                    'domain': domain,
                    'start_date': start_date,
                    'end_date': end_date
                }

                # Save temporary docx
                docx_path = Path(tmpdir) / f"{name}.docx"
                pdf_path = output_dir / f"{name}.pdf"
                template = DocxTemplate(template_file)
                template.render(context, jinja_env)
                template.save(docx_path)

                # Convert to PDF using pypandoc (Linux/Mac friendly)
                try:
                    pypandoc.convert_file(str(docx_path), 'pdf', outputfile=str(pdf_path))
                except Exception as e:
                    st.error(f"❌ PDF conversion failed for {name}: {e}")
                    continue

            # Show downloadable PDFs
            st.success("✅ Offer letters generated successfully!")
            for pdf in output_dir.glob("*.pdf"):
                with open(pdf, "rb") as f:
                    st.download_button(
                        label=f"⬇️ Download {pdf.name}",
                        data=f,
                        file_name=pdf.name,
                        mime="application/pdf"
                    )

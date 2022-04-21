import os
import shutil
import pandas as pd
from docxtpl import DocxTemplate
import streamlit as st



st.markdown("<h2 style='text-align: center; color: #ccc; background-color:#00000s; margin-top : 10px; margin-bottom: 40px'>Invoice Generator</h2>", unsafe_allow_html=True)

st.markdown("<h5 style='color: #ccc; margin-top : 0px; margin-bottom: -50px'>Upload an Excel File</h5>", unsafe_allow_html=True)
excel_file = st.file_uploader(" ")
st.markdown("<h5 style='color: #ccc; margin-top : 0px; margin-bottom: -50px'>Upload the Document Template</h5>", unsafe_allow_html=True)
template_fle = st.file_uploader("")
generate_btn = st.button("Generate Invoices")

if generate_btn:
    st.success("Invoices Generated Successfully")
    output_dir = os.mkdir("OUTPUT")

    df = pd.read_excel(excel_file, sheet_name = "Sheet1")

    for record in df.to_dict(orient="records"):
        doc = DocxTemplate(template_fle)
        doc.render(record)
        output_path = f"OUTPUT/{record['Name']}-invoice.docx"
        doc.save(output_path)
    shutil.make_archive("invoices", 'zip', "OUTPUT")
    with open("invoices.zip", "rb") as fp:
        btn = st.download_button(
            label="Download Generated Invoices",
            data=fp,
            file_name="invoices.zip",
            mime="application/zip"
        )
        

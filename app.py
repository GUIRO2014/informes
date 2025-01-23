from pydoc import doc
import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from docx.shared import Inches

import io
import pip
pip.main(["install","pydoc"])

def create_report(template_path,data,chart_data=None):
    st.write("Iniciando la creación del informe...")
    doc= Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if f'{{{{{key}}}}}' in paragraph.text:
                st.write(f"Reemplazando {key} con {value} en el informe")
            paragraph.text = paragraph.text.replace(f'{{{{{key}}}}}', str(value))
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    st.write("Informe creado con éxito.")
    return output





def main():
    st.title("Generador de Informes desde Plantillas")
    template_file = st.file_uploader("Cargar plantilla Word",type="docx")
    data_file = st.file_uploader("Cargar datos", type=["xlsx","csv"])
    if template_file and data_file:
        st.success("Archivos cargados correctamente")
        df = pd.read_csv(data_file) if data_file.name.endswith('.csv') else pd.read_excel(data_file)
        st.subheader("Datos cargados")
        st.dataframe(df)

        row_index= st.selectbox("Seleccionar fila para el informe",options=range(len(df)))
        selected_data = df.iloc[row_index].to_dict()
        

        if st.button("Generar Informe"):
            output = create_report(template_file,selected_data)
            st.download_button("Descargar informe",output, "Informe_generado.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
main()
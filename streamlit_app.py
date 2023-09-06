import streamlit as st
import shutil
import tempfile
import os
from actions.file_actions import extract_data_from_excel, generate_word_files_streamlit


# Define the template directory
template_dir = "note_templates/template.docx"
entidades = ["Provalor S.A.", "Progresar Corporation S.A.", "Sudameris Bank S.A.E.C.A.", "FACTORY"]
monedas = ["GS", "USD"]
productos = ["Préstamos de Consumo", "Sobregiros", "Tarjetas de Crédito", "Renovación Préstamos de Consumo"]

def main():
    st.title("Generación automática de Notas de Exclusión")
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])
    # Add a selectbox to choose the entidad
    entidad = st.selectbox("Seleccionar Entidad:", entidades)
    # Agregar un widget de radio para elegir entre GS y USD
    moneda = st.selectbox("Elegir Moneda:", monedas)
    # Add a selectbox to choose the product tipe
    producto = st.selectbox("Elegir Producto:", productos)
    # Declare output_dir with a default value
    output_dir = None

    if uploaded_file:
        data_list = extract_data_from_excel(uploaded_file)
        
        if st.button("Generar Notas"):
            output_dir = tempfile.mkdtemp()  # Create a temporary directory
            generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
            st.session_state.generated_files = output_dir  # Store the output directory in session state
            st.success("Notas generadas exitosamente!") 
            file_counter = 0       
            # Call the download_notes function here
            download_notes(output_dir)


def zip_notes(output_dir):
    # Create a ZIP file containing all generated notes
    zip_filename = os.path.join(tempfile.gettempdir(), "notas_generadas.zip")
    shutil.make_archive(zip_filename.replace('.zip', ''), 'zip', output_dir)
    return zip_filename


def download_notes(output_dir):
    zip_filename = zip_notes(output_dir)    
    # Provide a download button for the ZIP file
    with open(zip_filename, "rb") as f:
        if st.download_button("Descargar Notas", f.read(), file_name="notas_generadas.zip"):
            st.write("¡Gracias por descargar las notas!")


def get_generated_files(output_dir):
    generated_files = []
    for root, dirs, files in os.walk(output_dir):
        for file in files:
            if file.endswith(".docx"):
                generated_files.append(os.path.join(root, file))
    return generated_files


if __name__ == "__main__":
    main()
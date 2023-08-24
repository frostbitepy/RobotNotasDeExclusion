import streamlit as st
import shutil
import tempfile
from actions.file_actions import extract_data_from_excel, generate_word_files, generate_word_files_streamlit

# Define the template directory
template_dir = "note_templates"

def main():
    st.title("Generación automática de Notas de Exclusión")
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])
    if uploaded_file:
        data_list = extract_data_from_excel(uploaded_file)
        
        if st.button("Generar Notas"):
            output_dir = tempfile.mkdtemp()  # Create a temporary directory
            generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file)
            #generate_word_files(data_list, template_dir, output_dir)
            st.success("Notas generadas exitosamente!")
            
            if st.button("Descargar Notas"):
                # Provide download link for the generated notes
                download_notes(output_dir)


def zip_notes(output_dir):
    # Create a ZIP file containing all generated notes
    zip_filename = tempfile.mktemp(suffix=".zip")
    shutil.make_archive(zip_filename, 'zip', output_dir)
    return zip_filename


def download_notes(zip_filename):
    # Provide a download button for the ZIP file
    with open(zip_filename, "rb") as f:
        if st.download_button("Descargar Todas las Notas", f.read(), file_name="notas_generadas.zip"):
            st.write("¡Gracias por descargar las notas!")


if __name__ == "__main__":
    main()
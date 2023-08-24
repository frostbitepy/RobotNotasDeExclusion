import streamlit as st
import shutil
import tempfile
import os
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
            st.success("Notas generadas exitosamente!")  

        # Display the generated notes
        st.write("Notas Generadas:")
        for file in os.listdir(output_dir):
            if file.endswith(".docx"):
                st.write(f"- [{file}]({os.path.join(output_dir, file)})")
            
        if st.button("Descargar Notas"):
            # Create and download the ZIP file
            zip_filename = zip_notes(output_dir)
            download_notes(zip_filename)

        

def zip_notes(output_dir):
    # Create a ZIP file containing all generated notes
    zip_filename = tempfile.mktemp(suffix=".zip")
    shutil.make_archive(zip_filename, 'zip', output_dir)
    return zip_filename


def download_notes(zip_filename):
    # Provide a download button for the ZIP file
    with open(zip_filename, "rb") as f:
        if st.button("Descargar Todas las Notas"):
            st.download_button("Descargar Todas las Notas", f.read(), file_name="notas_generadas.zip")
            st.write("¡Gracias por descargar las notas!")


if __name__ == "__main__":
    main()
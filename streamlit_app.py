import streamlit as st
import shutil
import tempfile
import os
from actions.file_actions import extract_data_from_excel, generate_word_files_streamlit, multi_generate_word_files_streamlit, group_data_by_code



# Define the template directory
template_dir = "note_templates/template.docx"
entidades = ["Creditos Paraná S.A.", "Factory S.A.", "Progresar Corporation S.A.", "Provalor S.A.", "Sudameris Bank S.A.E.C.A."]
monedas = ["GS", "USD"]
productos = ["Préstamos de Consumo", "Renovación de Préstamos", "Renovación Préstamos de Consumo", "Sobregiros", "Tarjetas de Crédito", "Casa Matriz", "la sucursal de Santa Rita", "la sucursal de Fram", "la sucursaln de Ayolas"]
tipo_notas = ["Una exclusión por nota", "Múltiples exclusiones por nota"]


def main():
    st.title("Generación automática de Notas de Exclusión")
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])
    # Add a selectbox to choose the entidad
    entidad = st.selectbox("Seleccionar Entidad:", entidades)
    # Agregar un widget de radio para elegir entre GS y USD
    moneda = st.selectbox("Elegir Moneda:", monedas)
    if moneda == "GS":
        moneda = "guaraníes"
    elif moneda == "USD":
        moneda = "dólares americanos"
    # Add a selectbox to choose the product tipe
    producto = st.selectbox("Elegir Producto:", productos)
    # Add a selectbox to choose the type of note
    tipo_nota = st.selectbox("Elegir formato de nota:", tipo_notas)
    # Declare output_dir with a default value
    output_dir = None

    if uploaded_file:
        data_list = extract_data_from_excel(uploaded_file)
        
        if st.button("Generar Notas"):
            output_dir = tempfile.mkdtemp()  # Create a temporary directory
            # generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
            if tipo_nota == "Una exclusión por nota":
                generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
            elif tipo_nota == "Múltiples exclusiones por nota":
                data_dict = group_data_by_code(data_list)
                multi_generate_word_files_streamlit(data_dict, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
    
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

def test():
    st.title("Generación automática de Notas de Exclusión")
    uploaded_file = st.file_uploader("Cargar archivo Excel", type=["xlsx"])
    # Add a selectbox to choose the entidad
    entidad = st.selectbox("Seleccionar Entidad:", entidades)
    # Agregar un widget de radio para elegir entre GS y USD
    moneda = st.selectbox("Elegir Moneda:", monedas)
    if moneda == "GS":
        moneda = "guaraníes"
    elif moneda == "USD":
        moneda = "dólares americanos"
    # Add a selectbox to choose the product tipe
    producto = st.selectbox("Elegir Producto:", productos)
    # Add a selectbox to choose the type of note
    tipo_nota = st.selectbox("Elegir formato de nota:", tipo_notas)
    # Declare output_dir with a default value
    output_dir = None

    if uploaded_file:
        data_list = extract_data_from_excel(uploaded_file)
        
        if st.button("Generar Notas"):
            output_dir = tempfile.mkdtemp()  # Create a temporary directory
            # generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
            if tipo_nota == "Una exclusión por nota":
                generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
            elif tipo_nota == "Múltiples exclusiones por nota":
                data_dict = group_data_by_code(data_list)
                multi_generate_word_files_streamlit(data_dict, template_dir, output_dir, uploaded_file, entidad, moneda, producto)
    
            st.session_state.generated_files = output_dir  # Store the output directory in session state
            st.success("Notas generadas exitosamente!") 
            file_counter = 0       
            # Call the download_notes function here
            download_notes(output_dir)

if __name__ == "__main__":
    main()
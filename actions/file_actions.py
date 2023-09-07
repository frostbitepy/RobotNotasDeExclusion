import openpyxl
import os
import datetime
import streamlit as st
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from actions.word_template_generator import generate_template_with_content


def extract_data_from_excel(excel_file_path):
    """Extracts data from an Excel file and returns it as a list of lists."""
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active
    data_list = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data_list.append(list(row))
    workbook.close()
    #data_list = reorder_values_for_entity(data_list, entidad)
    return data_list


def is_date(value):
    """Comprueba si el valor es una fecha."""
    return isinstance(value, datetime.datetime)
    

def format_date(date_obj):
    """Formatea un objeto datetime en formato DD/MM/AAAA."""
    formatted_date = date_obj.strftime("%d/%m/%Y")
    return formatted_date
        

def get_receptor_segun_entidad(entidad):
    """Returns the name of the receptor according to the entity."""
    switch_case = {
        "Provalor S.A.": "Sra. Viviana Trociuk",
        "Progresar Corporation S.A.": "Sra. Viviana Vergara",
        "Sudameris Bank S.A.E.C.A.": "Sra. Alicia González",
        "Factory S.A.": "Sra. Rocío González",
        "Creditos Paraná S.A.": "Sr. Andrés Servián"
    }    
    return switch_case.get(entidad, "Nombre del Receptor")


def get_formatted_date():
    """Returns the current date in the format "18 de enero de 2023"."""
    today = datetime.date.today()
    formatted_date = today.strftime("%d de %B de %Y")
    # Translate the month name to Spanish
    month_name = translate_month_to_spanish(today.strftime("%B"))    
    formatted_date = formatted_date.replace(today.strftime("%B"), month_name)
    return formatted_date


def get_current_month():
    """Returns the name of the current month."""
    today = datetime.date.today()
    return today.strftime("%B")


def translate_month_to_spanish(month):
    """Translates the name of a month from English to Spanish."""
    switch_case = {
        "January": "enero",
        "February": "febrero",
        "March": "marzo",
        "April": "abril",
        "May": "mayo",
        "June": "junio",
        "July": "julio",
        "August": "agosto",
        "September": "septiembre",
        "October": "octubre",
        "November": "noviembre",
        "December": "diciembre"
    }
    
    return switch_case.get(month, "Mes no válido")


def generate_word_files_streamlit(data_list, template_dir, output_dir, uploaded_file, entidad, moneda, producto):
    """Generates a Word file for each row's data using the appropriate template. Streamlit app."""
    # List all files in the template directory
    # template_files = os.listdir(template_dir)  

    generated_files = []  # List to store paths of generated files

    for index, data_row in enumerate(data_list):
        # Assuming the first element is the template name 
        # #template_name = data_row[14] + '.docx'
        doc = Document(template_dir)
        generate_template_with_content(doc, entidad, moneda, producto, data_row)
        output_word_path = f"{output_dir}/{entidad}_{data_row[14]}_document_{index + 1}.docx"
        doc.save(output_word_path)
        generated_files.append(output_word_path)  # Store generated file path


    # Store the generated files in session state
    st.session_state.generated_files = generated_files
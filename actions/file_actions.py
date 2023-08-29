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


def reorder_values_for_entity(data_list, entidad):
    
    # Implement the logic to reorder the data_list based on the entidad
    # For example, you can use a dictionary to define the order for each entidad
    entity_order = {
        "PROVALOR": [0, 1, 2, 3],  # Replace with the correct order of indexes
        "PROGRESAR": [2, 0, 1, 3],
        "SUDAMERIS": [1, 3, 0, 2],
        "FACTORY": [3, 2, 1, 0]
    }  
    if entidad in entity_order:
        order = entity_order[entidad]
        reordered_data_list = [data_list[i] for i in order]
        return reordered_data_list
    else:
        # Return the original data_list if the entidad is not recognized
        return data_list


def is_date(value):
    """Comprueba si el valor es una fecha."""
    return isinstance(value, datetime.datetime)
    

def format_date(date_obj):
    """Formatea un objeto datetime en formato DD/MM/AAAA."""
    formatted_date = date_obj.strftime("%d/%m/%Y")
    return formatted_date


def replace_placeholders_in_word_template(doc, data_row):
    """Replaces placeholders in a Word template with data from a row."""
    for i, value in enumerate(data_row):
        placeholder = f"{{Value{i+1}}}"
        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                paragraph.text = paragraph.text.replace(placeholder, str(value))
         

def replace_placeholders_in_table(doc, data_row):
    """Replaces placeholders in runs within a table with data from a row."""
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        for i, value in enumerate(data_row):
                            placeholder = f"{{Value{i+1}}}"
                            if is_date(value):
                                value = format_date(value)  # Format if it is a date
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(value))


def replace_additional_placeholders(doc, entidad, moneda):
    """Replaces additional placeholders in a Word document with specific data."""
    add_text_to_document(doc, "Encarnación, " + get_formatted_date())
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Replace placeholders with corresponding data
            #run.text = run.text.replace("{fechanota}", "Encarncación, " + get_formatted_date())
            run.text = run.text.replace("{entidad}", entidad)
            run.text = run.text.replace("{receptor}", get_receptor_segun_entidad(entidad))
            run.text = run.text.replace("{Val0}", moneda)
            run.text = run.text.replace("{mes}", translate_month_to_spanish(get_current_month()))      
            
    
def add_text_to_document(doc, new_text):
    # Obtiene el primer párrafo original
    first_paragraph = doc.paragraphs[0]
    # Borra el contenido del primer párrafo
    for run in first_paragraph.runs:
        run.clear()
    # Agrega el nuevo texto en el primer párrafo
    first_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run = first_paragraph.add_run(new_text)


def get_receptor_segun_entidad(entidad):
    """Returns the name of the receptor according to the entity."""
    switch_case = {
        "PROVALOR": "Sra. Viviana Trociuk",
        "PROGRESAR": "Sra. Raisa Gutmann",
        "SUDAMERIS": "Sra. Roxana Arias",
        "FACTORY": "Sra. Rocío González"
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


def generate_word_files(data_list, template_dir, output_dir):
    """Generates a Word file for each row's data using the appropriate template."""
    template_files = os.listdir(template_dir)  # List all files in the template directory

    for index, data_row in enumerate(data_list):
        template_name = data_row[14] + '.docx' # Assuming the first element is the template name

        if template_name in template_files:
            selected_template_path = os.path.join(template_dir, template_name)

            doc = Document(selected_template_path)
            replace_placeholders_in_word_template(doc, data_row)
            replace_placeholders_in_table(doc, data_row)
            replace_additional_placeholders(doc)  # Add entity-specific placeholders
            
            output_word_path = f"{output_dir}/{data_row[14]}_document_{index + 1}.docx"
            doc.save(output_word_path)


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
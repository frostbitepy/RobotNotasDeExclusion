import openpyxl
import os
import datetime
from docx import Document


def extract_data_from_excel(excel_file_path):
    """Extracts data from an Excel file and returns it as a list of lists."""
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    data_list = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data_list.append(list(row))

    workbook.close()
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


def replace_additional_placeholders(doc, entidad):
    """Replaces additional placeholders in a Word document with specific data."""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Replace placeholders with corresponding data
            run.text = run.text.replace("{fechanota}", get_formatted_date())
            run.text = run.text.replace("{entidad}", entidad)
            run.text = run.text.replace("{receptor}", "Nombre del Receptor")
            run.text = run.text.replace("{mes}", translate_month_to_spanish(get_current_month()))


def get_formatted_date():
    """Returns the current date in the format "18 de enero de 2023"."""
    today = datetime.date.today()
    formatted_date = today.strftime("%d de %B de %Y")
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
        template_name = data_row[0] + '.docx' # Assuming the first element is the template name

        if template_name in template_files:
            selected_template_path = os.path.join(template_dir, template_name)

            doc = Document(selected_template_path)
            replace_placeholders_in_word_template(doc, data_row)
            replace_placeholders_in_table(doc, data_row)
            replace_additional_placeholders(doc, "PROVALOR")  # Add entity-specific placeholders
            
            output_word_path = f"{output_dir}/{data_row[0]}_document_{index + 1}.docx"
            doc.save(output_word_path)
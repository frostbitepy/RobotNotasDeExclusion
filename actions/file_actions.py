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
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, str(value))

                        # Add the new placeholders and fill them with data
                        run.text = run.text.replace("{fecha_nota}", get_formatted_date())
                        run.text = run.text.replace("{entidad}", "PROVALOR S.A.")
                        run.text = run.text.replace("{receptor}", "Viviana Trociuk")
                        run.text = run.text.replace("{mes}", get_current_month())

def get_formatted_date():
    """Returns the current date in the format "18 de enero de 2023"."""
    today = datetime.date.today()
    formatted_date = today.strftime("%d de %B de %Y")
    return formatted_date


def get_current_month():
    """Returns the name of the current month."""
    today = datetime.date.today()
    return today.strftime("%B")


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
            
            output_word_path = f"{output_dir}/output_document_{index + 1}.docx"
            doc.save(output_word_path)
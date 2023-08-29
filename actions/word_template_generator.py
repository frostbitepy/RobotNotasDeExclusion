import datetime
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


#template_dir = "note_templates/template.docx"


def generate_cumulo_template(doc, data_row, currency, producto):
    from actions.file_actions import translate_month_to_spanish, get_current_month
    # Add texto exclusion
    exclusion_paragraph = doc.add_paragraph("Por la presente se informa la exclusión del Prestatario indicado a continuación, de la póliza de Seguro de Vida Colectivo para Cancelación de Deudas, por superar el capital de Gs. 3.000.000.000, establecido como cúmulo máximo por Asegurado.")
    exclusion_paragraph = doc.add_paragraph(("La operación corresponde a la planilla de {producto} de casa Matriz en moneda {moneda} del mes de {mes}.").format(producto=producto, moneda=currency, mes=translate_month_to_spanish(get_current_month())))
    exclusion_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # Add a table asegurado
    asegurado_table = doc.add_table(rows=3, cols=2)
    asegurado_table.style = 'Table Grid'
    nombre_cells = asegurado_table.rows[0].cells
    nombre_cells[0].text = 'Asegurado:'
    nombre_cells[1].text = data_row[0]
    documento_cells = asegurado_table.rows[1].cells
    documento_cells[0].text = 'Documento:'
    documento_cells[1].text = data_row[1]
    nacimiento_cells = asegurado_table.rows[2].cells
    nacimiento_cells[0].text = 'Fecha de Nacimiento:'
    nacimiento_cells[1].text = str(data_row[2])

    # Add a section break
    doc.add_paragraph(" ")
    
    # Add a table operacion
    operacion_table = doc.add_table(rows=2, cols=4)
    operacion_table.style = 'Table Grid'
    hdr_cells = operacion_table.rows[0].cells
    hdr_cells[0].text = 'Nro. Operación'
    hdr_cells[1].text = 'Monto'
    hdr_cells[2].text = 'Costo'
    hdr_cells[3].text = 'Fecha Vencimiento'
    for cell in operacion_table.rows[0].cells:
        cell.paragraphs[0].runs[0].font.bold = True
    row_cells = operacion_table.rows[1].cells
    row_cells[0].text = str(data_row[4])
    row_cells[1].text = str(data_row[6])
    row_cells[2].text = str(data_row[8])
    row_cells[3].text = str(data_row[11])   
    
    # Definir el color gris (RGB: 192, 192, 192)
    gray_color = RGBColor(240, 240, 240)

    # Agregar fondo de color gris a la celda
    for cell in hdr_cells:
        tcPr = cell._tc.get_or_add_tcPr()
        shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{gray_color}" />')
        tcPr.append(shading_elm)
   
    # Add a section break
    doc.add_paragraph(" ")


def generate_fallecimiento_template(doc, data_row, currency, producto):
    pass


def generate_falta_ds_template(doc, data_row, currency, producto):
    pass


def generate_ds_incompleta_template(doc, data_row, currency, producto):
    pass


def generate_sin_capital_template(doc, data_row, currency, producto):
    pass


def generate_template_with_content(doc, entity_name, currency, producto, data_row):
    from actions.file_actions import get_formatted_date, get_receptor_segun_entidad
    """Generates a Word file using the appropriate template."""
    # Add date
    date_paragraph = doc.add_paragraph("Encarnación, " + get_formatted_date())
    date_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # Add entidad
    entidad_paragraph = doc.add_paragraph("Señores")
    entidad_paragraph = doc.add_paragraph(entity_name)
    entidad_paragraph = doc.add_paragraph("Presente")
    entidad_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add receptor
    receptor_paragraph = doc.add_paragraph("Atn: " + get_receptor_segun_entidad(entity_name))
    receptor_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    receptor_paragraph = doc.add_paragraph("Ref.: Exclusión en Seguro de Vida Cancelación de Deudas")
    receptor_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    receptor_paragraph = doc.add_paragraph("Nota.  N°: /2023")
    receptor_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    template_name = data_row[14]    # Lugar de la lista donde se encuentra el motivo de la exclusion
    if template_name == "cumulo":
        generate_cumulo_template(doc, data_row, currency, producto)
    elif template_name == "fallecimiento":
        generate_fallecimiento_template(doc, data_row, currency, producto)
    elif template_name == "falta_ds":
        generate_falta_ds_template(doc, data_row, currency, producto)
    elif template_name == "ds_incompleta":
        generate_ds_incompleta_template(doc, data_row, currency, producto)
    elif template_name == "sin_capital":
        generate_sin_capital_template(doc, data_row, currency, producto)
        

    # Add despedoda
    despedida_paragraph = doc.add_paragraph("Sin otro particular nos despedimos de usted, atentamente.")
    despedida_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY 

    # Save the document
    template_filename = f"{entity_name}_{template_name}_{currency}_template.docx"
    doc.save(template_filename)
    return template_filename

# database_insert.py

import requests

def insert_data(data):
    url = "https://app.regionalseguros.com.py/ords/rs/insertGET/insertGET"

    # Modificar estos datos con los que deben insertarse en la base de datos
    data = {
        "p_dptno": 101,
        "p_dname": "Ventas",
        "p_loc": "Nueva York"
    }
    response = requests.post(url, data=data)

    if response.status_code == 200:
        return "Inserción exitosa"
    else:
        return f"Error al insertar datos. Código de estado: {response.status_code}"

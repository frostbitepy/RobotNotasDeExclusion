# database_insert.py
# url2 = "https://app.regionalseguros.com.py/ords/rs/insert/"
# url3 = "https://app.regionalseguros.com.py/ords/rs/insertGET/insertGET?p_loc=Encar&p_dname=Ventas&p_dptno=13"

import requests

def insert_data():
    url = "https://app.regionalseguros.com.py/ords/rs/insertGET/insertGET?p_loc=Encar&p_dname=Ventas&p_dptno=15"
 
    data = {
        "p_dptno": 15,
        "p_dname": "Ventas",
        "p_loc": "Encarnación"
    }
    response = requests.post(url, data=data)

    if response.status_code == 200:
        try:
            # Intenta convertir el contenido de la respuesta a un número
            result = int(response.text)
            return result
        except ValueError:
            return "Error: El servidor no devolvió un número válido"
    else:
        return f"Error al insertar datos. Código de estado: {response.status_code}"

def insert_data2():
    url = "https://app.regionalseguros.com.py/ords/rs/insertGET/insertGET?p_loc=Encar&p_dname=Ventas&p_dptno=15"

    # Define los datos que se enviarán al servidor
    # data = data 
    data = {
        "p_dptno": 15,
        "p_dname": "Ventas",
        "p_loc": "Encarnacion"
    }
    response = requests.post(url, data=data)
    return response

if __name__ == "__main__":
    # Puedes ejecutar la función directamente desde este archivo para probarla
    result = insert_data2()
    print(result)
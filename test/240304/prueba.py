# cliente_para_servicio_excel.py
import requests
import json
import os # <--- IMPORTACIÓN AÑADIDA

# URL del microservicio Flask
SERVICE_URL = "http://localhost:9898/generar_excel" # Asegúrate que el puerto coincida

def solicitar_generacion_excel(ruta_del_proyecto):
    """
    Envía una solicitud POST al microservicio para generar el archivo Excel.

    Args:
        ruta_del_proyecto (str): La ruta a la carpeta principal del proyecto
                                 que contiene 'datos.json' y las subcarpetas.
    """
    payload = {
        "ruta_proyecto": ruta_del_proyecto
    }
    headers = {
        'Content-Type': 'application/json'
    }

    print(f"\nEnviando solicitud a: {SERVICE_URL}")
    print(f"Payload: {json.dumps(payload, indent=2)}")

    try:
        # Se puede ajustar el timeout según sea necesario.
        # El timeout aquí es para la respuesta del servidor, no para la generación completa del Excel
        # si la generación es muy larga y el servidor responde inmediatamente con un "procesando".
        # Nuestro servidor actual espera a que la generación termine o timeoutee.
        response = requests.post(SERVICE_URL, headers=headers, json=payload, timeout=190) # Un poco más que el timeout del servidor

        # Verificar si la solicitud al servicio fue exitosa (código 2xx)
        response.raise_for_status() 

        # Intentar decodificar la respuesta JSON del servicio
        try:
            respuesta_servicio = response.json()
            print("\n--- Respuesta del Servicio ---")
            print(f"Estado: {respuesta_servicio.get('status')}")
            print(f"Mensaje: {respuesta_servicio.get('message')}")
            if respuesta_servicio.get('status') == 'success':
                print(f"Archivo Excel generado en: {respuesta_servicio.get('file_path')}")
        except json.JSONDecodeError:
            print("\n--- Error ---")
            print("No se pudo decodificar la respuesta JSON del servicio.")
            print(f"Respuesta recibida (texto): {response.text}")

    except requests.exceptions.HTTPError as http_err:
        print("\n--- Error HTTP ---")
        print(f"Error HTTP al contactar el servicio: {http_err}")
        if http_err.response is not None:
            print(f"Código de estado: {http_err.response.status_code}")
            try:
                error_details = http_err.response.json()
                print(f"Detalles del error del servidor: {error_details}")
            except json.JSONDecodeError:
                print(f"Respuesta del servidor (error texto): {http_err.response.text}")
    except requests.exceptions.ConnectionError as conn_err:
        print("\n--- Error de Conexión ---")
        print(f"No se pudo conectar al servicio en {SERVICE_URL}: {conn_err}")
        print("Asegúrate de que el microservicio Flask (app.py) esté corriendo.")
    except requests.exceptions.Timeout as timeout_err:
        print("\n--- Error de Timeout ---")
        print(f"La solicitud al servicio excedió el tiempo de espera: {timeout_err}")
    except requests.exceptions.RequestException as req_err:
        print("\n--- Error en la Solicitud ---")
        print(f"Ocurrió un error con la solicitud: {req_err}")
    except Exception as e:
        print("\n--- Error Inesperado ---")
        print(f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    print("Cliente para el Servicio de Generación de Excel")
    print("----------------------------------------------")
    
    # Solicitar la ruta al usuario
    ruta_ingresada = input("Por favor, ingresa la ruta completa de la carpeta del proyecto (ej. C:\\ruta\\a\\240304): ")
    
    if not ruta_ingresada:
        print("No se ingresó ninguna ruta. Saliendo.")
    else:
        # Normalizar la ruta por si se usan barras invertidas simples
        ruta_normalizada = os.path.normpath(ruta_ingresada)
        solicitar_generacion_excel(ruta_normalizada)

    print("\n--- Script cliente finalizado ---")

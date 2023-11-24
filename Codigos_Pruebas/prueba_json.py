import json

# Ruta del archivo JSON de entrada y salida
ruta_archivo = 'json/archivo.json'

# Función para cargar y modificar el archivo JSON
def modificar_json(ruta_archivo):
    with open(ruta_archivo, 'r', encoding='utf-8') as archivo:
        data = json.load(archivo)

    # Filtrar las entradas con 'Numero_de_documento' no nulo o vacío
    data_filtrada = [registro for registro in data if registro.get('Numero_de_documento') not in [None, ""]]

    # Sobrescribir el archivo original con los datos filtrados y estructurados
    with open(ruta_archivo, 'w', encoding='utf-8') as archivo:
        json.dump(data_filtrada, archivo, ensure_ascii=False, indent=4)

# Llamar a la función para modificar el archivo JSON
modificar_json(ruta_archivo)

print(f'Json modificado')

import openai
from docx import Document
from pathlib import Path
import os

# Configuración de la clave de la API de OpenAI
<<<<<<< HEAD
API_KEY = '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'
=======
API_KEY = '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'
>>>>>>> 07757c7e23417c855c84eda7f4a847e73c9a8d83
openai.api_key = API_KEY


# Leer los primeros cuatro párrafos del documento DOCX
def leer_primeros_parrafos(doc_path, num_parrafos=4):
    doc = Document(doc_path)
    texto = []
    for i, paragraph in enumerate(doc.paragraphs):
        if i >= num_parrafos:
            break
        texto.append(paragraph.text)
    return "\n".join(texto)

def main():
    try:
        # Obtener una lista de archivos .docx en la carpeta documentos, excluyendo los archivos temporales
        archivos_docx = [archivo for archivo in os.listdir("documentos") if archivo.lower().endswith('.docx') and not archivo.startswith('~$')]
        print("Buscando documentos .docx")
        # Verificar si se encontraron archivos .docx en la carpeta
        if not archivos_docx:
            print("No se encontraron archivos .docx en la carpeta especificada.")
            return

        for archivo_docx in archivos_docx:
            # Obtener la ruta completa del archivo .docx
            documento_path = os.path.join("documentos", archivo_docx)

            # Obtener el texto de los dos primeros párrafos del documento
            texto_documento = leer_primeros_parrafos(documento_path)


            # Generar el JSON utilizando OpenAI GPT-3.5 Turbo (modelo de chat)
            prompt = f'''Toma el siguiente texto y extrae de él una lista de objetos JSON de cada una de las personas que se encuentran allí. 
            Estos objetos deben tener las siguientes propiedades: 
            Generar un Id a cada uno, empieza por el 1.
            Tipo (natural o jurídica) 
            Nombre 
            Nacionalidad 
            Tipo_de_documento(DNI o RUC) 
            Numero_de_documento
            Estado_civil (casado, divorciado, viudo, soltero o sus equivalentes femeninos) 
            Domicilio
            (Guarda los nombres de las propiedades tal cual tomando en cuenta las mayúsculas)
            Denominación que se le dará en el documento (después de 'a quien en adelante se le denominará') 
            Si existen representantes, deben ser una lista de enteros indicando la posición de cada representante en la lista de objetos. 
            También debe haber un campo de relaciones para relaciones conyugales, que será una lista de objetos con propiedades: 
            Índice (corresponde al índice del objeto relacionado) 
            Tipo de relación (solo para relaciones maritales u otros tipos) 
            Las propiedades que no se encuentren deben ser NULL o vacías en el caso de propiedades de lista. 
            Por favor, responde solo con el archivo JSON solicitado, sin añadir contexto adicional'''

            print("Enviando solicitud a OpenAI para generar el JSON...")
<<<<<<< HEAD
            respuesta_openai = openai.chat.completions.create(
=======
            respuesta_openai = openai.ChatCompletion.create(
>>>>>>> 07757c7e23417c855c84eda7f4a847e73c9a8d83
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt + "\n\n" + texto_documento}
                ]
            )
<<<<<<< HEAD
            json_generado = respuesta_openai.choices[0].message.content.strip()
=======
            json_generado = respuesta_openai.choices[0].message["content"].strip()
>>>>>>> 07757c7e23417c855c84eda7f4a847e73c9a8d83

            print("JSON generado por OpenAI:")
            print(json_generado)  # Imprime el JSON generado por OpenAI
            # Definir la ruta del archivo JSON de salida
            ruta_json = Path("json/archivo.json")

            # Guardar el JSON en el archivo
            with open(ruta_json, "w", encoding="utf-8") as archivo_json:
                archivo_json.write(json_generado)

            print(f"JSON guardado en '{ruta_json}'.")

    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()

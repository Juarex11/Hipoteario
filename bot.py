from pathlib import Path
import win32com.client 
from docx import Document
from docx.shared import Inches, Cm

def main():
    # Configuración de rutas
    directorio_actual = Path(__file__).parent if "__file__" in locals() else Path.cwd()
    directorio_entrada = directorio_actual / "kardex"
    directorio_salida = directorio_actual / "modificados"
    directorio_salida.mkdir(parents=True, exist_ok=True)

    # Configuración de búsqueda
    cadena_a_buscar = "17 de abril del 2023"

    # Solicitar al usuario que ingrese la cadena a reemplazar
    cadena_a_reemplazar = input("Ingrese la fecha de hoy: ")

    # Configuración de reemplazo en Word
    reemplazo_word = 2
    continuar_busqueda_word = 1

    # Inicializar la aplicación Word
    word_app = win32com.client.DispatchEx("Word.Application")
    word_app.Visible = False
    word_app.DisplayAlerts = False

    try:
        # Recorrer todos los documentos en el directorio de entrada
        for archivo in Path(directorio_entrada).rglob("*.*"):
            # Obtener la extensión del archivo
            ext = archivo.suffix.lower()

            if ext in (".docx", ".doc"):
                if ext == ".docx":
                    # Procesar archivos .docx utilizando la biblioteca docx
                    doc = Document(archivo)

                    
                    for para in doc.paragraphs:
                        para.alignment = 3  # 3 representa justificación (centro). Cambiar a 1 para alinear a la izquierda, 2 para alinear a la derecha, 0 para justificar
                        para.vertical_alignment = 1

                    # Configurar los márgenes
                    for section in doc.sections:
                        section.page_width = Cm(21)  # Ancho de página
                        section.page_height = Cm(29.7)  # Alto de página (A4)
                        section.left_margin = Cm(1.1)  # Margen izquierdo (1.1 cm)
                        section.right_margin = Cm(4)  # Margen derecho (4 cm)
                        section.top_margin = Cm(5)  # Margen superior (5 cm)
                        section.bottom_margin = Cm(1.1)  # Margen inferior (1.1 cm)
                        

                    doc.save(archivo)  # Guardar los cambios en los márgenes
                    # Configurar el documento en una sola columna
                    

                # Procesar archivos .doc utilizando pywin32
                word_app.Documents.Open(str(archivo))
                print(f"Procesando: {archivo.name}")

                # Seleccionar todo el texto en el documento
                word_app.Selection.WholeStory()

                # Cambiar el tamaño de fuente a 9
                word_app.Selection.Font.Size = 9

                # Cambiar la fuente a "Arial Narrow"
                word_app.Selection.Font.Name = "Arial Narrow"

                # Configurar el documento en una sola columna
                for section in word_app.ActiveDocument.Sections:
                    text_columns = section.PageSetup.TextColumns
                    text_columns.SetCount(1)  # Una columna
                    text_columns.LineBetween = 0  # Sin líneas entre columnas


                # Realizar la búsqueda y reemplazo en el texto del documento
                word_app.Selection.Find.Execute(
                    FindText=cadena_a_buscar,
                    ReplaceWith=cadena_a_reemplazar,
                    Replace=reemplazo_word,
                    Forward=True,
                    MatchCase=True,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Wrap=continuar_busqueda_word,
                    Format=False,
                )
                # Eliminar todos los espacios vacíos entre párrafos
                word_app.Selection.WholeStory()
                word_app.Selection.Find.Execute(
                    FindText="^p^p",  # Dos párrafos vacíos consecutivos
                    ReplaceWith="^p",  # Reemplazar por un solo párrafo
                    Replace=reemplazo_word,
                    Forward=True,
                    MatchCase=False,
                    MatchWholeWord=False,
                    MatchWildcards=False,
                    MatchSoundsLike=False,
                    MatchAllWordForms=False,
                    Wrap=continuar_busqueda_word,
                    Format=False,
                )

            # Guardar el nuevo archivo con la extensión .docx si es .doc
                if archivo.suffix.lower() == ".doc":
                    archivo_salida = directorio_salida / (archivo.stem + ".docx")
                else:
                    archivo_salida = directorio_salida / archivo.name
                word_app.ActiveDocument.SaveAs(str(archivo_salida), FileFormat=16)  # FileFormat=16 para .docx
                word_app.ActiveDocument.Close(SaveChanges=False)
                print(f"Procesamiento de {archivo.name} finalizado")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        # Cerrar la aplicación de Word cuando se ha terminado
        word_app.Application.Quit()
        #SI FUNCIONA LEE LOS PROGRAMAS .DOC .DOCX
if __name__ == "__main__":
    main()

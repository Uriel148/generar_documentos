from docx import Document

def reemplazar_palabra_en_docx():
    # Solicitar datos al usuario
    documento_entrada = input("Ingresa la ruta del documento Word (.docx): ")
    palabra_buscar = input("Ingresa la palabra a buscar (ej. XXXXX): ")
    palabra_reemplazo = input("Ingresa la palabra de reemplazo: ")
    documento_salida = input("Ingresa el nombre para el nuevo documento: ")

    # Cargar el documento
    doc = Document(documento_entrada)
    
    # Función para reemplazar en elementos principales
    def reemplazar_en_elemento(elemento):
        if hasattr(elemento, 'runs'):
            for run in elemento.runs:
                if palabra_buscar in run.text:
                    run.text = run.text.replace(palabra_buscar, palabra_reemplazo)
    
    # Procesar párrafos principales
    for párrafo in doc.paragraphs:
        reemplazar_en_elemento(párrafo)
    
    # Procesar tablas
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for párrafo in celda.paragraphs:
                    reemplazar_en_elemento(párrafo)
    
    # Procesar encabezados y pies de página
    for sección in doc.sections:
        for encabezado in [sección.header, sección.first_page_header]:
            if encabezado:
                for párrafo in encabezado.paragraphs:
                    reemplazar_en_elemento(párrafo)
        
        for pie in [sección.footer, sección.first_page_footer]:
            if pie:
                for párrafo in pie.paragraphs:
                    reemplazar_en_elemento(párrafo)
    
    # Guardar el documento modificado
    doc.save(documento_salida)
    print(f"\nDocumento modificado guardado como: {documento_salida}")

if __name__ == "__main__":
    reemplazar_palabra_en_docx()
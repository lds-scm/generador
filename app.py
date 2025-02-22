import streamlit as st
import datetime
from docx import Document

# Función para generar documento con los datos del alumno
def generar_documento(fecha, nombre, dni, grado, tipo_documento, año_escolar):
    # Asegurarse de que los valores sean cadenas de texto
    dni = str(dni)
    grado = str(grado)  # Convertimos también grado a string si es necesario
    año_escolar = str(año_escolar)
    
    # Convertir la fecha a cadena en el formato deseado
    if isinstance(fecha, datetime.date):  # Verificamos si 'fecha' es un objeto datetime.date
        fecha = fecha.strftime("%d/%m/%Y")  # Formato de fecha como DD/MM/AAAA
    
    # Determinar el nombre del archivo de plantilla basado en el tipo de documento
    archivo_template = f"{tipo_documento}.docx"
    
    try:
        # Cargar el documento de plantilla
        doc = Document(archivo_template)
        
        # Reemplazar las marcas de texto con los datos del alumno
        for p in doc.paragraphs:
            for run in p.runs:
                if '[[NOMBRE]]' in run.text:
                    run.text = run.text.replace('[[NOMBRE]]', nombre)
                if '[[DNI]]' in run.text:
                    run.text = run.text.replace('[[DNI]]', dni)
                if '[[GRADO]]' in run.text:
                    run.text = run.text.replace('[[GRADO]]', grado)
                if '[[FECHA]]' in run.text:
                    run.text = run.text.replace('[[FECHA]]', fecha)
                if '[[AÑO]]' in run.text:
                    run.text = run.text.replace('[[AÑO]]', año_escolar)

        # Guardar el nuevo documento con un nombre personalizado
        nuevo_nombre = f"{tipo_documento}_{nombre}.docx"
        doc.save(nuevo_nombre)
        return nuevo_nombre
    
    except Exception as e:
        return f"Error al generar el documento: {e}"


# Crear la interfaz Streamlit
st.title('Generar Documento del Alumno')

# Formulario Streamlit
nombre = st.text_input('Nombre:')
dni = st.text_input('DNI:')
grado = st.selectbox('Grado:', ['1ro Primaria', '2do Primaria', '3ro Primaria', '4to Primaria', '5to Primaria', '6to Primaria', '1ro Secundaria', '2do Secundaria', '3ro Secundaria', '4to Secundaria', '5to Secundaria'])
tipo_documento = st.multiselect('Tipo de Documento:', ['Constancia de Estudios', 'Constancia de Notas', 'Constancia de No Adeudamiento'])
fecha = st.date_input('Fecha:', datetime.date.today())
año_escolar = st.text_input('Año Escolar:')

# Botón para generar el documento
if st.button('Generar Documento'):
    documentos_generados = []
    
    for doc_tipo in tipo_documento:
        # Generar documento
        documento = generar_documento(fecha, nombre, dni, grado, doc_tipo, año_escolar)
        
        if documento and "Error" not in documento:
            documentos_generados.append(documento)
        elif "Error" in documento:
            st.error(documento)
    
    if documentos_generados:
        # Descargar el primer documento generado
        with open(documentos_generados[0], "rb") as f:
            st.download_button(
                label="Descargar Documento",
                data=f,
                file_name=documentos_generados[0],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error('Error al generar el documento.')

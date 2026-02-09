import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Test de Integridad Endalia", page_icon="🧪")

st.title("🧪 Fase 1: Prueba de Integridad")
st.markdown("""
Sube tu plantilla original de Endalia. El programa la leerá y te permitirá descargarla 
**sin hacer ningún cambio**. 
    
**Objetivo:** Verificar que el archivo descargado (`test_integridad.xlsx`) sigue teniendo las flechitas de los desplegables en Excel.
""")

# Cargador de archivo
uploaded_file = st.file_uploader("Sube la plantilla de Endalia aquí", type=["xlsx"])

if uploaded_file:
    try:
        # Paso 1: Leer el archivo de forma "cruda"
        # keep_vba=True mantiene estructuras complejas y macros si las hubiera
        # data_only=False asegura que no perdamos las fórmulas/validaciones
        wb = load_workbook(uploaded_file, data_only=False, keep_vba=True)
        
        st.success("✅ Archivo cargado en memoria correctamente.")
        
        # Paso 2: Guardarlo en un buffer (en memoria) sin modificar nada
        output = BytesIO()
        wb.save(output)
        processed_data = output.getvalue()
        
        # Paso 3: Botón de descarga
        st.download_button(
            label="📥 Descargar copia de prueba",
            data=processed_data,
            file_name="test_integridad.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("💡 Instrucciones: Descarga el archivo, ábrelo en tu ordenador y confirma si puedes ver los desplegables en las columnas correspondientes.")

    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")

import streamlit as st
from openpyxl import load_workbook
from io import BytesIO

# Configuración de la página
st.set_page_config(page_title="Test de Integridad Endalia", page_icon="🧪")

st.title("🧪 Fase 1: Prueba de Espejo")
st.markdown("""
Esta versión intenta devolverte el archivo **exactamente** como entró, 
sin que Excel detecte que ha sido manipulado por un software externo.
""")

uploaded_file = st.file_uploader("Sube la plantilla de Endalia aquí", type=["xlsx"])

if uploaded_file:
    try:
        # Cargamos el archivo original
        # keep_vba=True es crucial para que no borre las validaciones ocultas
        # data_only=False evita que se pierdan las fórmulas
        wb = load_workbook(uploaded_file, data_only=False, keep_vba=True)
        
        st.success("✅ Archivo cargado en memoria.")
        
        # Guardamos en un buffer intermedio
        output = BytesIO()
        wb.save(output)
        
        # Forzamos que el puntero vuelva al inicio para que Streamlit lea el archivo completo
        output.seek(0)
        processed_data = output.read()

        st.download_button(
            label="📥 Descargar copia de prueba",
            data=processed_data,
            file_name="test_espejo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("💡 Si este archivo abre y tiene los desplegables, ya podemos meter la lógica de los 14 tramos.")

    except Exception as e:
        st.error(f"Error técnico: {e}")

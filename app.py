import streamlit as st
import openpyxl
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Reparador de Integridad", page_icon="🛠️")

st.title("🛠️ Fase 1: Prueba de Flujo Limpio")
st.markdown("""
Si el error de 'archivo dañado' persiste, es probable que la plantilla tenga una estructura protegida. 
Este código utiliza un método de carga y guardado directo para intentar mantener la compatibilidad total con Excel.
""")

# Cargador de archivos
uploaded_file = st.file_uploader("Sube la plantilla de Endalia", type=["xlsx"])

if uploaded_file:
    try:
        # Leemos el archivo en un buffer de entrada
        file_bytes = uploaded_file.read()
        input_buffer = BytesIO(file_bytes)
        
        # Cargamos el libro de trabajo (workbook)
        # keep_vba=True es esencial para mantener macros o validaciones avanzadas
        # data_only=False asegura que las fórmulas no se conviertan en valores estáticos
        wb = openpyxl.load_workbook(input_buffer, data_only=False, keep_vba=True)
        
        st.success("✅ Archivo cargado correctamente en memoria.")
        
        # Preparamos el buffer de salida
        output_buffer = BytesIO()
        
        # Guardamos el archivo directamente al buffer
        wb.save(output_buffer)
        
        # Resetear el puntero al inicio para la descarga
        output_buffer.seek(0)
        
        st.download_button(
            label="📥 Descargar y Probar en Excel",
            data=output_buffer,
            file_name="plantilla_test_integridad.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.info("💡 Por favor, abre el archivo descargado. Si Excel dice que está dañado, prueba a darle a 'Sí' en reparar y mira si los desplegables siguen vivos.")

    except Exception as e:
        st.error(f"Error crítico al procesar: {e}")

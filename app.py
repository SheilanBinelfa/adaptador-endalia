import streamlit as st
import openpyxl
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Editor Fiel Endalia", page_icon="📝")

st.title("📝 Prueba de Edición Mínima")
st.markdown("""
Esta prueba intenta realizar un cambio invisible en la celda **Z1** (lejos de tus datos). 
Si este archivo se abre correctamente en tu Excel, significa que ya podemos avanzar con la lógica de los empleados.
""")

# Cargador de archivos
uploaded_file = st.file_uploader("Sube la plantilla de Endalia", type=["xlsx"])

if uploaded_file:
    try:
        # 1. Leer los bytes directamente del archivo subido
        bytes_data = uploaded_file.getvalue()
        input_buffer = BytesIO(bytes_data)
        
        # 2. Carga simplificada
        # No usamos keep_vba ni otros parámetros complejos que suelen corromper archivos protegidos
        wb = openpyxl.load_workbook(input_buffer, data_only=False)
        
        # 3. Acceder a la hoja "Registros de jornada"
        if "Registros de jornada" in wb.sheetnames:
            ws = wb["Registros de jornada"]
            # Escribimos algo en una celda vacía para que la librería genere una nueva firma de archivo
            ws['Z1'] = " "
            st.success("✅ Hoja 'Registros de jornada' localizada y editada.")
        else:
            st.warning("⚠️ No se encontró la hoja 'Registros de jornada', editando hoja activa.")
            ws = wb.active
            ws['Z1'] = " "

        # 4. Guardar el resultado en el buffer de salida
        output_buffer = BytesIO()
        wb.save(output_buffer)
        output_buffer.seek(0)
        
        st.download_button(
            label="📥 Descargar y Probar en Excel",
            data=output_buffer,
            file_name="plantilla_editada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error técnico durante el proceso: {e}")

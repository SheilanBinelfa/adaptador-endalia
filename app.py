import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Real", page_icon="🎯", layout="wide")

st.title("🎯 Adaptador Endalia: Inyección Total")
st.markdown("""
Esta versión está diseñada para **no tocar** los desplegables. 
1. Abre tu plantilla original.
2. Localiza a los empleados del registro.
3. Escribe los datos *dentro* de las celdas existentes.
""")

# --- BARRA LATERAL ---
st.sidebar.header("⚙️ Configuración")
bulk_time = st.sidebar.time_input("Hora de Cierre masivo", datetime.time(18, 0))
tz_text = st.sidebar.text_input("Zona Horaria", "(UTC+01:00) Bruselas, Copenhague, Madrid, París")

# --- CARGA ---
col1, col2 = st.columns(2)
with col1:
    f_plantilla = st.file_uploader("1. Sube la Plantilla con Desplegables", type=["xlsx"])
with col2:
    f_datos = st.file_uploader("2. Sube tus 14 tramos (o los que tengas)", type=["xlsx", "csv"])

if f_plantilla and f_datos:
    try:
        # Cargar datos a importar
        df_in = pd.read_excel(f_datos) if f_datos.name.endswith('xlsx') else pd.read_csv(f_datos)
        
        if st.button("🚀 GENERAR EXCEL CON DATOS Y DESPLEGABLES"):
            # CARGA QUIRÚRGICA: Abrimos el archivo real
            # keep_vba=True ayuda a que Excel no crea que es un archivo 'limpio'
            wb = load_workbook(f_plantilla, keep_vba=True, data_only=False)
            ws = wb["Registros de jornada"]
            
            # Identificar columnas por nombre exacto
            headers = {cell.value: i+1 for i, cell in enumerate(ws[1]) if cell.value is not None}
            
            # Columnas críticas en la plantilla de Endalia
            cols_map = {
                "id": headers.get("Nº doc. Identificador") or headers.get("Código empleado"),
                "fec": headers.get("Fecha de referencia"),
                "ini": headers.index.get("Inicio") if "Inicio" in headers else headers.get("Inicio"),
                "fin": headers.get("Fin"),
                "tipo": headers.get("Tipo de tramo"),
                "zona": headers.get("Zona Horaria"),
                "sob": headers.get("Sobrescritura")
            }
            
            # Como los nombres de cabecera pueden variar, buscamos por posición si fallan
            col_id = 1 # A (Identificador)
            col_fec = 4 # D (Fecha)
            col_ini = 6 # F (Inicio)
            col_fin = 7 # G (Fin)
            col_tipo = 8 # H (Tipo de tramo -> DESPLEGABLE)
            col_sob = 9 # I (Sobrescritura -> DESPLEGABLE)

            count = 0
            # Iterar sobre tus 14 tramos
            for _, row_data in df_in.iterrows():
                search_val = str(row_data['Empleado']).strip().upper()
                
                # Buscar en la plantilla (Columna A o B)
                for r in range(2, ws.max_row + 1):
                    cell_val = str(ws.cell(row=r, column=1).value).strip().upper()
                    
                    if search_val in cell_val:
                        # INYECCIÓN: Solo cambiamos el .value de la celda
                        # Esto NO borra la validación de datos (el desplegable)
                        ws.cell(row=r, column=col_fec).value = str(row_data['Fecha'])
                        ws.cell(row=r, column=col_ini).value = str(row_data['Hora inicio'])
                        
                        h_fin = str(row_data['Hora fin']) if pd.notnull(row_data['Hora fin']) and "00:00" not in str(row_data['Hora fin']) else bulk_time.strftime("%H:%M")
                        ws.cell(row=r, column=col_fin).value = h_fin
                        
                        # Valores que deben coincidir con el desplegable
                        ws.cell(row=r, column=col_tipo).value = "Trabajo" 
                        ws.cell(row=r, column=5).value = tz_text # Zona Horaria
                        ws.cell(row=r, column=col_sob).value = "SÍ"
                        
                        count += 1
                        break

            # GUARDADO ESPECIAL
            output = BytesIO()
            wb.save(output)
            processed_data = output.getvalue()
            
            st.success(f"¡Hecho! Se han inyectado {count} tramos manteniendo los desplegables.")
            
            st.download_button(
                label="📥 DESCARGAR PLANTILLA FINAL",
                data=processed_data,
                file_name="Endalia_Importacion_OK.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}. Asegúrate de que los nombres de las columnas en tu archivo de datos sean 'Empleado', 'Fecha', 'Hora inicio'.")

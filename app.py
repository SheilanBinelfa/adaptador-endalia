import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Pro", page_icon="📊", layout="wide")

# Estilos para que la interfaz sea limpia y profesional
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    .stButton>button { 
        width: 100%; 
        border-radius: 8px; 
        height: 3em; 
        font-weight: bold; 
        background-color: #2563eb; 
        color: white; 
        border: none;
    }
    .stButton>button:hover { 
        background-color: #1d4ed8; 
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🚀 Adaptador Quirúrgico Endalia")
st.info("Este motor edita la plantilla original celda por celda para garantizar que los menús desplegables y formatos de Endalia permanezcan intactos.")

# --- BARRA LATERAL: CONFIGURACIÓN ---
st.sidebar.header("⚙️ Parámetros de Inyección")
bulk_end_time = st.sidebar.time_input("Hora de Cierre para tramos abiertos", datetime.time(18, 0))
global_timezone = st.sidebar.text_input("Zona Horaria Exacta", "(UTC+01:00) Bruselas, Copenhague, Madrid, París")
global_overwrite = st.sidebar.selectbox("¿Sobrescribir datos existentes?", ["SÍ", "NO"], index=0)

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plantilla de Endalia")
    file_plantilla = st.file_uploader("Sube el Excel ORIGINAL (con desplegables)", type=["xlsx"], key="p")

with col2:
    st.subheader("2. Registro de Tramos")
    file_registros = st.file_uploader("Sube el archivo con los tramos a importar", type=["xlsx", "csv"], key="r")

if file_plantilla and file_registros:
    try:
        # 1. Leer los registros que queremos importar
        if file_registros.name.endswith('.csv'):
            df_registros = pd.read_csv(file_registros)
        else:
            df_registros = pd.read_excel(file_registros)
        
        def normalize(name):
            return str(name).strip().upper() if pd.notnull(name) else ""

        st.success(f"Se han cargado {len(df_registros)} registros para procesar.")

        if st.button("💉 INYECTAR DATOS Y MANTENER DESPLEGABLES"):
            # 2. CARGA QUIRÚRGICA: Abrimos el archivo sin evaluar fórmulas para no perder metadatos
            wb = load_workbook(file_plantilla, data_only=False)
            
            if "Registros de jornada" not in wb.sheetnames:
                st.error("No se encontró la hoja 'Registros de jornada' en la plantilla.")
            else:
                ws = wb["Registros de jornada"]
                
                # Mapear columnas de la fila 1 para saber dónde escribir
                headers = [str(cell.value) for cell in ws[1]]
                try:
                    m = {
                        "emp": headers.index("Empleado") + 1,
                        "fec": headers.index("Fecha de referencia") + 1,
                        "ini": headers.index("Inicio") + 1,
                        "fin": headers.index("Fin") + 1,
                        "tip": headers.index("Tipo de tramo") + 1,
                        "zon": headers.index("Zona Horaria") + 1,
                        "sob": headers.index("Sobrescritura") + 1
                    }
                except ValueError as e:
                    st.error(f"La plantilla no tiene el formato estándar de Endalia. Falta: {e}")
                    st.stop()

                log_cambios = []
                
                # 3. Procesar cada tramo del archivo de entrada
                for _, reg in df_registros.iterrows():
                    nombre_buscado = normalize(reg['Empleado'])
                    encontrado = False
                    
                    # Buscamos la fila correspondiente en la plantilla
                    for r in range(2, ws.max_row + 1):
                        nombre_en_plantilla = normalize(ws.cell(row=r, column=m["emp"]).value)
                        
                        if nombre_en_plantilla == nombre_buscado:
                            # Determinamos la hora de fin (si no viene en el archivo, usamos la masiva)
                            h_fin = str(reg['Hora fin']) if pd.notnull(reg['Hora fin']) and str(reg['Hora fin']) != "00:00" else bulk_end_time.strftime("%H:%M")
                            
                            # Escribimos solo el VALOR. Al no tocar la celda completa, el desplegable se mantiene.
                            ws.cell(row=r, column=m["fec"]).value = str(reg['Fecha'])
                            ws.cell(row=r, column=m["ini"]).value = str(reg['Hora inicio'])
                            ws.cell(row=r, column=m["fin"]).value = h_fin
                            ws.cell(row=r, column=m["tip"]).value = str(reg.get('Tipo de tramo', 'Trabajo'))
                            ws.cell(row=r, column=m["zon"]).value = global_timezone
                            ws.cell(row=r, column=m["sob"]).value = global_overwrite
                            
                            log_cambios.append({"Empleado": reg['Empleado'], "Fila": r, "Estado": "✅ Inyectado"})
                            encontrado = True
                            break
                    
                    if not encontrado:
                        log_cambios.append({"Empleado": reg['Empleado'], "Fila": "-", "Estado": "⚠️ No encontrado"})

                # 4. Mostrar resumen y generar descarga
                st.subheader("Resumen de la Inyección")
                st.dataframe(pd.DataFrame(log_cambios))
                
                output = BytesIO()
                wb.save(output)
                
                st.download_button(
                    label="💾 DESCARGAR EXCEL CON DESPLEGABLES",
                    data=output.getvalue(),
                    file_name=f"Endalia_Final_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
    except Exception as e:
        st.error(f"Error durante el proceso: {e}")
else:
    st.info("Esperando archivos para procesar...")

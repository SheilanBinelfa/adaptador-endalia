import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Real", page_icon="🎯", layout="wide")

st.title("🚀 Adaptador Endalia: Versión Ultra-Fiel")
st.info("Esta versión edita 'dentro' de tu plantilla original. Esto garantiza que las flechitas (desplegables) y las reglas de validación de Endalia no se borren.")

# --- SIDEBAR: PARÁMETROS ---
st.sidebar.header("⚙️ Configuración de Inyección")
bulk_time = st.sidebar.time_input("Hora de Cierre por defecto", datetime.time(18, 0))
timezone_val = st.sidebar.text_input("Zona Horaria", "(UTC+01:00) Bruselas, Copenhague, Madrid, París")
overwrite_val = st.sidebar.selectbox("Sobrescritura", ["SÍ", "NO"], index=0)

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plantilla de Endalia")
    f_plantilla = st.file_uploader("Sube el Excel ORIGINAL (el que tiene los desplegables)", type=["xlsx"])

with col2:
    st.subheader("2. Registro de Tramos")
    f_registros = st.file_uploader("Sube el archivo con los 14 tramos a importar", type=["xlsx", "csv"])

if f_plantilla and f_registros:
    try:
        # Cargar registros a importar
        if f_registros.name.endswith('.csv'):
            df_in = pd.read_csv(f_registros)
        else:
            df_in = pd.read_excel(f_registros)
        
        # Función para limpiar y comparar nombres
        def clean(val):
            return str(val).strip().upper() if pd.notnull(val) else ""

        st.success(f"Se han cargado {len(df_in)} tramos para procesar.")

        if st.button("💉 INYECTAR DATOS Y MANTENER DESPLEGABLES"):
            # CARGA QUIRÚRGICA: Abrimos el archivo original
            # keep_vba=True ayuda a mantener la estructura compleja del Excel
            wb = load_workbook(f_plantilla, data_only=False, keep_vba=True)
            
            if "Registros de jornada" not in wb.sheetnames:
                st.error("Error: No se encuentra la pestaña 'Registros de jornada'.")
            else:
                ws = wb["Registros de jornada"]
                
                # Detectar columnas por el encabezado de la fila 1
                headers = [str(cell.value) for cell in ws[1]]
                
                try:
                    # Buscamos los índices (base 1 para openpyxl)
                    idx_emp = headers.index("Empleado") + 1
                    idx_fec = headers.index("Fecha de referencia") + 1
                    idx_ini = headers.index("Inicio") + 1
                    idx_fin = headers.index("Fin") + 1
                    idx_tipo = headers.index("Tipo de tramo") + 1
                    idx_zona = headers.index("Zona Horaria") + 1
                    idx_sob = headers.index("Sobrescritura") + 1
                except ValueError as e:
                    st.error(f"Formato de plantilla no reconocido. Falta columna: {e}")
                    st.stop()

                log_resultados = []

                # Procesar cada fila de tus 14 registros
                for _, row in df_in.iterrows():
                    nombre_buscado = clean(row['Empleado'])
                    encontrado = False
                    
                    # Buscar al empleado en la plantilla original fila por fila
                    for r in range(2, ws.max_row + 1):
                        nombre_celda = clean(ws.cell(row=r, column=idx_emp).value)
                        
                        if nombre_buscado in nombre_celda or nombre_celda in nombre_buscado:
                            # Inyección de valores: SOLO editamos el .value
                            # Esto deja intacta la "Validación de Datos" de la celda
                            ws.cell(row=r, column=idx_fec).value = str(row['Fecha'])
                            ws.cell(row=r, column=idx_ini).value = str(row['Hora inicio'])
                            
                            # Gestión de Hora Fin
                            h_fin = str(row['Hora fin']) if pd.notnull(row['Hora fin']) and "00:00" not in str(row['Hora fin']) else bulk_time.strftime("%H:%M")
                            ws.cell(row=r, column=idx_fin).value = h_fin
                            
                            # Valores que deben coincidir con las opciones de los desplegables de Endalia
                            ws.cell(row=r, column=idx_tipo).value = "Trabajo" 
                            ws.cell(row=r, column=idx_zona).value = timezone_val
                            ws.cell(row=r, column=idx_sob).value = overwrite_val
                            
                            log_resultados.append({"Empleado": row['Empleado'], "Resultado": "✅ Inyectado en fila " + str(r)})
                            encontrado = True
                            break
                    
                    if not encontrado:
                        log_resultados.append({"Empleado": row['Empleado'], "Resultado": "❌ No encontrado en plantilla"})

                st.subheader("Resumen del Proceso")
                st.table(pd.DataFrame(log_resultados))

                # Guardado binario manteniendo la integridad de los metadatos
                output = BytesIO()
                wb.save(output)
                
                st.download_button(
                    label="📥 DESCARGAR EXCEL CON DESPLEGABLES ACTIVOS",
                    data=output.getvalue(),
                    file_name=f"Endalia_Final_Con_Desplegables.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Error crítico durante el procesado: {e}")
else:
    st.info("Sube los archivos para activar el motor de inyección.")

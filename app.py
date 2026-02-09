import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Pro", page_icon="📊", layout="wide")

# --- ESTILOS CSS CORREGIDOS (Se eliminó el error de escritura en unsafe_allow_html) ---
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
    .stInfo { border-left: 5px solid #2563eb; }
    </style>
    """, unsafe_allow_html=True)

st.title("🚀 Adaptador de Tramos Endalia")
st.info("Herramienta profesional para inyectar cierres de jornada manteniendo la integridad binaria de la plantilla original.")

# --- BARRA LATERAL: CONFIGURACIÓN ---
st.sidebar.header("⚙️ Configuración Global")
bulk_end_time = st.sidebar.time_input("Hora de Cierre masivo", datetime.time(18, 0))
global_timezone = st.sidebar.text_input("Zona Horaria Exacta", "(UTC+01:00) Bruselas, Copenhague, Madrid, París")
global_overwrite = st.sidebar.selectbox("Sobrescritura por defecto", ["SÍ", "NO"], index=0)

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plantilla de Endalia")
    file_plantilla = st.file_uploader("Sube el Excel original descargado de Endalia", type=["xlsx"], key="plantilla")

with col2:
    st.subheader("2. Registro de Tramos")
    file_registros = st.file_uploader("Sube tu archivo de datos (Excel/CSV)", type=["xlsx", "csv"], key="registros")

if file_plantilla and file_registros:
    try:
        # Lectura de los registros de entrada
        if file_registros.name.endswith('.csv'):
            df_registros = pd.read_csv(file_registros)
        else:
            df_registros = pd.read_excel(file_registros)
        
        def normalize_name(name):
            return str(name).strip().upper() if pd.notnull(name) else ""

        # Detectar tramos que necesitan ser cerrados (sin hora de fin)
        tramos_abiertos = df_registros[
            (df_registros['Hora fin'].isna()) | 
            (df_registros['Hora fin'].astype(str).str.contains("00:00")) |
            (df_registros['Hora fin'].astype(str) == "")
        ].copy()

        st.success(f"Se han detectado {len(tramos_abiertos)} tramos pendientes de cierre.")

        if st.button("🔍 INICIAR INYECCIÓN DE DATOS"):
            # CARGA QUIRÚRGICA: Cargamos el archivo tal cual para no perder metadatos
            wb = load_workbook(file_plantilla, data_only=False)
            
            if "Registros de jornada" not in wb.sheetnames:
                st.error("Error crítico: No se encontró la pestaña 'Registros de jornada' en la plantilla.")
            else:
                ws = wb["Registros de jornada"]
                # Leemos cabeceras para saber en qué columna está cada dato
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
                    st.error(f"La plantilla no tiene el formato esperado. Falta la columna: {e}")
                    st.stop()

                cambios_log = []
                no_encontrados = []

                # Procesar cada tramo abierto
                for _, reg in tramos_abiertos.iterrows():
                    nombre_objetivo = normalize_name(reg['Empleado'])
                    hallado = False
                    
                    # Buscar en la plantilla original fila por fila
                    for r in range(2, ws.max_row + 1):
                        nombre_plantilla = normalize_name(ws.cell(row=r, column=m["emp"]).value)
                        
                        if nombre_plantilla == nombre_objetivo:
                            hora_fin_str = bulk_end_time.strftime("%H:%M")
                            
                            # Modificamos SOLO el valor de las celdas
                            ws.cell(row=r, column=m["fec"]).value = str(reg['Fecha'])
                            ws.cell(row=r, column=m["ini"]).value = str(reg['Hora inicio'])
                            ws.cell(row=r, column=m["fin"]).value = hora_fin_str
                            ws.cell(row=r, column=m["tip"]).value = str(reg.get('Tipo de tramo', 'Trabajo'))
                            ws.cell(row=r, column=m["zon"]).value = global_timezone
                            ws.cell(row=r, column=m["sob"]).value = global_overwrite
                            
                            cambios_log.append({"Empleado": reg['Empleado'], "Fila Excel": r, "Hora Cierre": hora_fin_str})
                            hallado = True
                            break
                    
                    if not hallado:
                        no_encontrados.append(reg['Empleado'])

                # Mostrar resumen y habilitar descarga
                if cambios_log:
                    st.subheader("✅ Inyección completada")
                    st.dataframe(pd.DataFrame(cambios_log))
                    
                    output = BytesIO()
                    wb.save(output) 
                    
                    st.download_button(
                        label="💾 DESCARGAR PLANTILLA LISTA PARA ENDALIA",
                        data=output.getvalue(),
                        file_name=f"Endalia_Cierre_Masivo_{datetime.date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                if no_encontrados:
                    with st.expander("⚠️ Empleados no localizados"):
                        for emp in list(set(no_encontrados)):
                            st.write(f"- {emp}")

    except Exception as e:
        st.error(f"Error técnico durante el procesado: {e}")
else:
    st.info("Por favor, sube ambos archivos para activar el motor de mapeo.")

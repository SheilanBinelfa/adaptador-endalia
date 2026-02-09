import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Pro", page_icon="🎯", layout="wide")

st.title("🎯 Adaptador Endalia: Inyección y Restauración")
st.info("Este motor inyecta los datos y 'dibuja' de nuevo los desplegables para que Excel no los pierda.")

# --- CONFIGURACIÓN DE LAS LISTAS DE LOS DESPLEGABLES ---
OPCIONES_TRAMO = ["Trabajo", "Pausa", "Comida", "Viaje", "Formación"]
OPCIONES_SOBRESCRIBIR = ["SÍ", "NO"]
ZONA_DEFECTO = ["(UTC+01:00) Bruselas, Copenhague, Madrid, París"]

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)
with col1:
    f_plantilla = st.file_uploader("1. Sube la Plantilla de Endalia", type=["xlsx"])
with col2:
    f_datos = st.file_uploader("2. Sube el archivo con los 14 tramos", type=["xlsx", "csv"])

if f_plantilla and f_datos:
    try:
        # Cargar datos
        df_registros = pd.read_excel(f_datos) if f_datos.name.endswith('xlsx') else pd.read_csv(f_datos)
        df_plantilla = pd.read_excel(f_plantilla) 

        if st.button("🚀 PROCESAR E INYECTAR DESPLEGABLES"):
            output = BytesIO()
            
            # Usamos XlsxWriter para poder crear validaciones de datos (desplegables)
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final = df_plantilla.copy()
                
                log_cambios = []
                # Lógica de Match de Empleados
                for idx, row_reg in df_registros.iterrows():
                    emp_buscado = str(row_reg['Empleado']).strip().upper()
                    # Buscamos en la primera columna de la plantilla
                    mask = df_final.iloc[:, 0].astype(str).str.strip().str.upper().str.contains(emp_buscado, na=False)
                    
                    if mask.any():
                        f_idx = df_final[mask].index[0]
                        
                        # Inyección de datos en celdas
                        df_final.at[f_idx, 'Fecha de referencia'] = row_reg['Fecha']
                        df_final.at[f_idx, 'Inicio'] = row_reg['Hora inicio']
                        
                        # Hora fin (si es 00:00 o nula, ponemos 18:00)
                        h_fin = str(row_reg['Hora fin']) if pd.notnull(row_reg['Hora fin']) and "00:00" not in str(row_reg['Hora fin']) else "18:00"
                        df_final.at[f_idx, 'Fin'] = h_fin
                        
                        # Valores para los desplegables
                        df_final.at[f_idx, 'Tipo de tramo'] = "Trabajo"
                        df_final.at[f_idx, 'Zona Horaria'] = ZONA_DEFECTO[0]
                        df_final.at[f_idx, 'Sobrescritura'] = "SÍ"
                        
                        log_cambios.append(f"✅ {emp_buscado}: Inyectado")
                    else:
                        log_cambios.append(f"⚠️ {emp_buscado}: No encontrado")

                # Escribir el DataFrame
                df_final.to_excel(writer, sheet_name='Registros de jornada', index=False)
                
                workbook  = writer.book
                worksheet = writer.sheets['Registros de jornada']

                # --- RECONSTRUCCIÓN DE DESPLEGABLES ---
                # Definimos el rango hasta la fila 500 (puedes ampliarlo)
                
                # 1. Columna Zona Horaria (Columna E -> índice 4)
                worksheet.data_validation('E2:E500', {
                    'validate': 'list',
                    'source': ZONA_DEFECTO
                })

                # 2. Columna Tipo de tramo (Columna H -> índice 7)
                worksheet.data_validation('H2:H500', {
                    'validate': 'list',
                    'source': OPCIONES_TRAMO
                })

                # 3. Columna Sobrescritura (Columna I -> índice 8)
                worksheet.data_validation('I2:I500', {
                    'validate': 'list',
                    'source': OPCIONES_SOBRESCRIBIR
                })

                st.write("### Log de operaciones")
                for item in log_cambios:
                    st.text(item)

            # Botón de descarga
            st.download_button(
                label="📥 Descargar Excel con Desplegables",
                data=output.getvalue(),
                file_name=f"Endalia_Corregido_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error en el proceso: {e}")

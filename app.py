import streamlit as st
import openpyxl
from openpyxl import load_workbook
import pandas as pd
from io import BytesIO
import datetime

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Adaptador Endalia Pro", page_icon="📊", layout="wide")

st.title("🚀 Adaptador Endalia: Inyección con Desplegables")
st.info("Esta versión edita 'dentro' de la plantilla original para asegurar que los menús desplegables y las reglas de validación de Endalia no se pierdan.")

# --- SIDEBAR: PARÁMETROS ---
st.sidebar.header("⚙️ Configuración de Inyección")
bulk_time = st.sidebar.time_input("Hora de Cierre por defecto", datetime.time(18, 0))
timezone_val = st.sidebar.text_input("Zona Horaria", "(UTC+01:00) Bruselas, Copenhague, Madrid, París")
overwrite_val = st.sidebar.selectbox("Sobrescritura", ["SÍ", "NO"], index=0)

# --- CARGA DE ARCHIVOS ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Plantilla de Endalia")
    f_plantilla = st.file_uploader("Sube el Excel ORIGINAL (con desplegables)", type=["xlsx"])

with col2:
    st.subheader("2. Registro de Tramos")
    f_registros = st.file_uploader("Sube el archivo con los 14 tramos", type=["xlsx", "csv"])

if f_plantilla and f_registros:
    try:
        # Cargar registros a importar
        if f_registros.name.endswith('.csv'):
            df_in = pd.read_csv(f_registros)
        else:
            df_in = pd.read_excel(f_registros)
        
        def clean(val):
            return str(val).strip().upper() if pd.notnull(val) else ""

        st.success(f"Se han cargado {len(df_in)} tramos para procesar.")

        if st.button("💉 INYECTAR DATOS SIN ROMPER DESPLEGABLES"):
            # ABRIR ARCHIVO ORIGINAL (keep_vba=True ayuda a mantener la estructura compleja)
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
                except ValueError as e

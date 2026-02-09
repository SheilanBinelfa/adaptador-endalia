import streamlit as st
import openpyxl
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import re
from datetime import datetime

SPREADSHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', SPREADSHEET_NS)
ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
ET.register_namespace('xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
ET.register_namespace('xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2')
ET.register_namespace('xr3', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3')
ET.register_namespace('xr6', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision6')
ET.register_namespace('xr10', 'http://schemas.microsoft.com/office/spreadsheetml/2018/revision10')


def extract_all_validations(file_bytes):
    """Extrae dataValidations de TODAS las hojas del xlsx (cualquier nombre)."""
    validations = {}
    ns = '{' + SPREADSHEET_NS + '}'
    with ZipFile(BytesIO(file_bytes), 'r') as zf:
        for name in zf.namelist():
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                xml_bytes = zf.read(name)
                tree = ET.fromstring(xml_bytes)
                dv_node = tree.find(ns + 'dataValidations')
                if dv_node is not None:
                    validations[name] = ET.tostring(dv_node, encoding='unicode')
    return validations


def extract_ext_lst_x14_validations(file_bytes):
    """
    Extrae validaciones extendidas x14:dataValidations dentro de <extLst>.
    Algunas versiones de Excel usan este formato alternativo para desplegables.
    """
    ext_validations = {}
    with ZipFile(BytesIO(file_bytes), 'r') as zf:
        for name in zf.namelist():
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                raw = zf.read(name).decode('utf-8')
                # Buscar bloque x14:dataValidations dentro del XML crudo
                pattern = r'(<(?:x14|ext):dataValidations[\s\S]*?</(?:x14|ext):dataValidations>)'
                matches = re.findall(pattern, raw)
                if matches:
                    ext_validations[name] = matches
    return ext_validations


def patch_zip_with_validations(output_bytes, original_bytes):
    """
    Estrategia agresiva: para cada hoja, toma el XML generado por openpyxl
    pero REEMPLAZA/INYECTA las secciones de validacion del archivo original.
    Tambien copia archivos que openpyxl pudo haber eliminado.
    """
    ns = '{' + SPREADSHEET_NS + '}'

    # Leer ambos ZIPs
    original_zip = ZipFile(BytesIO(original_bytes), 'r')
    output_zip_in = ZipFile(BytesIO(output_bytes), 'r')

    # Mapear hojas originales: extraer XML completo
    original_sheets = {}
    for name in original_zip.namelist():
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            original_sheets[name] = original_zip.read(name)

    result_buffer = BytesIO()
    result_zip = ZipFile(result_buffer, 'w', ZIP_DEFLATED)

    # Archivos ya procesados
    processed = set()

    for item in output_zip_in.namelist():
        data = output_zip_in.read(item)
        processed.add(item)

        if item in original_sheets:
            # Esta es una hoja - necesitamos inyectar validaciones del original
            try:
                output_tree = ET.fromstring(data)
                original_tree = ET.fromstring(original_sheets[item])

                # 1. Eliminar dataValidations del output (puede estar vacio o corrupto)
                for dv in output_tree.findall(ns + 'dataValidations'):
                    output_tree.remove(dv)

                # 2. Copiar dataValidations del original
                original_dv = original_tree.find(ns + 'dataValidations')
                if original_dv is not None:
                    # Encontrar posicion correcta para insertar
                    insert_before = [
                        ns + 'pageMargins', ns + 'pageSetup', ns + 'headerFooter',
                        ns + 'drawing', ns + 'legacyDrawing', ns + 'tableParts',
                        ns + 'extLst',
                    ]
                    inserted = False
                    for tag in insert_before:
                        target = output_tree.find(tag)
                        if target is not None:
                            idx = list(output_tree).index(target)
                            output_tree.insert(idx, original_dv)
                            inserted = True
                            break
                    if not inserted:
                        output_tree.append(original_dv)

                # 3. Tambien preservar extLst del original (puede tener x14:dataValidations)
                original_ext = original_tree.find(ns + 'extLst')
                if original_ext is not None:
                    # Quitar extLst del output y reemplazar con el original
                    for ext in output_tree.findall(ns + 'extLst'):
                        output_tree.remove(ext)
                    output_tree.append(original_ext)

                data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                data += ET.tostring(output_tree, encoding='unicode').encode('utf-8')

            except ET.ParseError:
                pass  # Si falla el parseo, dejar el archivo como esta

        result_zip.writestr(item, data)

    # Copiar archivos del original que openpyxl pudo haber eliminado
    for item in original_zip.namelist():
        if item not in processed:
            result_zip.writestr(item, original_zip.read(item))

    original_zip.close()
    output_zip_in.close()
    result_zip.close()
    return result_buffer.getvalue()


def find_column_indices(ws, header_row=1):
    columns = {}
    keywords = {
        'nif': ['doc. identificador', 'nif', 'dni', 'identificador', 'documento'],
        'codigo': ['codigo empleado', 'codigo', 'code', 'cod'],
        'nombre': ['nombre', 'empleado', 'name', 'trabajador', 'colaborador'],
        'fecha_ref': ['fecha de referencia', 'referencia', 'fecha ref'],
        'zona': ['zona horaria', 'zona', 'timezone'],
        'inicio': ['inicio', 'start', 'desde'],
        'fin': ['fin', 'end', 'hasta'],
        'tipo_tramo': ['tipo de tramo', 'tipo tramo', 'tramo', 'type'],
        'sobrescritura': ['sobrescritura', 'overwrite', 'sobreescritura'],
    }
    for col_idx in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=header_row, column=col_idx).value
        if cell_value is None:
            continue
        cell_lower = str(cell_value).strip().lower()
        for key, terms in keywords.items():
            for term in terms:
                if term in cell_lower:
                    if key not in columns:
                        columns[key] = col_idx
                    break
    return columns


def extract_records(wb_records, sheet_name=None):
    ws = wb_records[sheet_name] if sheet_name else wb_records.active
    cols = find_column_indices(ws)

    # Necesitamos al menos NIF o codigo para hacer match
    match_key = None
    if 'nif' in cols:
        match_key = 'nif'
    elif 'codigo' in cols:
        match_key = 'codigo'
    elif 'nombre' in cols:
        match_key = 'nombre'

    if match_key is None:
        return None, None, "No se encontro columna identificadora en los registros."

    records = []
    for row in range(2, ws.max_row + 1):
        id_value = ws.cell(row=row, column=cols[match_key]).value
        if id_value is None or str(id_value).strip() == '':
            continue
        record = {'_match_value': str(id_value).strip()}
        for key, col_idx in cols.items():
            val = ws.cell(row=row, column=col_idx).value
            if val is not None:
                record[key] = val
        records.append(record)

    return records, match_key, None


def normalize(val):
    if val is None:
        return ''
    return re.sub(r'\s+', ' ', str(val).strip().lower())


def match_and_update(ws_template, records, match_key, header_row=1):
    cols_template = find_column_indices(ws_template, header_row)

    # Determinar columna de match en la plantilla
    if match_key not in cols_template:
        # Intentar fallback
        for alt in ['nif', 'codigo', 'nombre']:
            if alt in cols_template:
                match_key = alt
                break
        else:
            return 0, [], "No se encontro columna de match en la plantilla."

    # Indice de registros
    rec_index = {}
    for rec in records:
        key = normalize(rec['_match_value'])
        rec_index[key] = rec

    # Columnas actualizables (solo las que tienen datos en registros)
    updatable = ['fecha_ref', 'zona', 'inicio', 'fin', 'tipo_tramo', 'sobrescritura']

    updated = 0
    not_found = []

    for row in range(header_row + 1, ws_template.max_row + 1):
        cell_val = ws_template.cell(row=row, column=cols_template[match_key]).value
        if cell_val is None or str(cell_val).strip() == '':
            continue

        key = normalize(cell_val)
        match = rec_index.get(key)

        if match is None:
            # Matching parcial
            for rk, rv in rec_index.items():
                if rk in key or key in rk:
                    match = rv
                    break

        if match is None:
            not_found.append(str(cell_val).strip())
            continue

        changed = False
        for field in updatable:
            if field in cols_template and field in match:
                val = match[field]
                if val is not None:
                    ws_template.cell(row=row, column=cols_template[field]).value = val
                    changed = True

        if changed:
            updated += 1

    return updated, not_found, None


def main():
    st.set_page_config(page_title="Editor Plantilla Excel", page_icon="📊", layout="wide")
    st.title("📊 Editor de Plantilla Excel")
    st.markdown("**Preserva Validaciones de Datos (Desplegables)** — Actualiza datos sin destruir la estructura.")
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ Plantilla Excel")
        st.caption("Archivo .xlsx con desplegables que quieres preservar.")
        template_file = st.file_uploader("Sube la plantilla", type=['xlsx'], key='template')
    with col2:
        st.subheader("2️⃣ Archivo de Registros")
        st.caption("Archivo con los datos a volcar en la plantilla.")
        records_file = st.file_uploader("Sube los registros", type=['xlsx'], key='records')

    st.divider()

    with st.expander("⚙️ Configuracion avanzada"):
        header_row = st.number_input("Fila de encabezados", min_value=1, value=1)
        template_sheet = st.text_input("Hoja de la plantilla (vacio = primera hoja con datos)", value="Registros de jornada")
        records_sheet = st.text_input("Hoja de registros (vacio = hoja activa)", value="")

    if template_file and records_file:
        if st.button("🚀 Procesar y Generar Archivo", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    template_bytes = template_file.read()
                    records_bytes = records_file.read()

                    # === PASO 1: Diagnostico de validaciones originales ===
                    original_validations = extract_all_validations(template_bytes)
                    st.info(f"🔍 Hojas con validaciones en el original: **{len(original_validations)}**")

                    if original_validations:
                        with st.expander("📋 XML de validaciones originales"):
                            for path, xml_str in original_validations.items():
                                st.write(f"**{path}**")
                                st.code(xml_str[:3000], language='xml')

                    # === PASO 2: Cargar con openpyxl ===
                    wb_template = openpyxl.load_workbook(BytesIO(template_bytes), data_only=False)
                    wb_records = openpyxl.load_workbook(BytesIO(records_bytes), data_only=True)

                    # Seleccionar hoja de plantilla
                    if template_sheet and template_sheet in wb_template.sheetnames:
                        ws_template = wb_template[template_sheet]
                    else:
                        ws_template = wb_template.active

                    st.write(f"📄 Hoja de plantilla: **{ws_template.title}**")
                    st.write(f"📄 Hojas disponibles: {wb_template.sheetnames}")

                    # Mostrar columnas detectadas
                    cols = find_column_indices(ws_template, header_row)
                    with st.expander("🔎 Columnas detectadas en la plantilla"):
                        for key, idx in cols.items():
                            header = ws_template.cell(row=header_row, column=idx).value
                            st.write(f"**{key}** → Columna {idx} ({header})")

                    # === PASO 3: Extraer registros ===
                    records, match_key, error = extract_records(
                        wb_records, records_sheet if records_sheet else None
                    )
                    if error:
                        st.error(f"❌ {error}")
                        return

                    st.write(f"👥 Registros encontrados: **{len(records)}** (match por: **{match_key}**)")

                    with st.expander("👀 Vista previa de registros (primeros 5)"):
                        for rec in records[:5]:
                            st.json(rec)

                    # === PASO 4: Match y actualizacion (solo .value) ===
                    updated, not_found, err = match_and_update(ws_template, records, match_key, header_row)

                    if err:
                        st.error(f"❌ {err}")
                        return

                    st.success(f"✅ **{updated}** filas actualizadas en la plantilla.")

                    if not_found:
                        with st.expander(f"⚠️ {len(not_found)} sin match"):
                            for name in not_found[:50]:
                                st.write(f"- {name}")

                    # === PASO 5: Guardar con openpyxl ===
                    output_buffer = BytesIO()
                    wb_template.save(output_buffer)
                    output_bytes = output_buffer.getvalue()

                    # === PASO 6: PARCHE XML — re-inyectar validaciones del original ===
                    st.info("🔧 Parcheando XML: re-inyectando validaciones del archivo original...")
                    output_bytes = patch_zip_with_validations(output_bytes, template_bytes)

                    # === PASO 7: Verificacion ===
                    final_validations = extract_all_validations(output_bytes)
                    if final_validations:
                        st.success(f"🎯 **Verificacion OK**: {len(final_validations)} hoja(s) con validaciones en el archivo final.")
                        with st.expander("📋 XML de validaciones en archivo final"):
                            for path, xml_str in final_validations.items():
                                st.write(f"**{path}**")
                                st.code(xml_str[:3000], language='xml')
                    else:
                        st.warning("⚠️ No se detectaron validaciones en el archivo final.")

                    # === PASO 8: Descarga ===
                    st.divider()
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"plantilla_actualizada_{timestamp}.xlsx"

                    st.download_button(
                        label="📥 Descargar Archivo Actualizado",
                        data=output_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet",
                        type="primary",
                        use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    st.exception(e)

    st.divider()
    with st.expander("📖 Como funciona"):
        st.markdown("""
### Estrategia de preservacion

1. **Backup completo**: Antes de tocar nada, se lee el ZIP original y se extraen
   los nodos XML `<dataValidations>` y `<extLst>` de TODAS las hojas.

2. **Edicion minima**: openpyxl solo modifica `.value` de celdas especificas.
   No se tocan formatos, estilos ni estructura.

3. **Parche XML post-guardado**: Despues de que openpyxl guarda, se abre el ZIP
   resultante y se REEMPLAZA el XML de validaciones con el del archivo original.
   Tambien se preserva `<extLst>` que puede contener validaciones x14.

4. **Verificacion**: Se re-lee el archivo final para confirmar que las validaciones existen.

### Columnas reconocidas
- Nº doc. Identificador / Codigo empleado (para match)
- Fecha de referencia, Zona Horaria, Inicio, Fin, Tipo de tramo, Sobrescritura
        """)


main()

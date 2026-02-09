import streamlit as st
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
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


def extract_validations_from_zip(file_bytes):
    validations = {}
    ns = '{' + SPREADSHEET_NS + '}'
    with ZipFile(BytesIO(file_bytes), 'r') as zf:
        for name in zf.namelist():
            if re.match(r'xl/worksheets/sheet\d+\.xml', name):
                xml_bytes = zf.read(name)
                tree = ET.fromstring(xml_bytes)
                dv_node = tree.find(ns + 'dataValidations')
                if dv_node is not None:
                    validations[name] = ET.tostring(dv_node, encoding='unicode')
    return validations


def reinject_validations_into_zip(output_bytes, original_validations):
    ns = '{' + SPREADSHEET_NS + '}'
    input_zip = ZipFile(BytesIO(output_bytes), 'r')
    output_buffer = BytesIO()
    output_zip = ZipFile(output_buffer, 'w', ZIP_DEFLATED)

    for item in input_zip.namelist():
        data = input_zip.read(item)
        if item in original_validations:
            tree = ET.fromstring(data)
            existing = tree.find(ns + 'dataValidations')
            if existing is not None:
                tree.remove(existing)
            original_dv = ET.fromstring(original_validations[item])
            insert_before_tags = [
                ns + 'pageMargins',
                ns + 'pageSetup',
                ns + 'headerFooter',
                ns + 'drawing',
                ns + 'legacyDrawing',
                ns + 'tableParts',
                ns + 'extLst',
            ]
            inserted = False
            for tag in insert_before_tags:
                target = tree.find(tag)
                if target is not None:
                    idx = list(tree).index(target)
                    tree.insert(idx, original_dv)
                    inserted = True
                    break
            if not inserted:
                tree.append(original_dv)
            data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            data += ET.tostring(tree, encoding='unicode').encode('utf-8')
        output_zip.writestr(item, data)

    input_zip.close()
    output_zip.close()
    return output_buffer.getvalue()


def snapshot_openpyxl_validations(ws):
    snapshot = []
    if hasattr(ws, 'data_validations') and ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            snapshot.append({
                'type': dv.type,
                'formula1': dv.formula1,
                'formula2': dv.formula2,
                'allow_blank': dv.allow_blank,
                'showErrorMessage': dv.showErrorMessage,
                'showInputMessage': dv.showInputMessage,
                'errorTitle': dv.errorTitle,
                'error': dv.error,
                'promptTitle': dv.promptTitle,
                'prompt': dv.prompt,
                'sqref': str(dv.sqref),
                'showDropDown': dv.showDropDown,
            })
    return snapshot


def restore_openpyxl_validations(ws, snapshot):
    if not snapshot:
        return
    ws.data_validations.dataValidation = []
    for item in snapshot:
        dv = DataValidation(
            type=item['type'],
            formula1=item['formula1'],
            formula2=item['formula2'],
            allow_blank=item['allow_blank'],
            showErrorMessage=item['showErrorMessage'],
            showInputMessage=item['showInputMessage'],
            errorTitle=item['errorTitle'],
            error=item['error'],
            promptTitle=item['promptTitle'],
            prompt=item['prompt'],
            showDropDown=item['showDropDown'],
        )
        dv.sqref = item['sqref']
        ws.add_data_validation(dv)


def find_column_indices(ws, header_row=1):
    columns = {}
    keywords = {
        'nombre': ['nombre', 'empleado', 'name', 'trabajador', 'colaborador'],
        'inicio': ['inicio', 'fecha inicio', 'start', 'desde', 'begin'],
        'fin': ['fin', 'fecha fin', 'end', 'hasta', 'término', 'termino'],
        'fecha': ['fecha', 'date', 'día', 'dia'],
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


def extract_employees_from_records(wb_records, sheet_name=None):
    ws = wb_records[sheet_name] if sheet_name else wb_records.active
    cols = find_column_indices(ws)
    if 'nombre' not in cols:
        return None, "No se encontro columna de Nombre/Empleado en el archivo de registros."
    employees = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=cols['nombre']).value
        if name is None or str(name).strip() == '':
            continue
        record = {'nombre': str(name).strip()}
        if 'inicio' in cols:
            record['inicio'] = ws.cell(row=row, column=cols['inicio']).value
        if 'fin' in cols:
            record['fin'] = ws.cell(row=row, column=cols['fin']).value
        if 'fecha' in cols:
            record['fecha'] = ws.cell(row=row, column=cols['fecha']).value
        employees.append(record)
    return employees, None


def normalize_name(name):
    if name is None:
        return ''
    return re.sub(r'\s+', ' ', str(name).strip().lower())


def match_and_update(ws_template, employees, header_row=1):
    cols = find_column_indices(ws_template, header_row)
    if 'nombre' not in cols:
        return 0, "No se encontro columna de nombre en la plantilla."
    emp_index = {}
    for emp in employees:
        key = normalize_name(emp['nombre'])
        emp_index[key] = emp
    updated = 0
    not_found = []
    for row in range(header_row + 1, ws_template.max_row + 1):
        cell_name = ws_template.cell(row=row, column=cols['nombre']).value
        if cell_name is None or str(cell_name).strip() == '':
            continue
        key = normalize_name(cell_name)
        match = emp_index.get(key)
        if match is None:
            for emp_key, emp_data in emp_index.items():
                if emp_key in key or key in emp_key:
                    match = emp_data
                    break
        if match is None:
            not_found.append(str(cell_name).strip())
            continue
        changed = False
        if 'inicio' in cols and 'inicio' in match and match['inicio'] is not None:
            ws_template.cell(row=row, column=cols['inicio']).value = match['inicio']
            changed = True
        if 'fin' in cols and 'fin' in match and match['fin'] is not None:
            ws_template.cell(row=row, column=cols['fin']).value = match['fin']
            changed = True
        if 'fecha' in cols and 'fecha' in match and match['fecha'] is not None:
            ws_template.cell(row=row, column=cols['fecha']).value = match['fecha']
            changed = True
        if changed:
            updated += 1
    return updated, not_found


def main():
    st.set_page_config(page_title="Editor de Plantilla Excel", page_icon="📊", layout="wide")
    st.title("📊 Editor de Plantilla Excel")
    st.markdown("**Preserva Validaciones de Datos (Desplegables)** — Actualiza fechas de empleados sin destruir la estructura de la plantilla.")
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ Plantilla Excel")
        st.caption("El archivo .xlsx con desplegables/validaciones que quieres preservar.")
        template_file = st.file_uploader("Sube la plantilla", type=['xlsx'], key='template')
    with col2:
        st.subheader("2️⃣ Archivo de Registros")
        st.caption("El archivo con los datos actualizados de los empleados.")
        records_file = st.file_uploader("Sube los registros", type=['xlsx'], key='records')

    st.divider()

    with st.expander("⚙️ Configuracion avanzada"):
        header_row = st.number_input("Fila de encabezados en la plantilla", min_value=1, value=1)
        template_sheet = st.text_input("Nombre de hoja en plantilla (vacio = hoja activa)", value="")
        records_sheet = st.text_input("Nombre de hoja en registros (vacio = hoja activa)", value="")
        use_xml_fallback = st.checkbox("Usar re-inyeccion XML como fallback (recomendado)", value=True)

    if template_file and records_file:
        if st.button("🚀 Procesar y Generar Archivo", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    template_bytes = template_file.read()
                    records_bytes = records_file.read()

                    original_validations = extract_validations_from_zip(template_bytes)
                    st.info(f"🔍 Validaciones detectadas en **{len(original_validations)}** hoja(s) del archivo original.")

                    if original_validations:
                        with st.expander("📋 Detalle de validaciones encontradas"):
                            for sheet_path, xml_str in original_validations.items():
                                st.code(xml_str[:2000], language='xml')

                    wb_template = openpyxl.load_workbook(BytesIO(template_bytes), data_only=False)
                    wb_records = openpyxl.load_workbook(BytesIO(records_bytes), data_only=True)

                    if template_sheet and template_sheet in wb_template.sheetnames:
                        ws_template = wb_template[template_sheet]
                    else:
                        ws_template = wb_template.active

                    dv_snapshot = snapshot_openpyxl_validations(ws_template)
                    st.write(f"📌 Validaciones capturadas por openpyxl: **{len(dv_snapshot)}**")

                    employees, error = extract_employees_from_records(wb_records, records_sheet if records_sheet else None)
                    if error:
                        st.error(f"❌ Error en registros: {error}")
                        return

                    st.write(f"👥 Empleados encontrados en registros: **{len(employees)}**")

                    with st.expander("👀 Vista previa de registros"):
                        for emp in employees[:10]:
                            st.write(emp)
                        if len(employees) > 10:
                            st.caption(f"... y {len(employees) - 10} mas")

                    updated, not_found = match_and_update(ws_template, employees, header_row)
                    st.success(f"✅ **{updated}** empleados actualizados en la plantilla.")

                    if not_found:
                        with st.expander(f"⚠️ {len(not_found)} empleados sin match"):
                            for name in not_found:
                                st.write(f"- {name}")

                    restore_openpyxl_validations(ws_template, dv_snapshot)

                    output_buffer = BytesIO()
                    wb_template.save(output_buffer)
                    output_bytes = output_buffer.getvalue()

                    if use_xml_fallback and original_validations:
                        st.info("🔧 Aplicando re-inyeccion XML de validaciones...")
                        output_bytes = reinject_validations_into_zip(output_bytes, original_validations)
                        st.success("✅ Validaciones re-inyectadas exitosamente en el XML.")

                    final_validations = extract_validations_from_zip(output_bytes)
                    if final_validations:
                        st.success(f"🎯 **Verificacion final**: {len(final_validations)} hoja(s) con validaciones intactas.")
                    else:
                        st.warning("⚠️ No se detectaron validaciones en el archivo final.")

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
                    st.error(f"❌ Error durante el procesamiento: {str(e)}")
                    st.exception(e)


main()

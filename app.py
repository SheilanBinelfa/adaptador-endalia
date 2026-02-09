import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import re
from datetime import datetime, time, date
from collections import defaultdict
from copy import copy

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

# ─────────────────────────────────────────────────
# XML / ZIP — Preservar validaciones
# ─────────────────────────────────────────────────

def extract_all_validations(file_bytes):
    validations = {}
    ns = '{' + SPREADSHEET_NS + '}'
    with ZipFile(BytesIO(file_bytes), 'r') as zf:
        for name in zf.namelist():
            if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                tree = ET.fromstring(zf.read(name))
                dv_node = tree.find(ns + 'dataValidations')
                if dv_node is not None:
                    validations[name] = ET.tostring(dv_node, encoding='unicode')
    return validations


def patch_zip_with_validations(output_bytes, original_bytes):
    ns = '{' + SPREADSHEET_NS + '}'
    original_zip = ZipFile(BytesIO(original_bytes), 'r')
    output_zip_in = ZipFile(BytesIO(output_bytes), 'r')
    original_sheets = {}
    for name in original_zip.namelist():
        if name.startswith('xl/worksheets/') and name.endswith('.xml'):
            original_sheets[name] = original_zip.read(name)
    result_buffer = BytesIO()
    result_zip = ZipFile(result_buffer, 'w', ZIP_DEFLATED)
    processed = set()
    for item in output_zip_in.namelist():
        data = output_zip_in.read(item)
        processed.add(item)
        if item in original_sheets:
            try:
                output_tree = ET.fromstring(data)
                original_tree = ET.fromstring(original_sheets[item])
                # Quitar dataValidations del output
                for dv in output_tree.findall(ns + 'dataValidations'):
                    output_tree.remove(dv)
                # Copiar del original
                original_dv = original_tree.find(ns + 'dataValidations')
                if original_dv is not None:
                    # Actualizar sqref para cubrir nuevas filas
                    output_tree_sd = output_tree.find(ns + 'sheetData')
                    if output_tree_sd is not None:
                        all_rows = output_tree_sd.findall(ns + 'row')
                        if all_rows:
                            max_row = max(int(r.get('r', '1')) for r in all_rows)
                            # Expandir rangos de validacion
                            for dv_item in original_dv.findall(ns + 'dataValidation'):
                                sqref = dv_item.get('sqref', '')
                                new_sqref = expand_sqref(sqref, max_row)
                                dv_item.set('sqref', new_sqref)
                    insert_before = [
                        ns + 'pageMargins', ns + 'pageSetup', ns + 'headerFooter',
                        ns + 'drawing', ns + 'legacyDrawing', ns + 'tableParts', ns + 'extLst',
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
                # Preservar extLst
                original_ext = original_tree.find(ns + 'extLst')
                if original_ext is not None:
                    for ext in output_tree.findall(ns + 'extLst'):
                        output_tree.remove(ext)
                    output_tree.append(original_ext)
                data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                data += ET.tostring(output_tree, encoding='unicode').encode('utf-8')
            except ET.ParseError:
                pass
        result_zip.writestr(item, data)
    for item in original_zip.namelist():
        if item not in processed:
            result_zip.writestr(item, original_zip.read(item))
    original_zip.close()
    output_zip_in.close()
    result_zip.close()
    return result_buffer.getvalue()


def expand_sqref(sqref, max_row):
    """Expande rangos como D2:D500 a D2:D{max_row} si max_row es mayor."""
    parts = sqref.split(' ')
    new_parts = []
    for part in parts:
        if ':' in part:
            start, end = part.split(':')
            col_end = re.match(r'([A-Z]+)', end)
            row_end = re.search(r'(\d+)', end)
            if col_end and row_end:
                old_row = int(row_end.group(1))
                if max_row > old_row:
                    new_parts.append(f"{start}:{col_end.group(1)}{max_row}")
                else:
                    new_parts.append(part)
            else:
                new_parts.append(part)
        else:
            new_parts.append(part)
    return ' '.join(new_parts)


# ─────────────────────────────────────────────────
# Lectura del archivo de Registros de Tramos
# ─────────────────────────────────────────────────

def normalize(val):
    if val is None:
        return ''
    return re.sub(r'\s+', ' ', str(val).strip().lower())


def read_tramos(wb):
    ws = wb.active
    # Detectar columnas
    col_map = {}
    keywords = {
        'fecha': ['fecha'],
        'hora_inicio': ['hora inicio', 'inicio'],
        'hora_fin': ['hora fin', 'fin'],
        'duracion': ['duracion', 'duración'],
        'tipo_tramo': ['tipo de tramo', 'tipo tramo', 'tipo de tra'],
        'empleado': ['empleado'],
        'validado': ['validado'],
        'timezone_name': ['timezonename', 'timezone'],
        'timezone_offset': ['timezoneoffset', 'offset'],
    }
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val is None:
            continue
        val_lower = str(val).strip().lower()
        for key, terms in keywords.items():
            for term in terms:
                if term in val_lower and key not in col_map:
                    col_map[key] = col
                    break

    if 'empleado' not in col_map:
        return None, col_map, "No se encontro columna 'Empleado' en los registros."

    tramos = []
    for row in range(2, ws.max_row + 1):
        emp = ws.cell(row=row, column=col_map['empleado']).value
        if emp is None or str(emp).strip() == '':
            continue
        tramo = {
            'row': row,
            'empleado': str(emp).strip(),
        }
        if 'fecha' in col_map:
            tramo['fecha'] = ws.cell(row=row, column=col_map['fecha']).value
        if 'hora_inicio' in col_map:
            tramo['hora_inicio'] = ws.cell(row=row, column=col_map['hora_inicio']).value
        if 'hora_fin' in col_map:
            tramo['hora_fin'] = ws.cell(row=row, column=col_map['hora_fin']).value
        if 'tipo_tramo' in col_map:
            tramo['tipo_tramo'] = ws.cell(row=row, column=col_map['tipo_tramo']).value
        if 'timezone_name' in col_map:
            tramo['timezone'] = ws.cell(row=row, column=col_map['timezone_name']).value

        tramos.append(tramo)

    return tramos, col_map, None


def is_missing_hora_fin(hora_fin):
    """Retorna True si hora_fin esta vacia o es 00:00."""
    if hora_fin is None:
        return True
    if isinstance(hora_fin, time):
        return hora_fin == time(0, 0)
    if isinstance(hora_fin, datetime):
        return hora_fin.hour == 0 and hora_fin.minute == 0
    s = str(hora_fin).strip()
    return s == '' or s == '00:00' or s == '0:00' or s == '00:00:00'


def format_time_for_display(val):
    """Formatea un valor de tiempo para mostrar en la interfaz."""
    if val is None:
        return ''
    if isinstance(val, time):
        return val.strftime('%H:%M')
    if isinstance(val, datetime):
        return val.strftime('%H:%M')
    return str(val)


def format_date_for_display(val):
    if val is None:
        return ''
    if isinstance(val, (datetime, date)):
        return val.strftime('%d/%m/%Y')
    return str(val)


# ─────────────────────────────────────────────────
# Escritura en la Plantilla Endalia
# ─────────────────────────────────────────────────

def find_plantilla_columns(ws, header_row=1):
    cols = {}
    keywords = {
        'nif': ['doc. identificador', 'nif', 'dni', 'identificador'],
        'codigo': ['codigo empleado', 'código empleado', 'codigo', 'código'],
        'empleado': ['empleado'],
        'fecha_ref': ['fecha de referencia', 'fecha ref'],
        'zona': ['zona horaria'],
        'inicio': ['inicio'],
        'fin': ['fin'],
        'tipo_tramo': ['tipo de tramo'],
        'sobrescritura': ['sobrescritura'],
    }
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        val_lower = str(val).strip().lower()
        for key, terms in keywords.items():
            for term in terms:
                if term in val_lower and key not in cols:
                    cols[key] = col
                    break
    return cols


def copy_cell_style(source_cell, target_cell):
    """Copia formato de una celda a otra."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = source_cell.number_format
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)


def build_and_write_plantilla(wb_template, tramos_to_inject, template_sheet_name, header_row=1):
    """
    Reconstruye la hoja de la plantilla:
    - Empleados con match: tantas filas como tramos tengan
    - Empleados sin match: se eliminan
    """
    if template_sheet_name and template_sheet_name in wb_template.sheetnames:
        ws = wb_template[template_sheet_name]
    else:
        ws = wb_template.active

    cols = find_plantilla_columns(ws, header_row)

    if 'empleado' not in cols:
        return 0, 0, "No se encontro columna 'Empleado' en la plantilla."

    # Leer empleados de la plantilla
    plantilla_employees = []
    for row in range(header_row + 1, ws.max_row + 1):
        emp_name = ws.cell(row=row, column=cols['empleado']).value
        if emp_name is None or str(emp_name).strip() == '':
            continue
        nif = ws.cell(row=row, column=cols.get('nif', 1)).value
        codigo = ws.cell(row=row, column=cols.get('codigo', 2)).value
        plantilla_employees.append({
            'row': row,
            'nombre': str(emp_name).strip(),
            'nombre_norm': normalize(emp_name),
            'nif': nif,
            'codigo': codigo,
        })

    # Agrupar tramos por empleado normalizado
    tramos_by_emp = defaultdict(list)
    for t in tramos_to_inject:
        key = normalize(t['empleado'])
        tramos_by_emp[key] = tramos_by_emp.get(key, [])
        tramos_by_emp[key].append(t)

    # Hacer match
    matched_data = []  # Lista de (nif, codigo, nombre, tramo_data)
    matched_employees = set()
    unmatched_tramos = []

    for emp_norm, emp_tramos in tramos_by_emp.items():
        # Buscar en plantilla
        found = None
        for pe in plantilla_employees:
            if pe['nombre_norm'] == emp_norm:
                found = pe
                break
        if found is None:
            # Match parcial
            for pe in plantilla_employees:
                if emp_norm in pe['nombre_norm'] or pe['nombre_norm'] in emp_norm:
                    found = pe
                    break

        if found is None:
            for t in emp_tramos:
                unmatched_tramos.append(t['empleado'])
            continue

        matched_employees.add(found['nombre_norm'])
        for t in emp_tramos:
            matched_data.append({
                'nif': found['nif'],
                'codigo': found['codigo'],
                'nombre': found['nombre'],
                'fecha_ref': t.get('fecha'),
                'zona': t.get('zona', '(UTC+01:00) Bruselas, Copenhague, Madrid, París'),
                'inicio': t.get('hora_inicio'),
                'fin': t.get('hora_fin'),
                'tipo_tramo': t.get('tipo_tramo'),
                'sobrescritura': t.get('sobrescritura', 'Sí'),
            })

    # Empleados sin match en tramos -> se eliminan de la plantilla
    removed = 0
    for pe in plantilla_employees:
        if pe['nombre_norm'] not in matched_employees:
            removed += 1

    # Guardar estilos de la fila 2 (primera fila de datos) para copiarlos
    style_row = header_row + 1
    styles = {}
    for col_idx in range(1, ws.max_column + 1):
        styles[col_idx] = ws.cell(row=style_row, column=col_idx)

    # Limpiar todas las filas de datos existentes
    for row in range(header_row + 1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            ws.cell(row=row, column=col).value = None

    # Escribir las filas con match
    current_row = header_row + 1
    for item in matched_data:
        # Copiar estilo de la fila modelo
        for col_idx in range(1, ws.max_column + 1):
            copy_cell_style(styles[col_idx], ws.cell(row=current_row, column=col_idx))

        # Escribir datos
        if 'nif' in cols:
            ws.cell(row=current_row, column=cols['nif']).value = item['nif']
        if 'codigo' in cols:
            ws.cell(row=current_row, column=cols['codigo']).value = item['codigo']
        if 'empleado' in cols:
            ws.cell(row=current_row, column=cols['empleado']).value = item['nombre']
        if 'fecha_ref' in cols:
            ws.cell(row=current_row, column=cols['fecha_ref']).value = item['fecha_ref']
        if 'zona' in cols:
            ws.cell(row=current_row, column=cols['zona']).value = item['zona']
        if 'inicio' in cols:
            ws.cell(row=current_row, column=cols['inicio']).value = item['inicio']
        if 'fin' in cols:
            ws.cell(row=current_row, column=cols['fin']).value = item['fin']
        if 'tipo_tramo' in cols:
            ws.cell(row=current_row, column=cols['tipo_tramo']).value = item['tipo_tramo']
        if 'sobrescritura' in cols:
            ws.cell(row=current_row, column=cols['sobrescritura']).value = item['sobrescritura']

        current_row += 1

    return len(matched_data), removed, unmatched_tramos


# ─────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Conciliador Endalia", page_icon="📊", layout="wide")
    st.title("📊 Conciliador de Tramos → Plantilla Endalia")
    st.markdown("Importa tramos de fichaje a la plantilla Endalia preservando desplegables y formato.")
    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1️⃣ Registros de Tramos")
        st.caption("Excel con los fichajes (Fecha, Hora inicio, Hora fin, Empleado...)")
        tramos_file = st.file_uploader("Sube registros de tramos", type=['xlsx'], key='tramos')
    with col2:
        st.subheader("2️⃣ Plantilla Endalia")
        st.caption("Plantilla .xlsx con desplegables a preservar.")
        plantilla_file = st.file_uploader("Sube la plantilla", type=['xlsx'], key='plantilla')

    st.divider()

    with st.expander("⚙️ Configuracion"):
        header_row = st.number_input("Fila de encabezados en la plantilla", min_value=1, value=1)
        template_sheet = st.text_input("Hoja de la plantilla", value="Registros de jornada")
        zona_default = st.selectbox("Zona horaria por defecto", [
            "(UTC+01:00) Bruselas, Copenhague, Madrid, París",
            "(UTC+00:00) Dublín, Edimburgo, Lisboa, Londres",
            "(UTC+02:00) Atenas, Bucarest, Estambul",
            "(UTC-05:00) Hora del este (EE.UU. y Canadá)",
            "(UTC-06:00) Hora central (EE.UU. y Canadá)",
        ])
        sobrescritura_default = st.selectbox("Sobrescritura por defecto", ["Sí", "No"])

    if not tramos_file or not plantilla_file:
        st.info("👆 Sube ambos archivos para continuar.")
        return

    # ═══════════════════════════════════════════
    # PASO 1: Leer tramos
    # ═══════════════════════════════════════════
    tramos_bytes = tramos_file.read()
    wb_tramos = openpyxl.load_workbook(BytesIO(tramos_bytes), data_only=True)
    tramos, tramos_cols, error = read_tramos(wb_tramos)

    if error:
        st.error(f"❌ {error}")
        return

    st.success(f"✅ {len(tramos)} tramos leidos del archivo de registros.")

    # ═══════════════════════════════════════════
    # PASO 2: Detectar tramos sin Hora Fin
    # ═══════════════════════════════════════════
    tramos_sin_fin = [t for t in tramos if is_missing_hora_fin(t.get('hora_fin'))]
    tramos_con_fin = [t for t in tramos if not is_missing_hora_fin(t.get('hora_fin'))]

    if tramos_sin_fin:
        st.warning(f"⚠️ **{len(tramos_sin_fin)}** tramos sin Hora Fin detectados. Introduce la hora fin para cada uno:")

        st.markdown("---")
        hora_fin_inputs = {}

        for i, t in enumerate(tramos_sin_fin):
            col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 2])
            with col_a:
                st.write(f"**{t['empleado']}**")
            with col_b:
                st.write(f"📅 {format_date_for_display(t.get('fecha'))}")
            with col_c:
                st.write(f"🕐 Inicio: {format_time_for_display(t.get('hora_inicio'))}")
            with col_d:
                hora_fin_inputs[i] = st.time_input(
                    f"Hora Fin",
                    value=time(17, 0),
                    key=f"hora_fin_{i}",
                )
        st.markdown("---")

        # Aplicar las horas fin introducidas
        for i, t in enumerate(tramos_sin_fin):
            t['hora_fin'] = hora_fin_inputs[i]

        # Juntar todos los tramos
        all_tramos = tramos_con_fin + tramos_sin_fin
    else:
        all_tramos = tramos
        st.info("✅ Todos los tramos tienen Hora Fin.")

    # Aplicar valores por defecto
    for t in all_tramos:
        if not t.get('zona'):
            t['zona'] = zona_default
        t['sobrescritura'] = sobrescritura_default

    # ═══════════════════════════════════════════
    # PASO 3: Preview
    # ═══════════════════════════════════════════
    with st.expander(f"👀 Vista previa de {len(all_tramos)} tramos a procesar"):
        for t in all_tramos[:20]:
            st.write(
                f"**{t['empleado']}** | "
                f"{format_date_for_display(t.get('fecha'))} | "
                f"{format_time_for_display(t.get('hora_inicio'))} → "
                f"{format_time_for_display(t.get('hora_fin'))} | "
                f"{t.get('tipo_tramo', '-')}"
            )
        if len(all_tramos) > 20:
            st.caption(f"... y {len(all_tramos) - 20} mas")

    # ═══════════════════════════════════════════
    # PASO 4: Procesar
    # ═══════════════════════════════════════════
    if st.button("🚀 Generar Plantilla", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                plantilla_bytes = plantilla_file.read()

                # Backup de validaciones XML
                original_validations = extract_all_validations(plantilla_bytes)
                st.info(f"🔍 {len(original_validations)} hoja(s) con validaciones detectadas.")

                # Cargar plantilla
                wb_template = openpyxl.load_workbook(BytesIO(plantilla_bytes), data_only=False)

                # Escribir datos
                written, removed, unmatched = build_and_write_plantilla(
                    wb_template, all_tramos, template_sheet, header_row
                )

                st.success(f"✅ **{written}** filas escritas en la plantilla.")
                if removed:
                    st.info(f"🗑️ **{removed}** empleados sin tramos eliminados de la plantilla.")
                if unmatched:
                    unique_unmatched = list(set(unmatched))
                    with st.expander(f"⚠️ {len(unique_unmatched)} empleados de tramos sin match en plantilla"):
                        for name in sorted(unique_unmatched):
                            st.write(f"- {name}")

                # Guardar
                output_buffer = BytesIO()
                wb_template.save(output_buffer)
                output_bytes = output_buffer.getvalue()

                # Parche XML para preservar validaciones
                st.info("🔧 Re-inyectando validaciones XML...")
                output_bytes = patch_zip_with_validations(output_bytes, plantilla_bytes)

                # Verificacion
                final_vals = extract_all_validations(output_bytes)
                if final_vals:
                    st.success(f"🎯 Verificacion OK: {len(final_vals)} hoja(s) con desplegables intactos.")
                else:
                    st.warning("⚠️ No se detectaron validaciones en el archivo final.")

                # Descarga
                st.divider()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📥 Descargar Plantilla Completada",
                    data=output_bytes,
                    file_name=f"endalia_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)


main()

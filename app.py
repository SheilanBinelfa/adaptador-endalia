import streamlit as st
import openpyxl
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import re
from datetime import datetime, time, date
from copy import copy
import unicodedata

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

ZONA_DEFAULT = "(UTC+01:00) Bruselas, Copenhague, Madrid, París"
SOBRESCRITURA_DEFAULT = "Sí"
TEMPLATE_SHEET = "Registros de jornada"
HEADER_ROW = 1


# ═══════════════════════════════════════════════════
# XML / ZIP — Preservar validaciones
# ═══════════════════════════════════════════════════

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


def expand_sqref(sqref, max_row):
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
                for dv in output_tree.findall(ns + 'dataValidations'):
                    output_tree.remove(dv)
                original_dv = original_tree.find(ns + 'dataValidations')
                if original_dv is not None:
                    output_sd = output_tree.find(ns + 'sheetData')
                    if output_sd is not None:
                        all_rows = output_sd.findall(ns + 'row')
                        if all_rows:
                            max_row = max(int(r.get('r', '1')) for r in all_rows)
                            for dv_item in original_dv.findall(ns + 'dataValidation'):
                                sqref = dv_item.get('sqref', '')
                                dv_item.set('sqref', expand_sqref(sqref, max_row))
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


# ═══════════════════════════════════════════════════
# Utilidades
# ═══════════════════════════════════════════════════

def remove_accents(text):
    nfkd = unicodedata.normalize('NFKD', text)
    return ''.join(c for c in nfkd if not unicodedata.category(c).startswith('M'))


def normalize_name(name):
    if name is None:
        return ''
    s = str(name).strip().lower()
    s = remove_accents(s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s


def extract_time_part(val):
    if val is None:
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    s = str(val).strip()
    m = re.match(r'(\d{1,2}):(\d{2})', s)
    if m:
        return time(int(m.group(1)), int(m.group(2)))
    return None


def extract_date_part(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date):
        return val
    s = str(val).strip()
    for fmt in ['%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y']:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def combine_date_time(fecha_val, hora_val):
    d = extract_date_part(fecha_val)
    t = extract_time_part(hora_val)
    if d is None or t is None:
        return None
    return datetime.combine(d, t)


def fmt_time(val):
    if val is None:
        return ''
    if isinstance(val, time):
        return val.strftime('%H:%M')
    if isinstance(val, datetime):
        return val.strftime('%H:%M')
    return str(val)


def fmt_date(val):
    if val is None:
        return ''
    if isinstance(val, (datetime, date)):
        return val.strftime('%d/%m/%Y')
    return str(val)


def is_missing_hora_fin(hora_fin):
    if hora_fin is None:
        return True
    if isinstance(hora_fin, time):
        return hora_fin == time(0, 0)
    if isinstance(hora_fin, datetime):
        return hora_fin.hour == 0 and hora_fin.minute == 0
    s = str(hora_fin).strip()
    return s == '' or s in ('00:00', '0:00', '00:00:00')


def read_tramos(wb):
    ws = wb.active
    col_map = {}
    keywords = {
        'fecha': ['fecha'],
        'hora_inicio': ['hora inicio'],
        'hora_fin': ['hora fin'],
        'tipo_tramo': ['tipo de tramo', 'tipo de tra', 'tipo tramo'],
        'empleado': ['empleado'],
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
        return None, "No se encontró la columna 'Empleado' en el archivo de registros."
    tramos = []
    for row in range(2, ws.max_row + 1):
        emp = ws.cell(row=row, column=col_map['empleado']).value
        if emp is None or str(emp).strip() == '':
            continue
        tramo = {'empleado': str(emp).strip()}
        for key in ['fecha', 'hora_inicio', 'hora_fin', 'tipo_tramo']:
            if key in col_map:
                tramo[key] = ws.cell(row=row, column=col_map[key]).value
        tramos.append(tramo)
    return tramos, None


def find_plantilla_columns(ws):
    cols = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=HEADER_ROW, column=col).value
        if val is None:
            continue
        h = str(val).strip().lower()
        if 'doc. identificador' in h or h == 'nif' or h == 'dni':
            cols.setdefault('nif', col)
        elif 'codigo empleado' in h or 'código empleado' in h:
            cols.setdefault('codigo', col)
        elif h == 'empleado':
            cols.setdefault('empleado', col)
        elif 'fecha de referencia' in h:
            cols.setdefault('fecha_ref', col)
        elif 'zona horaria' in h:
            cols.setdefault('zona', col)
        elif h == 'inicio':
            cols.setdefault('inicio', col)
        elif h == 'fin':
            cols.setdefault('fin', col)
        elif 'tipo de tramo' in h:
            cols.setdefault('tipo_tramo', col)
        elif 'sobrescritura' in h:
            cols.setdefault('sobrescritura', col)
    return cols


def copy_cell_style(src, tgt):
    if src.has_style:
        tgt.font = copy(src.font)
        tgt.border = copy(src.border)
        tgt.fill = copy(src.fill)
        tgt.number_format = src.number_format
        tgt.protection = copy(src.protection)
        tgt.alignment = copy(src.alignment)


# ═══════════════════════════════════════════════════
# Motor de conciliación
# ═══════════════════════════════════════════════════

def conciliar(wb_template, tramos):
    if TEMPLATE_SHEET in wb_template.sheetnames:
        ws = wb_template[TEMPLATE_SHEET]
    else:
        ws = wb_template.active

    cols = find_plantilla_columns(ws)
    if 'empleado' not in cols:
        return None, None, None, "No se encontró la columna 'Empleado' en la plantilla."

    plantilla_emps = []
    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        emp_val = ws.cell(row=row, column=cols['empleado']).value
        if emp_val is None or str(emp_val).strip() == '':
            continue
        plantilla_emps.append({
            'nombre': str(emp_val).strip(),
            'nombre_norm': normalize_name(emp_val),
            'nif': ws.cell(row=row, column=cols.get('nif', 1)).value,
            'codigo': ws.cell(row=row, column=cols.get('codigo', 2)).value,
        })

    plantilla_index = {pe['nombre_norm']: pe for pe in plantilla_emps}

    tramos_by_emp = {}
    for t in tramos:
        key = normalize_name(t['empleado'])
        if key not in tramos_by_emp:
            tramos_by_emp[key] = []
        tramos_by_emp[key].append(t)

    output_rows = []
    matched_norms = set()
    unmatched = []

    for tramo_norm, emp_tramos in tramos_by_emp.items():
        found = plantilla_index.get(tramo_norm)
        if found is None:
            for pe_norm, pe in plantilla_index.items():
                if tramo_norm in pe_norm or pe_norm in tramo_norm:
                    found = pe
                    break
        if found is None:
            for t in emp_tramos:
                unmatched.append(t['empleado'])
            continue

        matched_norms.add(found['nombre_norm'])
        for t in emp_tramos:
            fecha_ref = t.get('fecha')
            output_rows.append({
                'nif': found['nif'],
                'codigo': found['codigo'],
                'nombre': found['nombre'],
                'fecha_ref': extract_date_part(fecha_ref),
                'zona': ZONA_DEFAULT,
                'inicio': combine_date_time(fecha_ref, t.get('hora_inicio')),
                'fin': combine_date_time(fecha_ref, t.get('hora_fin')),
                'tipo_tramo': t.get('tipo_tramo'),
                'sobrescritura': SOBRESCRITURA_DEFAULT,
            })

    removed = sum(1 for pe in plantilla_emps if pe['nombre_norm'] not in matched_norms)

    style_row = HEADER_ROW + 1
    styles = {}
    max_col = ws.max_column
    for c in range(1, max_col + 1):
        styles[c] = ws.cell(row=style_row, column=c)

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(row=row, column=c).value = None

    for i, item in enumerate(output_rows):
        row = HEADER_ROW + 1 + i
        for c in range(1, max_col + 1):
            copy_cell_style(styles[c], ws.cell(row=row, column=c))
        if 'nif' in cols:
            ws.cell(row=row, column=cols['nif']).value = item['nif']
        if 'codigo' in cols:
            ws.cell(row=row, column=cols['codigo']).value = item['codigo']
        if 'empleado' in cols:
            ws.cell(row=row, column=cols['empleado']).value = item['nombre']
        if 'fecha_ref' in cols:
            cell = ws.cell(row=row, column=cols['fecha_ref'])
            cell.value = item['fecha_ref']
            cell.number_format = 'DD/MM/YYYY'
        if 'zona' in cols:
            ws.cell(row=row, column=cols['zona']).value = item['zona']
        if 'inicio' in cols:
            cell = ws.cell(row=row, column=cols['inicio'])
            cell.value = item['inicio']
            cell.number_format = 'DD/MM/YYYY HH:MM'
        if 'fin' in cols:
            cell = ws.cell(row=row, column=cols['fin'])
            cell.value = item['fin']
            cell.number_format = 'DD/MM/YYYY HH:MM'
        if 'tipo_tramo' in cols:
            ws.cell(row=row, column=cols['tipo_tramo']).value = item['tipo_tramo']
        if 'sobrescritura' in cols:
            ws.cell(row=row, column=cols['sobrescritura']).value = item['sobrescritura']

    return output_rows, list(set(unmatched)), removed, None


# ═══════════════════════════════════════════════════
# INTERFAZ
# ═══════════════════════════════════════════════════

def init_state():
    if 'horas_fin' not in st.session_state:
        st.session_state.horas_fin = {}
    if 'resultado' not in st.session_state:
        st.session_state.resultado = None


def apply_mass_hora(indices, hora):
    for i in indices:
        st.session_state.horas_fin[i] = hora


def main():
    st.set_page_config(page_title="Adaptador Endalia", page_icon="📋", layout="centered")
    init_state()

    st.markdown("""
    <div style="text-align:center; padding: 1.5rem 0 0.5rem 0;">
        <h1 style="margin-bottom:0.2rem;">📋 Adaptador Endalia</h1>
        <p style="color:#888;">Completa los tramos sin hora de fin y genera la plantilla de importación</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("---")

    # ── 1. Subida de archivos ──
    c1, c2 = st.columns(2)
    with c1:
        tramos_file = st.file_uploader("📂 Registro de tramos", type=['xlsx'], key='tramos',
                                        help="Excel exportado con los fichajes")
    with c2:
        plantilla_file = st.file_uploader("📂 Plantilla Endalia", type=['xlsx'], key='plantilla',
                                           help="Plantilla .xlsx de Endalia con los empleados")

    if not tramos_file or not plantilla_file:
        st.markdown("<div style='text-align:center;padding:3rem 0;color:#bbb;font-size:1.05rem;'>"
                    "Sube ambos archivos para comenzar</div>", unsafe_allow_html=True)
        return

    # ── Leer tramos ──
    tramos_bytes = tramos_file.read()
    wb_tramos = openpyxl.load_workbook(BytesIO(tramos_bytes), data_only=True)
    tramos_all, error = read_tramos(wb_tramos)
    if error:
        st.error(error)
        return

    tramos_sin_fin = [t for t in tramos_all if is_missing_hora_fin(t.get('hora_fin'))]
    if not tramos_sin_fin:
        st.success("✅ Todos los tramos ya tienen hora de fin. No hay nada pendiente.")
        return

    # Inicializar horas en session_state
    for i in range(len(tramos_sin_fin)):
        if i not in st.session_state.horas_fin:
            st.session_state.horas_fin[i] = time(17, 0)

    # ── 2. Completar horas ──
    st.markdown("---")
    st.markdown(f"### 🕐 {len(tramos_sin_fin)} tramos pendientes")

    # Aplicar hora en bloque
    st.markdown("##### Aplicar misma hora a varios tramos")
    empleados_unicos = sorted(set(t['empleado'] for t in tramos_sin_fin))

    m1, m2 = st.columns([3, 3])
    with m1:
        hora_masiva = st.time_input("Hora fin a aplicar", value=time(17, 0), key="masa_hora")
    with m2:
        seleccion = st.multiselect("Empleados", options=["— Todos —"] + empleados_unicos, key="masa_sel")

    if st.button("Aplicar hora", use_container_width=True):
        if seleccion:
            if "— Todos —" in seleccion:
                indices = list(range(len(tramos_sin_fin)))
            else:
                indices = [i for i, t in enumerate(tramos_sin_fin) if t['empleado'] in seleccion]
            for i in indices:
                st.session_state.horas_fin[i] = hora_masiva
                st.session_state[f"hf_{i}"] = hora_masiva
            st.rerun()

    st.markdown("")
    st.markdown("##### Detalle por tramo")
    st.caption("Puedes ajustar la hora de cada tramo de forma individual.")
    st.markdown("")

    # Cabecera de tabla
    hdr1, hdr2, hdr3, hdr4 = st.columns([3, 2, 1.5, 2])
    with hdr1:
        st.markdown("**Empleado**")
    with hdr2:
        st.markdown("**Fecha**")
    with hdr3:
        st.markdown("**Inicio**")
    with hdr4:
        st.markdown("**Hora fin**")

    # Filas
    for i, t in enumerate(tramos_sin_fin):
        c1, c2, c3, c4 = st.columns([3, 2, 1.5, 2])
        with c1:
            st.text(t['empleado'])
        with c2:
            st.text(fmt_date(t.get('fecha')))
        with c3:
            st.text(fmt_time(t.get('hora_inicio')))
        with c4:
            new_val = st.time_input(
                "fin", value=st.session_state.horas_fin[i],
                key=f"hf_{i}", label_visibility="collapsed"
            )
            st.session_state.horas_fin[i] = new_val

    # ── 3. Generar ──
    st.markdown("---")

    if st.button("🚀 Generar plantilla", type="primary", use_container_width=True):
        # Aplicar horas finales
        for i, t in enumerate(tramos_sin_fin):
            t['hora_fin'] = st.session_state.horas_fin[i]

        with st.spinner("Generando plantilla..."):
            try:
                plantilla_bytes = plantilla_file.read()
                wb_template = openpyxl.load_workbook(BytesIO(plantilla_bytes), data_only=False)
                output_rows, unmatched, removed, err = conciliar(wb_template, tramos_sin_fin)

                if err:
                    st.error(err)
                    return

                buf = BytesIO()
                wb_template.save(buf)
                output_bytes = patch_zip_with_validations(buf.getvalue(), plantilla_bytes)

                st.session_state.resultado = {
                    'bytes': output_bytes,
                    'rows': output_rows,
                    'unmatched': unmatched,
                }
            except Exception as e:
                st.error(f"Ha ocurrido un error: {str(e)}")
                return

    # ── 4. Resultado ──
    if st.session_state.resultado:
        res = st.session_state.resultado
        st.markdown("---")
        st.markdown(f"### ✅ {len(res['rows'])} registros listos para importar")

        if res['unmatched']:
            st.warning(f"Empleados no encontrados en la plantilla: {', '.join(sorted(res['unmatched']))}")

        preview = []
        for r in res['rows']:
            preview.append({
                'Empleado': r['nombre'],
                'Fecha': fmt_date(r['fecha_ref']),
                'Inicio': fmt_time(r['inicio']) if r['inicio'] else '',
                'Fin': fmt_time(r['fin']) if r['fin'] else '',
                'Tipo de tramo': r['tipo_tramo'],
            })

        st.dataframe(preview, use_container_width=True, hide_index=True)

        st.markdown("")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="📥 Descargar plantilla",
            data=res['bytes'],
            file_name=f"endalia_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet",
            type="primary",
            use_container_width=True,
        )


main()

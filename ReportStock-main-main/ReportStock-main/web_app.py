from __future__ import annotations

import json
import os
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from pypdf import PdfWriter
from werkzeug.utils import secure_filename

from reportstock_core import (
    generar_pdf_agotados,
    listar_vendedores_desde_excel,
    obtener_agotados_por_bodega,
)


BASE_DIR = Path(__file__).resolve().parent
TMP_DIR = BASE_DIR / 'tmp_web'
UPLOAD_DIR = TMP_DIR / 'uploads'
REPORTS_DIR = TMP_DIR / 'reportes'
FINAL_REPORTS_DIR = TMP_DIR / 'reportes_finales'
HISTORY_FILE = TMP_DIR / 'historial_pdf.json'

for directory in (UPLOAD_DIR, REPORTS_DIR, FINAL_REPORTS_DIR):
    directory.mkdir(parents=True, exist_ok=True)

if not HISTORY_FILE.exists():
    HISTORY_FILE.write_text(json.dumps({'reportes': []}, ensure_ascii=False, indent=2), encoding='utf-8')

app = Flask(__name__, template_folder='web_templates')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024
app.secret_key = 'reportstock-web-secret'

DELETE_HISTORY_KEY = os.getenv('HISTORIAL_DELETE_KEY', 'admin123')

PREFERRED_PAGE_LOGO = 'lamar_optical_logo_no_bg.png'


def _find_page_logo() -> Path | None:
    preferred = BASE_DIR / PREFERRED_PAGE_LOGO
    if preferred.exists() and preferred.suffix.lower() in {'.png', '.jpg', '.jpeg', '.webp', '.svg'}:
        return preferred

    candidatos = [
        'LAMMAR OPTICAL LOGO NO BG.png',
        'LAMAR OPTICAL LOGO NO BG.png',
        'LAMAR OPTICAL - LOGO.png',
        'LAMAR OPTICAL - LOGO 2.png',
    ]
    for name in candidatos:
        path = BASE_DIR / name
        if path.exists() and path.suffix.lower() in {'.png', '.jpg', '.jpeg', '.webp', '.svg'}:
            return path

    for pattern in ('*lamar*optical*logo*.png', '*lamar*optical*logo*.jpg', '*lamar*optical*logo*.jpeg'):
        files = sorted(BASE_DIR.glob(pattern))
        if files:
            return files[0]

    return None


def _find_lamar_logo() -> Path | None:
    candidatos = sorted(BASE_DIR.glob('LAMAR OPTICAL - LOGO*.pdf'))
    return candidatos[0] if candidatos else None


LAMAR_LOGO_FILE = _find_lamar_logo()
PAGE_LOGO_FILE = _find_page_logo()


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in {'.xls', '.xlsx'}


def _load_history() -> dict:
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
            if isinstance(data, dict) and 'reportes' in data:
                return data
    except Exception:
        pass
    return {'reportes': []}


def _save_history(data: dict) -> None:
    with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _add_history_record(record: dict) -> None:
    data = _load_history()
    data.setdefault('reportes', []).append(record)
    data['reportes'] = sorted(data['reportes'], key=lambda x: x.get('fecha', ''), reverse=True)
    _save_history(data)


def _merge_pdfs(pdf_paths: list[Path], output_path: Path) -> None:
    writer = PdfWriter()
    for pdf_path in sorted(pdf_paths):
        writer.append(str(pdf_path))
    with open(output_path, 'wb') as f:
        writer.write(f)


def _sanitize_label(value: str) -> str:
    safe = ''.join(ch if ch.isalnum() or ch in {'-', '_'} else '_' for ch in (value or '').strip())
    return safe or 'Sin_Empresa'


def _build_zip(pdf_paths: list[Path], output_path: Path) -> None:
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for pdf_path in sorted(pdf_paths):
            zip_file.write(pdf_path, arcname=pdf_path.name)


def _copy_individual_pdfs(pdf_paths: list[Path], empresa_file: str, timestamp: str) -> list[dict]:
    individuales: list[dict] = []
    for pdf_file in sorted(pdf_paths):
        bodega = pdf_file.stem.split('_agotados_')[0]
        destino_nombre = f"reporte_{empresa_file}_{bodega}_{timestamp}.pdf"
        destino = FINAL_REPORTS_DIR / destino_nombre
        destino.write_bytes(pdf_file.read_bytes())
        individuales.append({'bodega': bodega, 'archivo': destino_nombre})
    return individuales


def _parse_bodegas(raw: str) -> list[str]:
    partes = [p.strip() for p in (raw or '').replace('\n', ',').split(',')]
    limpias = [p for p in partes if p]
    return list(dict.fromkeys(limpias))


def _parse_filtros_por_bodega(raw: str) -> dict[str, int]:
    filtros: dict[str, int] = {}
    if not raw:
        return filtros

    lineas = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    for linea in lineas:
        if ':' not in linea:
            raise ValueError(
                "Formato inválido en filtros por bodega. Usa una línea por bodega con formato: BODEGA:MINIMO"
            )
        nombre, minimo = linea.split(':', 1)
        nombre = nombre.strip()
        minimo = minimo.strip()
        if not nombre:
            raise ValueError("Hay una línea de filtro sin nombre de bodega.")
        try:
            minimo_int = int(minimo)
            if minimo_int < 0:
                raise ValueError
        except Exception:
            raise ValueError(f"El mínimo para '{nombre}' debe ser un entero >= 0")
        filtros[nombre] = minimo_int

    return filtros


@app.route('/', methods=['GET'])
def index():
    return render_template(
        'index.html',
        logo_available=bool(LAMAR_LOGO_FILE),
        page_logo_available=bool(PAGE_LOGO_FILE),
    )


@app.route('/api/bodegas', methods=['POST'])
def api_bodegas():
    excel_file = request.files.get('archivo_excel')

    if not excel_file or not excel_file.filename:
        return jsonify({'ok': False, 'message': 'Debes seleccionar un archivo Excel.'}), 400

    if not allowed_file(excel_file.filename):
        return jsonify({'ok': False, 'message': 'Formato inválido. Solo se permiten archivos .xls o .xlsx.'}), 400

    token = uuid.uuid4().hex
    filename = secure_filename(excel_file.filename)
    upload_path = UPLOAD_DIR / f'preview_{token}_{filename}'

    try:
        excel_file.save(upload_path)
        bodegas = listar_vendedores_desde_excel(upload_path)
        return jsonify({'ok': True, 'bodegas': bodegas})
    except Exception as exc:
        return jsonify({'ok': False, 'message': f'Error leyendo archivo: {exc}'}), 500


@app.route('/historial', methods=['GET'])
def historial():
    empresa = (request.args.get('empresa') or '').strip()
    historial_data = _load_history().get('reportes', [])
    if empresa:
        historial_data = [item for item in historial_data if item.get('empresa') == empresa]

    pdf_items = []
    for reporte in historial_data:
        for idx, indiv in enumerate(reporte.get('archivos_vendedor') or []):
            pdf_items.append(
                {
                    'reporte_id': reporte.get('id'),
                    'item_index': idx,
                    'fecha': reporte.get('fecha'),
                    'empresa': reporte.get('empresa'),
                    'bodega': indiv.get('bodega') or '-',
                    'archivo_pdf': indiv.get('archivo') or '-',
                }
            )

    pdf_items = sorted(pdf_items, key=lambda x: x.get('fecha', ''), reverse=True)
    return render_template(
        'historial.html',
        items=pdf_items,
        empresa_filtro=empresa,
        logo_available=bool(LAMAR_LOGO_FILE),
        page_logo_available=bool(PAGE_LOGO_FILE),
    )


@app.route('/logo-pagina', methods=['GET'])
def logo_pagina():
    if not PAGE_LOGO_FILE or not PAGE_LOGO_FILE.exists():
        flash('No se encontró el logo de la página.', 'error')
        return redirect(url_for('index'))
    return send_file(
        PAGE_LOGO_FILE,
        as_attachment=False,
        download_name=PAGE_LOGO_FILE.name,
    )


@app.route('/logo-lamar', methods=['GET'])
def logo_lamar():
    if not LAMAR_LOGO_FILE or not LAMAR_LOGO_FILE.exists():
        flash('No se encontró el logo de Lamar.', 'error')
        return redirect(url_for('index'))
    return send_file(
        LAMAR_LOGO_FILE,
        mimetype='application/pdf',
        as_attachment=False,
        download_name=LAMAR_LOGO_FILE.name,
    )


@app.route('/descargar/<reporte_id>/<tipo>', methods=['GET'])
def descargar_reporte(reporte_id: str, tipo: str):
    historial_data = _load_history().get('reportes', [])
    reporte = next((item for item in historial_data if item.get('id') == reporte_id), None)
    if not reporte:
        flash('Reporte no encontrado en historial.', 'error')
        return redirect(url_for('historial'))

    tipo = (tipo or '').strip().lower()
    if tipo not in {'pdf', 'zip'}:
        flash('Tipo de descarga inválido.', 'error')
        return redirect(url_for('historial'))

    if tipo == 'pdf':
        nombre_archivo = reporte.get('archivo_pdf', '') or reporte.get('archivo', '')
        mimetype = 'application/pdf'
    else:
        nombre_archivo = reporte.get('archivo_zip', '')
        mimetype = 'application/zip'

    ruta_archivo = FINAL_REPORTS_DIR / nombre_archivo
    if not nombre_archivo or not ruta_archivo.exists():
        flash(f'El archivo {tipo.upper()} ya no existe en el servidor.', 'error')
        return redirect(url_for('historial'))

    return send_file(
        ruta_archivo,
        mimetype=mimetype,
        as_attachment=True,
        download_name=ruta_archivo.name,
    )


@app.route('/descargar-vendedor/<reporte_id>/<int:item_index>', methods=['GET'])
def descargar_vendedor(reporte_id: str, item_index: int):
    historial_data = _load_history().get('reportes', [])
    reporte = next((item for item in historial_data if item.get('id') == reporte_id), None)
    if not reporte:
        flash('Reporte no encontrado en historial.', 'error')
        return redirect(url_for('historial'))

    items = reporte.get('archivos_vendedor') or []
    if item_index < 0 or item_index >= len(items):
        flash('PDF individual no encontrado.', 'error')
        return redirect(url_for('historial'))

    item = items[item_index]
    nombre_archivo = item.get('archivo', '')
    ruta_archivo = FINAL_REPORTS_DIR / nombre_archivo
    if not nombre_archivo or not ruta_archivo.exists():
        flash('El PDF individual ya no existe en el servidor.', 'error')
        return redirect(url_for('historial'))

    return send_file(
        ruta_archivo,
        mimetype='application/pdf',
        as_attachment=True,
        download_name=ruta_archivo.name,
    )


@app.route('/historial/eliminar-item/<reporte_id>/<int:item_index>', methods=['POST'])
def eliminar_item_historial(reporte_id: str, item_index: int):
    empresa_filtro = (request.form.get('empresa') or '').strip()
    delete_key = (request.form.get('delete_key') or '').strip()

    if not delete_key or delete_key != DELETE_HISTORY_KEY:
        flash('Clave inválida para eliminar.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    historial = _load_history()
    reportes = historial.get('reportes', [])

    reporte = next((item for item in reportes if item.get('id') == reporte_id), None)
    if not reporte:
        flash('No se encontró el registro a eliminar.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    items = reporte.get('archivos_vendedor') or []
    if item_index < 0 or item_index >= len(items):
        flash('No se encontró el PDF individual a eliminar.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    item = items.pop(item_index)
    nombre_pdf = (item or {}).get('archivo', '')
    if nombre_pdf:
        ruta_pdf = FINAL_REPORTS_DIR / nombre_pdf
        if ruta_pdf.exists():
            try:
                ruta_pdf.unlink()
            except Exception:
                pass

    if not items:
        nombre_zip = (reporte or {}).get('archivo_zip', '')
        if nombre_zip:
            ruta_zip = FINAL_REPORTS_DIR / nombre_zip
            if ruta_zip.exists():
                try:
                    ruta_zip.unlink()
                except Exception:
                    pass
        historial['reportes'] = [r for r in reportes if r.get('id') != reporte_id]

    _save_history(historial)
    flash('Registro eliminado del historial.', 'success')
    return redirect(url_for('historial', empresa=empresa_filtro or None))


@app.route('/historial/eliminar-seleccion', methods=['POST'])
def eliminar_seleccion_historial():
    empresa_filtro = (request.form.get('empresa') or '').strip()
    delete_key = (request.form.get('delete_key') or '').strip()

    if not delete_key or delete_key != DELETE_HISTORY_KEY:
        flash('Clave inválida para eliminar.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    seleccion = request.form.getlist('selected_items')
    if not seleccion:
        flash('Debes seleccionar al menos una línea.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    seleccion_por_reporte: dict[str, set[int]] = {}
    for token in seleccion:
        if '::' not in token:
            continue
        reporte_id, index_raw = token.split('::', 1)
        reporte_id = (reporte_id or '').strip()
        try:
            item_index = int(index_raw)
        except ValueError:
            continue
        if not reporte_id:
            continue
        seleccion_por_reporte.setdefault(reporte_id, set()).add(item_index)

    if not seleccion_por_reporte:
        flash('Selección inválida.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    historial = _load_history()
    reportes = historial.get('reportes', [])
    ids_a_remover: set[str] = set()
    eliminados = 0

    for reporte in reportes:
        reporte_id = str(reporte.get('id') or '').strip()
        if reporte_id not in seleccion_por_reporte:
            continue

        items = reporte.get('archivos_vendedor') or []
        indices = sorted(seleccion_por_reporte[reporte_id], reverse=True)
        for item_index in indices:
            if item_index < 0 or item_index >= len(items):
                continue
            item = items.pop(item_index)
            nombre_pdf = (item or {}).get('archivo', '')
            if nombre_pdf:
                ruta_pdf = FINAL_REPORTS_DIR / nombre_pdf
                if ruta_pdf.exists():
                    try:
                        ruta_pdf.unlink()
                    except Exception:
                        pass
            eliminados += 1

        if not items:
            nombre_zip = (reporte or {}).get('archivo_zip', '')
            if nombre_zip:
                ruta_zip = FINAL_REPORTS_DIR / nombre_zip
                if ruta_zip.exists():
                    try:
                        ruta_zip.unlink()
                    except Exception:
                        pass
            ids_a_remover.add(reporte_id)

    if ids_a_remover:
        historial['reportes'] = [r for r in reportes if str(r.get('id') or '').strip() not in ids_a_remover]

    _save_history(historial)

    if eliminados > 0:
        flash(f'Se eliminaron {eliminados} línea(s) del historial.', 'success')
    else:
        flash('No se eliminaron líneas (revisa la selección).', 'error')

    return redirect(url_for('historial', empresa=empresa_filtro or None))


@app.route('/historial/eliminar-todo', methods=['POST'])
def eliminar_todo_historial():
    empresa_filtro = (request.form.get('empresa') or '').strip()
    delete_key = (request.form.get('delete_key') or '').strip()

    if not delete_key or delete_key != DELETE_HISTORY_KEY:
        flash('Clave inválida para eliminar.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    historial = _load_history()
    reportes = historial.get('reportes', [])

    if empresa_filtro:
        a_eliminar = [r for r in reportes if r.get('empresa') == empresa_filtro]
        restantes = [r for r in reportes if r.get('empresa') != empresa_filtro]
    else:
        a_eliminar = list(reportes)
        restantes = []

    if not a_eliminar:
        flash('No hay registros para eliminar con ese filtro.', 'error')
        return redirect(url_for('historial', empresa=empresa_filtro or None))

    for reporte in a_eliminar:
        for item in (reporte.get('archivos_vendedor') or []):
            nombre_pdf = (item or {}).get('archivo', '')
            if not nombre_pdf:
                continue
            ruta_pdf = FINAL_REPORTS_DIR / nombre_pdf
            if ruta_pdf.exists():
                try:
                    ruta_pdf.unlink()
                except Exception:
                    pass

        nombre_zip = (reporte or {}).get('archivo_zip', '')
        if nombre_zip:
            ruta_zip = FINAL_REPORTS_DIR / nombre_zip
            if ruta_zip.exists():
                try:
                    ruta_zip.unlink()
                except Exception:
                    pass

    historial['reportes'] = restantes
    _save_history(historial)

    if empresa_filtro:
        flash(f'Historial de {empresa_filtro} eliminado correctamente.', 'success')
    else:
        flash('Se eliminó todo el historial correctamente.', 'success')

    return redirect(url_for('historial', empresa=empresa_filtro or None))


@app.route('/generar', methods=['POST'])
def generar():
    excel_file = request.files.get('archivo_excel')
    empresa = (request.form.get('empresa') or '').strip()
    bodegas_raw = (request.form.get('bodegas') or '').strip()
    filtros_por_bodega_raw = (request.form.get('filtros_por_bodega') or '').strip()

    try:
        cantidad_minima = int((request.form.get('cantidad_minima') or '3').strip())
        if cantidad_minima < 0:
            raise ValueError
    except ValueError:
        flash('La cantidad mínima debe ser un número entero mayor o igual a 0.', 'error')
        return render_template('index.html'), 400

    if not excel_file or not excel_file.filename:
        flash('Debes seleccionar un archivo Excel.', 'error')
        return render_template('index.html'), 400

    if not allowed_file(excel_file.filename):
        flash('Formato inválido. Solo se permiten archivos .xls o .xlsx.', 'error')
        return render_template('index.html'), 400

    if empresa not in {'Inversiones', 'Lamar'}:
        flash('Debes elegir una empresa.', 'error')
        return render_template('index.html'), 400

    bodegas_seleccionadas = _parse_bodegas(bodegas_raw)
    if not bodegas_seleccionadas:
        flash('Debes elegir los vendedores.', 'error')
        return render_template('index.html'), 400

    try:
        filtros_por_bodega = _parse_filtros_por_bodega(filtros_por_bodega_raw)
    except ValueError as parse_exc:
        flash(str(parse_exc), 'error')
        return render_template('index.html'), 400

    token = uuid.uuid4().hex
    filename = secure_filename(excel_file.filename)
    upload_path = UPLOAD_DIR / f'{token}_{filename}'
    excel_file.save(upload_path)

    try:
        vendedores_disponibles = listar_vendedores_desde_excel(upload_path)
        disponibles_set = {str(v).strip() for v in vendedores_disponibles}

        bodegas_invalidas = [b for b in bodegas_seleccionadas if b not in disponibles_set]
        if bodegas_invalidas:
            flash(
                f"Estas bodegas no existen en el archivo: {', '.join(bodegas_invalidas)}",
                'error',
            )
            return render_template('index.html'), 400

        filtros_invalidos = [b for b in filtros_por_bodega.keys() if b not in disponibles_set]
        if filtros_invalidos:
            flash(
                f"Hay filtros para bodegas inexistentes: {', '.join(filtros_invalidos)}",
                'error',
            )
            return render_template('index.html'), 400

        agotados = obtener_agotados_por_bodega(
            upload_path,
            cantidad_minima=cantidad_minima,
            vendedores_filtrados=bodegas_seleccionadas,
            cantidad_minima_por_bodega=filtros_por_bodega,
        )

        report_folder = REPORTS_DIR / token
        report_folder.mkdir(parents=True, exist_ok=True)
        pdf_paths = generar_pdf_agotados(agotados, carpeta_salida=report_folder, empresa=empresa)

        if not pdf_paths:
            flash('No se generaron reportes. Revisa el archivo de entrada.', 'error')
            return render_template('index.html'), 400

        empresa_file = _sanitize_label(empresa if empresa in {'Inversiones', 'Lamar', 'Sin_Empresa'} else empresa)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        zip_name = f"reporte_agotados_{empresa_file}_{timestamp}.zip"
        zip_path = FINAL_REPORTS_DIR / zip_name
        _build_zip(pdf_paths, zip_path)
        archivos_vendedor = _copy_individual_pdfs(pdf_paths, empresa_file, timestamp)

        _add_history_record(
            {
                'id': token,
                'fecha': datetime.now().isoformat(),
                'empresa': empresa,
                'vendedor': ', '.join(bodegas_seleccionadas),
                'cantidad_minima': cantidad_minima,
                'filtros_por_bodega': filtros_por_bodega,
                'archivo_zip': zip_name,
                'archivos_vendedor': archivos_vendedor,
                'total_bodegas': len(pdf_paths),
            }
        )

        return send_file(
            zip_path,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_name,
        )
    except Exception as exc:
        flash(f'Error procesando archivo: {exc}', 'error')
        return render_template('index.html'), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8071, debug=True)
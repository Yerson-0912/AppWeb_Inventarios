from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

VENDEDORA_CON_COLUMNA_PRINCIPAL = 'MALETA NATALIA REYES'


def _find_company_logo(empresa: str) -> Path | None:
    base_dir = Path(__file__).resolve().parent
    empresa_key = (empresa or '').strip().lower()

    if empresa_key == 'lamar':
        candidates = [
            'lamar_optical_logo_no_bg.png',
            'LAMMAR OPTICAL LOGO NO BG.png',
            'LAMAR OPTICAL LOGO NO BG.png',
        ]
        for name in candidates:
            path = base_dir / name
            if path.exists() and path.suffix.lower() in {'.png', '.jpg', '.jpeg', '.webp'}:
                return path
        for path in base_dir.glob('*lamar*logo*no*bg*.png'):
            return path

    if empresa_key == 'inversiones':
        candidates = [
            'LOGO OPTICAL SHOP EDITABLE 1.png',
            'LOGO OPTICAL SHOP EDITABLE.png',
        ]
        for name in candidates:
            path = base_dir / name
            if path.exists() and path.suffix.lower() in {'.png', '.jpg', '.jpeg', '.webp'}:
                return path
        for path in base_dir.glob('*LOGO*OPTICAL*SHOP*EDITABLE*.png'):
            return path

    return None


def read_excel_smart(ruta: Path) -> pd.DataFrame:
    suffix = ruta.suffix.lower()
    if suffix == '.xlsx':
        return pd.read_excel(ruta, engine='openpyxl')
    if suffix == '.xls':
        return pd.read_excel(ruta, engine='xlrd')
    return pd.read_excel(ruta)


def obtener_unicos(archivo: Path, columna: str) -> list:
    df = read_excel_smart(archivo)
    if columna not in df.columns:
        logging.warning("Columna %s no encontrada en %s", columna, archivo)
        return []
    return df[columna].dropna().astype(str).tolist()


def obtener_agotados_por_bodega(
    ruta_archivo: Path,
    cantidad_minima: int = 3,
    vendedores_filtrados: list[str] | None = None,
    cantidad_minima_por_bodega: dict[str, int] | None = None,
) -> dict:
    df = read_excel_smart(ruta_archivo)

    columnas_requeridas = {'BODEGA', 'REFERENCIA', 'CANTIDAD', 'NOMBRE DEL GRUPO'}
    faltantes = columnas_requeridas.difference(set(df.columns))
    if faltantes:
        raise ValueError(
            f"El archivo no contiene las columnas requeridas: {', '.join(sorted(faltantes))}"
        )

    principal = df[df['BODEGA'].astype(str).str.upper().str.strip() == 'PRINCIPAL']
    stock_principal = {}
    for _, fila in principal.iterrows():
        referencia = fila.get('REFERENCIA')
        cantidad = fila.get('CANTIDAD', 0)
        try:
            cantidad = int(cantidad) if pd.notna(cantidad) else 0
        except Exception:
            cantidad = 0
        if referencia is not None:
            stock_principal[referencia] = cantidad

    bodegas_vendedores = (
        df[df['BODEGA'].astype(str).str.upper().str.strip() != 'PRINCIPAL']['BODEGA']
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    if vendedores_filtrados:
        vendedores_set = {str(v).strip() for v in vendedores_filtrados if str(v).strip()}
        bodegas_vendedores = [b for b in bodegas_vendedores if str(b).strip() in vendedores_set]

    filtros_por_bodega = {}
    if cantidad_minima_por_bodega:
        for bodega, minimo in cantidad_minima_por_bodega.items():
            bodega_key = str(bodega).strip()
            try:
                filtros_por_bodega[bodega_key] = max(0, int(minimo))
            except Exception:
                continue

    agotados_por_bodega = {}
    for bodega in bodegas_vendedores:
        minimo_aplicable = filtros_por_bodega.get(str(bodega).strip(), cantidad_minima)
        productos_vendedor = df[df['BODEGA'].astype(str) == str(bodega)]
        lista_agotados = []

        for _, fila in productos_vendedor.iterrows():
            referencia = fila.get('REFERENCIA')
            cantidad_vendedor = fila.get('CANTIDAD', 0)
            nombre_grupo = fila.get('NOMBRE DEL GRUPO') or ''

            try:
                cantidad_vendedor = int(cantidad_vendedor) if pd.notna(cantidad_vendedor) else 0
            except Exception:
                cantidad_vendedor = 0

            cantidad_principal = stock_principal.get(referencia, 0)
            if referencia is not None and cantidad_vendedor > 0 and cantidad_principal <= minimo_aplicable:
                lista_agotados.append(
                    {
                        'referencia': referencia,
                        'nombre_grupo': nombre_grupo,
                        'cantidad_vendedor': cantidad_vendedor,
                        'cantidad_principal': cantidad_principal,
                    }
                )

        lista_agotados = sorted(
            lista_agotados,
            key=lambda x: (x.get('nombre_grupo') or '', x.get('referencia') or ''),
        )
        agotados_por_bodega[str(bodega)] = lista_agotados

    return agotados_por_bodega


def generar_pdf_agotados(agotados: dict, carpeta_salida: str | Path = 'reportes', empresa: str = '') -> list[Path]:
    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")

    default_font = 'Helvetica'
    try:
        font_file = Path(__file__).parent / 'fonts' / 'DejaVuSans.ttf'
        if font_file.is_file():
            pdfmetrics.registerFont(TTFont('DejaVuSans', str(font_file)))
            default_font = 'DejaVuSans'
    except Exception:
        pass

    empresa_prefix = ''
    if empresa:
        empresa_clean = re.sub(r'[^0-9A-Za-z-_]', '_', empresa.strip())
        empresa_prefix = f"{empresa_clean}_"

    nombres_empresas = {
        'Inversiones': 'INVERSIONES RUEDA S.A.S',
        'Lamar': 'LAMAR OPTICAL S.A.S',
        'Sin_Empresa': 'REPORTE DE PRODUCTOS AGOTADOS',
    }

    logo_path = _find_company_logo(empresa)

    rutas_generadas: list[Path] = []
    for bodega, productos in agotados.items():
        mostrar_columna_principal = str(bodega).strip().upper() == VENDEDORA_CON_COLUMNA_PRINCIPAL
        bodega_clean = re.sub(r'[^0-9A-Za-z-_]', '_', bodega.strip())
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ruta_pdf = carpeta / f"{empresa_prefix}{bodega_clean}_agotados_{timestamp}.pdf"

        doc = SimpleDocTemplate(
            str(ruta_pdf),
            pagesize=letter,
            rightMargin=50,
            leftMargin=50,
            topMargin=50,
            bottomMargin=50,
            title=f'Reporte de Agotados - {bodega}',
            author='Sistema de Reportes',
            subject=f'Productos agotados en {bodega}',
            creator='ReportStock',
        )

        estilos = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle(
            'CustomTitle',
            parent=estilos['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName=default_font,
        )
        estilo_subtitulo = ParagraphStyle(
            'CustomSubtitle',
            parent=estilos['Heading2'],
            fontSize=12,
            textColor=colors.HexColor('#34495e'),
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName=default_font,
        )
        estilo_grupo = ParagraphStyle(
            'GrupoHeader',
            parent=estilos['Heading3'],
            fontSize=11,
            textColor=colors.HexColor('#2980b9'),
            spaceAfter=10,
            spaceBefore=15,
            fontName=default_font,
        )

        elementos = []
        if logo_path and logo_path.exists():
            try:
                logo = Image(str(logo_path))
                max_logo_width = 2.4 * inch
                scale = max_logo_width / float(logo.imageWidth) if logo.imageWidth else 1
                if scale < 1:
                    logo.drawWidth = logo.imageWidth * scale
                    logo.drawHeight = logo.imageHeight * scale
                logo.hAlign = 'CENTER'
                elementos.append(logo)
                elementos.append(Spacer(1, 10))
            except Exception:
                pass

        titulo_text = nombres_empresas.get(empresa, 'REPORTE DE PRODUCTOS AGOTADOS')
        elementos.append(Paragraph(f"{titulo_text}", estilo_titulo))
        elementos.append(Paragraph(f"<b>Bodega:</b> {bodega}", estilo_subtitulo))
        elementos.append(Paragraph(f"<b>Fecha:</b> {fecha_actual}", estilo_subtitulo))
        elementos.append(Spacer(1, 20))
        elementos.append(Paragraph(f"<b>Total de productos agotados:</b> {len(productos)}", estilos['Normal']))
        elementos.append(Spacer(1, 20))

        if productos:
            grupos = {}
            for prod in productos:
                grupo = prod.get('nombre_grupo') or 'Sin Grupo'
                grupos.setdefault(grupo, []).append(prod)

            for grupo in sorted(grupos.keys()):
                elementos.append(Paragraph(f"<b>{grupo}</b>", estilo_grupo))
                if mostrar_columna_principal:
                    datos_tabla = [['Referencia', 'Cant. Vendedor', 'Cant. Principal']]
                else:
                    datos_tabla = [['Referencia', 'Cant. Vendedor']]

                for prod in grupos[grupo]:
                    if mostrar_columna_principal:
                        datos_tabla.append(
                            [
                                str(prod['referencia']),
                                str(prod['cantidad_vendedor']),
                                str(prod['cantidad_principal']),
                            ]
                        )
                    else:
                        datos_tabla.append([str(prod['referencia']), str(prod['cantidad_vendedor'])])

                if mostrar_columna_principal:
                    tabla = Table(datos_tabla, colWidths=[2.2 * inch, 1.3 * inch, 1.3 * inch])
                else:
                    tabla = Table(datos_tabla, colWidths=[2.5 * inch, 1.5 * inch])

                tabla.setStyle(
                    TableStyle(
                        [
                            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 11),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
                            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                            ('FONTSIZE', (0, 1), (-1, -1), 9),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black),
                            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                        ]
                    )
                )
                elementos.append(tabla)
                elementos.append(Spacer(1, 15))
        else:
            elementos.append(Paragraph("<b>No hay productos agotados en esta bodega</b>", estilos['Normal']))

        doc.build(elementos)
        rutas_generadas.append(ruta_pdf)

    return rutas_generadas


def listar_vendedores_desde_excel(ruta_archivo: Path) -> list[str]:
    df = read_excel_smart(ruta_archivo)
    if 'BODEGA' not in df.columns:
        return []
    vendedores = (
        df['BODEGA']
        .dropna()
        .astype(str)
        .loc[lambda s: s.str.upper().str.strip() != 'PRINCIPAL']
        .unique()
        .tolist()
    )
    return sorted([v for v in vendedores if str(v).strip()])
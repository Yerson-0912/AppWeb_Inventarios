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


def _normalizar_token_referencia(token: str, indice: int) -> str:
    token_limpio = str(token or '').strip().upper()
    if indice > 0:
        match = re.fullmatch(r'([A-Z]+)0*(\d+)', token_limpio)
        if match:
            prefijo, numero = match.groups()
            return f'{prefijo}{int(numero)}'
    return token_limpio


def normalizar_referencia(referencia: object) -> str:
    referencia_base = re.sub(r'\s+', ' ', str(referencia or '').strip()).upper()
    if not referencia_base:
        return ''
    tokens = referencia_base.split(' ')
    return ' '.join(_normalizar_token_referencia(token, indice) for indice, token in enumerate(tokens))


NUEVA_COLECCION_REFERENCIAS = frozenset(
    referencia.strip()
    for referencia in """
L1290 C01
L1290 C02
L1290 C03
L1290 C04
L1291 C01
L1291 C02
L1291 C03
L1291 C04
L1292 C01
L1292 C02
L1292 C03
L1292 C04
L1293 C01
L1293 C02
L1293 C03
L1293 C04
L1294 C01
L1294 C02
L1294 C03
L1294 C04
L1295 C01
L1295 C02
L1295 C03
L1295 C04
L1296 C01
L1296 C02
L1296 C03
L1296 C04
L1297 C01
L1297 C02
L1297 C03
L1297 C04
L1297 C05
L1298 C01
L1298 C02
L1298 C03
L1298 C04
L1299 C01
L1299 C02
L1299 C03
L1299 C04
L1300 C01
L1300 C02
L1300 C03
L1300 C04
L1301 C01
L1301 C02
L1301 C03
L1301 C04
L1302 C01
L1302 C02
L1302 C03
L1302 C04
L1303 C01
L1303 C02
L1303 C03
L1303 C04
L1304 C01
L1304 C02
L1304 C03
L1304 C04
L1305 C01
L1305 C02
L1305 C03
L1305 C04
L1306 C01
L1306 C02
L1306 C03
L1306 C04
L1307 C01
L1307 C02
L1307 C03
L1307 C04
L1308 C01
L1308 C02
L1308 C03
L1308 C04
L1309 C01
L1309 C02
L1309 C03
L1309 C04
L1310 C01
L1310 C02
L1310 C03
L1310 C04
L1311 C01
L1311 C02
L1311 C03
L1311 C04
OB070 C1
OB070 C2
OB070 C3
OB070 C4
OB071 C1
OB071 C2
OB071 C3
OB071 C4
OB072 C1
OB072 C2
OB072 C3
OB072 C4
OB073 C1
OB073 C2
OB073 C3
OB073 C4
OB074 C1
OB074 C2
OB074 C3
OB074 C4
OB075 C1
OB075 C2
OB075 C3
OB075 C4
OB076 C1
OB076 C2
OB076 C3
OB076 C4
OB077 C1
OB077 C2
OB077 C3
OB077 C4
LJ090 C01
LJ090 C02
LJ090 C03
LJ090 C04
LJ091 C01
LJ091 C02
LJ091 C03
LJ091 C04
LJ092 C01
LJ092 C02
LJ092 C03
LJ092 C04
LJ093 C01
LJ093 C02
LJ093 C03
LJ093 C04
LJ094 C01
LJ094 C02
LJ094 C03
LJ094 C04
LJ095 C01
LJ095 C02
LJ095 C03
LJ095 C04
LJ096 C01
LJ096 C02
LJ096 C03
LJ096 C04
LJ097 C01
LJ097 C02
LJ097 C03
LJ097 C04
LJ098 C01
LJ098 C02
LJ098 C03
LJ098 C04
LJ099 C01
LJ099 C02
LJ099 C03
LJ099 C04
LJ100 C01
LJ100 C02
LJ100 C03
LJ100 C04
""".splitlines()
    if referencia.strip()
)

PALETAS_RESALTADO = {
    'Amarillo suave': {'fondo': '#FFF200', 'texto': '#1F1F1F', 'borde': '#D4B100'},
    'Amarillo intenso': {'fondo': '#FFF000', 'texto': '#111111', 'borde': '#C9A800'},
    'Verde menta': {'fondo': '#D9FBE8', 'texto': '#0F5132', 'borde': '#34A853'},
    'Azul cielo': {'fondo': '#DCEEFF', 'texto': '#0B3C5D', 'borde': '#4A90E2'},
    'Coral suave': {'fondo': '#FFE2DB', 'texto': '#7A271A', 'borde': '#F97360'},
    'Lavanda': {'fondo': '#EFE4FF', 'texto': '#4B2E83', 'borde': '#9B6DDB'},
}
DEFAULT_PALETA_RESALTADO = 'Amarillo suave'
NUEVA_COLECCION_REFERENCIAS_NORMALIZADAS = frozenset(
    normalizar_referencia(referencia)
    for referencia in NUEVA_COLECCION_REFERENCIAS
    if normalizar_referencia(referencia)
)


def referencia_en_nueva_coleccion(referencia: object) -> bool:
    return normalizar_referencia(referencia) in NUEVA_COLECCION_REFERENCIAS_NORMALIZADAS


def obtener_paleta_resaltado(nombre_paleta: str | None = None) -> dict[str, str]:
    paleta_final = nombre_paleta if nombre_paleta in PALETAS_RESALTADO else DEFAULT_PALETA_RESALTADO
    return {'nombre': paleta_final, **PALETAS_RESALTADO[paleta_final]}


def construir_configuracion_resaltado(
    habilitado: bool = False,
    nombre_paleta: str | None = None,
) -> dict[str, object]:
    paleta = obtener_paleta_resaltado(nombre_paleta)
    return {
        'habilitado': bool(habilitado),
        'nombre_paleta': paleta['nombre'],
        'fondo_hex': paleta['fondo'],
        'texto_hex': paleta['texto'],
        'borde_hex': paleta['borde'],
        'fondo': colors.HexColor(paleta['fondo']),
        'texto': colors.HexColor(paleta['texto']),
        'borde': colors.HexColor(paleta['borde']),
    }


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


def generar_pdf_agotados(
    agotados: dict,
    carpeta_salida: str | Path = 'reportes',
    empresa: str = '',
    resaltar_nueva_coleccion: bool = False,
    paleta_resaltado: str | None = None,
) -> list[Path]:
    carpeta = Path(carpeta_salida)
    carpeta.mkdir(parents=True, exist_ok=True)

    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    configuracion_resaltado = construir_configuracion_resaltado(
        habilitado=resaltar_nueva_coleccion,
        nombre_paleta=paleta_resaltado,
    )

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

                estilos_tabla = [
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

                if configuracion_resaltado['habilitado']:
                    for fila_idx, prod in enumerate(grupos[grupo], start=1):
                        if referencia_en_nueva_coleccion(prod.get('referencia')):
                            estilos_tabla.extend(
                                [
                                    ('BACKGROUND', (0, fila_idx), (-1, fila_idx), configuracion_resaltado['fondo']),
                                    ('TEXTCOLOR', (0, fila_idx), (-1, fila_idx), configuracion_resaltado['texto']),
                                    ('LINEBELOW', (0, fila_idx), (-1, fila_idx), 1, configuracion_resaltado['borde']),
                                ]
                            )

                tabla.setStyle(TableStyle(estilos_tabla))
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
"""
Sistema de Gestión de Inventario y Reportes de Agotados
Este script analiza el inventario de productos y genera reportes PDF 
detallando los productos con bajo stock en la bodega principal.
"""

import sys
import os
import pandas as pd
from pathlib import Path
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER
from datetime import datetime
from tkinter import filedialog
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox
try:
    import customtkinter as ctk
except Exception:
    ctk = None
import threading
import re
import logging
import csv
import platform
import subprocess
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportstock_core import (
    DEFAULT_PALETA_RESALTADO,
    NUEVA_COLECCION_REFERENCIAS,
    PALETAS_RESALTADO,
    construir_configuracion_resaltado,
    obtener_paleta_resaltado,
    referencia_en_nueva_coleccion,
)

def get_base_dir() -> Path:
    """
    Obtiene el directorio base de la aplicación, funciona tanto para el script
    como para el ejecutable generado por PyInstaller.
    """
    if getattr(sys, 'frozen', False):
        # Ejecutando como exe (PyInstaller)
        return Path(os.path.dirname(sys.executable))
    else:
        # Ejecutando como script
        return Path(os.path.dirname(os.path.abspath(__file__)))


# Logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

VENDEDORA_CON_COLUMNA_PRINCIPAL = 'MALETA NATALIA REYES'
TREEVIEW_TAG_NUEVA_COLECCION = 'nueva_coleccion'


def read_excel_smart(ruta: Path) -> pd.DataFrame:
    """
    Lee un archivo Excel seleccionando el engine apropiado según la extensión.
    Soporta .xlsx (openpyxl) y .xls (xlrd). Si no se reconoce la extensión,
    intenta leer sin especificar engine y propaga la excepción si falla.
    """
    suffix = ruta.suffix.lower()
    if suffix == '.xlsx':
        return pd.read_excel(ruta, engine='openpyxl')
    elif suffix == '.xls':
        return pd.read_excel(ruta, engine='xlrd')
    else:
        # Dejar que pandas elija o falle con un error claro
        return pd.read_excel(ruta)

# Configuración de rutas
BASE_DIR = get_base_dir()
RUTA_ARCHIVO = BASE_DIR / 'data' / 'ACTUALIZACION INVERSIONES.xls'
CARPETA_REPORTES = BASE_DIR / 'reportes'

def obtener_unicos(archivo: Path, columna: str) -> list:
    """
    Obtiene valores únicos de una columna específica del archivo Excel.
    
    Args:
        archivo (Path): Ruta al archivo Excel
        columna (str): Nombre de la columna a analizar
    
    Returns:
        list: Lista de valores únicos sin valores nulos
    """
    df = read_excel_smart(archivo)
    if columna not in df.columns:
        logging.warning("Columna %s no encontrada en %s", columna, archivo)
        return []
    return df[columna].dropna().unique().tolist()



def obtener_agotados_por_bodega(ruta_archivo: Path, cantidad_minima: int = 3, progress_callback=None, vendedores_filtrados=None) -> dict:
    """
    Identifica productos con bajo stock en la bodega PRINCIPAL para cada bodega de vendedor.
    
    El proceso sigue los siguientes pasos:
    1. Lee el archivo Excel de inventario
    2. Extrae el inventario de la bodega PRINCIPAL
    3. Identifica todas las bodegas de vendedores
    4. Para cada bodega de vendedor:
       - Analiza su inventario
       - Compara con el stock en PRINCIPAL
       - Identifica productos bajo el mínimo
    
    Args:
        ruta_archivo (Path): Ruta al archivo Excel de inventario
        cantidad_minima (int): Cantidad mínima requerida en PRINCIPAL (default: 3)
        progress_callback (callable): Función(bodega, índice, total) para actualizar progreso
    
    Returns:
        dict: Diccionario con la siguiente estructura:
            {
                'BODEGA_VENDEDOR': [
                    {
                        'referencia': 'REF123',
                        'nombre_grupo': 'GRUPO_A',
                        'cantidad_vendedor': 5,
                        'cantidad_principal': 2
                    },
                    ...
                ],
                ...
            }
    """
    # Leer el archivo
    df = read_excel_smart(ruta_archivo)
    
    # Obtener inventario de PRINCIPAL
    principal = df[df['BODEGA'] == 'PRINCIPAL']
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
    
    # Obtener bodegas de vendedores (todas menos PRINCIPAL)
    bodegas_vendedores = df[df['BODEGA'] != 'PRINCIPAL']['BODEGA'].unique().tolist()

    # Aplicar filtro opcional por vendedor
    if vendedores_filtrados:
        vendedores_set = {str(v).strip() for v in vendedores_filtrados if str(v).strip()}
        bodegas_vendedores = [b for b in bodegas_vendedores if str(b).strip() in vendedores_set]
    
    # Diccionario para almacenar agotados por bodega
    agotados_por_bodega = {}
    
    for idx, bodega in enumerate(bodegas_vendedores):
        if progress_callback:
            progress_callback(bodega, idx, len(bodegas_vendedores))
        # Obtener productos de esta bodega
        productos_vendedor = df[df['BODEGA'] == bodega]
        
        # Lista de productos agotados para este vendedor
        lista_agotados = []
        
        for _, fila in productos_vendedor.iterrows():
            referencia = fila.get('REFERENCIA')
            cantidad_vendedor = fila.get('CANTIDAD', 0)
            nombre_grupo = fila.get('NOMBRE DEL GRUPO') or ''

            try:
                cantidad_vendedor = int(cantidad_vendedor) if pd.notna(cantidad_vendedor) else 0
            except Exception:
                cantidad_vendedor = 0

            # Obtener cantidad en PRINCIPAL (0 si no existe)
            cantidad_principal = stock_principal.get(referencia, 0)

            # Si en PRINCIPAL hay menor o igual a la cantidad mínima, agregar a agotados
            if referencia is not None and cantidad_vendedor > 0 and cantidad_principal <= cantidad_minima:
                lista_agotados.append({
                    'referencia': referencia,
                    'nombre_grupo': nombre_grupo,
                    'cantidad_vendedor': cantidad_vendedor,
                    'cantidad_principal': cantidad_principal
                })

        # Ordenar la lista una sola vez por bodega
        lista_agotados = sorted(lista_agotados, key=lambda x: (x.get('nombre_grupo') or '', x.get('referencia') or ''))


        agotados_por_bodega[bodega] = lista_agotados
    
    if progress_callback:
        progress_callback('Completado', len(bodegas_vendedores), len(bodegas_vendedores))
    return agotados_por_bodega








def generar_pdf_agotados(
    agotados: dict,
    carpeta_salida: str = 'reportes',
    empresa: str = '',
    resaltar_nueva_coleccion: bool = False,
    paleta_resaltado: str | None = None,
) -> None:
    """
    Genera un PDF por cada bodega con su informe de productos agotados
    
    Args:
        agotados (dict): Diccionario con bodegas y sus productos agotados
        carpeta_salida (str): Carpeta donde guardar los PDFs
        empresa (str): Nombre de la empresa (para prefijo de archivo)
    """
    # Crear carpeta si no existe
    Path(carpeta_salida).mkdir(exist_ok=True)
    
    # Fecha actual
    fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M")
    configuracion_resaltado = construir_configuracion_resaltado(
        habilitado=resaltar_nueva_coleccion,
        nombre_paleta=paleta_resaltado,
    )

    # Registrar fuente TTF si existe (permite acentos y símbolos unicode)
    default_font = 'Helvetica'
    try:
        # Buscar fuente local en carpeta 'fonts' junto al script
        font_file = Path(__file__).parent / 'fonts' / 'DejaVuSans.ttf'
        if font_file.is_file():
            pdfmetrics.registerFont(TTFont('DejaVuSans', str(font_file)))
            default_font = 'DejaVuSans'
            logging.info('Registered TTF font: %s', font_file)
    except Exception as e:
        logging.warning('No TTF font registered (%s). Using fallback font.', e)
    
    # Sanitize empresa prefix
    empresa_prefix = ''
    if empresa:
        # Keep alphanum, dash and underscore
        empresa_clean = re.sub(r'[^0-9A-Za-z-_]', '_', empresa.strip())
        empresa_prefix = f"{empresa_clean}_"

    for bodega, productos in agotados.items():
        mostrar_columna_principal = str(bodega).strip().upper() == VENDEDORA_CON_COLUMNA_PRINCIPAL
        # Nombre del archivo PDF (limpio, sin caracteres especiales) con timestamp
        bodega_clean = re.sub(r'[^0-9A-Za-z-_]', '_', bodega.strip())
        empresa_prefix = ''
        if empresa:
            # Keep alphanum, dash and underscore
            empresa_clean = re.sub(r'[^0-9A-Za-z-_]', '_', empresa.strip())
            empresa_prefix = f"{empresa_clean}_"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"{empresa_prefix}{bodega_clean}_agotados_{timestamp}.pdf"
        ruta_pdf = Path(carpeta_salida) / nombre_archivo
        
        # Crear documento PDF
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
            creator='ReportStock'
        )
        
        # Estilos
        estilos = getSampleStyleSheet()
        estilo_titulo = ParagraphStyle(
            'CustomTitle',
            parent=estilos['Heading1'],
            fontSize=18,
            textColor=colors.HexColor('#2c3e50'),
            spaceAfter=30,
            alignment=TA_CENTER,
            fontName=default_font
        )
        
        estilo_subtitulo = ParagraphStyle(
            'CustomSubtitle',
            parent=estilos['Heading2'],
            fontSize=12,
            textColor=colors.HexColor('#34495e'),
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName=default_font
        )
        
        estilo_grupo = ParagraphStyle(
            'GrupoHeader',
            parent=estilos['Heading3'],
            fontSize=11,
            textColor=colors.HexColor('#2980b9'),
            spaceAfter=10,
            spaceBefore=15,
            fontName=default_font
        )
        
        # Contenedor de elementos
        elementos = []
        
        # Mapeo de empresas a nombres completos
        nombres_empresas = {
            'Inversiones': 'INVERSIONES RUEDA S.A.S',
            'Lamar': 'LAMAR OPTICAL S.A.S',
            'Sin_Empresa': 'REPORTE DE PRODUCTOS AGOTADOS'
        }
        
        # Título principal: nombre completo de la empresa
        titulo_text = nombres_empresas.get(empresa, 'REPORTE DE PRODUCTOS AGOTADOS')
        titulo = Paragraph(f"{titulo_text}", estilo_titulo)
        elementos.append(titulo)
        
        # Información de la bodega
        subtitulo = Paragraph(f"<b>Bodega:</b> {bodega}", estilo_subtitulo)
        elementos.append(subtitulo)
        
        fecha = Paragraph(f"<b>Fecha:</b> {fecha_actual}", estilo_subtitulo)
        elementos.append(fecha)
        
        elementos.append(Spacer(1, 20))
        
        # Resumen
        total = len(productos)
        resumen = Paragraph(
            f"<b>Total de productos agotados:</b> {total}",
            estilos['Normal']
        )
        elementos.append(resumen)
        elementos.append(Spacer(1, 20))
        
        if productos:
            # Agrupar productos por nombre_grupo
            grupos = {}
            for prod in productos:
                grupo = prod.get('nombre_grupo') or 'Sin Grupo'
                if grupo not in grupos:
                    grupos[grupo] = []
                grupos[grupo].append(prod)
            
            # Iterar por grupos ordenados
            for grupo in sorted(grupos.keys()):
                elementos.append(Paragraph(f"<b>{grupo}</b>", estilo_grupo))
                
                # Crear tabla para este grupo
                if mostrar_columna_principal:
                    datos_tabla = [
                        ['Referencia', 'Cant. Vendedor', 'Cant. Principal']
                    ]
                else:
                    datos_tabla = [
                        ['Referencia', 'Cant. Vendedor']
                    ]
                
                for prod in grupos[grupo]:
                    if mostrar_columna_principal:
                        datos_tabla.append([
                            str(prod['referencia']),
                            str(prod['cantidad_vendedor']),
                            str(prod['cantidad_principal'])
                        ])
                    else:
                        datos_tabla.append([
                            str(prod['referencia']),
                            str(prod['cantidad_vendedor'])
                        ])
                
                if mostrar_columna_principal:
                    tabla = Table(datos_tabla, colWidths=[2.2*inch, 1.3*inch, 1.3*inch])
                else:
                    tabla = Table(datos_tabla, colWidths=[2.5*inch, 1.5*inch])
                
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
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
                ]

                if configuracion_resaltado['habilitado']:
                    for fila_idx, prod in enumerate(grupos[grupo], start=1):
                        if referencia_en_nueva_coleccion(prod.get('referencia')):
                            estilos_tabla.extend([
                                ('BACKGROUND', (0, fila_idx), (-1, fila_idx), configuracion_resaltado['fondo']),
                                ('TEXTCOLOR', (0, fila_idx), (-1, fila_idx), configuracion_resaltado['texto']),
                                ('LINEBELOW', (0, fila_idx), (-1, fila_idx), 1, configuracion_resaltado['borde']),
                            ])

                tabla.setStyle(TableStyle(estilos_tabla))
                
                elementos.append(tabla)
                elementos.append(Spacer(1, 15))
        else:
            # Sin productos agotados
            mensaje = Paragraph(
                "<b>No hay productos agotados en esta bodega</b>",
                estilos['Normal']
            )
            elementos.append(mensaje)
        
        # Generar PDF
        try:
            doc.build(elementos)
            logging.info('PDF generado: %s', ruta_pdf)
        except Exception as e:
            logging.exception('Error generando PDF %s: %s', ruta_pdf, e)


def exportar_csv(datos: list, ruta_archivo: Path) -> None:
    """
    Exporta lista de productos a un archivo CSV.
    
    Args:
        datos (list): Lista de diccionarios con keys: referencia, nombre_grupo, cantidad_vendedor, cantidad_principal
        ruta_archivo (Path): Ruta de salida del archivo CSV
    """
    if not datos:
        logging.warning('No hay datos para exportar a CSV')
        return
    
    try:
        with open(ruta_archivo, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['nombre_grupo', 'referencia', 'cantidad_vendedor', 'cantidad_principal'])
            writer.writeheader()
            writer.writerows(datos)
        logging.info('CSV exportado: %s', ruta_archivo)
    except Exception as e:
        logging.exception('Error exportando CSV: %s', e)


def abrir_carpeta(ruta: str) -> None:
    """
    Abre una carpeta en el explorador del sistema operativo.
    """
    try:
        if platform.system() == 'Windows':
            os.startfile(ruta)
        elif platform.system() == 'Darwin':
            subprocess.Popen(['open', ruta])
        else:
            subprocess.Popen(['xdg-open', ruta])
        logging.info('Abriendo carpeta: %s', ruta)
    except Exception as e:
        logging.exception('Error abriendo carpeta: %s', e)


# Uso completo
def generar_reporte_completo(ruta: Path, cantidad_minima: int = 3, carpeta_salida: str = 'reportes', empresa: str = '', progress_callback=None) -> None:
    """
    Función principal que genera todos los reportes de productos agotados.
    
    Args:
        ruta (Path): Ruta al archivo Excel de inventario
        cantidad_minima (int): Cantidad mínima requerida en PRINCIPAL
        carpeta_salida (str): Carpeta donde guardar los PDFs
        empresa (str): Nombre de la empresa (para prefijo de archivo)
    """
    # Obtener agotados
    agotados = obtener_agotados_por_bodega(ruta, cantidad_minima, progress_callback=progress_callback)
    
    # Generar PDFs
    generar_pdf_agotados(agotados, carpeta_salida=carpeta_salida, empresa=empresa)
    
    # Mostrar resumen
    print("\n" + "="*60)
    print("RESUMEN DE REPORTES GENERADOS")
    print("="*60)
    for bodega, productos in agotados.items():
        print(f"{bodega}: {len(productos)} productos agotados")


if __name__ == '__main__':
    if ctk is None:
        message = (
            "La librería 'customtkinter' no está instalada.\n"
            "Instala la dependencia con: pip install customtkinter"
        )
        print(message)
        messagebox.showerror("Dependencia faltante", message)
    else:
        # Inicializar CustomTkinter
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # Ventana principal CTk
        ventana = ctk.CTk()
        ventana.title("Sistema de Reportes - Actualización Maletas")
        ventana.geometry("1000x680")

        # Colores definidos (para compatibilidad visual con CTk)
        COLORS = {
            'primary': '#2c3e50',
            'accent': '#3498db',
            'success': '#27ae60',
            'card': '#ffffff',
            'muted': '#7f8c8d',
            'bg': '#f5f7fa'
        }

        # Header
        header = ctk.CTkFrame(ventana, corner_radius=0, fg_color=COLORS['primary'])
        header.pack(fill='x')
        header.configure(height=88)
        titulo = ctk.CTkLabel(header, text='Reporte de Productos Agotados', text_color='white', font=('Segoe UI', 18, 'bold'))
        subtitulo = ctk.CTkLabel(header, text='Sistema de Gestión de Inventario', text_color='#ecf0f1', font=('Segoe UI', 11))
        titulo.pack(pady=(12, 0))
        subtitulo.pack(pady=(0, 12))

        # Main container
        main_container = ctk.CTkFrame(ventana, fg_color=COLORS['bg'], corner_radius=0)
        main_container.pack(fill='both', expand=True, padx=16, pady=12)

        # Left controls frame
        left_frame = ctk.CTkFrame(main_container, width=320, fg_color=COLORS['card'], corner_radius=8)
        left_frame.pack(side='left', fill='y', padx=(0, 12), pady=6)

        # Empresa
        empresa_var = tk.StringVar(value='')
        lbl_empresa = ctk.CTkLabel(left_frame, text='Empresa', anchor='w', font=('Segoe UI', 11, 'bold'))
        lbl_empresa.pack(fill='x', padx=16, pady=(14, 6))
        rb1 = ctk.CTkRadioButton(left_frame, text='Inversiones Rueda', variable=empresa_var, value='Inversiones')
        rb2 = ctk.CTkRadioButton(left_frame, text='Lamar Optical', variable=empresa_var, value='Lamar')
        rb1.pack(anchor='w', padx=16)
        rb2.pack(anchor='w', padx=16, pady=(0, 8))

        # Archivo
        label_ruta = ctk.CTkLabel(left_frame, text='No has seleccionado ningún archivo', wraplength=260, anchor='w')
        boton_archivo = ctk.CTkButton(left_frame, text='Cargar archivo Excel', fg_color=COLORS['accent'], hover_color='#5aaef0')
        boton_archivo.pack(fill='x', padx=16, pady=(8, 6))
        label_ruta.pack(fill='x', padx=16)

        # Filtro por vendedor (se carga después de seleccionar archivo)
        vendedor_var = tk.StringVar(value='Todos')
        vendedores_disponibles = ['Todos']
        lbl_vendedor = ctk.CTkLabel(left_frame, text='Vendedor', anchor='w')
        lbl_vendedor.pack(fill='x', padx=16, pady=(10, 4))
        combo_vendedor = ttk.Combobox(left_frame, textvariable=vendedor_var, values=vendedores_disponibles, state='readonly')
        combo_vendedor.pack(fill='x', padx=16, pady=(0, 6))
        
        # Progress bar (determinado)
        progress_var = tk.DoubleVar(value=0.0)
        progress = ttk.Progressbar(left_frame, mode='determinate', variable=progress_var, maximum=100)
        progress.pack(fill='x', padx=16, pady=(8, 6))
        progress_label = ctk.CTkLabel(left_frame, text='', text_color=COLORS['muted'], font=('Segoe UI', 10))
        progress_label.pack(fill='x', padx=16, pady=(0, 6))

        # Carpeta
        label_carpeta = ctk.CTkLabel(left_frame, text='Carpeta de destino no seleccionada', wraplength=260, anchor='w')
        boton_carpeta = ctk.CTkButton(left_frame, text='Elegir carpeta de salida', fg_color='#34495e', hover_color='#435866')
        boton_carpeta.pack(fill='x', padx=16, pady=(12, 6))
        label_carpeta.pack(fill='x', padx=16)

        # Filtro
        cantidad_var = tk.IntVar(value=3)
        lbl_cantidad = ctk.CTkLabel(left_frame, text='Filtro: Cantidad mínima en PRINCIPAL', anchor='w')
        lbl_cantidad.pack(fill='x', padx=16, pady=(12, 4))
        try:
            spin_cantidad = ttk.Spinbox(left_frame, from_=0, to=1000, textvariable=cantidad_var, width=10)
        except Exception:
            spin_cantidad = tk.Spinbox(left_frame, from_=0, to=1000, textvariable=cantidad_var, width=10)
        spin_cantidad.pack(anchor='w', padx=16, pady=(0, 8))

        # Resaltado de nueva coleccion
        resaltar_nueva_coleccion_var = tk.BooleanVar(value=False)
        paleta_resaltado_var = tk.StringVar(value=DEFAULT_PALETA_RESALTADO)
        lbl_resaltado = ctk.CTkLabel(left_frame, text='Nueva coleccion', anchor='w')
        lbl_resaltado.pack(fill='x', padx=16, pady=(6, 4))
        chk_resaltado = ctk.CTkCheckBox(
            left_frame,
            text=f'Resaltar referencias de la ultima coleccion ({len(NUEVA_COLECCION_REFERENCIAS)})',
            variable=resaltar_nueva_coleccion_var,
            onvalue=True,
            offvalue=False,
        )
        chk_resaltado.pack(fill='x', padx=16, pady=(0, 6))
        combo_paleta = ttk.Combobox(
            left_frame,
            textvariable=paleta_resaltado_var,
            values=list(PALETAS_RESALTADO.keys()),
            state='disabled',
        )
        combo_paleta.pack(fill='x', padx=16, pady=(0, 4))
        lbl_paleta = ctk.CTkLabel(
            left_frame,
            text='Elige la gama de colores para las filas resaltadas.',
            wraplength=260,
            anchor='w',
            text_color=COLORS['muted'],
            font=('Segoe UI', 10),
        )
        lbl_paleta.pack(fill='x', padx=16, pady=(0, 8))

        # Acción principal
        status_var = tk.StringVar(value='Listo')
        boton_generar = ctk.CTkButton(left_frame, text='Generar PDF', fg_color=COLORS['success'], hover_color='#33bc54', corner_radius=6)
        boton_generar.pack(fill='x', padx=16, pady=(8, 12))
        status_label = ctk.CTkLabel(left_frame, textvariable=status_var, anchor='w', text_color=COLORS['muted'])
        status_label.pack(fill='x', padx=16, pady=(0, 12))
        
        # Botones auxiliares: CSV y Abrir carpeta
        btn_frame = ctk.CTkFrame(left_frame, fg_color=COLORS['card'], corner_radius=0)
        btn_frame.pack(fill='x', padx=16, pady=(6, 12))
        boton_csv = ctk.CTkButton(btn_frame, text='Exportar CSV', fg_color='#16a34a', hover_color='#22c55e', font=('Segoe UI', 10), height=32)
        boton_csv.pack(side='left', fill='both', expand=True, padx=(0, 6))
        boton_abrir = ctk.CTkButton(btn_frame, text='Abrir Carpeta', fg_color='#8b5cf6', hover_color='#a78bfa', font=('Segoe UI', 10), height=32)
        boton_abrir.pack(side='left', fill='both', expand=True)
        
        # Buscar en la vista previa
        buscar_var = tk.StringVar(value='')
        lbl_buscar = ctk.CTkLabel(left_frame, text='Buscar referencia/grupo', anchor='w')
        lbl_buscar.pack(fill='x', padx=16)
        entry_buscar = ctk.CTkEntry(left_frame, textvariable=buscar_var)
        entry_buscar.pack(fill='x', padx=16, pady=(4, 8))

        # Right / main content
        right_frame = ctk.CTkFrame(main_container, fg_color=COLORS['card'], corner_radius=8)
        right_frame.pack(side='left', fill='both', expand=True, pady=6)

        preview_title = ctk.CTkLabel(right_frame, text='Vista previa (vacía hasta generar)', anchor='w', font=('Segoe UI', 12, 'bold'))
        preview_title.pack(fill='x', padx=12, pady=(12, 6))

        # Treeview dentro de un Frame compatible
        tree_container = tk.Frame(right_frame, bg=COLORS['card'])
        tree_container.pack(fill='both', expand=True, padx=12, pady=(0,12))

        columns = ('grupo', 'referencia', 'cant_v', 'cant_p')
        tree = ttk.Treeview(tree_container, columns=columns, show='headings', selectmode='browse')
        tree.heading('grupo', text='Nombre del Grupo')
        tree.heading('referencia', text='Referencia')
        tree.heading('cant_v', text='Cant. Vendedor')
        tree.heading('cant_p', text='Cant. Principal')
        tree.column('grupo', width=320)
        tree.column('referencia', width=160)
        tree.column('cant_v', width=120, anchor='center')
        tree.column('cant_p', width=120, anchor='center')
        tree.pack(fill='both', expand=True)

        # Keep preview data for filtering and sorting
        preview_data = []

        def configure_tree_highlight_tag():
            paleta = obtener_paleta_resaltado(paleta_resaltado_var.get())
            tree.tag_configure(
                TREEVIEW_TAG_NUEVA_COLECCION,
                background=paleta['fondo'],
                foreground=paleta['texto'],
            )

        def insert_preview_row(prod):
            tags = ()
            if resaltar_nueva_coleccion_var.get() and referencia_en_nueva_coleccion(prod.get('referencia')):
                tags = (TREEVIEW_TAG_NUEVA_COLECCION,)
            tree.insert(
                '',
                'end',
                values=(prod['nombre_grupo'], prod['referencia'], prod['cantidad_vendedor'], prod['cantidad_principal']),
                tags=tags,
            )

        def refresh_preview():
            q = buscar_var.get().strip().lower()
            configure_tree_highlight_tag()
            for item in tree.get_children():
                tree.delete(item)
            for prod in preview_data:
                if not q or q in str(prod['nombre_grupo']).lower() or q in str(prod['referencia']).lower():
                    insert_preview_row(prod)

        def update_highlight_controls(*_args):
            combo_paleta.configure(state='readonly' if resaltar_nueva_coleccion_var.get() else 'disabled')
            refresh_preview()

        # Sorting helper
        def treeview_sort_column(tv, col, reverse=False):
            try:
                l = [(tv.set(k, col), k) for k in tv.get_children('')]
                try:
                    l.sort(key=lambda t: float(t[0]) if t[0].replace('.', '', 1).isdigit() else t[0], reverse=reverse)
                except Exception:
                    l.sort(key=lambda t: t[0], reverse=reverse)
                for index, (_, k) in enumerate(l):
                    tv.move(k, '', index)
                # reverse sort next time
                tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))
            except Exception as e:
                logging.exception('Error sorting column %s: %s', col, e)

        # Attach sorting to headings
        for col in columns:
            tree.heading(col, text=tree.heading(col)['text'], command=lambda _col=col: treeview_sort_column(tree, _col, False))

        # Filter helper
        def apply_filter():
            refresh_preview()

        entry_buscar.bind('<KeyRelease>', lambda e: apply_filter())
        combo_paleta.bind('<<ComboboxSelected>>', update_highlight_controls)
        chk_resaltado.configure(command=update_highlight_controls)

        # Footer
        footer = ctk.CTkLabel(ventana, text='Desarrollado por: Yerson Vargas', fg_color='transparent', text_color=COLORS['muted'])
        footer.pack(side='bottom', fill='x', pady=(6, 6))

        # Funciones de selección (mantienen filedialog de tkinter)
        def seleccionar_archivo():
            archivo = filedialog.askopenfilename(
                title="Selecciona un archivo",
                filetypes=[
                    ("Archivos Excel", "*.xls *.xlsx"),
                    ("Todos los archivos", "*.*")
                ]
            )
            if archivo:
                label_ruta.configure(text=archivo)
                try:
                    df_tmp = read_excel_smart(Path(archivo))
                    if 'BODEGA' in df_tmp.columns:
                        vendedores = (
                            df_tmp['BODEGA']
                            .dropna()
                            .astype(str)
                            .loc[lambda s: s.str.upper().str.strip() != 'PRINCIPAL']
                            .unique()
                            .tolist()
                        )
                    else:
                        vendedores = []

                    vendedores = sorted([v for v in vendedores if str(v).strip()])
                    nuevos_valores = ['Todos'] + vendedores
                    combo_vendedor.configure(values=nuevos_valores)
                    vendedor_var.set('Todos')
                except Exception as e:
                    combo_vendedor.configure(values=['Todos'])
                    vendedor_var.set('Todos')
                    logging.warning('No se pudieron cargar vendedores del archivo: %s', e)

        def seleccionar_carpeta():
            carpeta = filedialog.askdirectory(title="Seleccionar carpeta de destino")
            if carpeta:
                label_carpeta.configure(text=carpeta)

        boton_archivo.configure(command=seleccionar_archivo)
        boton_carpeta.configure(command=seleccionar_carpeta)

        # Worker para generar reportes (reusa lógica existente)
        def progress_handler(bodega, idx, total):
            """Actualiza barra de progreso desde callback de obtener_agotados_por_bodega"""
            if total > 0:
                porcentaje = (idx / total) * 100
                progress_var.set(porcentaje)
                progress_label.configure(text=f'{bodega}: {int(porcentaje)}%')
                ventana.update_idletasks()

        def export_csv_handler():
            """Exporta preview_data a CSV"""
            if not preview_data:
                messagebox.showwarning("CSV", "No hay datos para exportar. Genera un reporte primero.")
                return
            carpeta = label_carpeta.cget('text')
            if carpeta.startswith('Carpeta de destino'):
                carpeta = str(CARPETA_REPORTES)
            ruta_csv = Path(carpeta) / f"agotados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            exportar_csv(preview_data, ruta_csv)
            messagebox.showinfo("CSV", f"Archivo exportado a:\n{ruta_csv}")

        def abrir_carpeta_handler():
            """Abre la carpeta de destino"""
            carpeta = label_carpeta.cget('text')
            if carpeta.startswith('Carpeta de destino'):
                carpeta = str(CARPETA_REPORTES)
            if Path(carpeta).is_dir():
                abrir_carpeta(carpeta)
            else:
                messagebox.showerror("Error", "La carpeta de destino no existe.")

        boton_csv.configure(command=export_csv_handler)
        boton_abrir.configure(command=abrir_carpeta_handler)

        # Worker para generar reportes (reusa lógica existente)
        def _worker_generar_reporte():
            try:
                status_var.set('Generando reportes...')
                boton_generar.configure(state='disabled')
                progress_var.set(0)
                progress_label.configure(text='Iniciando...')

                selected_file = label_ruta.cget('text')
                carpeta_destino = label_carpeta.cget('text')
                empresa_sel = empresa_var.get()

                if not selected_file or selected_file.startswith('No has seleccionado') or not Path(selected_file).is_file():
                    messagebox.showerror("Error", "Seleccione un archivo Excel válido antes de generar el reporte.")
                    status_var.set('Listo')
                    boton_generar.configure(state='normal')
                    return

                if not empresa_sel:
                    empresa_sel = "Sin_Empresa"

                if not carpeta_destino or carpeta_destino.startswith('Carpeta de destino'):
                    carpeta_destino = str(CARPETA_REPORTES)
                Path(carpeta_destino).mkdir(parents=True, exist_ok=True)

                try:
                    cantidad_minima = int(cantidad_var.get())
                    if cantidad_minima < 0:
                        raise ValueError()
                except Exception:
                    messagebox.showerror("Error", "El filtro de cantidad debe ser un número entero mayor o igual a 0.")
                    status_var.set('Listo')
                    boton_generar.configure(state='normal')
                    return

                vendedor_sel = vendedor_var.get().strip()
                vendedores_filtrados = None if not vendedor_sel or vendedor_sel == 'Todos' else [vendedor_sel]

                agotados = obtener_agotados_por_bodega(
                    Path(selected_file),
                    cantidad_minima,
                    progress_callback=progress_handler,
                    vendedores_filtrados=vendedores_filtrados
                )
                generar_pdf_agotados(
                    agotados,
                    carpeta_salida=carpeta_destino,
                    empresa=empresa_sel,
                    resaltar_nueva_coleccion=resaltar_nueva_coleccion_var.get(),
                    paleta_resaltado=paleta_resaltado_var.get(),
                )

                # Prepare preview_data and schedule GUI update on main thread
                preview = []
                for bodega, productos in agotados.items():
                    for prod in productos:
                        preview.append(prod)
                    break

                def update_preview():
                    # update the outer preview_data list in-place to avoid nonlocal
                    preview_data.clear()
                    preview_data.extend(preview)
                    del preview_data[200:]
                    refresh_preview()
                    status_var.set('Generación finalizada')
                    messagebox.showinfo("Reporte", "Generación de reportes finalizada.")

                ventana.after(50, update_preview)
            except Exception as e:
                status_var.set('Error')
                messagebox.showerror("Error", f"Error generando reportes:\n{e}")
            finally:
                boton_generar.configure(state='normal')
                progress_label.configure(text='')

        def on_generar_clicked():
            threading.Thread(target=_worker_generar_reporte, daemon=True).start()

        boton_generar.configure(command=on_generar_clicked)
        update_highlight_controls()

        ventana.mainloop()
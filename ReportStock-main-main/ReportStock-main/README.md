# Main.py - Sistema de Gestión de Inventario y Reportes de Agotados

## 📋 Descripción General

`main.py` es la aplicación principal que proporciona una **interfaz gráfica completa (GUI) con Tkinter** para generar reportes PDF automáticos de productos con bajo stock en la bodega principal. Es la versión recomendada para el uso diario.

**Autor**: Yerson Vargas

## 🚀 Inicio Rápido

### Instalación de dependencias
```bash
pip install -r requirements.txt
```

### Ejecutar la aplicación
```bash
python main.py
```

### Ejecutar versión web (nueva)
```bash
python web_app.py
```

Abre en el navegador:

```text
http://localhost:8071
```

## 📦 Dependencias

```
pandas>=1.3.0      # Lectura y manipulación de datos Excel
xlrd>=2.0.1        # Soporte para archivos .xls
reportlab>=3.6.0   # Generación de reportes PDF
tkinter            # (Incluido en Python) GUI
Flask>=3.0.0       # Backend web
gunicorn>=22.0.0   # Servidor WSGI para producción
```

## 🌐 Migración a Web

Se agregó una versión web mínima en el archivo `web_app.py` usando Flask.

### Flujo web
1. Subir archivo Excel (`.xls` o `.xlsx`)
2. Definir empresa, cantidad mínima y vendedor opcional
3. Generar reportes PDF por bodega
4. Descargar un `.zip` con todos los PDFs

### Archivos nuevos para web
- `reportstock_core.py`: lógica de negocio reutilizable (sin GUI)
- `web_app.py`: backend Flask
- `web_templates/index.html`: formulario web

### Despliegue sugerido (Linux)

Para producción:

```bash
gunicorn -w 2 -b 0.0.0.0:8071 web_app:app
```

Puedes publicarlo en servicios como Render, Railway, VPS con Nginx, o similar.

## 🎯 Funcionalidades Principales

### 1. Interfaz Gráfica Intuitiva
- Aplicación de escritorio con ventana Tkinter
- Diseño limpio y organizado con ttk (themed widgets)
- Responsive y fácil de usar

### 2. Selección de Empresa
Elige entre dos opciones:
- **Inversiones Rueda**: Genera PDFs con título "INVERSIONES RUEDA S.A.S"
- **Lamar Optical**: Genera PDFs con título "LAMAR OPTICAL S.A.S"
- **(Opcional)** Sin seleccionar: Usa "Sin_Empresa" como prefijo

### 3. Selección de Archivo Excel
- Abre un diálogo para seleccionar archivo
- Soporta formatos `.xls` y `.xlsx`
- Valida que el archivo exista antes de procesarlo

### 4. Selección de Carpeta de Destino
- Elige dónde guardar los reportes PDF
- Si no seleccionas, usa la carpeta `reportes/` por defecto
- Se crea automáticamente si no existe

### 5. Filtro de Cantidad Mínima
- Spinbox ajustable (0 a 1000)
- Valor por defecto: 3
- Incluye productos con: `cantidad_principal <= cantidad_minima`

### 6. Generación de Reportes
- Botón "Generar Reporte" que inicia el proceso
- Se ejecuta en hilo separado (no bloquea la UI)
- Muestra estado en tiempo real
- Validación completa de inputs

## 📊 Estructura de la Aplicación

### Funciones Principales

#### `get_base_dir() -> Path`
Obtiene el directorio base de ejecución. Funciona tanto como script Python como ejecutable generado por PyInstaller.

#### `obtener_unicos(archivo: Path, columna: str) -> list`
Extrae valores únicos de una columna del Excel (no utilizada actualmente, disponible para futuras extensiones).

#### `obtener_agotados_por_bodega(ruta_archivo: Path, cantidad_minima: int = 3) -> dict`
**Función core del análisis**:
1. Lee el archivo Excel completo
2. Extrae inventario de bodega PRINCIPAL
3. Identifica todas las bodegas de vendedores
4. Para cada bodega de vendedor:
   - Filtra productos donde `cantidad_principal <= cantidad_minima`
   - Ordena alfabéticamente por nombre de grupo
5. Retorna diccionario: `{bodega: [productos_agotados]}`

#### `generar_pdf_agotados(agotados: dict, carpeta_salida: str, empresa: str) -> None`
**Genera los PDFs**:
- Mapea nombres de empresa a títulos completos (S.A.S)
- Para cada bodega crea un PDF con:
  - Título personalizado (nombre empresa)
  - Información bodega y fecha
  - Tabla profesional con productos
  - Estilos y colores predefinidos
- Nombres de archivo: `[Empresa]_[Bodega]_agotados.pdf`

#### `generar_reporte_completo(ruta: Path, cantidad_minima: int, carpeta_salida: str, empresa: str) -> None`
Orquestador que:
1. Llama a `obtener_agotados_por_bodega()`
2. Llama a `generar_pdf_agotados()`
3. Imprime resumen en consola

### Interfaz Tkinter (if __name__ == '__main__')

**Ventana principal**:
- Título: "Sistema de Reportes - Actualización Maletas"
- Tamaño: 600x380 píxeles
- Tema: clam (moderno)

**Componentes principales**:

1. **Título**: Etiqueta grande "Reporte de Productos Agotados"

2. **Sección Empresa** (LabelFrame):
   - Radio button: "Inversiones Rueda" → valor: 'Inversiones'
   - Radio button: "Lamar Optical" → valor: 'Lamar'

3. **Sección Archivos** (Frame con grid):
   - Botón: "Seleccionar archivo" → abre diálogo
   - Label: Muestra ruta seleccionada
   - Botón: "Seleccionar carpeta de destino" → abre diálogo
   - Label: Muestra carpeta seleccionada
   - Spinbox: Filtro cantidad (0-1000, default=3)

4. **Status Label**: Muestra estado actual ("Listo", "Generando...", "Error", etc.)

5. **Botón Acción**: "Generar Reporte" (alineado a la derecha)

6. **Firma**: "Desarrollado por: Yerson Vargas" (pie de ventana)

## 🔄 Flujo de Ejecución

```
┌─────────────────────────────────┐
│  Inicio de Aplicación           │
│  - Crear ventana Tkinter        │
│  - Mostrar UI                   │
└────────────┬────────────────────┘
             │
             ▼
    ┌─────────────────────┐
    │ Usuario selecciona: │
    │ - Empresa           │
    │ - Archivo Excel     │
    │ - Carpeta destino   │
    │ - Cantidad mínima   │
    └──────────┬──────────┘
             │
             ▼
    ┌─────────────────────────────┐
    │ Clic en "Generar Reporte"   │
    │ (Inicia hilo secundario)    │
    └──────────┬──────────────────┘
             │
             ▼
    ┌─────────────────────────┐
    │ Validar inputs:         │
    │ ✓ Archivo existe        │
    │ ✓ Cantidad >= 0         │
    └──────────┬──────────────┘
             │
             ▼
    ┌─────────────────────────┐
    │ obtener_agotados...()   │
    │ - Leer Excel            │
    │ - Filtrar por cantidad  │
    │ - Ordenar por grupo     │
    └──────────┬──────────────┘
             │
             ▼
    ┌─────────────────────────┐
    │ generar_pdf_agotados()  │
    │ - Crear PDF por bodega  │
    │ - Aplicar estilos       │
    │ - Guardar archivos      │
    └──────────┬──────────────┘
             │
             ▼
    ┌─────────────────────────┐
    │ Mostrar resultado:      │
    │ - Messagebox exito      │
    │ - Status actualizado    │
    └─────────────────────────┘
```

## 📄 Formato de Archivo Excel Esperado

### Columnas requeridas:
- **BODEGA**: Nombre bodega (PRINCIPAL, Vendedor 1, Vendedor 2, etc.)
- **REFERENCIA**: Código producto (REF001, P-123, etc.)
- **NOMBRE DEL GRUPO**: Categoría (GRUPO_A, LENTES, etc.)
- **CANTIDAD**: Stock disponible (número entero)

### Ejemplo:

| BODEGA | REFERENCIA | NOMBRE DEL GRUPO | CANTIDAD |
|--------|-----------|------------------|----------|
| PRINCIPAL | REF001 | GRUPO_A | 2 |
| PRINCIPAL | REF002 | GRUPO_B | 5 |
| Vendedor 1 | REF001 | GRUPO_A | 10 |
| Vendedor 1 | REF003 | GRUPO_C | 3 |
| Vendedor 2 | REF002 | GRUPO_B | 8 |
| Vendedor 2 | REF001 | GRUPO_A | 1 |

### Con `cantidad_minima=3`:
- **Vendedor 1**: REF001 aparece (2 < 3)
- **Vendedor 2**: REF001 y REF002 aparecen (2 < 3, 5 >= 3 pero igual incluir si <= 3)

## 📊 Salida: Archivos PDF Generados

### Nombre de archivo:
```
[Empresa]_[Bodega]_agotados.pdf
```

### Ejemplos:
- `Inversiones_Vendedor_1_agotados.pdf`
- `Inversiones_Vendedor_2_agotados.pdf`
- `Lamar_Vendedor_1_agotados.pdf`
- `Sin_Empresa_Bodega_agotados.pdf`

### Contenido de cada PDF:

```
┌────────────────────────────────┐
│  INVERSIONES RUEDA S.A.S       │  ← Título (nombre empresa)
├────────────────────────────────┤
│  Bodega: Vendedor 1            │
│  Fecha: 07/12/2025 14:30       │
│                                │
│  Total de productos agotados: 5│
├────────────────────────────────┤
│  Tabla:                        │
│  ┌──────┬──────┬───┬───┐      │
│  │Grupo │ Ref  │V1 │PRI│      │
│  ├──────┼──────┼───┼───┤      │
│  │GRP_A │REF01 │10 │ 2 │      │
│  │GRP_B │REF02 │ 5 │ 1 │      │
│  └──────┴──────┴───┴───┘      │
│  (V1=Cant. Vendedor, PRI=Principal)
└────────────────────────────────┘
```

## 🎨 Estilos y Colores

### Paleta de colores:
- **Título**: Gris oscuro (#2c3e50), 18pt, negrita
- **Subtítulos**: Gris azulado (#34495e), 12pt
- **Tabla encabezado**: Azul (#3498db) con texto blanco
- **Tabla filas**: Beige alternado con blanco/gris claro

### Fuentes:
- Helvetica (estándar en reportlab)
- Bold para encabezados

## ⚙️ Configuración Interna

### Variables globales:
```python
BASE_DIR = get_base_dir()  # Directorio de ejecución
RUTA_ARCHIVO = BASE_DIR / 'data' / 'ACTUALIZACION INVERSIONES.xls'
CARPETA_REPORTES = BASE_DIR / 'reportes'
```

### Mapeo de empresas:
```python
nombres_empresas = {
    'Inversiones': 'INVERSIONES RUEDA S.A.S',
    'Lamar': 'LAMAR OPTICAL S.A.S',
    'Sin_Empresa': 'REPORTE DE PRODUCTOS AGOTADOS'
}
```

## 🐛 Manejo de Errores

La aplicación valida:

1. **Archivo Excel**:
   - ✓ Que exista el archivo
   - ✓ Que sea accesible
   - ✓ Que contenga datos válidos

2. **Entrada del usuario**:
   - ✓ Cantidad mínima es número >= 0
   - ✓ Carpeta de destino es válida
   - ✓ Empresa está seleccionada (o usa default)

3. **Generación de PDF**:
   - ✓ Permisos de escritura en carpeta
   - ✓ Creación automática de carpetas
   - ✓ Manejo de caracteres especiales en nombres

## 🔄 Ejecución en Hilo Secundario

El botón "Generar Reporte" ejecuta en un hilo daemon:
```python
threading.Thread(target=_worker_generar_reporte, daemon=True).start()
```

**Ventajas**:
- La UI no se congela durante la generación
- El usuario puede cerrar la ventana en cualquier momento
- Estado actualiza en tiempo real

## 💾 Rutas de archivos

### Lectura:
```
p:\acutalizacion_maletas\
└── data\
    └── ACTUALIZACION INVERSIONES.xls
```

### Escritura (reportes):
```
p:\acutalizacion_maletas\
└── reportes\
    ├── Inversiones_Vendedor_1_agotados.pdf
    ├── Inversiones_Vendedor_2_agotados.pdf
    └── ...
```

## 🧪 Prueba Rápida

```bash
# 1. Instalar dependencias
pip install -r requirements.txt

# 2. Ejecutar
python main.py

# 3. En la GUI:
#    - Selecciona "Inversiones Rueda"
#    - Selecciona archivo Excel
#    - Deja carpeta por defecto
#    - Filtro: 3 (default)
#    - Click "Generar Reporte"

# 4. Verificar PDFs en carpeta "reportes/"
```

## 📝 Notas Técnicas

### Compatibilidad
- ✅ Python 3.7+
- ✅ Windows, macOS, Linux
- ✅ Ejecutable con PyInstaller

### Limitaciones
- Excel: Máx. 1 millón de filas (limitación de pandas)
- PDF: Sin encriptación (reportlab community edition)
- Bodegas: Sin límite, pero para mejor rendimiento < 100 bodegas

### Futuras mejoras
- [ ] Exportar a Excel además de PDF
- [ ] Gráficos de tendencias
- [ ] Email automático
- [ ] Base de datos para persistencia
- [ ] Traducción a otros idiomas

## 📞 Soporte

Para reportar bugs o sugerencias, contactar a: **Yerson Vargas**
Correo electronico: **yervargas6@gmail.com**

---

**Versión**: 2.0  
**Última actualización**: Diciembre 2025  
**Estado**: Producción ✅

**Cambios en v2.0:**
- Lectura Excel robusta (soporta .xlsx y .xls)
- Barra de progreso determinista con callbacks
- Exportación a CSV
- Botón "Abrir carpeta" multiplataforma
- PDFs con metadata y timestamps
- Agrupamiento de productos por grupo
- Tests unitarios incluidos
- PyInstaller spec para generar .exe
- Este README actualizado


# CHANGELOG - ReportStock

## [2.0.0] - 2025-12-20

### ✨ Nuevas características

- **Lectura Excel robusta**: Soporte automático para `.xlsx` (openpyxl) y `.xls` (xlrd)
- **Barra de progreso determinista**: Muestra porcentaje real por bodega durante generación
- **Exportación a CSV**: Botón para exportar vista previa a CSV con timestamp
- **Botón "Abrir Carpeta"**: Abre carpeta de destino multiplataforma (Windows/macOS/Linux)
- **PDFs mejorados**:
  - Agrupamiento de productos por "NOMBRE DEL GRUPO"
  - Metadata PDF (título, autor, asunto)
  - Timestamps en nombres de archivo: `empresa_bodega_agotados_YYYYMMDD_HHMMSS.pdf`
  - Soporte para fuentes Unicode (DejaVuSans.ttf)
- **Búsqueda en tiempo real**: Filtrado instantáneo en vista previa (grupo + referencia)
- **Orden por columna**: Click en encabezado de Treeview para ordenar
- **Logging detallado**: Usa módulo logging en lugar de print
- **Tests unitarios**: pytest para `obtener_agotados_por_bodega` y `obtener_unicos`
- **PyInstaller spec**: Archivo `.spec` para generar `.exe` standalone
- **GitHub Actions CI/CD**: 
  - Tests automáticos en Python 3.10, 3.11, 3.12
  - Linting (flake8, black, isort)
  - Build automático de .exe en releases
- **README comprensivo**: Instalación, uso, troubleshooting, ejemplos avanzados

### 🔧 Mejoras internas

- Refactorización de `obtener_agotados_por_bodega` con callback de progreso
- Manejo seguro de datos Excel (`.get()`, conversión de tipos)
- Ordenación de lista de agotados **fuera** del bucle (mejora performance)
- GUI thread-safe usando `ventana.after()` y `update_idletasks()`
- Mejor separación de concerns (funciones helper: `exportar_csv`, `abrir_carpeta`, `read_excel_smart`)
- Validación de archivo antes de procesar
- Manejo de excepciones mejorado en generación de PDFs

### 📦 Dependencias

Actualizado `requirements.txt`:
- pandas >= 1.3.0
- xlrd >= 2.0.1
- reportlab >= 3.6.0
- customtkinter >= 5.0.0
- **openpyxl >= 3.0.0** (nuevo)
- **pytest >= 7.0.0** (nuevo)

### 🐛 Bugs corregidos

- Error al acceder a columnas inexistentes en fila
- Ordenación repetida innecesariamente dentro de bucle
- Falta de manejo de valores NaN en cantidades
- Interfaz bloqueada durante generación de reportes (solucionado con threading mejorado)

### 📚 Documentación

- README.md completamente reescrito
- CHANGELOG.md (este archivo)
- Docstrings mejorados en funciones principales
- Ejemplos de uso en README y código

### 🏗️ Infraestructura

- Agregado `reportstock.spec` para PyInstaller
- Agregado `.github/workflows/tests.yml` para CI/CD en GitHub
- Estructura de directorios mejorada

### 📝 Notas de migración (desde v1.0)

1. Se recomienda usar `.xlsx` (más compatible que `.xls`)
2. Instalar nuevas dependencias: `pip install openpyxl pytest`
3. Para mejores PDFs, copiar DejaVuSans.ttf a carpeta `fonts/`
4. Variables de GUI ahora más robustas (menos crashes)
5. Archivos PDF incluyen timestamp (nombres más únicos)

### 🚀 Próximas versiones (roadmap)

- [ ] Interfaz web (Flask/Django)
- [ ] Base de datos para histórico de reportes
- [ ] Configuración persistente (config.yaml)
- [ ] Email automático de reportes
- [ ] Gráficos de tendencias de stock
- [ ] Soporte para múltiples archivos simultáneamente
- [ ] Traducción a inglés

---

## [1.0.0] - 2025-12-07 (versión anterior)

Versión inicial con funcionalidad base:
- GUI con tkinter para seleccionar archivos y generar PDFs
- Lectura básica de Excel
- Generación de PDFs por bodega
- Filtrado por cantidad mínima

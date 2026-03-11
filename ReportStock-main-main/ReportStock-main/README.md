<p align="center">
	<img src="lamar_optical_logo_no_bg.png" alt="Logo Lamar Optical" width="320" />
</p>

<h1 align="center">ReportStock</h1>

<p align="center">
	Sistema de gestion de inventario para detectar productos agotados o con bajo stock en bodega PRINCIPAL,<br>
	con salida en PDF por vendedor y soporte para escritorio y web.
</p>

<p align="center">
	<a href="main.py">Aplicacion de escritorio (Tkinter)</a> | <a href="web_app.py">Aplicacion web (Flask)</a>
</p>

<p align="center">
	<img src="https://img.shields.io/badge/Python-3.8%2B-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python" />
	<img src="https://img.shields.io/badge/GUI-Tkinter-1E3A8A?style=for-the-badge" alt="Tkinter" />
	<img src="https://img.shields.io/badge/Web-Flask-000000?style=for-the-badge&logo=flask&logoColor=white" alt="Flask" />
	<img src="https://img.shields.io/badge/PDF-ReportLab-E34F26?style=for-the-badge" alt="ReportLab" />
	<img src="https://img.shields.io/badge/Data-Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas" />
</p>

<p align="center">
	<img src="LOGO%20OPTICAL%20SHOP%20EDITABLE%201.png" alt="Logo Optical Shop" width="180" />
</p>

---

## Tabla de contenido

- [Descripcion](#descripcion)
- [Caracteristicas principales](#caracteristicas-principales)
- [Inicio rapido](#inicio-rapido)
- [Dependencias](#dependencias)
- [Flujo funcional](#flujo-funcional)
- [Formato esperado del Excel](#formato-esperado-del-excel)
- [Arquitectura del proyecto](#arquitectura-del-proyecto)
- [Endpoints web relevantes](#endpoints-web-relevantes)
- [Ejecucion en produccion](#ejecucion-en-produccion)
- [Calidad y pruebas](#calidad-y-pruebas)
- [Soporte](#soporte)
- [Estado del proyecto](#estado-del-proyecto)

---

## Descripcion

ReportStock procesa un archivo Excel de inventario, compara la disponibilidad por referencia contra la bodega PRINCIPAL y genera reportes PDF por cada bodega de vendedor.

El proyecto incluye:

- Motor de negocio desacoplado en [reportstock_core.py](reportstock_core.py)
- Interfaz de escritorio lista para operacion diaria en [main.py](main.py)
- Interfaz web para cargas y descargas en [web_app.py](web_app.py)
- Pruebas de backend en [tests/test_backend.py](tests/test_backend.py)

---

## Caracteristicas principales

| Modulo | Capacidad |
|---|---|
| Desktop (Tkinter) | Seleccion de archivo, empresa, carpeta destino y minimo por cantidad |
| Web (Flask) | Carga de Excel, filtros de bodegas, descarga en PDF o ZIP |
| Core | Lectura robusta de `.xls` y `.xlsx`, deteccion de agotados por vendedor |
| PDF | Generacion por bodega con estilos, metadata y timestamp |
| Historial | Registro de reportes web en [tmp_web/historial_pdf.json](tmp_web/historial_pdf.json) |

---

## Inicio rapido

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 2. Ejecutar version escritorio

```bash
python main.py
```

### 3. Ejecutar version web

```bash
python web_app.py
```

Abrir en navegador:

```text
http://localhost:8071
```

---

## Dependencias

Definidas en [requirements.txt](requirements.txt):

- pandas
- xlrd
- openpyxl
- reportlab
- customtkinter
- Flask
- gunicorn
- pypdf
- pytest

---

## Flujo funcional

1. Cargar Excel de inventario (`.xls` o `.xlsx`).
2. Localizar inventario de bodega PRINCIPAL.
3. Comparar por referencia contra cada bodega de vendedor.
4. Filtrar por minimo global o minimo por bodega.
5. Generar un PDF por vendedor y exportar PDF/ZIP (modo web).

---

## Formato esperado del Excel

Columnas obligatorias:

- BODEGA
- REFERENCIA
- NOMBRE DEL GRUPO
- CANTIDAD

Regla base aplicada:

- Un producto entra al reporte si en PRINCIPAL su cantidad es menor o igual al minimo configurado y en la bodega del vendedor tiene cantidad mayor a 0.

---

## Arquitectura del proyecto

```text
ReportStock-main/
|-- main.py
|-- web_app.py
|-- reportstock_core.py
|-- requirements.txt
|-- tests/
|   |-- test_backend.py
|-- web_templates/
|   |-- index.html
|   |-- historial.html
|-- tmp_web/
|   |-- uploads/
|   |-- reportes/
|   |-- reportes_finales/
|   |-- historial_pdf.json
|-- reportes/
```

---

## Endpoints web relevantes

Definidos en [web_app.py](web_app.py):

- `GET /` interfaz principal
- `POST /api/bodegas` lectura de bodegas desde Excel
- `GET /historial` historial de reportes
- `GET /descargar/<reporte_id>/<tipo>` descarga de PDF o ZIP
- `POST /generar` generacion de reportes (formulario principal)

---

## Ejecucion en produccion

Comando sugerido:

```bash
gunicorn -w 2 -b 0.0.0.0:8071 web_app:app
```

Variable opcional de seguridad para borrar historial:

```bash
export HISTORIAL_DELETE_KEY="tu_clave_segura"
```

---

## Calidad y pruebas

Para validar backend:

```bash
pytest -q
```

Pipeline de CI disponible en [.github/workflows/tests.yml](.github/workflows/tests.yml).

---

## Soporte

Autor: Yerson Vargas Vargas

- Correo: yervargas6@gmail.com
- GitHub: https://github.com/Yerson-0912

---

## Estado del proyecto

- Version funcional: 2.x
- Estado: Produccion
- Ultima actualizacion de README: Marzo 2026

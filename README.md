# CIMA Planning v3.0

Mini-ERP para la veterinaria **CIMA**. Integra planificación de compras, dashboard analítico y gestión de datos maestros en una única interfaz web.

---

## Estructura de carpetas

```
cima-planning-v3/
│
├── backend/                   ← Servidor FastAPI (Python)
│   ├── main.py
│   ├── processor.py
│   ├── pdf_generator.py
│   ├── requirements.txt
│   └── reports/               ← Reportes generados (se crea automáticamente)
│
├── frontend/
│   └── index.html             ← Interfaz web (servida por FastAPI en GET /)
│
├── assets/                    ← ⚠ ARCHIVOS MAESTROS OBLIGATORIOS (ver abajo)
│
├── static/                    ← ⚠ LOGOS OBLIGATORIOS (ver abajo)
│
└── INICIAR_CIMA.bat           ← Script de inicio (Windows)
```

---

## Archivos obligatorios en `assets/`

> **Importante:** Los nombres de archivo deben ser EXACTAMENTE estos, incluyendo mayúsculas, minúsculas y espacios.

| Archivo | Descripción |
|---|---|
| `Planning_CIMA_rev03.xlsm` | Maestro de artículos con parámetros de compra |
| `Forecast.xlsx` | Pronóstico de ventas por código, mes y año |
| `compras 2023-2025 v2.xlsx` | Historial de compras (maestro acumulativo) |
| `Ventas 2023-2025 (todas las categorías).xlsx` | Historial de ventas (maestro acumulativo) |
| `Maestro de Productos.xlsx` | Catálogo de productos activos |
| `stock medicamentos.XLS` | Stock actual de medicamentos |
| `stock accesorios.XLS` | Stock actual de accesorios |
| `stock balanceados.XLS` | Stock actual de balanceados |

---

## Archivos obligatorios en `static/`

| Archivo | Descripción |
|---|---|
| `logo_cima.png` | Logo CIMA (aparece en el PDF y en el header web) |
| `logo_roadmap.png` | Logo Roadmap (aparece en el PDF y en el footer web) |

---

## Instrucciones de inicio (Windows)

### Primera vez

1. Asegurarse de tener **Python 3.11+** instalado y en el PATH.
2. Copiar todos los archivos maestros en `assets/` con los nombres exactos indicados arriba.
3. Copiar los logos en `static/`.
4. Hacer doble clic en **`INICIAR_CIMA.bat`**.

El script instalará automáticamente las dependencias y abrirá el navegador en `http://localhost:8000`.

### Uso normal (luego de la primera vez)

Hacer doble clic en **`INICIAR_CIMA.bat`** nuevamente. Las dependencias ya estarán instaladas.

---

## Funcionalidades principales

### Tab 1: Planning
- Subir los 3 archivos de stock del día (Medicamentos, Accesorios, Balanceados).
- Ver un **resumen visual** del estado del stock: torta de proporciones (Sin Stock / Faltante / Normal / Sobrestock) y tabla por familia.
- Generar **Reporte PDF** de compras o **Excel ERP** con los artículos a pedir.
- Descargar el **Excel de Status** completo con todos los artículos y colores por estado.

### Tab 2: Dashboard
- Análisis histórico de Ventas, Compras y Stock.
- Filtros dinámicos por año.
- Sección especial de Servicios CIMA.
- Los datos se cargan automáticamente desde los maestros en `assets/`.

### Tab 3: Actualizar Datos
- Subir nuevos archivos de **Compras** (se hace append al maestro; se eliminan subtotales automáticamente).
- Subir nuevas **Ventas** por categoría (se hace append al maestro mixto).
- Subir **Maestro de Productos** (solo agrega códigos nuevos, no pisa existentes).
- Subir nuevo **Forecast**.
- Ajustar **Parámetros de Reposición** (Lead Time y Stock de Seguridad por familia).

---

## API disponible (FastAPI)

Documentación interactiva en: `http://localhost:8000/docs`

```
GET  /                              → Frontend (index.html)
GET  /api/obtener-parametros        → Parámetros + fecha forecast
POST /api/guardar-parametros        → Guarda config.json
GET  /api/archivos-status           → Fechas de modificación de maestros
GET  /api/dashboard/data            → JSON completo para el dashboard
POST /api/planning/preview          → Resumen de status de stock
POST /api/generar-reporte           → PDF o Excel de pedidos
POST /api/planning/status-stock     → Excel completo con todos los artículos
GET  /api/mdm/download/{tipo}       → Descarga cualquier maestro
POST /api/mdm/upload-compras        → Append compras
POST /api/mdm/upload-ventas         → Append ventas
POST /api/mdm/upload-maestro        → Append productos nuevos
POST /api/mdm/upload-forecast       → Reemplaza forecast
```

---

## Soporte técnico

Desarrollado por **Roadmap Analytics**.  
Para consultas o errores, revisar la consola del servidor (ventana negra del `.bat`).

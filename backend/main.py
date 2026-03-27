"""
Roadmap Analytics — main.py v4.0
==================================
v4.0:
  - MOD 4: Caché del pronóstico default (Todas/Todas/sin artículo).
    Se pre-computa en startup y se sirve instantáneamente al frontend.
    Se invalida y recalcula en background cuando se suben Ventas.
  - Sin cambios en la lógica de negocio de los endpoints existentes.
"""

from __future__ import annotations

import asyncio
import io
import json
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Optional

import numpy as np
import pandas as pd

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from processor import DEFAULT_PARAMS, ForecastEngine, StockProcessor
from pdf_generator import generate_pdf

# ─── Rutas del proyecto ───────────────────────────────────────────────────────
BACKEND_DIR  = Path(__file__).parent
PROJECT_ROOT = BACKEND_DIR.parent
ASSETS_DIR   = PROJECT_ROOT / "assets"
STATIC_DIR   = PROJECT_ROOT / "static"
FRONTEND_DIR = PROJECT_ROOT / "frontend"
OUTPUT_DIR   = BACKEND_DIR  / "reports"

XLSM_PATH     = ASSETS_DIR / "Planning_CIMA_rev03.xlsm"
FORECAST_PATH = ASSETS_DIR / "Forecast.xlsx"
CONFIG_PATH   = ASSETS_DIR / "config.json"
LOGO_PATH     = STATIC_DIR / "logo_cima.png"
ROADMAP_LOGO  = STATIC_DIR / "logo_roadmap.png"
COMPRAS_PATH  = ASSETS_DIR / "compras 2023-2025 v2.xlsx"
VENTAS_PATH   = ASSETS_DIR / "Ventas 2023-2025 (todas las categorías).xlsx"
MAESTRO_PATH  = ASSETS_DIR / "Maestro de Productos.xlsx"
STOCK_PATHS   = {
    "med": ASSETS_DIR / "stock medicamentos.XLS",
    "acc": ASSETS_DIR / "stock accesorios.XLS",
    "bal": ASSETS_DIR / "stock balanceados.XLS",
}
MASTER_DOWNLOAD_MAP = {
    "compras":  COMPRAS_PATH,
    "ventas":   VENTAS_PATH,
    "maestro":  MAESTRO_PATH,
    "forecast": FORECAST_PATH,
    "planning": XLSM_PATH,
}

for _d in (ASSETS_DIR, STATIC_DIR, OUTPUT_DIR, FRONTEND_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# ─── MOD 4: Caché del Dashboard ───────────────────────────────────────────────
# Se invalida al subir/sobrescribir Compras o Ventas.
_dashboard_cache: Optional[dict] = None


def _invalidate_dashboard_cache() -> None:
    global _dashboard_cache
    _dashboard_cache = None
    print("[CACHE] Dashboard cache invalidado.", flush=True)


# ─── App ──────────────────────────────────────────────────────────────────────
app = FastAPI(title="Roadmap Analytics API")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/")
async def root():
    index = FRONTEND_DIR / "index.html"
    if index.exists():
        return FileResponse(str(index))
    return JSONResponse({"status": "Roadmap Analytics API", "docs": "/docs"})


# ─── Utilidades generales ─────────────────────────────────────────────────────

def _load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return DEFAULT_PARAMS.copy()


def _save_config(data: dict) -> None:
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _get_mtime_str(path: Path) -> str:
    if path.exists():
        try:
            return datetime.fromtimestamp(os.path.getmtime(str(path))).strftime("%d/%m/%Y %H:%M")
        except Exception:
            pass
    return "Sin archivo"


def _get_forecast_fecha() -> str:
    return _get_mtime_str(FORECAST_PATH).split(" ")[0] if FORECAST_PATH.exists() else "Sin archivo"


def _cleanup(tmp_dir: Path) -> None:
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass


def _read_excel_auto(path: Path, **kwargs) -> Optional[pd.DataFrame]:
    if not path.exists():
        return None
    ext    = path.suffix.lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        return pd.read_excel(path, engine=engine, **kwargs)
    except Exception as e:
        print(f"[WARN] No se pudo leer {path.name}: {e}", flush=True)
        return None


def _normalize_col_names(df: pd.DataFrame) -> pd.DataFrame:
    import unicodedata
    def _clean(c: str) -> str:
        c = str(c).strip()
        c = "".join(ch for ch in unicodedata.normalize("NFD", c)
                    if unicodedata.category(ch) != "Mn")
        return c.lower()
    df.columns = [_clean(c) for c in df.columns]
    return df


def _validate_overwrite_cols(new_data: bytes, existing_path: Path,
                              new_filename: str = "") -> None:
    """
    MOD 5 — Validación estricta de columnas para sobrescritura.
    Compara cantidad y nombres de columnas entre el archivo nuevo y el existente.
    Lanza ValueError con mensaje descriptivo si no coinciden.
    """
    if not existing_path.exists():
        return  # Sin archivo previo, cualquier estructura es válida

    # Detectar engine del archivo nuevo (por extensión)
    ext_new = Path(new_filename).suffix.lower() if new_filename else ".xlsx"
    engine_new = "xlrd" if ext_new == ".xls" else "openpyxl"

    # Detectar engine del archivo existente
    ext_exist = existing_path.suffix.lower()
    engine_exist = "xlrd" if ext_exist == ".xls" else "openpyxl"

    try:
        new_df    = pd.read_excel(io.BytesIO(new_data), engine=engine_new, nrows=0)
        exist_df  = pd.read_excel(existing_path, engine=engine_exist, nrows=0)
    except Exception as exc:
        raise ValueError(f"No se pudo leer uno de los archivos para validar: {exc}")

    new_cols   = [str(c).strip() for c in new_df.columns]
    exist_cols = [str(c).strip() for c in exist_df.columns]

    if len(new_cols) != len(exist_cols):
        raise ValueError(
            f"Columnas incompatibles: el archivo tiene {len(new_cols)} columnas, "
            f"se esperaban {len(exist_cols)}. "
            f"Verificá que estés subiendo el archivo correcto."
        )

    mismatched = [
        (i + 1, n, e)
        for i, (n, e) in enumerate(zip(new_cols, exist_cols))
        if n != e
    ]
    if mismatched:
        examples = "; ".join(
            f"col {i}: '{n}' ≠ '{e}'"
            for i, n, e in mismatched[:4]
        )
        raise ValueError(
            f"Nombres de columnas incompatibles ({len(mismatched)} diferencias): {examples}. "
            f"El archivo debe tener exactamente las mismas columnas que el original."
        )


# ─── Parámetros ───────────────────────────────────────────────────────────────

@app.get("/api/obtener-parametros")
async def obtener_parametros():
    config = _load_config()
    fecha  = await asyncio.to_thread(_get_forecast_fecha)
    return JSONResponse({"params": config, "forecast_fecha": fecha})


@app.post("/api/guardar-parametros")
async def guardar_parametros(body: dict):
    valid_seg = {"2 semanas", "3 semanas", "1 mes", "2 meses"}
    for fam, cfg in body.items():
        if not isinstance(cfg, dict):
            raise HTTPException(400, f"Parámetros inválidos para '{fam}'")
        lt = cfg.get("lead_time")
        ss = cfg.get("stock_seguridad")
        if lt is not None and (not isinstance(lt, (int, float)) or lt < 0):
            raise HTTPException(400, f"lead_time inválido: {lt}")
        if ss is not None and ss not in valid_seg:
            raise HTTPException(400, f"stock_seguridad inválido: '{ss}'")
    await asyncio.to_thread(_save_config, body)
    return JSONResponse({"ok": True, "mensaje": "Parámetros guardados correctamente."})


@app.get("/api/archivos-status")
async def archivos_status():
    """
    MOD 3: Incluye fechas de los archivos de stock de assets/
    para que el usuario decida si necesita subir nuevos.
    """
    return JSONResponse({
        "compras":   _get_mtime_str(COMPRAS_PATH),
        "ventas":    _get_mtime_str(VENTAS_PATH),
        "maestro":   _get_mtime_str(MAESTRO_PATH),
        "forecast":  _get_mtime_str(FORECAST_PATH),
        "planning":  _get_mtime_str(XLSM_PATH),
        # stock files en assets/ — para info en UI de Planning
        "stock_med": _get_mtime_str(STOCK_PATHS["med"]),
        "stock_acc": _get_mtime_str(STOCK_PATHS["acc"]),
        "stock_bal": _get_mtime_str(STOCK_PATHS["bal"]),
    })


# ─── Planning: Resolución de archivos de stock ────────────────────────────────

def _resolve_stock_path(
    tmp: Path,
    uploaded_bytes: Optional[bytes],
    uploaded_name: str,
    fallback_path: Path,
    label: str,
) -> Path:
    """
    MOD 3 — Si se recibió un archivo uploaded, lo guarda en tmp y devuelve esa ruta.
    Si no, devuelve la ruta del archivo en assets/ (fallback).
    """
    if uploaded_bytes:
        dest = tmp / Path(uploaded_name).name
        dest.write_bytes(uploaded_bytes)
        return dest
    # Fallback a assets/
    if not fallback_path.exists():
        raise ValueError(
            f"No se subió el archivo de {label} y tampoco existe en assets/. "
            f"Subí el archivo o verificá la carpeta assets/."
        )
    print(f"[PLANNING] Usando stock de assets/ para {label}: {fallback_path.name}", flush=True)
    return fallback_path


# ─── Planning: Preview ───────────────────────────────────────────────────────

@app.post("/api/planning/preview")
async def planning_preview(
    med:     Optional[UploadFile] = File(None),
    acc:     Optional[UploadFile] = File(None),
    bal:     Optional[UploadFile] = File(None),
    familia: str = Form(default="Todas"),
):
    """
    MOD 3: archivos de stock opcionales.
    Si no se suben, usa los .XLS de assets/.
    """
    if not XLSM_PATH.exists():
        raise HTTPException(500, "Planning_CIMA_rev03.xlsm no encontrado en assets/.")

    med_b    = await med.read() if med and med.filename else None
    acc_b    = await acc.read() if acc and acc.filename else None
    bal_b    = await bal.read() if bal and bal.filename else None
    med_name = (med.filename or "stock_med.xls") if med else "stock_med.xls"
    acc_name = (acc.filename or "stock_acc.xls") if acc else "stock_acc.xls"
    bal_name = (bal.filename or "stock_bal.xls") if bal else "stock_bal.xls"
    params   = _load_config()

    def _run():
        tmp = Path(tempfile.mkdtemp(prefix="cima_prev_"))
        try:
            p_med = _resolve_stock_path(tmp, med_b, med_name, STOCK_PATHS["med"], "Medicamentos")
            p_acc = _resolve_stock_path(tmp, acc_b, acc_name, STOCK_PATHS["acc"], "Accesorios")
            p_bal = _resolve_stock_path(tmp, bal_b, bal_name, STOCK_PATHS["bal"], "Balanceados")

            proc   = StockProcessor(str(XLSM_PATH), str(FORECAST_PATH))
            df_all = proc.process_all(p_med, p_acc, p_bal, familia, params)

            def _status(row):
                q, sm = row["CANTIDAD"], row["S_MIN"]
                if q <= 0:               return "Sin Stock"
                if q < sm:               return "Faltante"
                if sm > 0 and q >= 3*sm: return "Sobrestock"
                return "Normal"

            df_all["STATUS"] = df_all.apply(_status, axis=1)
            sc    = df_all["STATUS"].value_counts().to_dict()
            total = len(df_all)
            fam_breakdown = []
            for fn, grp in df_all.groupby("FAMILIA"):
                s = grp["STATUS"].value_counts().to_dict()
                fam_breakdown.append({
                    "familia":    fn,
                    "sin_stock":  int(s.get("Sin Stock", 0)),
                    "faltante":   int(s.get("Faltante", 0)),
                    "normal":     int(s.get("Normal", 0)),
                    "sobrestock": int(s.get("Sobrestock", 0)),
                    "total":      len(grp),
                    "a_pedir":    int(grp["PEDIR"].sum()),
                })
            total_stock = int(df_all["CANTIDAD"].clip(lower=0).sum())
            return {
                "total_articulos": total,
                "total_stock":     total_stock,
                "sin_stock":       int(sc.get("Sin Stock", 0)),
                "faltante":        int(sc.get("Faltante", 0)),
                "normal":          int(sc.get("Normal", 0)),
                "sobrestock":      int(sc.get("Sobrestock", 0)),
                "articulos_pedir": int((df_all["PEDIR"] > 0).sum()),
                "unidades_pedir":  int(df_all["PEDIR"].sum()),
                "fam_breakdown":   fam_breakdown,
                "forecast_fecha":  _get_forecast_fecha(),
                "uso_stock_base":  not bool(med_b or acc_b or bal_b),
            }
        finally:
            _cleanup(tmp)

    try:
        result = await asyncio.to_thread(_run)
    except ValueError as exc:
        raise HTTPException(422, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error en preview: {exc}")
    return JSONResponse(result)


# ─── Planning: Generar Reporte ────────────────────────────────────────────────

@app.post("/api/generar-reporte")
async def generar_reporte(
    med:           Optional[UploadFile] = File(None),
    acc:           Optional[UploadFile] = File(None),
    bal:           Optional[UploadFile] = File(None),
    familia:       str = Form(default="Todas"),
    output_format: str = Form(default="pdf"),
):
    if output_format not in ("pdf", "excel"):
        raise HTTPException(400, "output_format debe ser 'pdf' o 'excel'")
    if not XLSM_PATH.exists():
        raise HTTPException(500, "Planning_CIMA_rev03.xlsm no encontrado en assets/.")

    ts       = datetime.now().strftime("%Y-%m-%d_%H%M")
    med_b    = await med.read() if med and med.filename else None
    acc_b    = await acc.read() if acc and acc.filename else None
    bal_b    = await bal.read() if bal and bal.filename else None
    med_name = (med.filename or "stock_med.xls") if med else "stock_med.xls"
    acc_name = (acc.filename or "stock_acc.xls") if acc else "stock_acc.xls"
    bal_name = (bal.filename or "stock_bal.xls") if bal else "stock_bal.xls"
    params   = _load_config()
    fc_fecha = _get_forecast_fecha()

    def _run():
        tmp = Path(tempfile.mkdtemp(prefix="cima_rep_"))
        try:
            p_med = _resolve_stock_path(tmp, med_b, med_name, STOCK_PATHS["med"], "Medicamentos")
            p_acc = _resolve_stock_path(tmp, acc_b, acc_name, STOCK_PATHS["acc"], "Accesorios")
            p_bal = _resolve_stock_path(tmp, bal_b, bal_name, STOCK_PATHS["bal"], "Balanceados")

            proc = StockProcessor(str(XLSM_PATH), str(FORECAST_PATH))
            df   = proc.process(p_med, p_acc, p_bal, familia, params, str(tmp))
            if len(df) == 0:
                raise ValueError("No hay artículos para pedir. Verificá el Pronóstico y los parámetros.")

            if output_format == "pdf":
                content = generate_pdf(
                    df,
                    logo_path         = str(LOGO_PATH)    if LOGO_PATH.exists()    else None,
                    roadmap_logo_path = str(ROADMAP_LOGO) if ROADMAP_LOGO.exists() else None,
                    single_family     = familia.upper() != "TODAS",
                    forecast_fecha    = fc_fecha,
                )
                fname = f"Reporte_CIMA_{ts}.pdf"
                media = "application/pdf"
            else:
                content = proc.to_excel_bytes(df)
                fc_slug = fc_fecha.replace("/", "-") if fc_fecha != "Sin archivo" else "sin-fecha"
                fname   = f"Reporte_CIMA_{ts}_fc{fc_slug}.xlsx"
                media   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            out = OUTPUT_DIR / fname
            out.write_bytes(content)
            debug = tmp / "debug_matematica.csv"
            if debug.exists():
                shutil.copy2(str(debug), str(OUTPUT_DIR / f"debug_{ts}.csv"))
            return out, fname, media
        except Exception as exc:
            _cleanup(tmp); raise exc

    try:
        out, fname, media = await asyncio.to_thread(_run)
    except ValueError as exc:
        raise HTTPException(422, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error procesando: {exc}")

    return FileResponse(str(out), media_type=media, filename=fname,
                        headers={"Content-Disposition": f'attachment; filename="{fname}"'})


# ─── Planning: Status de Stock ────────────────────────────────────────────────

@app.post("/api/planning/status-stock")
async def status_stock_excel(
    med:     Optional[UploadFile] = File(None),
    acc:     Optional[UploadFile] = File(None),
    bal:     Optional[UploadFile] = File(None),
    familia: str = Form(default="Todas"),
):
    if not XLSM_PATH.exists():
        raise HTTPException(500, "Planning_CIMA_rev03.xlsm no encontrado en assets/.")

    ts     = datetime.now().strftime("%Y-%m-%d_%H%M")
    med_b  = await med.read() if med and med.filename else None
    acc_b  = await acc.read() if acc and acc.filename else None
    bal_b  = await bal.read() if bal and bal.filename else None
    med_name = (med.filename or "stock_med.xls") if med else "stock_med.xls"
    acc_name = (acc.filename or "stock_acc.xls") if acc else "stock_acc.xls"
    bal_name = (bal.filename or "stock_bal.xls") if bal else "stock_bal.xls"
    params = _load_config()

    def _run():
        tmp = Path(tempfile.mkdtemp(prefix="cima_sts_"))
        try:
            p_med = _resolve_stock_path(tmp, med_b, med_name, STOCK_PATHS["med"], "Medicamentos")
            p_acc = _resolve_stock_path(tmp, acc_b, acc_name, STOCK_PATHS["acc"], "Accesorios")
            p_bal = _resolve_stock_path(tmp, bal_b, bal_name, STOCK_PATHS["bal"], "Balanceados")

            proc    = StockProcessor(str(XLSM_PATH), str(FORECAST_PATH))
            content = proc.to_status_excel_bytes(p_med, p_acc, p_bal, familia, params)
            fname   = f"Status_Stock_CIMA_{ts}.xlsx"
            out     = OUTPUT_DIR / fname
            out.write_bytes(content)
            return out, fname
        finally:
            _cleanup(tmp)

    try:
        out, fname = await asyncio.to_thread(_run)
    except ValueError as exc:
        raise HTTPException(422, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error generando status: {exc}")

    media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return FileResponse(str(out), media_type=media, filename=fname,
                        headers={"Content-Disposition": f'attachment; filename="{fname}"'})


# ─── Forecast: Filtros ────────────────────────────────────────────────────────

@app.get("/api/forecast/filtros")
async def forecast_filtros():
    def _run():
        return ForecastEngine.get_forecast_filtros(
            forecast_path = FORECAST_PATH,
            maestro_path  = MAESTRO_PATH if MAESTRO_PATH.exists() else None,
        )
    try:
        result = await asyncio.to_thread(_run)
    except Exception as exc:
        raise HTTPException(500, f"Error obteniendo filtros: {exc}")
    return JSONResponse(result)


# ─── Forecast: Serie de tiempo ────────────────────────────────────────────────

@app.post("/api/forecast/timeseries")
async def forecast_timeseries(
    familia:   str = Form(default="Todas"),
    categoria: str = Form(default="Todas"),
    articulo:  str = Form(default=""),
):
    def _run():
        return ForecastEngine.get_timeseries_for_chart(
            ventas_path   = VENTAS_PATH,
            forecast_path = FORECAST_PATH,
            maestro_path  = MAESTRO_PATH if MAESTRO_PATH.exists() else None,
            familia   = familia,
            categoria = categoria,
            articulo  = articulo,
        )
    try:
        result = await asyncio.to_thread(_run)
    except Exception as exc:
        raise HTTPException(500, f"Error generando timeseries: {exc}")
    return JSONResponse(result)


# ─── Dashboard — con caché ────────────────────────────────────────────────────

@app.get("/api/dashboard/data")
async def dashboard_data():
    """
    MOD 4: Devuelve el caché si está disponible. Solo recalcula si el caché
    fue invalidado (al subir/sobrescribir Compras o Ventas).
    """
    global _dashboard_cache
    if _dashboard_cache is not None:
        print("[CACHE] Dashboard servido desde caché.", flush=True)
        return JSONResponse(_dashboard_cache)
    try:
        D = await asyncio.to_thread(_compute_dashboard)
        _dashboard_cache = D
        print("[CACHE] Dashboard calculado y guardado en caché.", flush=True)
        return JSONResponse(D)
    except Exception as exc:
        print(f"[DASH ERROR] {type(exc).__name__}: {exc}", flush=True)
        raise HTTPException(500, f"Error computando dashboard: {exc}")


def _compute_dashboard() -> dict:
    MES = list(range(1, 13))

    # ── VENTAS — carga vectorizada ────────────────────────────────────────────
    dv_raw = ForecastEngine._get_ventas(VENTAS_PATH) if VENTAS_PATH.exists() else None

    if dv_raw is not None and len(dv_raw):
        dv = _normalize_col_names(dv_raw.copy())
        cols = list(dv.columns)

        fc  = next((c for c in cols if "fecha" in c or "emisi" in c), cols[0])
        qc  = next((c for c in cols if "cantidad" in c), "")
        fmc = next((c for c in cols if "familia" in c or "categor" in c), "")
        adc = next((c for c in cols if "articulo desc" in c or
                    ("articulo" in c and "desc" in c)), "")
        dc  = next((c for c in cols if c.startswith("descripc")
                    and c not in (fmc, adc)), "")
        clc = next((c for c in cols if "denominaci" in c or "razon" in c), "")

        # Parsear fechas y cantidad de forma vectorizada
        dv["_f"] = pd.to_datetime(dv[fc], dayfirst=True, errors="coerce")
        dv        = dv.dropna(subset=["_f"])

        if qc:
            dv["_q"] = pd.to_numeric(dv[qc], errors="coerce").fillna(0)
        else:
            dv["_q"] = 0.0

        dv["_yr"] = dv["_f"].dt.year.astype(int)
        dv["_mo"] = dv["_f"].dt.month.astype(int)
        dv["_ym"] = dv["_f"].dt.strftime("%Y-%m")

        # Columnas de texto — limpiar NaN → cadena vacía
        for alias, col in [("_fam", fmc), ("_ad", adc), ("_desc", dc), ("_cl", clc)]:
            if col:
                dv[alias] = dv[col].fillna("").astype(str)
            else:
                dv[alias] = ""

        dv = dv[dv["_yr"] >= 2020].copy()
        total_v_rows = len(dv)
        print(f"[DASH] Ventas: {total_v_rows:,} filas", flush=True)
    else:
        dv            = pd.DataFrame()
        total_v_rows  = 0
        print("[DASH] Ventas: sin datos", flush=True)

    # Años disponibles
    years_v = sorted(dv["_yr"].unique()) if len(dv) else [datetime.now().year]

    # Estructuras de salida — ventas
    kv: dict = {}
    ventas_mes_anio_dict: dict = {}
    ventas_fam_anio:      dict = {}
    ventas_cat_anio:      dict = {}
    ventas_cat3_anio:     dict = {}
    ventas_subcat_anio:   dict = {}
    top_prods_anio:       dict = {}
    total_ventas_anio:    dict = {}
    ventas_anio:          list = []
    cima_kv:              dict = {}
    cima_mes_anio:        dict = {}
    cima_subcat_anio:     dict = {}
    cima_top_anio:        dict = {}
    cima_clientes_anio:   dict = {}

    # ── helper: construir serie mensual completa (todos los meses 1-12) ───────
    def _mes_series(sub: pd.DataFrame) -> dict:
        """Devuelve {str(m): cantidad} para todos los meses 1-12."""
        if len(sub) == 0:
            return {str(m): 0 for m in MES}
        byM = sub.groupby("_mo")["_q"].sum()
        return {str(m): int(round(byM.get(m, 0))) for m in MES}

    # ── procesar por año ──────────────────────────────────────────────────────
    if len(dv):
        # Indicador CIMA en el DataFrame completo (vectorizado)
        fam_upper = dv["_fam"].str.upper()
        dv["_is_cima"] = fam_upper.str.contains("CIMA", na=False) | \
                         fam_upper.str.contains("SERVIC", na=False)

        for yr in sorted(years_v):
            sub = dv[dv["_yr"] == yr]
            if len(sub) == 0:
                continue

            # KPIs de ventas
            kv[str(yr)] = {
                "unidades": int(round(sub["_q"].sum())),
                "tx":       len(sub),
                "clientes": sub["_cl"].replace("", pd.NA).dropna().nunique(),
            }
            ventas_mes_anio_dict[str(yr)] = _mes_series(sub)
            total_ventas_anio[str(yr)]    = kv[str(yr)]["unidades"]

            # Ventas por familia (fam_anio, cat_anio, cat3)
            fam_d = (sub.assign(_fam2=sub["_fam"].replace("", "Sin Categoría"))
                       .groupby("_fam2")["_q"].sum()
                       .round().astype(int))
            ventas_fam_anio[str(yr)]  = fam_d.to_dict()
            ventas_cat_anio[str(yr)]  = (
                [{"categoria": k, "unidades": v}
                 for k, v in fam_d.sort_values(ascending=False).head(10).items()])
            top3_fam = fam_d.sort_values(ascending=False).head(3)
            ventas_cat3_anio[str(yr)] = [
                {"cat": k, "unidades": v} for k, v in top3_fam.items()]

            # Subcategorías (desc)
            subc = (sub.assign(_d2=sub["_desc"].replace("", "Sin Desc."))
                      .groupby("_d2")["_q"].sum()
                      .round().astype(int)
                      .sort_values(ascending=False).head(12))
            ventas_subcat_anio[str(yr)] = [
                {"subcat": k, "unidades": v} for k, v in subc.items()]

            # Top productos (ad)
            tpd = (sub.assign(_a2=sub["_ad"].replace("", "Sin Artículo"))
                     .groupby("_a2")["_q"].sum()
                     .round().astype(int)
                     .sort_values(ascending=False).head(15))
            top_prods_anio[str(yr)] = [
                {"producto": k, "unidades": v} for k, v in tpd.items()]

            # CIMA — subconjunto de este año
            csub = sub[sub["_is_cima"]]
            cima_kv[str(yr)] = {
                "unidades": int(round(csub["_q"].sum())),
                "tx":       len(csub),
                "clientes": csub["_cl"].replace("", pd.NA).dropna().nunique(),
            }
            cima_mes_anio[str(yr)] = _mes_series(csub)

            cs = (csub.assign(_d2=csub["_desc"].replace("", "Sin Desc."))
                      .groupby("_d2")["_q"].sum()
                      .round().astype(int)
                      .sort_values(ascending=False))
            cima_subcat_anio[str(yr)] = [
                {"subcat": k, "unidades": v} for k, v in cs.items()]

            ct = (csub.assign(_a2=csub["_ad"].replace("", "Sin Artículo"))
                      .groupby("_a2")["_q"].sum()
                      .round().astype(int)
                      .sort_values(ascending=False).head(10))
            cima_top_anio[str(yr)] = [
                {"producto": k, "unidades": v} for k, v in ct.items()]

            cc = (csub.assign(_c2=csub["_cl"].replace("", "Sin Cliente"))
                      .groupby("_c2")["_q"].sum()
                      .round().astype(int)
                      .sort_values(ascending=False).head(10))
            cima_clientes_anio[str(yr)] = [
                {"cliente": k, "unidades": v} for k, v in cc.items()]

    # ── totales "all" ─────────────────────────────────────────────────────────
    if len(dv):
        all_fam = (dv.assign(_fam2=dv["_fam"].replace("", "Sin Categoría"))
                    .groupby("_fam2")["_q"].sum()
                    .round().astype(int))
        kv["all"] = {
            "unidades": int(round(dv["_q"].sum())),
            "tx":       len(dv),
            "clientes": dv["_cl"].replace("", pd.NA).dropna().nunique(),
        }
        total_ventas_anio["all"] = kv["all"]["unidades"]
        ventas_cat_anio["all"]   = [
            {"categoria": k, "unidades": v}
            for k, v in all_fam.sort_values(ascending=False).head(10).items()]
        ventas_cat3_anio["all"]  = [
            {"cat": k, "unidades": v}
            for k, v in all_fam.sort_values(ascending=False).head(3).items()]

        all_subc = (dv.assign(_d2=dv["_desc"].replace("", "Sin Desc."))
                     .groupby("_d2")["_q"].sum()
                     .round().astype(int)
                     .sort_values(ascending=False).head(12))
        ventas_subcat_anio["all"] = [
            {"subcat": k, "unidades": v} for k, v in all_subc.items()]

        all_tp = (dv.assign(_a2=dv["_ad"].replace("", "Sin Artículo"))
                   .groupby("_a2")["_q"].sum()
                   .round().astype(int)
                   .sort_values(ascending=False).head(15))
        top_prods_anio["all"] = [
            {"producto": k, "unidades": v} for k, v in all_tp.items()]

        ventas_anio = [
            {"Año": yr, "unidades": kv[str(yr)]["unidades"]}
            for yr in sorted(years_v)]

        # Serie mensual "all" = dict ym → cantidad
        all_ym = (dv.groupby("_ym")["_q"].sum()
                    .round().astype(int).to_dict())
        ventas_mes_anio_dict["all"] = all_ym

        # CIMA "all"
        cima_all = dv[dv["_is_cima"]]
        cima_kv["all"] = {
            "unidades": int(round(cima_all["_q"].sum())),
            "tx":       len(cima_all),
            "clientes": cima_all["_cl"].replace("", pd.NA).dropna().nunique(),
        }
        cima_mes_anio["all"] = all_ym   # igual que ventas all

        cs_all = (cima_all.assign(_d2=cima_all["_desc"].replace("", "Sin Desc."))
                           .groupby("_d2")["_q"].sum()
                           .round().astype(int)
                           .sort_values(ascending=False))
        cima_subcat_anio["all"] = [
            {"subcat": k, "unidades": v} for k, v in cs_all.items()]

        ct_all = (cima_all.assign(_a2=cima_all["_ad"].replace("", "Sin Artículo"))
                           .groupby("_a2")["_q"].sum()
                           .round().astype(int)
                           .sort_values(ascending=False).head(10))
        cima_top_anio["all"] = [
            {"producto": k, "unidades": v} for k, v in ct_all.items()]

        cc_all = (cima_all.assign(_c2=cima_all["_cl"].replace("", "Sin Cliente"))
                           .groupby("_c2")["_q"].sum()
                           .round().astype(int)
                           .sort_values(ascending=False).head(10))
        cima_clientes_anio["all"] = [
            {"cliente": k, "unidades": v} for k, v in cc_all.items()]
    else:
        kv["all"] = {"unidades": 0, "tx": 0, "clientes": 0}
        total_ventas_anio["all"] = 0
        ventas_cat_anio["all"] = []; ventas_cat3_anio["all"] = []
        ventas_subcat_anio["all"] = []; top_prods_anio["all"] = []
        ventas_anio = []; ventas_mes_anio_dict["all"] = {}
        cima_kv["all"] = {"unidades": 0, "tx": 0, "clientes": 0}
        cima_mes_anio["all"] = {}
        cima_subcat_anio["all"] = []; cima_top_anio["all"] = []
        cima_clientes_anio["all"] = []

    cima_anio = [
        {"Año": yr, "unidades": cima_kv[str(yr)]["unidades"]}
        for yr in sorted(years_v) if str(yr) in cima_kv]

    # ── COMPRAS — carga vectorizada ───────────────────────────────────────────
    dc_raw = ForecastEngine._get_compras(COMPRAS_PATH) if COMPRAS_PATH.exists() else None

    kc:               dict = {}
    compras_mes_anio: dict = {}
    top_provs_anio:   dict = {}

    if dc_raw is not None and len(dc_raw):
        dc = _normalize_col_names(dc_raw.copy())
        cols = list(dc.columns)
        fc_c = next((c for c in cols if "emisi" in c or "fecha" in c), cols[0])
        qc_c = next((c for c in cols if "cantidad" in c), "")
        pc_c = next((c for c in cols if "nombre" in c or "proveedor" in c), "")

        dc["_f"]  = pd.to_datetime(dc[fc_c], dayfirst=True, errors="coerce")
        dc        = dc.dropna(subset=["_f"])
        dc["_q"]  = pd.to_numeric(dc[qc_c], errors="coerce").fillna(0) if qc_c else 0.0
        dc["_yr"] = dc["_f"].dt.year.astype(int)
        dc["_mo"] = dc["_f"].dt.month.astype(int)
        dc["_ym"] = dc["_f"].dt.strftime("%Y-%m")
        dc["_pv"] = dc[pc_c].fillna("").astype(str) if pc_c else ""

        dc = dc[dc["_yr"] >= 2020].copy()
        print(f"[DASH] Compras: {len(dc):,} filas", flush=True)

        years_c = sorted(dc["_yr"].unique())
        all_years_c = set(years_v) | set(years_c)

        for yr in sorted(all_years_c):
            sub = dc[dc["_yr"] == yr]
            compras_mes_anio[str(yr)] = _mes_series(sub)
            kc[str(yr)] = {
                "unidades":    int(round(sub["_q"].sum())),
                "tx":          len(sub),
                "proveedores": sub["_pv"].replace("", pd.NA).dropna().nunique(),
            }
            pv = (sub.assign(_p2=sub["_pv"].replace("", "Sin Proveedor"))
                    .groupby("_p2")["_q"].sum()
                    .round().astype(int)
                    .sort_values(ascending=False).head(10))
            top_provs_anio[str(yr)] = [
                {"proveedor": k, "unidades": v} for k, v in pv.items()]

        # "all"
        kc["all"] = {
            "unidades":    int(round(dc["_q"].sum())),
            "tx":          len(dc),
            "proveedores": dc["_pv"].replace("", pd.NA).dropna().nunique(),
        }
        all_ym_c = (dc.groupby("_ym")["_q"].sum()
                      .round().astype(int).to_dict())
        compras_mes_anio["all"] = all_ym_c

        pv_all = (dc.assign(_p2=dc["_pv"].replace("", "Sin Proveedor"))
                    .groupby("_p2")["_q"].sum()
                    .round().astype(int)
                    .sort_values(ascending=False).head(10))
        top_provs_anio["all"] = [
            {"proveedor": k, "unidades": v} for k, v in pv_all.items()]
    else:
        print("[DASH] Compras: sin datos", flush=True)
        years_c  = years_v
        all_years_c = set(years_v)
        kc["all"] = {"unidades": 0, "tx": 0, "proveedores": 0}
        compras_mes_anio["all"] = {}
        top_provs_anio["all"]   = []

    # ── STOCK ─────────────────────────────────────────────────────────────────
    # El stock son archivos pequeños (< 5000 filas) — se mantiene iterrows
    # porque cada archivo ya tarda < 100 ms. No hay ganancia que justifique
    # la complejidad de vectorizar aquí.
    stock_items: list = []
    for spath in STOCK_PATHS.values():
        df_s = _read_excel_auto(spath)
        if df_s is None or not len(df_s):
            continue
        df_s = _normalize_col_names(df_s)
        cols_s = list(df_s.columns)
        ac  = next((c for c in cols_s if c.startswith("articulo") and "desc" not in c),
                   cols_s[0] if cols_s else "")
        adc = next((c for c in cols_s if "articulo desc" in c or
                    ("articulo" in c and "desc" in c)), "")
        qc  = next((c for c in cols_s if c == "cantidad"), "")
        mc  = next((c for c in cols_s if "minimo" in c or c.startswith("min")), "")
        for _, row in df_s.iterrows():
            q  = max(0.0, float(row[qc])  if qc and pd.notna(row.get(qc))  else 0.0)
            mn = max(0.0, float(row[mc])  if mc and pd.notna(row.get(mc))  else 0.0)
            stock_items.append({
                "art":  str(row.get(ac, "") or ""),
                "desc": str(row.get(adc, "") or "") if adc else "",
                "q": q, "min": mn, "dif": q - mn,
            })

    # Lookup ventas por descripción de artículo (vectorizado)
    if len(dv) and stock_items:
        vh_series = (dv.assign(_a2=dv["_ad"].replace("", ""))
                       .groupby("_a2")["_q"].sum())
        vh = vh_series.to_dict()
    else:
        vh = {}

    if stock_items:
        stock_kpis = {
            "total_productos": len(stock_items),
            "total_unidades":  int(sum(r["q"] for r in stock_items)),
            "sin_stock":       sum(1 for r in stock_items if r["q"] <= 0),
            "por_reponer":     sum(1 for r in stock_items if r["dif"] < 0),
            "en_exceso":       sum(1 for r in stock_items
                                   if r["min"] > 0 and r["q"] >= 3 * r["min"]),
        }
        stock_critico = sorted(
            [{"Artículo descripción": r["desc"], "Cantidad": r["q"],
              "Mínimo rep.": r["min"], "Diferencia": r["dif"],
              "ventas_total": round(vh.get(r["desc"], 0))}
             for r in stock_items if r["dif"] < 0],
            key=lambda x: -x["ventas_total"],
        )[:20]
    else:
        stock_kpis    = {"total_productos": 0, "total_unidades": 0,
                         "sin_stock": 0, "por_reponer": 0, "en_exceso": 0}
        stock_critico = []

    available_years = sorted(set(years_v) | set(years_c))

    raw_result = {
        "available_years":       [str(y) for y in available_years],
        "kv":                    kv,
        "ventas_mes_anio_dict":  ventas_mes_anio_dict,
        "ventas_fam_anio":       ventas_fam_anio,
        "ventas_cat_anio":       ventas_cat_anio,
        "ventas_cat3_anio":      ventas_cat3_anio,
        "ventas_subcat_anio":    ventas_subcat_anio,
        "top_prods_anio":        top_prods_anio,
        "ventas_anio":           ventas_anio,
        "total_ventas_anio":     total_ventas_anio,
        "kc":                    kc,
        "compras_mes_anio":      compras_mes_anio,
        "top_provs_anio":        top_provs_anio,
        "stock_kpis":            stock_kpis,
        "stock_critico":         stock_critico,
        "cima_kv":               cima_kv,
        "cima_mes_anio":         cima_mes_anio,
        "cima_subcat_anio":      cima_subcat_anio,
        "cima_top_anio":         cima_top_anio,
        "cima_clientes_anio":    cima_clientes_anio,
        "cima_anio":             cima_anio,
    }

    # Función limpiadora para purgar los numpy ints/floats
    def clean_numpy(obj):
        if isinstance(obj, dict):
            return {k: clean_numpy(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [clean_numpy(v) for v in obj]
        elif isinstance(obj, (np.integer, np.int32, np.int64)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float32, np.float64)):
            return float(obj)
        return obj

    return clean_numpy(raw_result)


# ─── MDM: Descarga ────────────────────────────────────────────────────────────

@app.get("/api/mdm/download/{tipo}")
async def mdm_download(tipo: str):
    path = MASTER_DOWNLOAD_MAP.get(tipo)
    if path is None:
        raise HTTPException(404, f"Tipo desconocido: '{tipo}'")
    if not path.exists():
        raise HTTPException(404, f"Archivo no encontrado: {path.name}")
    ext   = path.suffix.lower()
    media = ("application/vnd.ms-excel" if ext == ".xls"
             else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return FileResponse(str(path), filename=path.name, media_type=media,
                        headers={"Content-Disposition": f'attachment; filename="{path.name}"'})


# ─── MDM: APPEND Compras ─────────────────────────────────────────────────────

@app.post("/api/mdm/upload-compras")
async def mdm_upload_compras(file: UploadFile = File(...)):
    data = await file.read()
    def _process():
        ext    = Path(file.filename).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        new_df = pd.read_excel(io.BytesIO(data), engine=engine)
        emision_col = None
        for c in new_df.columns:
            if "emisi" in str(c).lower() or "fecha" in str(c).lower():
                emision_col = c; break
        if emision_col is None and len(new_df.columns) > 0:
            emision_col = new_df.columns[0]
        if emision_col is not None:
            mask_sub = new_df[emision_col].astype(str).str.lower().str.contains("subtotal", na=False)
            new_df   = new_df[~mask_sub].reset_index(drop=True)
        if COMPRAS_PATH.exists():
            master = pd.read_excel(COMPRAS_PATH, engine="openpyxl")
            if len(new_df.columns) != len(master.columns):
                raise ValueError(f"Columnas incompatibles: {len(new_df.columns)} vs {len(master.columns)}")
            new_df.columns = master.columns
            combined = pd.concat([master, new_df], ignore_index=True)
        else:
            combined = new_df
        combined.to_excel(COMPRAS_PATH, index=False, engine="openpyxl")
        return len(new_df)
    try:
        n = await asyncio.to_thread(_process)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error procesando compras: {exc}")
    _invalidate_dashboard_cache()
    ForecastEngine.invalidate_caches()
    return JSONResponse({"ok": True, "filas_agregadas": n, "ultima_mod": _get_mtime_str(COMPRAS_PATH)})


# ─── MDM: OVERWRITE Compras (con validación de columnas) ─────────────────────

@app.post("/api/mdm/overwrite-compras")
async def mdm_overwrite_compras(file: UploadFile = File(...)):
    """MOD 5: Reemplaza compras con validación estricta de columnas."""
    data = await file.read()
    def _validate():
        _validate_overwrite_cols(data, COMPRAS_PATH, file.filename or "")
    try:
        await asyncio.to_thread(_validate)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    await asyncio.to_thread(lambda: COMPRAS_PATH.write_bytes(data))
    _invalidate_dashboard_cache()
    ForecastEngine.invalidate_caches()
    return JSONResponse({"ok": True, "mensaje": "Archivo de compras sobrescrito.",
                         "ultima_mod": _get_mtime_str(COMPRAS_PATH)})


# ─── MDM: APPEND Ventas ──────────────────────────────────────────────────────

@app.post("/api/mdm/upload-ventas")
async def mdm_upload_ventas(file: UploadFile = File(...)):
    data = await file.read()
    def _process():
        ext    = Path(file.filename).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        new_df = pd.read_excel(io.BytesIO(data), engine=engine)
        if VENTAS_PATH.exists():
            master = pd.read_excel(VENTAS_PATH, engine="openpyxl")
            if len(new_df.columns) != len(master.columns):
                raise ValueError(f"Columnas incompatibles: {len(new_df.columns)} vs {len(master.columns)}")
            new_df.columns = master.columns
            combined = pd.concat([master, new_df], ignore_index=True)
        else:
            combined = new_df
        combined.to_excel(VENTAS_PATH, index=False, engine="openpyxl")
        return len(new_df)
    try:
        n = await asyncio.to_thread(_process)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error procesando ventas: {exc}")

    _invalidate_dashboard_cache()
    ForecastEngine.invalidate_caches()

    if MAESTRO_PATH.exists():
        async def _recalc():
            def _run_fc():
                try:
                    ForecastEngine.run_and_save(VENTAS_PATH, MAESTRO_PATH, FORECAST_PATH)
                    print("[MDM-VENTAS] Pronóstico recalculado.", flush=True)
                except Exception as exc2:
                    print(f"[MDM-VENTAS] Warn: {exc2}", flush=True)
            await asyncio.to_thread(_run_fc)
        asyncio.create_task(_recalc())

    return JSONResponse({"ok": True, "filas_agregadas": n, "ultima_mod": _get_mtime_str(VENTAS_PATH)})


# ─── MDM: OVERWRITE Ventas (con validación de columnas) ──────────────────────

@app.post("/api/mdm/overwrite-ventas")
async def mdm_overwrite_ventas(file: UploadFile = File(...)):
    """MOD 5: Reemplaza ventas con validación estricta de columnas."""
    data = await file.read()
    def _validate():
        _validate_overwrite_cols(data, VENTAS_PATH, file.filename or "")
    try:
        await asyncio.to_thread(_validate)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    await asyncio.to_thread(lambda: VENTAS_PATH.write_bytes(data))
    _invalidate_dashboard_cache()
    ForecastEngine.invalidate_caches()

    if MAESTRO_PATH.exists():
        async def _recalc():
            def _run_fc():
                try:
                    ForecastEngine.run_and_save(VENTAS_PATH, MAESTRO_PATH, FORECAST_PATH)
                    print("[MDM-OVERWRITE] Pronóstico recalculado.", flush=True)
                except Exception as exc2:
                    print(f"[MDM-OVERWRITE] Warn: {exc2}", flush=True)
            await asyncio.to_thread(_run_fc)
        asyncio.create_task(_recalc())

    return JSONResponse({"ok": True, "mensaje": "Archivo de ventas sobrescrito.",
                         "ultima_mod": _get_mtime_str(VENTAS_PATH)})


# ─── MDM: APPEND Maestro ─────────────────────────────────────────────────────

@app.post("/api/mdm/upload-maestro")
async def mdm_upload_maestro(file: UploadFile = File(...)):
    data = await file.read()
    def _process():
        ext    = Path(file.filename).suffix.lower()
        engine = "xlrd" if ext == ".xls" else "openpyxl"
        new_df = pd.read_excel(io.BytesIO(data), engine=engine)
        new_df.columns = [str(c).strip() for c in new_df.columns]
        def _find_cod(cols):
            import unicodedata
            for c in cols:
                cn = "".join(ch for ch in unicodedata.normalize("NFD", str(c).lower())
                             if unicodedata.category(ch) != "Mn")
                if "articulo" in cn and "desc" not in cn:
                    return c
            return cols[0] if cols else None
        cod_new = _find_cod(list(new_df.columns))
        if cod_new:
            new_df[cod_new] = new_df[cod_new].astype(str).str.strip()
        if MAESTRO_PATH.exists():
            master = pd.read_excel(MAESTRO_PATH, engine="openpyxl")
            master.columns = [str(c).strip() for c in master.columns]
            if len(new_df.columns) != len(master.columns):
                raise ValueError(f"Columnas incompatibles: {len(new_df.columns)} vs {len(master.columns)}")
            cod_m    = _find_cod(list(master.columns))
            existing = set(master[cod_m].astype(str).str.strip().tolist()) if cod_m else set()
            new_df.columns = master.columns
            novelty  = new_df[~new_df[master.columns[0]].isin(existing)]
            combined = pd.concat([master, novelty], ignore_index=True)
            n_new    = len(novelty)
        else:
            combined = new_df; n_new = len(new_df)
        combined.to_excel(MAESTRO_PATH, index=False, engine="openpyxl")
        return n_new
    try:
        n = await asyncio.to_thread(_process)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error procesando maestro: {exc}")
    ForecastEngine.invalidate_caches()
    return JSONResponse({"ok": True, "codigos_nuevos_agregados": n,
                         "ultima_mod": _get_mtime_str(MAESTRO_PATH)})


# ─── MDM: OVERWRITE Maestro (con validación de columnas) ─────────────────────

@app.post("/api/mdm/overwrite-maestro")
async def mdm_overwrite_maestro(file: UploadFile = File(...)):
    """MOD 5: Reemplaza Maestro de Productos con validación estricta de columnas."""
    data = await file.read()
    def _validate():
        _validate_overwrite_cols(data, MAESTRO_PATH, file.filename or "")
    try:
        await asyncio.to_thread(_validate)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    await asyncio.to_thread(lambda: MAESTRO_PATH.write_bytes(data))
    ForecastEngine.invalidate_caches()
    return JSONResponse({"ok": True, "mensaje": "Maestro de Productos sobrescrito.",
                         "ultima_mod": _get_mtime_str(MAESTRO_PATH)})


# ─── MDM: Upload / Overwrite Forecast ────────────────────────────────────────

@app.post("/api/mdm/upload-forecast")
async def mdm_upload_forecast(file: UploadFile = File(...)):
    ext = Path(file.filename).suffix.lower()
    if ext not in (".xlsx", ".xls"):
        raise HTTPException(400, "El pronóstico debe ser .xlsx o .xls")
    data = await file.read()
    await asyncio.to_thread(lambda: FORECAST_PATH.write_bytes(data))
    ForecastEngine.invalidate_caches()
    fecha = _get_forecast_fecha()
    return JSONResponse({
        "ok": True, "mensaje": f"Pronóstico actualizado ({len(data):,} bytes).",
        "forecast_fecha": fecha, "ultima_mod": _get_mtime_str(FORECAST_PATH),
    })


@app.post("/api/mdm/overwrite-forecast")
async def mdm_overwrite_forecast(file: UploadFile = File(...)):
    """Sobrescribe Forecast.xlsx con validación de columnas."""
    ext = Path(file.filename).suffix.lower()
    if ext not in (".xlsx", ".xls"):
        raise HTTPException(400, "El pronóstico debe ser .xlsx o .xls")
    data = await file.read()
    def _validate():
        _validate_overwrite_cols(data, FORECAST_PATH, file.filename or "")
    try:
        await asyncio.to_thread(_validate)
    except ValueError as exc:
        raise HTTPException(400, str(exc))
    await asyncio.to_thread(lambda: FORECAST_PATH.write_bytes(data))
    ForecastEngine.invalidate_caches()
    fecha = _get_forecast_fecha()
    return JSONResponse({
        "ok": True, "mensaje": "Pronóstico sobrescrito.",
        "forecast_fecha": fecha, "ultima_mod": _get_mtime_str(FORECAST_PATH),
    })


# ─── MDM: Motor IA — Run Forecast ────────────────────────────────────────────

@app.post("/api/mdm/run-forecast")
async def mdm_run_forecast():
    for req_path, label in [
        (VENTAS_PATH,  "Ventas 2023-2025 (todas las categorías).xlsx"),
        (MAESTRO_PATH, "Maestro de Productos.xlsx"),
    ]:
        if not req_path.exists():
            raise HTTPException(404, f"Archivo requerido no encontrado: '{label}'.")

    def _run():
        return ForecastEngine.run_and_save(VENTAS_PATH, MAESTRO_PATH, FORECAST_PATH)

    try:
        stats = await asyncio.to_thread(_run)
    except RuntimeError as exc:
        raise HTTPException(503, str(exc))
    except Exception as exc:
        raise HTTPException(500, f"Error en motor de pronóstico: {exc}")

    return JSONResponse({
        "ok":      True,
        "mensaje": (f"Pronóstico regenerado (Holt-Winters Amortiguado). "
                    f"{stats['n_activos']} activos · {stats['n_inactivos']} inactivos · "
                    f"{stats['total_forecast']:,} uds. "
                    f"Período: {stats['periodo_inicio']} → {stats['periodo_fin']}"),
        "stats":          stats,
        "forecast_fecha": _get_forecast_fecha(),
        "ultima_mod":     _get_mtime_str(FORECAST_PATH),
    })
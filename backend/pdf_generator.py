"""
Roadmap Analytics — pdf_generator.py v4.0
==========================================
v4.0:
  - Subtítulo corporativo: "Veterinaria CIMA - App by Roadmap Analytics"
  - Footer: "Veterinaria CIMA - App by Roadmap Analytics"
  - Sin referencias a modelos IA ni emojis.
"""

from __future__ import annotations

import io
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    HRFlowable,
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ─── Paleta corporativa ───────────────────────────────────────────────────────
C_NAVY     = colors.HexColor("#0F2D54")
C_BLUE_LT  = colors.HexColor("#2E6BB0")
C_BLUE_PAL = colors.HexColor("#EEF4FB")
C_ACCENT   = colors.HexColor("#F0A500")
C_WHITE    = colors.white
C_GRAY     = colors.HexColor("#6B7280")
C_TEXT_LT  = colors.HexColor("#374151")
C_TEAL     = colors.HexColor("#0D9488")

PAGE_W, PAGE_H = portrait(A4)
LEFT_M  = 1.0 * cm
RIGHT_M = 1.0 * cm
TOP_M   = 1.5 * cm
BOT_M   = 1.5 * cm

_MESES_ES = [
    "", "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

_BRAND = "Veterinaria CIMA - App by Roadmap Analytics"


def _mes_anio_dinamico(dt: datetime) -> str:
    return f"{_MESES_ES[dt.month]} de {dt.year}"


# ─── Función principal ────────────────────────────────────────────────────────

def generate_pdf(
    df: pd.DataFrame,
    logo_path: Optional[str]         = None,
    roadmap_logo_path: Optional[str] = None,
    single_family: bool              = False,
    forecast_fecha: str              = "",
) -> bytes:
    buf = io.BytesIO()
    now = datetime.now()

    doc = SimpleDocTemplate(
        buf,
        pagesize=portrait(A4),
        leftMargin=LEFT_M,
        rightMargin=RIGHT_M,
        topMargin=TOP_M,
        bottomMargin=BOT_M,
        title=f"CIMA Planificación de Compras — {_mes_anio_dinamico(now)}",
    )

    styles = _build_styles()
    story: list = []

    story += _build_header(styles, logo_path, roadmap_logo_path, forecast_fecha, now)
    story.append(Spacer(1, 0.8 * cm))

    grand_total = 0
    family_totals: dict[str, int] = {}

    for prov in df["PROVEEDOR"].unique():
        df_prov = df[df["PROVEEDOR"] == prov]
        t_prov, prov_total, fam_tots = _build_provider_table(prov, df_prov, single_family)
        grand_total += prov_total
        for fn, ft in fam_tots.items():
            family_totals[fn] = family_totals.get(fn, 0) + ft
        story.append(t_prov)
        story.append(Spacer(1, 0.5 * cm))

    story.append(PageBreak())
    story += _build_summary(family_totals, grand_total, styles)

    doc.build(
        story,
        onFirstPage=_make_footer_fn(now),
        onLaterPages=_make_footer_fn(now),
    )
    buf.seek(0)
    return buf.read()


# ─── Encabezado ──────────────────────────────────────────────────────────────

def _build_header(styles, logo_path, roadmap_logo_path,
                  forecast_fecha: str, now: datetime) -> list:
    """
    v4.0: Subtítulo corporativo, sin referencias a modelos IA.
    """
    def _img(path_str, w_cm, h_cm):
        p = Path(path_str) if path_str else None
        if p and p.exists():
            return Image(str(p), width=w_cm * cm, height=h_cm * cm, kind="proportional")
        return Paragraph("", styles["subtitle"])

    logo_cell    = _img(logo_path,         3.8, 1.8)
    roadmap_cell = _img(roadmap_logo_path, 2.2, 1.2)

    title_para  = Paragraph("<b>PLANIFICACIÓN DE COMPRAS</b>", styles["doc_title"])
    periodo_str = _mes_anio_dinamico(now)

    # v4.0 — branding corporativo, sin IA
    brand_html = (
        f"<br/><font color='#0F2D54'>{_BRAND}</font>"
    )

    fecha_fc_html = ""
    if forecast_fecha and forecast_fecha not in ("", "Sin archivo"):
        fecha_fc_html = (
            f"<br/><font color='#1A4D8F'>"
            f"Última modificación de pronóstico: <b>{forecast_fecha}</b>"
            f"</font>"
        )

    subtitle_para = Paragraph(
        f"Período: <b>{periodo_str}</b> &nbsp;·&nbsp; "
        f"Generado el: {now.strftime('%d/%m/%Y')} a las {now.strftime('%H:%M')} hs."
        f"{fecha_fc_html}"
        f"{brand_html}",
        styles["subtitle"],
    )

    center_content = [title_para, Spacer(1, 3 * mm), subtitle_para]

    header_tbl = Table(
        [[logo_cell, center_content, roadmap_cell]],
        colWidths=["22%", "56%", "22%"],
    )
    header_tbl.setStyle(TableStyle([
        ("ALIGN",         (0, 0), (0, 0), "LEFT"),
        ("ALIGN",         (1, 0), (1, 0), "CENTER"),
        ("ALIGN",         (2, 0), (2, 0), "RIGHT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))

    return [
        header_tbl,
        Spacer(1, 3 * mm),
        HRFlowable(width="100%", thickness=2, color=C_NAVY, spaceAfter=0),
    ]


# ─── Tabla por proveedor ──────────────────────────────────────────────────────

def _build_provider_table(prov: str, df_prov: pd.DataFrame, single_family: bool):
    data: list  = []
    styles: list = []
    row_idx = 0

    data.append([f"PROVEEDOR: {prov.upper()}", "", "", "", ""])
    styles.extend([
        ("SPAN",          (0, row_idx), (-1, row_idx)),
        ("BACKGROUND",    (0, row_idx), (-1, row_idx), C_NAVY),
        ("TEXTCOLOR",     (0, row_idx), (-1, row_idx), C_WHITE),
        ("FONTNAME",      (0, row_idx), (-1, row_idx), "Helvetica-Bold"),
        ("FONTSIZE",      (0, row_idx), (-1, row_idx), 9),
        ("ALIGN",         (0, row_idx), (-1, row_idx), "LEFT"),
        ("TOPPADDING",    (0, row_idx), (-1, row_idx), 5),
        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 5),
    ])
    row_idx += 1

    data.append(["CÓDIGO", "DESCRIPCIÓN DEL ARTÍCULO", "STOCK ACT.", "STOCK MÍN.", "A PEDIR"])
    styles.extend([
        ("BACKGROUND",    (0, row_idx), (-1, row_idx), C_BLUE_LT),
        ("TEXTCOLOR",     (0, row_idx), (-1, row_idx), C_WHITE),
        ("FONTNAME",      (0, row_idx), (-1, row_idx), "Helvetica-Bold"),
        ("FONTSIZE",      (0, row_idx), (-1, row_idx), 8),
        ("ALIGN",         (0, row_idx), (-1, row_idx), "CENTER"),
        ("TOPPADDING",    (0, row_idx), (-1, row_idx), 4),
        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 4),
    ])
    row_idx += 1

    prov_total = 0
    fam_totals: dict = {}

    for familia in df_prov["FAMILIA"].unique():
        df_fam  = df_prov[df_prov["FAMILIA"] == familia]
        fam_tot = int(df_fam["PEDIR"].sum())
        prov_total          += fam_tot
        fam_totals[familia]  = fam_tot

        if not single_family:
            data.append([f"  Familia: {familia}", "", "", "", ""])
            styles.extend([
                ("SPAN",          (0, row_idx), (-1, row_idx)),
                ("BACKGROUND",    (0, row_idx), (-1, row_idx), C_BLUE_PAL),
                ("TEXTCOLOR",     (0, row_idx), (-1, row_idx), C_NAVY),
                ("FONTNAME",      (0, row_idx), (-1, row_idx), "Helvetica-Bold"),
                ("FONTSIZE",      (0, row_idx), (-1, row_idx), 8),
                ("ALIGN",         (0, row_idx), (-1, row_idx), "LEFT"),
                ("TOPPADDING",    (0, row_idx), (-1, row_idx), 3),
                ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 3),
            ])
            row_idx += 1

        for row in df_fam.itertuples():
            data.append([
                str(row.CODIGO),
                str(row.DESCRIPCION)[:65],
                str(int(row.CANTIDAD)),
                str(int(row.S_MIN)),
                str(int(row.PEDIR)),
            ])
            styles.extend([
                ("FONTNAME",  (0, row_idx), (-1, row_idx), "Helvetica"),
                ("FONTSIZE",  (0, row_idx), (-1, row_idx), 8),
                ("ALIGN",     (0, row_idx), (0, row_idx),  "CENTER"),
                ("ALIGN",     (2, row_idx), (-1, row_idx), "RIGHT"),
                ("FONTNAME",  (4, row_idx), (4, row_idx),  "Helvetica-Bold"),
                ("TEXTCOLOR", (4, row_idx), (4, row_idx),  C_NAVY),
                ("LINEBELOW", (0, row_idx), (-1, row_idx), 0.25, colors.HexColor("#E5EAF2")),
                ("TOPPADDING",    (0, row_idx), (-1, row_idx), 2),
                ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 2),
            ])
            row_idx += 1

        if not single_family:
            data.append(["", "", "", f"Subtotal {familia}:", str(fam_tot)])
            styles.extend([
                ("FONTNAME",  (0, row_idx), (-1, row_idx), "Helvetica-Bold"),
                ("FONTSIZE",  (0, row_idx), (-1, row_idx), 8),
                ("TEXTCOLOR", (0, row_idx), (-1, row_idx), C_NAVY),
                ("ALIGN",     (3, row_idx), (-1, row_idx), "RIGHT"),
                ("BACKGROUND",(0, row_idx), (-1, row_idx), colors.HexColor("#F8FAFC")),
                ("LINEBELOW", (0, row_idx), (-1, row_idx), 0.5, C_BLUE_LT),
                ("TOPPADDING",    (0, row_idx), (-1, row_idx), 3),
                ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 3),
            ])
            row_idx += 1

    data.append(["", "", "", f"TOTAL {prov.upper()}:", str(prov_total)])
    styles.extend([
        ("FONTNAME",  (0, row_idx), (-1, row_idx), "Helvetica-Bold"),
        ("FONTSIZE",  (0, row_idx), (-1, row_idx), 9),
        ("TEXTCOLOR", (0, row_idx), (-1, row_idx), C_NAVY),
        ("ALIGN",     (3, row_idx), (-1, row_idx), "RIGHT"),
        ("BACKGROUND",(0, row_idx), (-1, row_idx), colors.HexColor("#E2E8F0")),
        ("TOPPADDING",    (0, row_idx), (-1, row_idx), 5),
        ("BOTTOMPADDING", (0, row_idx), (-1, row_idx), 5),
    ])

    col_widths = [2.4 * cm, 10.2 * cm, 2.0 * cm, 2.2 * cm, 2.2 * cm]
    t = Table(data, colWidths=col_widths, repeatRows=2)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN",   (0, 0), (-1, -1), "MIDDLE"),
        *styles,
    ]))
    return t, prov_total, fam_totals


# ─── Resumen final ────────────────────────────────────────────────────────────

def _build_summary(family_totals: dict, grand_total: int, styles) -> list:
    elements: list = []
    elements.append(Paragraph("<b>RESUMEN GENERAL DE PEDIDOS</b>", styles["summary_title"]))
    elements.append(Spacer(1, 4 * mm))

    rows = []
    for familia, total in sorted(family_totals.items()):
        rows.append([
            Paragraph(f"Total {familia}:", styles["summary_label"]),
            Paragraph(f"<b>{total:,}</b>".replace(",", "."), styles["summary_val"]),
        ])
    rows.append([
        Paragraph("<b>TOTAL UNIDADES A PEDIR:</b>", styles["summary_grand_label"]),
        Paragraph(f"<b>{grand_total:,}</b>".replace(",", "."), styles["summary_grand_val"]),
    ])

    t  = Table(rows, colWidths=[12 * cm, 5 * cm])
    ts = TableStyle([
        ("FONT",          (0, 0),  (-1, -2), "Helvetica",      9),
        ("FONT",          (0, -1), (-1, -1), "Helvetica-Bold", 11),
        ("TOPPADDING",    (0, 0),  (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0),  (-1, -1), 4),
        ("LEFTPADDING",   (0, 0),  (-1, -1), 8),
        ("RIGHTPADDING",  (0, 0),  (-1, -1), 8),
        ("LINEBELOW",     (0, 0),  (-1, -2), 0.25, colors.HexColor("#CBD5E1")),
        ("BACKGROUND",    (0, -1), (-1, -1), C_NAVY),
        ("TEXTCOLOR",     (0, -1), (-1, -1), C_WHITE),
        ("LINEABOVE",     (0, -1), (-1, -1), 1, C_ACCENT),
    ])
    for i in range(len(rows) - 1):
        if i % 2 == 0:
            ts.add("BACKGROUND", (0, i), (-1, i), C_BLUE_PAL)
    t.setStyle(ts)

    wrapper = Table([[t]], colWidths=["100%"])
    wrapper.setStyle(TableStyle([
        ("BOX",           (0, 0), (-1, -1), 1, C_NAVY),
        ("TOPPADDING",    (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("LEFTPADDING",   (0, 0), (-1, -1), 0),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 0),
    ]))
    elements.append(wrapper)
    return elements


# ─── Footer ───────────────────────────────────────────────────────────────────

def _make_footer_fn(now: datetime):
    periodo = _mes_anio_dinamico(now)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(C_GRAY)
        canvas.drawString(
            LEFT_M, BOT_M * 0.6,
            f"Página {canvas.getPageNumber()} · Roadmap Analytics · {periodo}"
        )
        # v4.0 — branding corporativo
        canvas.drawRightString(
            PAGE_W - RIGHT_M, BOT_M * 0.6,
            _BRAND,
        )
        canvas.setStrokeColor(C_NAVY)
        canvas.setLineWidth(0.5)
        canvas.line(LEFT_M, BOT_M * 0.85, PAGE_W - RIGHT_M, BOT_M * 0.85)
        canvas.restoreState()

    return _footer


# ─── Estilos de texto ─────────────────────────────────────────────────────────

def _build_styles() -> dict:
    base = getSampleStyleSheet()
    s: dict = {}
    s["doc_title"] = ParagraphStyle(
        "doc_title", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=16,
        textColor=C_NAVY, alignment=TA_CENTER, spaceAfter=0,
    )
    s["subtitle"] = ParagraphStyle(
        "subtitle", parent=base["Normal"],
        fontName="Helvetica", fontSize=8,
        textColor=C_TEXT_LT, alignment=TA_CENTER, spaceAfter=0,
    )
    s["summary_title"] = ParagraphStyle(
        "summary_title", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=13,
        textColor=C_NAVY, alignment=TA_CENTER, spaceAfter=4,
    )
    s["summary_label"] = ParagraphStyle(
        "summary_label", parent=base["Normal"],
        fontName="Helvetica", fontSize=9,
        textColor=C_NAVY, alignment=TA_RIGHT,
    )
    s["summary_val"] = ParagraphStyle(
        "summary_val", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=9,
        textColor=C_NAVY, alignment=TA_RIGHT,
    )
    s["summary_grand_label"] = ParagraphStyle(
        "summary_grand_label", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=11,
        textColor=C_WHITE, alignment=TA_RIGHT,
    )
    s["summary_grand_val"] = ParagraphStyle(
        "summary_grand_val", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=12,
        textColor=C_ACCENT, alignment=TA_RIGHT,
    )
    return s
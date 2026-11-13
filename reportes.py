"""Generación de reportes: gráficos (PNG), Excel y PDF."""
from pathlib import Path

import matplotlib
import pandas as pd
matplotlib.use("Agg")  # sin ventana: sirve para ejecutar en servidores
import matplotlib.pyplot as plt
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

VERDE, DORADO, GRIS = "#1F4D36", "#F2B705", "#5A685E"


def ars(n) -> str:
    """Formato de pesos argentinos: $ 1.234.567"""
    return "$ " + f"{n:,.0f}".replace(",", ".")


# ---------- Gráficos ----------
def exportar_graficos(r: dict, carpeta: Path) -> dict:
    carpeta.mkdir(parents=True, exist_ok=True)
    rutas = {}
    plt.rcParams.update({"font.size": 10, "axes.spines.top": False, "axes.spines.right": False})

    fig, ax = plt.subplots(figsize=(9, 3.0))
    m = r["por_mes"]
    ax.bar(m["mes"], m["total"] / 1e6, color=VERDE)
    ax.set_title("Ventas por mes", loc="left", fontweight="bold")
    ax.set_ylabel("Millones de $")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    rutas["mes"] = carpeta / "ventas_por_mes.png"
    fig.savefig(rutas["mes"], dpi=160); plt.close(fig)

    fig, ax = plt.subplots(figsize=(5, 3.6))
    p = r["por_producto"].head(8).iloc[::-1]
    ax.barh(p["producto"].str.slice(0, 26), p["total"] / 1e6, color=DORADO)
    ax.set_title("Ventas por producto", loc="left", fontweight="bold")
    ax.set_xlabel("Millones de $")
    fig.tight_layout()
    rutas["producto"] = carpeta / "ventas_por_producto.png"
    fig.savefig(rutas["producto"], dpi=160); plt.close(fig)

    fig, ax = plt.subplots(figsize=(4, 3.6))
    c = r["por_categoria"]
    ax.pie(c["total"], labels=c["categoria"], autopct="%1.0f%%", startangle=90,
           colors=[VERDE, DORADO, "#6E9B7D", "#B9C9BD"], wedgeprops={"linewidth": 1, "edgecolor": "white"})
    ax.set_title("Por categoría", loc="left", fontweight="bold")
    fig.tight_layout()
    rutas["categoria"] = carpeta / "ventas_por_categoria.png"
    fig.savefig(rutas["categoria"], dpi=160); plt.close(fig)
    return rutas


# ---------- Excel ----------
def _dar_formato(ws, formatos: dict | None = None):
    encabezado = PatternFill("solid", fgColor=VERDE.lstrip("#"))
    for celda in ws[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = encabezado
        celda.alignment = Alignment(vertical="center")
    for col in ws.columns:
        letra = get_column_letter(col[0].column)
        ancho = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[letra].width = min(max(ancho + 3, 11), 48)
        if formatos and col[0].value in formatos:
            for c in col[1:]:
                c.number_format = formatos[col[0].value]
    ws.freeze_panes = "A2"


def exportar_excel(ruta: Path, limpios, descartados, r: dict):
    moneda, fecha = '"$" #,##0', "DD/MM/YYYY"
    hojas = {
        "Por mes": (r["por_mes"], {"total": moneda}),
        "Por producto": (r["por_producto"], {"total": moneda}),
        "Por categoría": (r["por_categoria"], {"total": moneda}),
        "Por vendedor": (r["por_vendedor"], {"total": moneda}),
        "Alertas (en revisión)": (r["alertas"], {"fecha": fecha, "total": moneda, "precio_unitario": moneda}),
        "Datos limpios": (limpios, {"fecha": fecha, "total": moneda, "precio_unitario": moneda}),
        "Descartados": (descartados, None),
    }
    with pd.ExcelWriter(ruta, engine="openpyxl") as xl:
        for nombre, (df, _) in hojas.items():
            df.to_excel(xl, sheet_name=nombre, index=False)
        wb = xl.book
        for nombre, (_, formatos) in hojas.items():
            _dar_formato(wb[nombre], formatos)

        ws = wb.create_sheet("Resumen", 0)
        ws["A1"] = "Reporte de ventas"; ws["A1"].font = Font(size=16, bold=True, color=VERDE.lstrip("#"))
        for i, (clave, valor) in enumerate(r["kpis"].items(), start=3):
            ws.cell(i, 1, clave).font = Font(bold=True)
            monetario = clave in ("Ventas totales", "Ticket promedio")
            ws.cell(i, 2, valor).number_format = moneda if monetario else "General"
            ws.cell(i, 2).alignment = Alignment(horizontal="left")
        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 32

        mes = wb["Por mes"]
        grafico = BarChart()
        grafico.title, grafico.height, grafico.width = "Ventas por mes", 8, 17
        grafico.y_axis.title = "Total ($)"
        grafico.add_data(Reference(mes, min_col=4, min_row=1, max_row=mes.max_row), titles_from_data=True)
        grafico.set_categories(Reference(mes, min_col=1, min_row=2, max_row=mes.max_row))
        grafico.legend = None
        ws.add_chart(grafico, "D3")


# ---------- PDF ----------
def exportar_pdf(ruta: Path, r: dict, graficos: dict):
    base = dict(fontName="Helvetica", fontSize=10, leading=14)
    titulo = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=colors.HexColor(VERDE))
    sub = ParagraphStyle("s", **base, textColor=colors.HexColor(GRIS))
    h2 = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=12, leading=16, spaceBefore=12, spaceAfter=4)
    cuerpo = ParagraphStyle("b", **base)
    k = r["kpis"]

    doc = SimpleDocTemplate(str(ruta), pagesize=A4, leftMargin=2*cm, rightMargin=2*cm, topMargin=1.3*cm,
                            bottomMargin=1.3*cm, title="Reporte de ventas")
    ancho = A4[0] - 4*cm
    hist = [
        [ars(k["Ventas totales"]), ars(k["Ticket promedio"]), f"{k['Unidades vendidas']}", k["Mejor mes"]],
        ["Ventas totales", "Ticket promedio", "Unidades vendidas", "Mejor mes"],
    ]
    kpi = Table(hist, colWidths=[ancho / 4] * 4)
    kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E8EEE4")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 13),
        ("FONTSIZE", (0, 1), (-1, 1), 8.5), ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor(GRIS)),
        ("TOPPADDING", (0, 0), (-1, 0), 10), ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
        ("LEFTPADDING", (0, 0), (-1, -1), 10)]))

    lado = Table([[Image(str(graficos["producto"]), width=ancho * 0.6, height=ancho * 0.6 * 3.6 / 5),
                   Image(str(graficos["categoria"]), width=ancho * 0.4, height=ancho * 0.4 * 3.6 / 4)]],
                 colWidths=[ancho * 0.6, ancho * 0.4])
    lado.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))

    filas_mot = [["Motivo de descarte", "Filas"]] + [[m, str(n)] for m, n in r["motivos"].itertuples(index=False)]
    mot = Table(filas_mot, colWidths=[ancho * 0.75, ancho * 0.25])
    mot.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(VERDE)),
                             ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                             ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                             ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6F0")]),
                             ("ALIGN", (1, 0), (1, -1), "RIGHT")]))

    top = r["por_producto"].head(3)
    texto_top = "; ".join(f"{p} ({ars(t)})" for p, t in zip(top["producto"], top["total"]))
    historia = [
        Paragraph("Reporte de ventas", titulo),
        Paragraph(f"Período: {k['Período']}", sub), Spacer(1, 10), kpi,
        Paragraph("Evolución mensual", h2), Image(str(graficos["mes"]), width=ancho, height=ancho * 3.0 / 9),
        Paragraph("Productos y categorías", h2), lado,
        Paragraph(f"<b>Los 3 productos con más facturación:</b> {texto_top}.", cuerpo),
        Paragraph("Calidad de los datos", h2),
        Paragraph(f"Se leyeron {k['Filas leídas']} filas: {k['Filas válidas']} válidas, {k['Filas descartadas']} descartadas y "
                  f"{k['Filas en revisión (fuera de totales)']} en revisión por tener cantidades atípicas (posibles errores de tipeo). "
                  "Estas últimas no se suman a los totales; el detalle está en la hoja &quot;Alertas&quot; del Excel.", cuerpo),
        Spacer(1, 6), mot,
    ]
    doc.build(historia)

"""Procesador de ventas: limpia un CSV desordenado y genera reportes.

Uso:
    python procesar_ventas.py ventas_sucias.csv --salida salida

Genera en la carpeta de salida:
    reporte_ventas.xlsx   Excel con resumen, tablas, gráfico, datos limpios y descartados
    reporte_ventas.pdf    Resumen ejecutivo de 2 páginas
    graficos/*.png        Gráficos para presentaciones
"""
import argparse
from pathlib import Path

import pandas as pd

from catalogo import CATALOGO, NOMBRES_CANONICOS
from reportes import exportar_excel, exportar_graficos, exportar_pdf

FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y")
COLUMNAS_OBLIGATORIAS = {"fecha", "producto", "cantidad", "precio_unitario"}


# ---------- 1. Lectura y conversión de valores ----------
def cargar(ruta: Path) -> pd.DataFrame:
    """Lee todo como texto para no perder información antes de limpiar."""
    df = pd.read_csv(ruta, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    df.columns = [c.strip().lower() for c in df.columns]
    faltantes = COLUMNAS_OBLIGATORIAS - set(df.columns)
    if faltantes:
        raise SystemExit(f"Faltan columnas obligatorias: {', '.join(sorted(faltantes))}")
    df.insert(0, "fila_original", range(2, len(df) + 2))  # fila 1 = encabezado
    return df


def parsear_fecha(texto: str):
    for formato in FORMATOS_FECHA:
        try:
            return pd.to_datetime(texto.strip(), format=formato)
        except ValueError:
            continue
    return pd.NaT


def parsear_monto(texto: str):
    """Convierte '$7.900,00', '7900.0' o '7900' en número. Devuelve NaN si no se puede."""
    s = texto.strip().replace("$", "").replace(" ", "")
    if not s:
        return float("nan")
    if "," in s:                       # formato argentino: punto = miles, coma = decimal
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return float("nan")


def normalizar_producto(texto: str):
    clave = " ".join(texto.split()).casefold()
    return NOMBRES_CANONICOS.get(clave)  # None si no está en el catálogo


# ---------- 2. Limpieza ----------
def limpiar(df: pd.DataFrame):
    """Devuelve (datos_limpios, descartados). Cada fila descartada trae su motivo."""
    d = df.copy()
    d["fecha"] = d["fecha"].map(parsear_fecha)
    d["producto"] = d["producto"].map(normalizar_producto)
    d["cantidad"] = d["cantidad"].map(parsear_monto)
    d["precio_unitario"] = d["precio_unitario"].map(parsear_monto)

    d["motivo"] = ""

    def marcar(mascara, texto):
        d.loc[mascara & (d["motivo"] == ""), "motivo"] = texto

    marcar(d["fecha"].isna(), "Fecha inválida o vacía")
    marcar(d["producto"].isna(), "Producto no reconocido")
    marcar(d["cantidad"].isna() | (d["cantidad"] <= 0), "Cantidad inválida")
    marcar(d["precio_unitario"].isna() | (d["precio_unitario"] <= 0), "Precio inválido")

    # duplicados: misma venta repetida exactamente (se conserva la primera)
    validas = d[d["motivo"] == ""]
    claves = ["fecha", "producto", "cantidad", "precio_unitario", "vendedor", "medio_pago"]
    claves = [c for c in claves if c in d.columns]
    dup = validas.duplicated(subset=claves, keep="first")
    d.loc[dup.index[dup], "motivo"] = "Fila duplicada"

    descartados = df.loc[d["motivo"] != ""].copy()   # conserva el texto original
    descartados["motivo"] = d.loc[d["motivo"] != "", "motivo"]

    limpios = d[d["motivo"] == ""].drop(columns="motivo").copy()
    limpios["cantidad"] = limpios["cantidad"].astype(int)
    limpios["categoria"] = limpios["producto"].map(lambda p: CATALOGO[p][0])  # completa vacías
    for col in ("vendedor", "medio_pago"):
        if col not in limpios:               # el archivo puede no traer estas columnas
            limpios[col] = "Sin dato"
        limpios[col] = limpios[col].str.strip().str.title().replace("", "Sin dato")
    limpios["total"] = limpios["cantidad"] * limpios["precio_unitario"]
    limpios["mes"] = limpios["fecha"].dt.strftime("%Y-%m")
    columnas = ["fila_original", "fecha", "mes", "producto", "categoria", "cantidad",
                "precio_unitario", "total", "vendedor", "medio_pago"]
    return limpios[columnas].sort_values("fecha").reset_index(drop=True), descartados.reset_index(drop=True)


# ---------- 3. Análisis ----------
def separar_atipicas(limpios: pd.DataFrame):
    """Separa ventas con cantidades atípicas (método IQR): suelen ser errores de tipeo.

    Quedan 'en revisión': no se suman a los totales hasta que alguien las confirme.
    """
    q1, q3 = limpios["cantidad"].quantile([0.25, 0.75])
    limite = q3 + 3 * (q3 - q1)
    es_atipica = limpios["cantidad"] > limite
    alertas = limpios[es_atipica].copy()
    alertas["motivo"] = f"Cantidad mayor a {limite:,.0f} (posible error de tipeo)"
    return limpios[~es_atipica].reset_index(drop=True), alertas.reset_index(drop=True)


def analizar(limpios: pd.DataFrame, descartados: pd.DataFrame, alertas: pd.DataFrame, filas_leidas: int) -> dict:
    total = limpios["total"].sum()
    por_mes = (limpios.groupby("mes")
               .agg(ventas=("total", "size"), unidades=("cantidad", "sum"), total=("total", "sum"))
               .reset_index())
    por_producto = (limpios.groupby(["producto", "categoria"])
                    .agg(unidades=("cantidad", "sum"), total=("total", "sum"))
                    .reset_index().sort_values("total", ascending=False))
    por_producto["participacion_%"] = (por_producto["total"] / total * 100).round(1)
    por_categoria = (limpios.groupby("categoria")
                     .agg(unidades=("cantidad", "sum"), total=("total", "sum"))
                     .reset_index().sort_values("total", ascending=False))
    por_vendedor = (limpios.groupby("vendedor")
                    .agg(ventas=("total", "size"), total=("total", "sum"))
                    .reset_index().sort_values("total", ascending=False))


    mejor_mes = por_mes.loc[por_mes["total"].idxmax()]
    kpis = {
        "Filas leídas": filas_leidas,
        "Filas válidas": len(limpios),
        "Filas descartadas": len(descartados),
        "Filas en revisión (fuera de totales)": len(alertas),
        "Ventas totales": total,
        "Ticket promedio": limpios["total"].mean(),
        "Unidades vendidas": int(limpios["cantidad"].sum()),
        "Producto más vendido ($)": por_producto.iloc[0]["producto"],
        "Mejor mes": f"{mejor_mes['mes']}",
        "Período": f"{limpios['fecha'].min():%d/%m/%Y} al {limpios['fecha'].max():%d/%m/%Y}",
    }
    motivos = descartados["motivo"].value_counts().rename_axis("motivo").reset_index(name="filas")
    return dict(kpis=kpis, por_mes=por_mes, por_producto=por_producto, por_categoria=por_categoria,
                por_vendedor=por_vendedor, alertas=alertas, motivos=motivos)


# ---------- 4. Programa principal ----------
def main():
    ap = argparse.ArgumentParser(description="Limpia un CSV de ventas y genera reportes.")
    ap.add_argument("entrada", type=Path, help="Archivo CSV de ventas")
    ap.add_argument("--salida", type=Path, default=Path("salida"), help="Carpeta de salida")
    args = ap.parse_args()

    if not args.entrada.exists():
        raise SystemExit(f"No existe el archivo: {args.entrada}")
    args.salida.mkdir(parents=True, exist_ok=True)

    crudo = cargar(args.entrada)
    limpios, descartados = limpiar(crudo)
    if limpios.empty:
        raise SystemExit("No quedaron filas válidas después de la limpieza.")
    limpios, alertas = separar_atipicas(limpios)
    r = analizar(limpios, descartados, alertas, len(crudo))

    graficos = exportar_graficos(r, args.salida / "graficos")
    exportar_excel(args.salida / "reporte_ventas.xlsx", limpios, descartados, r)
    exportar_pdf(args.salida / "reporte_ventas.pdf", r, graficos)

    k = r["kpis"]
    print(f"Filas leídas: {k['Filas leídas']} | válidas: {k['Filas válidas']} | "
          f"descartadas: {k['Filas descartadas']} | en revisión: {k['Filas en revisión (fuera de totales)']}")
    print(r["motivos"].to_string(index=False))
    print(f"Ventas totales: $ {k['Ventas totales']:,.0f}")
    print(f"Archivos generados en: {args.salida.resolve()}")


if __name__ == "__main__":
    main()

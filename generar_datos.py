"""Genera un CSV de ventas 'sucio' para probar el procesador.

Simula los problemas típicos de un archivo real: fechas en distintos formatos,
nombres de producto mal escritos, precios con símbolos y comas, celdas vacías,
cantidades inválidas, errores de tipeo y filas duplicadas.
"""
import csv
import random
from datetime import date, timedelta

from catalogo import CATALOGO

random.seed(42)
VENDEDORES = ["Ana", "Bruno", "Carla"]
MEDIOS = ["Efectivo", "Transferencia", "Tarjeta", "Mercado Pago"]
INICIO, DIAS = date(2026, 1, 1), 304  # enero a octubre de 2026


def ensuciar_nombre(nombre: str) -> str:
    return random.choice([nombre, nombre, nombre.lower(), nombre.upper(),
                          f"  {nombre} ", nombre.replace(" ", "  ")])


def formatear_precio(p: int) -> str:
    r = random.random()
    if r < 0.5:
        return str(p)
    if r < 0.8:  # formato argentino: $28.900,00
        return f"${p:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{p}.0"


def formatear_fecha(d: date) -> str:
    r = random.random()
    if r < 0.03:
        return "sin fecha"
    return d.isoformat() if r < 0.55 else d.strftime("%d/%m/%Y")


def generar_filas(n: int = 320) -> list[list[str]]:
    nombres = list(CATALOGO)
    filas = []
    for _ in range(n):
        nombre = random.choices(nombres, weights=[6, 5, 8, 7, 12, 6, 4, 2])[0]
        categoria, precio = CATALOGO[nombre]
        d = INICIO + timedelta(days=random.randrange(DIAS))
        cantidad = str(random.choices([1, 2, 3, 4, 5], weights=[50, 25, 12, 8, 5])[0])
        r = random.random()
        if r < 0.02:
            cantidad = ""
        elif r < 0.03:
            cantidad = "-1"
        precio_txt = "" if random.random() < 0.02 else formatear_precio(precio)
        filas.append([
            formatear_fecha(d), ensuciar_nombre(nombre),
            categoria if random.random() < 0.6 else "",
            cantidad, precio_txt, random.choice(VENDEDORES), random.choice(MEDIOS),
        ])
    for _ in range(2):                      # errores de tipeo: cantidades absurdas
        filas[random.randrange(len(filas))][3] = "250"
    filas[random.randrange(len(filas))][1] = "Mate de plastico"   # producto que no existe
    filas += [list(random.choice(filas)) for _ in range(12)]      # duplicados exactos
    random.shuffle(filas)
    return filas


if __name__ == "__main__":
    with open("ventas_sucias.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["fecha", "producto", "categoria", "cantidad", "precio_unitario", "vendedor", "medio_pago"])
        w.writerows(generar_filas())
    print("Archivo generado: ventas_sucias.csv")

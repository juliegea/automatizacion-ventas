# Automatización de reportes de ventas

Script en Python que toma un archivo CSV de ventas **desordenado** (como los que mandan los
clientes reales), lo limpia y genera automáticamente un **reporte en Excel**, un **PDF
ejecutivo** y **gráficos** listos para presentar.

## Qué problemas resuelve

| Problema del archivo | Qué hace el script |
|---|---|
| Fechas en distintos formatos (`2026-05-03`, `03/05/2026`) o vacías | Las unifica; descarta las inválidas |
| Nombres de producto mal escritos (`  MATE  camionero `) | Los normaliza contra el catálogo |
| Precios como `$28.900,00`, `28900.0` o `28900` | Los convierte a número |
| Categoría vacía | La completa desde el catálogo |
| Cantidades vacías o negativas, precios faltantes | Descarta la fila y registra el motivo |
| Filas duplicadas | Conserva la primera y descarta el resto |
| Errores de tipeo (cantidad `250` en vez de `2`) | Las aparta "en revisión" y no las suma a los totales |

Nada se pierde en silencio: cada fila descartada queda en la hoja **Descartados** con su motivo.

## Uso

```bash
pip install -r requirements.txt
python generar_datos.py                                  # crea ventas_sucias.csv de ejemplo
python procesar_ventas.py ventas_sucias.csv --salida salida
```

Con tus propios datos, el CSV necesita al menos estas columnas: `fecha`, `producto`,
`cantidad`, `precio_unitario` (opcionales: `categoria`, `vendedor`, `medio_pago`).

## Qué genera (carpeta `salida/`)

- `reporte_ventas.xlsx`: hojas Resumen (con gráfico), Por mes, Por producto, Por categoría,
  Por vendedor, Alertas (en revisión), Datos limpios y Descartados.
- `reporte_ventas.pdf`: resumen de una página con indicadores, gráficos y calidad de datos.
- `graficos/*.png`: ventas por mes, por producto y por categoría.

## Estructura

```
catalogo.py          Productos, categorías y precios de lista
generar_datos.py     Crea un CSV de prueba con errores típicos
procesar_ventas.py   Lectura, limpieza, detección de atípicos y análisis
reportes.py          Exportación a Excel, PDF y gráficos
```

## Supuestos y límites

- Un punto sin coma se interpreta como decimal (`7900.0`); con coma, el punto es separador de
  miles (`7.900,00`).
- Dos filas idénticas (misma fecha, producto, cantidad, precio, vendedor y medio de pago) se
  consideran duplicado. En un negocio real conviene confirmar esta regla con el cliente.
- El umbral de "cantidad atípica" usa el rango intercuartil (Q3 + 3 × IQR).
- Los datos de ejemplo son generados al azar; no corresponden a un comercio real.

## Posibles mejoras

Leer archivos `.xlsx` de entrada, configurar reglas por archivo YAML, enviar el reporte por
mail, y programar la ejecución diaria.

**Ejemplo del reporte generado:**

<img width="520" alt="Reporte de ventas en PDF con indicadores, gráficos y calidad de datos" src="https://github.com/user-attachments/assets/116875f5-5138-4bdf-940a-56e2af0a723a" />



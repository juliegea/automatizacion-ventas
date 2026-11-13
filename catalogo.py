"""Catálogo de productos del comercio (nombre canónico -> categoría y precio de lista)."""

CATALOGO = {
    "Mate imperial de calabaza": ("Mates", 28900),
    "Mate camionero": ("Mates", 18500),
    "Bombilla pico de loro": ("Bombillas", 9800),
    "Bombilla de acero inoxidable": ("Bombillas", 6400),
    "Yerba mate artesanal 1 kg": ("Yerbas", 7900),
    "Yerba compuesta con hierbas 500 g": ("Yerbas", 4600),
    "Termo de acero 1 litro": ("Termos", 45900),
    "Set matero completo": ("Mates", 52000),
}

# clave en minúsculas -> nombre canónico (para reconocer variantes mal escritas)
NOMBRES_CANONICOS = {nombre.casefold(): nombre for nombre in CATALOGO}

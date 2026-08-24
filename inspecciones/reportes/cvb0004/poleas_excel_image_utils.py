"""Ranuras de fotografías exclusivas de POLEAS CVB004.

Los rangos reproducen los cuadros visibles de la plantilla maestra. Las
dimensiones de filas y columnas nunca se modifican: cada imagen se adapta al
cuadro que le corresponde.
"""

from .image_utils import insertar_imagen_ajustada, medir_rango_celdas_px


PHOTO_SLOTS = {
    1: ("E97:AB114", "AD97:AW114"),
    2: ("E131:AB144", "AD131:AW144"),
    3: ("E162:AB181", "AD162:AW181", "E201:AB220", "AD201:AW220"),
    4: ("D224:AB241", "AD224:AW241"),
    5: ("D247:AB263", "AD247:AW263"),
    6: ("D281:AX303",),
    7: ("D323:AA342", "AC323:AX342"),
    8: ("D360:AB378", "AD360:AW378", "E395:AB414", "AD395:AW414"),
    9: ("D432:R450", "S432:AG450", "AH432:AX450"),
}


def insertar_fotos_polea_cvb0004(ws, numero_polea, fotos_con_ruta, margen_px=4):
    """Inserta las fotos en las ranuras maestras sin alterar la hoja."""
    ranuras = PHOTO_SLOTS[numero_polea]
    insertadas = 0
    for (_foto, ruta), rango in zip(fotos_con_ruta, ranuras):
        ancho, alto = medir_rango_celdas_px(ws, rango)
        if insertar_imagen_ajustada(ws, ruta, rango, ancho, alto, margen_px):
            insertadas += 1
    return insertadas

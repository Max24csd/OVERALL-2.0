import re
from collections import OrderedDict
from pathlib import Path

from inspecciones.reportes.cvb0003.mappings.faja import (
    EMPALME_PHOTO_GROUPS,
    EMPALME_PHOTO_SLOTS,
)


MEASUREMENT_FIELDS = tuple("abcdefg")
PHOTOS_PER_PAGE = 9


def resumen_medicion_faja(*grupos):
    candidatos = []
    for mediciones in grupos:
        for medicion in mediciones:
            componente = (
                getattr(medicion, "empalme", "")
                or getattr(medicion, "tramo", "")
                or "componente"
            )
            punto = (
                getattr(medicion, "posicion", "")
                or getattr(medicion, "medicion", "")
                or "-"
            )
            for campo in MEASUREMENT_FIELDS:
                valor = getattr(medicion, campo, None)
                if valor is not None:
                    candidatos.append(
                        (valor, componente, punto, campo.upper())
                    )
    if not candidatos:
        return "-"
    valor, componente, punto, columna = min(
        candidatos,
        key=lambda item: item[0],
    )
    return (
        f"Espesor mínimo hallado fue de {valor:.2f} mm en "
        f"{componente}, {punto}, punto {columna}."
    )


def _photo_anchor(photo):
    description = photo.descripcion or ""
    match = re.search(r"anclaje\s+([A-Z]+\d+)", description, re.IGNORECASE)
    if match:
        return match.group(1).upper()
    filename = Path(photo.imagen.name or "").name
    match = re.search(r"_(?:empalmes|carga|retorno)_([a-z]+\d+)_", filename)
    return match.group(1).upper() if match else None


def _empalme_identification(photo):
    anchor = _photo_anchor(photo)
    caption_by_anchor = {
        cell_range.split(":", 1)[0].upper(): caption
        for cell_range, caption in EMPALME_PHOTO_SLOTS
    }
    identification_by_caption = {
        group["caption"]: group["identification"]
        for group in EMPALME_PHOTO_GROUPS
    }
    if anchor in caption_by_anchor:
        return identification_by_caption.get(caption_by_anchor[anchor])
    text = f"{photo.codigo_dano or ''} {photo.descripcion or ''}"
    match = re.search(r"\bE-?0?(\d{1,2})\b", text, re.IGNORECASE)
    if match:
        number = int(match.group(1))
        for group in EMPALME_PHOTO_GROUPS:
            if re.search(rf"\bE-0?{number}\b", group["identification"]):
                return group["identification"]
    return "EMPALMES"


def paginas_fotograficas_faja(fotos, seccion):
    groups = OrderedDict()
    for photo in fotos:
        if seccion == "EMPALMES":
            identification = _empalme_identification(photo)
        elif seccion == "CARGA":
            identification = "TRAMOS DE CARGA"
        else:
            identification = "TRAMOS DE RETORNO"
        groups.setdefault(identification, []).append(photo)

    pages = []
    for identification, photos in groups.items():
        chunks = [
            photos[index:index + PHOTOS_PER_PAGE]
            for index in range(0, len(photos), PHOTOS_PER_PAGE)
        ]
        for index, chunk in enumerate(chunks):
            continuation = " (CONTINUACIÓN)" if index else ""
            pages.append({
                "titulo": f"INSPECCIÓN VISUAL – {identification}{continuation}",
                "fotos": chunk,
                "mostrar_resumen": index == len(chunks) - 1,
            })
    return pages

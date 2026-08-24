from __future__ import annotations

import re
from datetime import date

from django.db.models import Max


SUFIJOS_CVB0003 = {
    "FAJA": "FAJA",
    "POLEAS": "POLEAS",
    "LIFE_SHAFT": "LIFE-SHAFT",
}


def es_tag_cvb0003(tag: str | None) -> bool:
    tag_normalizado = re.sub(r"[^A-Z0-9]", "", (tag or "").upper())
    return tag_normalizado in {"CVB003", "CVB0003", "0220CVB003", "0220CVB0003"}


def generar_codigo_cvb0003(
    fecha_inspeccion: date,
    tipo_reporte: str,
) -> str:
    sufijo = SUFIJOS_CVB0003.get(tipo_reporte)
    if not sufijo:
        raise ValueError(f"Tipo de reporte CVB0003 no soportado: {tipo_reporte}")
    return (
        f"{fecha_inspeccion.strftime('%Y%m%d')}"
        f"-VTUT-CVB0003-{sufijo}"
    )


def _es_codigo_historico_emitido(inspeccion) -> bool:
    if not inspeccion.pk:
        return False

    anterior = (
        inspeccion.__class__.objects.filter(pk=inspeccion.pk)
        .values(
            "codigo_reporte",
            "fecha_inspeccion",
            "fecha_programada",
            "tipo",
        )
        .first()
    )
    if not anterior:
        return False

    fecha_anterior = anterior["fecha_inspeccion"]
    if fecha_anterior:
        codigo_segun_regla = generar_codigo_cvb0003(
            fecha_anterior,
            anterior["tipo"],
        )
        if anterior["codigo_reporte"] != codigo_segun_regla:
            return True

    ultima_fecha_programada = (
        inspeccion.__class__.objects.filter(faja_id=inspeccion.faja_id)
        .aggregate(fecha=Max("fecha_programada"))["fecha"]
    )
    return bool(
        ultima_fecha_programada
        and anterior["fecha_programada"]
        and anterior["fecha_programada"] < ultima_fecha_programada
    )


def sincronizar_codigo_cvb0003(
    inspeccion,
    *,
    preservar_historico: bool = True,
) -> bool:
    """
    Sincroniza el código si la inspección corresponde a CVB003.

    Retorna True cuando el registro pertenece a CVB003, incluso si se
    conserva el código histórico o todavía no existe fecha de inspección.
    """
    tag = getattr(getattr(inspeccion, "faja", None), "tag", "")
    if not es_tag_cvb0003(tag):
        return False

    if not inspeccion.fecha_inspeccion:
        return True

    if preservar_historico and _es_codigo_historico_emitido(inspeccion):
        return True

    inspeccion.codigo_reporte = generar_codigo_cvb0003(
        inspeccion.fecha_inspeccion,
        inspeccion.tipo,
    )
    return True

from decimal import Decimal

from inspecciones.models import Inspeccion
from inspecciones.reportes.campaign_utils import es_campana, resultados_por_modalidad


TAGS_CVB0001 = {
    "CVB0001",
    "CVB001",
    "0220-CVB-0001",
    "0220-CVB0001",
    "0220-CVB-001",
}


def es_cvb0001(inspeccion):
    return (inspeccion.faja.tag or "").upper().strip() in TAGS_CVB0001


def minimo_componente(componente):
    candidatos = []
    for medicion in componente.mediciones.order_by("orden", "punto"):
        for letra in "abcdefg":
            valor = getattr(medicion, letra)
            if valor is not None:
                candidatos.append((Decimal(valor), letra.upper(), medicion.punto))
    return min(candidatos, key=lambda item: item[0]) if candidatos else None


def condicion(componente):
    return dict(Inspeccion.Condicion.choices).get(
        componente.condicion, componente.condicion or "No medido"
    )


def combinar_textos(automatico, manual):
    automatico = (automatico or "").strip()
    manual = (manual or "").strip()
    if not automatico:
        return manual
    if not manual or manual.casefold() in automatico.casefold():
        return automatico
    if automatico.casefold() in manual.casefold():
        return manual
    return f"{automatico}\n\nOBSERVACIÓN MANUAL:\n{manual}"


def observacion_life_shaft(shaft):
    partes = []
    for modalidad, minimo in resultados_por_modalidad(shaft):
        if minimo is None:
            tecnico = "No existen mediciones registradas."
        else:
            valor, punto, _radial = minimo
            tecnico = f"El espesor mínimo hallado fue de {valor:.2f} mm en el punto {punto}. Condición {condicion(shaft).upper()}."
        partes.append(f"{modalidad}: {tecnico}" if es_campana(shaft) else tecnico)
    tecnico = "\n".join(partes)
    manuales = " ".join(
        texto.strip()
        for texto in (shaft.observacion_visual, shaft.observacion_medicion)
        if (texto or "").strip()
    )
    texto = f"LIVESHAFT #{shaft.numero:02d}:\n* {tecnico}"
    return combinar_textos(texto, manuales)


def observaciones_life_shaft(shafts, manual_general=""):
    automatico = "\n\n".join(observacion_life_shaft(shaft) for shaft in shafts)
    return combinar_textos(automatico, manual_general)


def conclusiones_life_shaft(shafts, manual_general=""):
    lineas = []
    for shaft in shafts:
        for modalidad, minimo in resultados_por_modalidad(shaft):
            detalle = f"; espesor mínimo {minimo[0]:.2f} mm" if minimo else ""
            prefijo = f"LIVESHAFT #{shaft.numero:02d}"
            if es_campana(shaft):
                prefijo += f" - {modalidad}"
            lineas.append(f"{prefijo} condición {condicion(shaft).upper()}{detalle}.")
        if (shaft.recomendaciones or "").strip():
            lineas.append(shaft.recomendaciones.strip())
    lineas.append(
        "Seguir con el plan de inspecciones programadas a fin de evaluar la condición del equipo."
    )
    return combinar_textos("\n".join(lineas), manual_general)

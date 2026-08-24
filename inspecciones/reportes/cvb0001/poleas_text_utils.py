from .text_utils import combinar_textos, condicion, minimo_componente
from inspecciones.reportes.campaign_utils import es_campana, resultados_por_modalidad


def observaciones_poleas(poleas, manual_general=""):
    bloques = []
    for polea in poleas:
        manual = " ".join(
            texto.strip()
            for texto in (polea.observacion_visual, polea.observacion_medicion)
            if (texto or "").strip()
        )
        partes = []
        for modalidad, minimo in resultados_por_modalidad(polea):
            tecnico = f"El espesor mínimo encontrado es de {minimo[0]:.2f} mm en el punto {minimo[1]}." if minimo else "Evaluación correspondiente a inspección visual."
            etiqueta = f" - {modalidad}" if es_campana(polea) else ""
            partes.append(f"Polea #{polea.numero:02d}{etiqueta}:\n* {tecnico} Condición {condicion(polea).upper()}.")
        bloque = "\n".join(partes)
        bloques.append(combinar_textos(bloque, manual))
    return combinar_textos("\n\n".join(bloques), manual_general)


def conclusiones_poleas(poleas, manual_general=""):
    lineas = []
    for polea in poleas:
        manual = " ".join(
            texto.strip()
            for texto in (
                polea.observacion_visual,
                polea.observacion_medicion,
                polea.recomendaciones,
            )
            if (texto or "").strip()
        )
        for modalidad, _minimo in resultados_por_modalidad(polea):
            etiqueta = f" - {modalidad}" if es_campana(polea) else ""
            linea = f"Polea {polea.numero:02d}{etiqueta} condición {condicion(polea).upper()}."
            lineas.append(f"{linea} {manual}".strip())
    lineas.append(
        "Continuar con el plan de inspecciones programadas a fin de evaluar la condición del equipo."
    )
    return combinar_textos("\n".join(lineas), manual_general)

from decimal import Decimal

from django import forms

from inspecciones.forms import InspeccionForm
from inspecciones.models import Inspeccion
from inspecciones.reportes.campaign_utils import es_campana, resultados_por_modalidad


TAGS_CVB0004 = {
    "CVB0004",
    "CVB004",
    "0220-CVB-0004",
    "0220-CVB0004",
    "0220-CVB-004",
}


class InspeccionCVB0004Form(InspeccionForm):
    class Meta(InspeccionForm.Meta):
        fields = [
            "planta",
            "proceso",
            "etapa",
            *InspeccionForm.Meta.fields,
        ]
        widgets = {
            **InspeccionForm.Meta.widgets,
            "planta": forms.TextInput(attrs={"class": "form-control"}),
            "proceso": forms.TextInput(attrs={"class": "form-control"}),
            "etapa": forms.TextInput(attrs={"class": "form-control"}),
        }


def es_cvb0004(inspeccion):
    return (inspeccion.faja.tag or "").upper().strip() in TAGS_CVB0004


def nombre_usuario(usuario):
    if usuario is None:
        return ""
    return (usuario.get_full_name() or usuario.username or "").strip()


def nombre_campo(inspeccion, campo, usuario_respaldo):
    return (getattr(inspeccion, campo, "") or "").strip() or nombre_usuario(
        usuario_respaldo
    )


def minimo_componente(componente):
    candidatos = []
    for medicion in componente.mediciones.order_by("orden", "punto"):
        for letra in "abcdefg":
            valor = getattr(medicion, letra)
            if valor is not None:
                candidatos.append((Decimal(valor), letra.upper(), medicion.punto))
    return min(candidatos, key=lambda item: item[0]) if candidatos else None


def etiqueta_condicion(componente):
    valor = componente.condicion
    etiquetas = dict(Inspeccion.Condicion.choices)
    return etiquetas.get(valor, valor or "No medido")


def generar_observacion_life_shaft(life_shaft):
    condicion = etiqueta_condicion(life_shaft)
    lineas = []
    for modalidad, minimo in resultados_por_modalidad(life_shaft):
        if minimo is None:
            continue
        valor, letra, _fila = minimo
        sufijo = f" - {modalidad}" if es_campana(life_shaft) else ""
        lineas.append(f"LIVESHAFT #{life_shaft.numero:02d}{sufijo}:\n* El espesor minimo hallado fue de {valor:.2f} mm en el punto {letra}. {condicion}.")
    return "\n".join(lineas)


def generar_observaciones_life_shaft(shafts):
    return "\n\n".join(
        texto for shaft in shafts if (texto := generar_observacion_life_shaft(shaft))
    )


def generar_conclusiones_life_shaft(shafts):
    lineas = []
    for shaft in shafts:
        for modalidad, _minimo in resultados_por_modalidad(shaft):
            sufijo = f" - {modalidad}" if es_campana(shaft) else ""
            lineas.append(f"LIVESHAFT #{shaft.numero:02d}{sufijo} {etiqueta_condicion(shaft)}")
    lineas.append(
        "Seguir con el plan de inspecciones programadas a fin de evaluar condicion."
    )
    return "\n".join(lineas)


def generar_observaciones_poleas(poleas):
    bloques = []
    for polea in poleas:
        for modalidad, minimo in resultados_por_modalidad(polea):
            if minimo is None:
                continue
            valor, letra, _fila = minimo
            sufijo = f" - {modalidad}" if es_campana(polea) else ""
            bloques.append(f"Polea #{polea.numero:02d}{sufijo}:\n* El espesor minimo encontrado es de {valor:.2f} mm en el punto {letra}.")
    return "\n\n".join(bloques)


def generar_conclusiones_poleas(poleas):
    lineas = []
    for polea in poleas:
        for modalidad, _minimo in resultados_por_modalidad(polea):
            sufijo = f" - {modalidad}" if es_campana(polea) else ""
            lineas.append(f"Polea #{polea.numero:02d}{sufijo}: Lagging de polea - Condicion {etiqueta_condicion(polea)}")
    lineas.append(
        "Seguir con el plan de inspecciones programadas a fin de evaluar condicion del equipo."
    )
    return "\n".join(lineas)

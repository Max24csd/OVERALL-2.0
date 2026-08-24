from inspecciones.forms import (
    crear_medicion_life_shaft_campana_formset,
    crear_medicion_polea_campana_formset,
)
from inspecciones.models import FaseCampana, TipoMedicionComponente


FASES_CAMPANA = (FaseCampana.INICIO, FaseCampana.FIN)


def es_campana(componente):
    return componente.tipo_medicion == TipoMedicionComponente.CAMPANA


def _crear_formset(
    request,
    componente,
    fase,
    prefijo,
    fabrica,
    cantidad_predeterminada,
    campo_ubicacion,
):
    queryset = componente.mediciones_campana.filter(fase=fase).order_by(
        "orden", "punto"
    )
    existentes = set(queryset.values_list("punto", flat=True))
    cantidad_normal = componente.mediciones.count()
    cantidad = max(cantidad_normal, cantidad_predeterminada)
    faltantes = [punto for punto in range(1, cantidad + 1) if punto not in existentes]
    FormSet = fabrica(extra=0 if request.method == "POST" else len(faltantes))
    iniciales = []
    for punto in faltantes:
        inicial = {"punto": punto, "orden": punto}
        if campo_ubicacion:
            inicial[campo_ubicacion] = (
                "PUNTOS SENTIDO RADIAL" if campo_ubicacion == "ubicacion" else ""
            )
        iniciales.append(inicial)
    return FormSet(
        request.POST or None,
        instance=componente,
        prefix=f"{prefijo}-{fase.lower()}",
        queryset=queryset,
        initial=iniciales,
    )


def formsets_polea_campana(request, polea):
    prefijo = f"campana-polea-{polea.id}"
    return {
        FaseCampana.INICIO: _crear_formset(
            request,
            polea,
            FaseCampana.INICIO,
            prefijo,
            crear_medicion_polea_campana_formset,
            5,
            "posicion",
        ),
        FaseCampana.FIN: _crear_formset(
            request,
            polea,
            FaseCampana.FIN,
            prefijo,
            crear_medicion_polea_campana_formset,
            5,
            "posicion",
        ),
    }


def formsets_life_shaft_campana(request, shaft):
    prefijo = f"campana-shaft-{shaft.id}"
    return {
        FaseCampana.INICIO: _crear_formset(
            request,
            shaft,
            FaseCampana.INICIO,
            prefijo,
            crear_medicion_life_shaft_campana_formset,
            3,
            "ubicacion",
        ),
        FaseCampana.FIN: _crear_formset(
            request,
            shaft,
            FaseCampana.FIN,
            prefijo,
            crear_medicion_life_shaft_campana_formset,
            3,
            "ubicacion",
        ),
    }


def formsets_campana_validos(componente, formsets):
    if not es_campana(componente):
        return True
    return all(formset.is_valid() for formset in formsets.values())


def guardar_formsets_campana(componente, formsets, campo_relacion):
    if not es_campana(componente):
        return
    for fase, formset in formsets.items():
        for medicion in formset.save(commit=False):
            setattr(medicion, campo_relacion, componente)
            medicion.fase = fase
            medicion.save()
        formset.save_m2m()


def mediciones_campana(componente, fase):
    return list(
        componente.mediciones_campana.filter(fase=fase).order_by("orden", "punto")
    )


def agregar_mediciones_campana_bloque(bloque, componente):
    bloque["es_campana"] = es_campana(componente)
    inicio = mediciones_campana(componente, FaseCampana.INICIO)
    fin = mediciones_campana(componente, FaseCampana.FIN)
    bloque["mediciones_inicio"] = inicio
    bloque["mediciones_fin"] = fin
    bloque["campana_mediciones"] = (("INICIO", inicio), ("FIN", fin))
    bloque["minimo_inicio"] = minimo_mediciones(inicio)
    bloque["minimo_fin"] = minimo_mediciones(fin)
    if bloque["minimo_inicio"] and bloque["minimo_fin"]:
        bloque["variacion_campana"] = abs(
            bloque["minimo_inicio"][0] - bloque["minimo_fin"][0]
        )
    else:
        bloque["variacion_campana"] = None
    return bloque


def minimo_mediciones(mediciones):
    candidatos = []
    for medicion in mediciones:
        for campo in "abcdefg":
            valor = getattr(medicion, campo, None)
            if valor is not None:
                candidatos.append((valor, campo.upper(), medicion.punto))
    return min(candidatos, key=lambda item: item[0]) if candidatos else None


def resultados_por_modalidad(componente):
    if not es_campana(componente):
        return (("NORMAL", minimo_mediciones(componente.mediciones.order_by("orden", "punto"))),)
    return (
        ("INICIO DE CAMPAÑA", minimo_mediciones(componente.mediciones_campana.filter(fase=FaseCampana.INICIO))),
        ("FIN DE CAMPAÑA", minimo_mediciones(componente.mediciones_campana.filter(fase=FaseCampana.FIN))),
    )

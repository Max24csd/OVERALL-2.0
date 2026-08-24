from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render

from inspecciones.forms import LifeShaftInspeccionForm, MedicionLifeShaftFormSet
from inspecciones.models import FotoLifeShaft, Inspeccion, LifeShaftInspeccion, MedicionLifeShaft
from inspecciones.views import (
    _procesar_accion_flujo,
    obtener_permisos_flujo,
    obtener_rol,
    usuario_puede_abrir_inspeccion,
)
from inspecciones.reportes.campaign_utils import (
    agregar_mediciones_campana_bloque,
    es_campana,
    formsets_campana_validos,
    formsets_life_shaft_campana,
    guardar_formsets_campana,
)

from .life_shaft_excel import generar_excel_life_shaft_cvb0004
from .utils import (
    InspeccionCVB0004Form,
    es_cvb0004,
    generar_conclusiones_life_shaft,
    generar_observaciones_life_shaft,
)


PUNTOS_POR_SHAFT = {1: 4, 2: 4, 3: 3, 4: 3}


def crear_estructura_life_shaft_cvb0004(inspeccion):
    for numero, cantidad_puntos in PUNTOS_POR_SHAFT.items():
        shaft, _ = LifeShaftInspeccion.objects.get_or_create(
            inspeccion=inspeccion,
            numero=numero,
            defaults={
                "orden": numero,
                "nombre": f"Life Shaft #{numero:02d}",
                "tag": f"CVB0004-LS{numero:02d}",
                "condicion": Inspeccion.Condicion.NO_MEDIDO,
            },
        )
        for punto in range(1, cantidad_puntos + 1):
            MedicionLifeShaft.objects.get_or_create(
                life_shaft=shaft,
                punto=punto,
                defaults={"orden": punto},
            )


def _bloques_reporte(inspeccion):
    return [
        agregar_mediciones_campana_bloque({
            "life_shaft": shaft,
            "mediciones": list(shaft.mediciones.order_by("orden", "punto")),
            "fotografias": list(shaft.fotografias.order_by("creada_en", "id")),
        }, shaft)
        for shaft in inspeccion.life_shafts.order_by("orden", "numero")[:4]
    ]


@transaction.atomic
def formulario_life_shaft_cvb0004(request, inspeccion):
    if not es_cvb0004(inspeccion):
        return HttpResponseForbidden("Este formulario corresponde unicamente a CVB0004.")
    crear_estructura_life_shaft_cvb0004(inspeccion)
    shafts = list(
        inspeccion.life_shafts.prefetch_related("mediciones", "mediciones_campana", "fotografias")
        .order_by("orden", "numero")[:4]
    )
    permisos = obtener_permisos_flujo(request.user, inspeccion)
    datos = request.POST.copy() if request.method == "POST" else None
    inicial = {}
    if request.method == "GET":
        if not inspeccion.observaciones:
            inicial["observaciones"] = generar_observaciones_life_shaft(shafts)
        if not inspeccion.recomendaciones:
            inicial["recomendaciones"] = generar_conclusiones_life_shaft(shafts)
    if datos is not None and not datos.get("inspeccion-condicion_equipo"):
        datos["inspeccion-condicion_equipo"] = inspeccion.condicion_equipo or "En operacion"
    formulario = InspeccionCVB0004Form(
        datos, instance=inspeccion, prefix="inspeccion", initial=inicial
    )
    bloques = []
    todo_valido = formulario.is_valid() if request.method == "POST" else False
    for shaft in shafts:
        formulario_shaft = LifeShaftInspeccionForm(
            request.POST or None, instance=shaft, prefix=f"shaft-{shaft.id}"
        )
        mediciones = MedicionLifeShaftFormSet(
            request.POST or None, instance=shaft, prefix=f"mediciones-{shaft.id}"
        )
        campana = formsets_life_shaft_campana(request, shaft)
        if request.method == "POST":
            shaft_valido = formulario_shaft.is_valid()
            normales_validas = mediciones.is_valid() if shaft_valido and not es_campana(shaft) else True
            campana_valida = formsets_campana_validos(shaft, campana) if shaft_valido else False
            todo_valido = todo_valido and shaft_valido and normales_validas and campana_valida
        bloques.append(
            {
                "life_shaft": shaft,
                "formulario_shaft": formulario_shaft,
                "mediciones_formset": mediciones,
                "campana_inicio_formset": campana["INICIO"],
                "campana_fin_formset": campana["FIN"],
                "campana_formsets": (("INICIO", campana["INICIO"]), ("FIN", campana["FIN"])),
                "es_campana": es_campana(shaft),
                "fotografias": list(shaft.fotografias.order_by("creada_en", "id")),
            }
        )
    if request.method == "POST":
        if not permisos["puede_editar"]:
            return HttpResponseForbidden("La inspeccion no esta habilitada para edicion.")
        if todo_valido:
            guardada = formulario.save()
            for bloque in bloques:
                shaft = bloque["formulario_shaft"].save(commit=False)
                shaft.inspeccion = guardada
                shaft.save()
                if es_campana(shaft):
                    guardar_formsets_campana(
                        shaft,
                        {"INICIO": bloque["campana_inicio_formset"], "FIN": bloque["campana_fin_formset"]},
                        "life_shaft",
                    )
                else:
                    formset = bloque["mediciones_formset"]
                    for medicion in formset.save(commit=False):
                        medicion.life_shaft = shaft
                        medicion.save()
                for indice in range(1, 11):
                    archivo = request.FILES.get(f"foto-{shaft.id}-{indice}")
                    if archivo:
                        FotoLifeShaft.objects.create(
                            life_shaft=shaft,
                            imagen=archivo,
                            codigo_dano="",
                            descripcion=request.POST.get(
                                f"descripcion-foto-{shaft.id}-{indice}", ""
                            ).strip(),
                            subida_por=request.user,
                        )
            shafts_actualizados = list(guardada.life_shafts.order_by("orden", "numero")[:4])
            campos = []
            if not guardada.observaciones.strip():
                guardada.observaciones = generar_observaciones_life_shaft(shafts_actualizados)
                campos.append("observaciones")
            if not guardada.recomendaciones.strip():
                guardada.recomendaciones = generar_conclusiones_life_shaft(shafts_actualizados)
                campos.append("recomendaciones")
            if campos:
                guardada.save(update_fields=campos)
            correcto, mensaje = _procesar_accion_flujo(request, guardada)
            (messages.success if correcto else messages.error)(request, mensaje)
            return redirect("formulario_life_shaft", inspeccion_id=guardada.id)
        messages.error(request, "No se pudo guardar Life Shaft CVB0004. Revisa los campos marcados.")
    return render(
        request,
        "inspecciones/formulario_life_shaft_cvb0004.html",
        {
            "inspeccion": inspeccion,
            "formulario": formulario,
            "bloques_shafts": bloques,
            "rangos_fotos": range(1, 6),
            **permisos,
        },
    )


def reporte_life_shaft_cvb0004(request, inspeccion):
    bloques = _bloques_reporte(inspeccion)
    shafts = [bloque["life_shaft"] for bloque in bloques]
    if not inspeccion.observaciones:
        inspeccion.observaciones = generar_observaciones_life_shaft(shafts)
    if not inspeccion.recomendaciones:
        inspeccion.recomendaciones = generar_conclusiones_life_shaft(shafts)
    return render(
        request,
        "inspecciones/reporte_life_shaft_cvb0004.html",
        {"inspeccion": inspeccion, "bloques_shafts": bloques, "rol": obtener_rol(request.user)},
    )


@login_required
def exportar_excel_life_shaft_cvb0004(request, inspeccion_id):
    inspeccion = get_object_or_404(
        Inspeccion.objects.select_related("faja", "inspector", "supervisor", "analista", "cliente")
        .prefetch_related("life_shafts__mediciones", "life_shafts__mediciones_campana", "life_shafts__fotografias"),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.LIFE_SHAFT,
    )
    if not usuario_puede_abrir_inspeccion(request.user, inspeccion):
        return HttpResponseForbidden("No tienes permiso para descargar este reporte.")
    if not es_cvb0004(inspeccion):
        return HttpResponseForbidden("Esta exportacion corresponde unicamente a CVB0004.")
    salida = generar_excel_life_shaft_cvb0004(inspeccion, _bloques_reporte(inspeccion))
    respuesta = HttpResponse(
        salida.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    respuesta["Content-Disposition"] = 'attachment; filename="REPORTE_INSPECCION_CVB0004_LIFE_SHAFT.xlsx"'
    return respuesta

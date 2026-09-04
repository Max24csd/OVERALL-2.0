from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render

from inspecciones.forms import FotoPoleaFormSet, MedicionPoleaFormSet, PoleaInspeccionForm
from inspecciones.models import Inspeccion, MedicionPolea, PoleaInspeccion
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
    formsets_polea_campana,
    guardar_formsets_campana,
)
from inspecciones.reportes.cvb0003.history import (
    historial_componente_visible,
    preparar_formset_historico,
    validar_formset_historico,
)

from .poleas_excel import generar_excel_poleas_cvb0004
from .utils import (
    InspeccionCVB0004Form,
    es_cvb0004,
    generar_conclusiones_poleas,
    generar_observaciones_poleas,
)


POLEAS_CERAMICAS = {4, 5}


def crear_estructura_poleas_cvb0004(inspeccion):
    for numero in range(1, 10):
        polea, _ = PoleaInspeccion.objects.get_or_create(
            inspeccion=inspeccion,
            numero=numero,
            defaults={
                "orden": numero,
                "nombre": f"Polea #{numero:02d}",
                "tag": f"CVB0004-P{numero:02d}",
                "condicion": Inspeccion.Condicion.NO_MEDIDO,
            },
        )
        if numero not in POLEAS_CERAMICAS:
            for punto in range(1, 6):
                MedicionPolea.objects.get_or_create(
                    polea=polea, punto=punto, defaults={"orden": punto}
                )


def _bloques_reporte(inspeccion):
    return [
        agregar_mediciones_campana_bloque({
            "polea": polea,
            "mediciones": list(polea.mediciones.order_by("orden", "punto")),
            "fotografias": list(polea.fotografias.order_by("creada_en", "id")),
        }, polea)
        for polea in inspeccion.poleas_inspeccionadas.order_by("orden", "numero")[:9]
    ]


@transaction.atomic
def formulario_poleas_cvb0004(request, inspeccion):
    if not es_cvb0004(inspeccion):
        return HttpResponseForbidden("Este formulario corresponde unicamente a CVB0004.")
    crear_estructura_poleas_cvb0004(inspeccion)
    poleas = list(
        inspeccion.poleas_inspeccionadas.prefetch_related("mediciones", "mediciones_campana", "fotografias")
        .order_by("orden", "numero")[:9]
    )
    permisos = obtener_permisos_flujo(request.user, inspeccion)
    datos = request.POST.copy() if request.method == "POST" else None
    inicial = {}
    if request.method == "GET":
        if not inspeccion.observaciones:
            inicial["observaciones"] = generar_observaciones_poleas(poleas)
        if not inspeccion.recomendaciones:
            inicial["recomendaciones"] = generar_conclusiones_poleas(poleas)
    if datos is not None and not datos.get("inspeccion-condicion_equipo"):
        datos["inspeccion-condicion_equipo"] = inspeccion.condicion_equipo or "En operacion"
    formulario = InspeccionCVB0004Form(datos, instance=inspeccion, prefix="inspeccion", initial=inicial)
    bloques = []
    todo_valido = formulario.is_valid() if request.method == "POST" else False
    for polea in poleas:
        form_polea = PoleaInspeccionForm(request.POST or None, instance=polea, prefix=f"polea-{polea.id}")
        mediciones = MedicionPoleaFormSet(request.POST or None, instance=polea, prefix=f"mediciones-{polea.id}")
        campana = formsets_polea_campana(request, polea)
        historial = historial_componente_visible(
            inspeccion, polea, inspeccion.fecha_inspeccion,
            "poleas_inspeccionadas",
        )
        valores_historicos = historial["valores"] if historial else {}
        preparar_formset_historico(mediciones, valores_historicos)
        preparar_formset_historico(campana["FIN"], valores_historicos)
        fotos = FotoPoleaFormSet(request.POST or None, request.FILES or None, instance=polea, prefix=f"fotografias-{polea.id}")
        if request.method == "POST":
            polea_valida = form_polea.is_valid()
            fotos_validas = fotos.is_valid()
            normales_validas = validar_formset_historico(mediciones, valores_historicos) if polea_valida and not es_campana(polea) and polea.numero not in POLEAS_CERAMICAS else True
            campana_valida = (formsets_campana_validos(polea, campana) and validar_formset_historico(campana["FIN"], valores_historicos)) if polea_valida else False
            todo_valido = todo_valido and polea_valida and fotos_validas and normales_validas and campana_valida
        bloques.append({"polea": polea, "formulario_polea": form_polea, "mediciones_formset": mediciones, "campana_inicio_formset": campana["INICIO"], "campana_fin_formset": campana["FIN"], "campana_formsets": (("INICIO", campana["INICIO"]), ("FIN", campana["FIN"])), "es_campana": es_campana(polea), "fotos_formset": fotos, "fotografias": list(polea.fotografias.order_by("creada_en", "id"))})
    if request.method == "POST":
        if not permisos["puede_editar"]:
            return HttpResponseForbidden("La inspeccion no esta habilitada para edicion.")
        if todo_valido:
            guardada = formulario.save()
            for bloque in bloques:
                polea = bloque["formulario_polea"].save(commit=False)
                polea.inspeccion = guardada
                polea.save()
                if es_campana(polea):
                    guardar_formsets_campana(polea, {"INICIO": bloque["campana_inicio_formset"], "FIN": bloque["campana_fin_formset"]}, "polea")
                elif polea.numero not in POLEAS_CERAMICAS:
                    for medicion in bloque["mediciones_formset"].save(commit=False):
                        medicion.polea = polea
                        medicion.save()
                formset_fotos = bloque["fotos_formset"]
                for foto in formset_fotos.save(commit=False):
                    foto.polea = polea
                    foto.codigo_dano = ""
                    if not foto.subida_por_id:
                        foto.subida_por = request.user
                    foto.save()
            poleas_actualizadas = list(guardada.poleas_inspeccionadas.order_by("orden", "numero")[:9])
            campos = []
            if not guardada.observaciones.strip():
                guardada.observaciones = generar_observaciones_poleas(poleas_actualizadas)
                campos.append("observaciones")
            if not guardada.recomendaciones.strip():
                guardada.recomendaciones = generar_conclusiones_poleas(poleas_actualizadas)
                campos.append("recomendaciones")
            if campos:
                guardada.save(update_fields=campos)
            correcto, mensaje = _procesar_accion_flujo(request, guardada)
            (messages.success if correcto else messages.error)(request, mensaje)
            return redirect("formulario_poleas", inspeccion_id=guardada.id)
        messages.error(request, "No se pudo guardar Poleas CVB0004. Revisa los campos marcados.")
    return render(request, "inspecciones/formulario_poleas_cvb0004.html", {"inspeccion": inspeccion, "formulario": formulario, "bloques_poleas": bloques, "numeros_poleas_ceramicas": POLEAS_CERAMICAS, **permisos})


def reporte_poleas_cvb0004(request, inspeccion):
    bloques = _bloques_reporte(inspeccion)
    poleas = [bloque["polea"] for bloque in bloques]
    if not inspeccion.observaciones:
        inspeccion.observaciones = generar_observaciones_poleas(poleas)
    if not inspeccion.recomendaciones:
        inspeccion.recomendaciones = generar_conclusiones_poleas(poleas)
    return render(request, "inspecciones/reporte_poleas_cvb0004.html", {"inspeccion": inspeccion, "bloques_poleas": bloques, "rol": obtener_rol(request.user)})

@login_required
def exportar_excel_poleas_cvb0004(request, inspeccion_id):

    print(
        ">>>>>>>> ENTRO A EXPORTAR CVB004 POLEAS <<<<<<<<",
        flush=True,
    )

    inspeccion = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja",
            "inspector",
            "supervisor",
            "analista",
        ),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.POLEAS,
    )

    if not usuario_puede_abrir_inspeccion(
        request.user,
        inspeccion,
    ):
        return HttpResponseForbidden(
            "No tienes permiso para descargar este reporte."
        )

    bloques = _bloques_reporte(inspeccion)

    print(
        ">>> POLEAS:",
        [
            (
                bloque["polea"].numero,
                bloque.get("es_campana", False),
            )
            for bloque in bloques
        ],
        flush=True,
    )

    salida = generar_excel_poleas_cvb0004(
        inspeccion,
        bloques,
    )

    print(
        ">>> EXCEL CVB004 GENERADO",
        flush=True,
    )

    response = HttpResponse(
        salida.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    response["Content-Disposition"] = (
        'attachment; filename="PRUEBA_CVB004_NUEVO.xlsx"'
    )

    response["Cache-Control"] = (
        "no-store, no-cache, must-revalidate, max-age=0"
    )
    response["Pragma"] = "no-cache"
    response["Expires"] = "0"

    print(
        ">>>>>>>> ENVIANDO CVB004 POLEAS <<<<<<<<",
        flush=True,
    )

    return response

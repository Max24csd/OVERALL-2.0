from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django import forms
from django.db import transaction
from django.forms import inlineformset_factory
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render

from inspecciones.forms import (
    InspeccionForm,
    LifeShaftInspeccionForm,
    MedicionLifeShaftFormSet,
)
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
from inspecciones.reportes.cvb0003.history import (
    historial_componente_visible,
    preparar_formset_historico,
    validar_formset_historico,
)

from .life_shaft_excel import generar_excel_life_shaft_cvb0001
from .text_utils import (
    conclusiones_life_shaft,
    es_cvb0001,
    observaciones_life_shaft,
)


class FotoLifeShaftCVB0001Form(forms.ModelForm):
    class Meta:
        model = FotoLifeShaft
        fields = ("imagen", "codigo_dano", "descripcion")
        widgets = {
            "imagen": forms.FileInput(attrs={"accept": "image/*"}),
            "codigo_dano": forms.TextInput(attrs={"placeholder": "Título opcional"}),
            "descripcion": forms.Textarea(attrs={"rows": 2, "placeholder": "Observación opcional"}),
        }


FotoLifeShaftCVB0001FormSet = inlineformset_factory(
    LifeShaftInspeccion,
    FotoLifeShaft,
    form=FotoLifeShaftCVB0001Form,
    extra=5,
    can_delete=False,
)


def crear_estructura_life_shaft_cvb0001(inspeccion):
    for numero in range(1, 3):
        shaft, _ = LifeShaftInspeccion.objects.get_or_create(
            inspeccion=inspeccion,
            numero=numero,
            defaults={
                "orden": numero,
                "nombre": f"Life Shaft #{numero:02d}",
                "tag": f"CVB0001-LS{numero:02d}",
                "condicion": Inspeccion.Condicion.NO_MEDIDO,
            },
        )
        for punto in range(1, 4):
            MedicionLifeShaft.objects.get_or_create(
                life_shaft=shaft, punto=punto, defaults={"orden": punto}
            )


def _bloques(inspeccion):
    return [
        agregar_mediciones_campana_bloque({
            "life_shaft": shaft,
            "mediciones": list(shaft.mediciones.order_by("orden", "punto")),
            "fotografias": list(shaft.fotografias.order_by("creada_en", "id")),
        }, shaft)
        for shaft in inspeccion.life_shafts.order_by("orden", "numero")[:2]
    ]


@transaction.atomic
def formulario_life_shaft_cvb0001(request, inspeccion):
    if not es_cvb0001(inspeccion):
        return HttpResponseForbidden("Este formulario corresponde únicamente a CVB001.")
    crear_estructura_life_shaft_cvb0001(inspeccion)
    shafts = list(
        inspeccion.life_shafts.prefetch_related("mediciones", "mediciones_campana", "fotografias")
        .order_by("orden", "numero")[:2]
    )
    permisos = obtener_permisos_flujo(request.user, inspeccion)
    datos = request.POST.copy() if request.method == "POST" else None
    if datos is not None and not datos.get("inspeccion-condicion_equipo"):
        datos["inspeccion-condicion_equipo"] = inspeccion.condicion_equipo or "En operación"
    formulario = InspeccionForm(datos, instance=inspeccion, prefix="inspeccion")
    bloques = []
    todo_valido = formulario.is_valid() if request.method == "POST" else False
    for shaft in shafts:
        form_shaft = LifeShaftInspeccionForm(
            request.POST or None, instance=shaft, prefix=f"shaft-{shaft.id}"
        )
        mediciones = MedicionLifeShaftFormSet(
            request.POST or None, instance=shaft, prefix=f"mediciones-{shaft.id}"
        )
        campana = formsets_life_shaft_campana(request, shaft)
        historial=historial_componente_visible(inspeccion,shaft,inspeccion.fecha_inspeccion,"life_shafts")
        valores_historicos=historial["valores"] if historial else {}
        preparar_formset_historico(mediciones,valores_historicos)
        preparar_formset_historico(campana["FIN"],valores_historicos)
        fotos = FotoLifeShaftCVB0001FormSet(
            request.POST or None,
            request.FILES or None,
            instance=shaft,
            prefix=f"fotografias-{shaft.id}",
        )
        if request.method == "POST":
            shaft_valido = form_shaft.is_valid()
            mediciones_validas = validar_formset_historico(mediciones,valores_historicos) if shaft_valido and not es_campana(shaft) else True
            campana_valida = (formsets_campana_validos(shaft,campana) and validar_formset_historico(campana["FIN"],valores_historicos)) if shaft_valido else False
            fotos_validas = fotos.is_valid()
            todo_valido = todo_valido and shaft_valido and mediciones_validas and campana_valida and fotos_validas
        bloques.append(
            {
                "life_shaft": shaft,
                "formulario_shaft": form_shaft,
                "mediciones_formset": mediciones,
                "campana_inicio_formset": campana["INICIO"],
                "campana_fin_formset": campana["FIN"],
                "campana_formsets": (("INICIO", campana["INICIO"]), ("FIN", campana["FIN"])),
                "es_campana": es_campana(shaft),
                "fotos_formset": fotos,
                "fotografias": list(shaft.fotografias.order_by("creada_en", "id")),
            }
        )
    if request.method == "POST":
        if not permisos["puede_editar"]:
            return HttpResponseForbidden("La inspección no está habilitada para edición.")
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
                    for medicion in bloque["mediciones_formset"].save(commit=False):
                        medicion.life_shaft = shaft
                        medicion.save()
                for foto in bloque["fotos_formset"].save(commit=False):
                    foto.life_shaft = shaft
                    if not foto.subida_por_id:
                        foto.subida_por = request.user
                    foto.save()
            correcto, mensaje = _procesar_accion_flujo(request, guardada)
            (messages.success if correcto else messages.error)(request, mensaje)
            return redirect("formulario_life_shaft", inspeccion_id=guardada.id)
        messages.error(request, "No se pudo guardar Life Shaft CVB001. Revisa los campos marcados.")
    return render(
        request,
        "inspecciones/formulario_life_shaft_cvb0001.html",
        {"inspeccion": inspeccion, "formulario": formulario, "bloques_shafts": bloques, **permisos},
    )


def reporte_life_shaft_cvb0001(request, inspeccion):
    bloques = _bloques(inspeccion)
    shafts = [b["life_shaft"] for b in bloques]
    inspeccion.observaciones = observaciones_life_shaft(shafts, inspeccion.observaciones)
    inspeccion.recomendaciones = conclusiones_life_shaft(shafts, inspeccion.recomendaciones)
    return render(
        request,
        "inspecciones/reporte_life_shaft_cvb0001.html",
        {"inspeccion": inspeccion, "bloques_shafts": bloques, "rol": obtener_rol(request.user)},
    )


@login_required
def exportar_excel_life_shaft_cvb0001(request, inspeccion_id):
    inspeccion = get_object_or_404(
        Inspeccion.objects.select_related("faja", "inspector", "supervisor", "analista", "cliente")
        .prefetch_related("life_shafts__mediciones", "life_shafts__mediciones_campana", "life_shafts__fotografias"),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.LIFE_SHAFT,
    )
    if not usuario_puede_abrir_inspeccion(request.user, inspeccion):
        return HttpResponseForbidden("No tienes permiso para descargar este reporte.")
    if not es_cvb0001(inspeccion):
        return HttpResponseForbidden("Esta exportación corresponde únicamente a CVB001.")
    salida = generar_excel_life_shaft_cvb0001(inspeccion, _bloques(inspeccion))
    respuesta = HttpResponse(salida.getvalue(), content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    respuesta["Content-Disposition"] = 'attachment; filename="REPORTE_INSPECCION_CVB0001_LIFE_SHAFT.xlsx"'
    return respuesta

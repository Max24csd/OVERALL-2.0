from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django import forms
from django.db import transaction
from django.forms import inlineformset_factory
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render

from inspecciones.forms import InspeccionForm, MedicionPoleaFormSet, PoleaInspeccionForm
from inspecciones.models import FotoPolea, Inspeccion, MedicionPolea, PoleaInspeccion
from inspecciones.views import _procesar_accion_flujo, obtener_permisos_flujo, obtener_rol, usuario_puede_abrir_inspeccion
from inspecciones.reportes.campaign_utils import (
    agregar_mediciones_campana_bloque,
    es_campana,
    formsets_campana_validos,
    formsets_polea_campana,
    guardar_formsets_campana,
)

from .poleas_excel import generar_excel_poleas_cvb0001
from .poleas_text_utils import conclusiones_poleas, observaciones_poleas
from .text_utils import es_cvb0001


POLEAS_CERAMICAS = {1}


class FotoPoleaCVB0001Form(forms.ModelForm):
    class Meta:
        model = FotoPolea
        fields = ("imagen", "descripcion")
        widgets = {
            "imagen": forms.FileInput(attrs={"accept": "image/*"}),
            "descripcion": forms.Textarea(attrs={"rows": 2, "placeholder": "Observación opcional"}),
        }


FotoPoleaCVB0001FormSet = inlineformset_factory(
    PoleaInspeccion,
    FotoPolea,
    form=FotoPoleaCVB0001Form,
    extra=5,
    can_delete=False,
)


def crear_estructura_poleas_cvb0001(inspeccion):
    for numero in range(1, 6):
        polea, _ = PoleaInspeccion.objects.get_or_create(
            inspeccion=inspeccion,
            numero=numero,
            defaults={"orden": numero, "nombre": f"Polea #{numero:02d}", "tag": f"CVB0001-P{numero:02d}", "condicion": Inspeccion.Condicion.NO_MEDIDO},
        )
        if numero not in POLEAS_CERAMICAS:
            for punto in range(1, 6):
                MedicionPolea.objects.get_or_create(polea=polea, punto=punto, defaults={"orden": punto})


def _bloques(inspeccion):
    return [agregar_mediciones_campana_bloque({"polea": polea, "mediciones": list(polea.mediciones.order_by("orden", "punto")), "fotografias": list(polea.fotografias.order_by("creada_en", "id"))}, polea) for polea in inspeccion.poleas_inspeccionadas.order_by("orden", "numero")[:5]]


@transaction.atomic
def formulario_poleas_cvb0001(request, inspeccion):
    if not es_cvb0001(inspeccion): return HttpResponseForbidden("Este formulario corresponde únicamente a CVB001.")
    crear_estructura_poleas_cvb0001(inspeccion)
    poleas=list(inspeccion.poleas_inspeccionadas.prefetch_related("mediciones","mediciones_campana","fotografias").order_by("orden","numero")[:5])
    permisos=obtener_permisos_flujo(request.user,inspeccion);datos=request.POST.copy() if request.method=="POST" else None
    if datos is not None and not datos.get("inspeccion-condicion_equipo"):datos["inspeccion-condicion_equipo"]=inspeccion.condicion_equipo or "En operación"
    formulario=InspeccionForm(datos,instance=inspeccion,prefix="inspeccion");bloques=[];todo_valido=formulario.is_valid() if request.method=="POST" else False
    for polea in poleas:
        form_polea=PoleaInspeccionForm(request.POST or None,instance=polea,prefix=f"polea-{polea.id}")
        mediciones=MedicionPoleaFormSet(request.POST or None,instance=polea,prefix=f"mediciones-{polea.id}")
        campana=formsets_polea_campana(request,polea)
        fotos=FotoPoleaCVB0001FormSet(request.POST or None,request.FILES or None,instance=polea,prefix=f"fotografias-{polea.id}")
        if request.method=="POST":
            polea_valida=form_polea.is_valid();fotos_validas=fotos.is_valid()
            normales_validas=mediciones.is_valid() if polea_valida and not es_campana(polea) and polea.numero not in POLEAS_CERAMICAS else True
            campana_valida=formsets_campana_validos(polea,campana) if polea_valida else False
            todo_valido=todo_valido and polea_valida and fotos_validas and normales_validas and campana_valida
        bloques.append({"polea":polea,"formulario_polea":form_polea,"mediciones_formset":mediciones,"campana_inicio_formset":campana["INICIO"],"campana_fin_formset":campana["FIN"],"campana_formsets":(("INICIO",campana["INICIO"]),("FIN",campana["FIN"])),"es_campana":es_campana(polea),"fotos_formset":fotos,"fotografias":list(polea.fotografias.order_by("creada_en","id"))})
    if request.method=="POST":
        if not permisos["puede_editar"]:return HttpResponseForbidden("La inspección no está habilitada para edición.")
        if todo_valido:
            guardada=formulario.save()
            for bloque in bloques:
                polea=bloque["formulario_polea"].save(commit=False);polea.inspeccion=guardada;polea.save()
                if es_campana(polea):
                    guardar_formsets_campana(polea,{"INICIO":bloque["campana_inicio_formset"],"FIN":bloque["campana_fin_formset"]},"polea")
                elif polea.numero not in POLEAS_CERAMICAS:
                    for medicion in bloque["mediciones_formset"].save(commit=False):medicion.polea=polea;medicion.save()
                for foto in bloque["fotos_formset"].save(commit=False):
                    foto.polea=polea;foto.codigo_dano=""
                    if not foto.subida_por_id:foto.subida_por=request.user
                    foto.save()
            correcto,mensaje=_procesar_accion_flujo(request,guardada);(messages.success if correcto else messages.error)(request,mensaje)
            return redirect("formulario_poleas",inspeccion_id=guardada.id)
        messages.error(request,"No se pudo guardar Poleas CVB001. Revisa los campos marcados.")
    return render(request,"inspecciones/formulario_poleas_cvb0001.html",{"inspeccion":inspeccion,"formulario":formulario,"bloques_poleas":bloques,"numeros_poleas_ceramicas":POLEAS_CERAMICAS,**permisos})


def reporte_poleas_cvb0001(request,inspeccion):
    bloques=_bloques(inspeccion);poleas=[b["polea"] for b in bloques];inspeccion.observaciones=observaciones_poleas(poleas,inspeccion.observaciones);inspeccion.recomendaciones=conclusiones_poleas(poleas,inspeccion.recomendaciones)
    return render(request,"inspecciones/reporte_poleas_cvb0001.html",{"inspeccion":inspeccion,"bloques_poleas":bloques,"rol":obtener_rol(request.user)})


@login_required
def exportar_excel_poleas_cvb0001(request,inspeccion_id):
    inspeccion=get_object_or_404(Inspeccion.objects.select_related("faja","inspector","supervisor","analista","cliente").prefetch_related("poleas_inspeccionadas__mediciones","poleas_inspeccionadas__mediciones_campana","poleas_inspeccionadas__fotografias"),id=inspeccion_id,tipo=Inspeccion.Tipo.POLEAS)
    if not usuario_puede_abrir_inspeccion(request.user,inspeccion):return HttpResponseForbidden("No tienes permiso para descargar este reporte.")
    if not es_cvb0001(inspeccion):return HttpResponseForbidden("Esta exportación corresponde únicamente a CVB001.")
    salida=generar_excel_poleas_cvb0001(inspeccion,_bloques(inspeccion));respuesta=HttpResponse(salida.getvalue(),content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");respuesta["Content-Disposition"]='attachment; filename="REPORTE_INSPECCION_CVB0001_POLEAS.xlsx"';return respuesta

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404

from inspecciones.models import Inspeccion
from inspecciones.reportes.campaign_utils import agregar_mediciones_campana_bloque

from .exporters.life_shaft import generar_excel_life_shaft_cvb0003_master
from .permissions import puede_acceder_inspeccion_cvb0003


TAGS_CVB0003 = {
    "CVB0003",
    "0220-CVB-0003",
    "0220-CVB0003",
}


@login_required
def exportar_excel_life_shaft_cvb0003(request, inspeccion_id):
    inspection = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja",
            "inspector",
            "supervisor",
            "analista",
            "cliente",
        ).prefetch_related(
            "life_shafts__mediciones",
            "life_shafts__mediciones_campana",
            "life_shafts__fotografias",
        ),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.LIFE_SHAFT,
    )
    if not puede_acceder_inspeccion_cvb0003(request.user, inspection, "ver"):
        return HttpResponseForbidden(
            "No tienes permiso para descargar este reporte."
        )
    tag = (inspection.faja.tag or "").upper().strip()
    if tag not in TAGS_CVB0003:
        return HttpResponseForbidden(
            "Esta exportación sólo corresponde al Life Shaft CVB0003."
        )

    blocks = [
        agregar_mediciones_campana_bloque({
            "life_shaft": shaft,
            "mediciones": list(shaft.mediciones.order_by("orden", "punto")),
            "fotografias": list(shaft.fotografias.order_by("id")),
        }, shaft)
        for shaft in inspection.life_shafts.order_by("orden", "numero")
    ]
    output = generar_excel_life_shaft_cvb0003_master(inspection, blocks)
    response = HttpResponse(
        output.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
    response["Content-Disposition"] = (
        'attachment; filename="REPORTE_INSPECCION_CVB0003_LIFE_SHAFT.xlsx"'
    )
    return response

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, render

from inspecciones.models import Inspeccion
from inspecciones.reportes.campaign_utils import agregar_mediciones_campana_bloque

from .exporters.poleas import generar_excel_poleas_cvb0003_master
from .permissions import obtener_rol_cvb0003, puede_acceder_inspeccion_cvb0003


TAGS_CVB0003 = {
    "CVB0003",
    "CVB003",
    "0220-CVB-0003",
    "0220-CVB0003",
    "0220-CVB-003",
}


def _role(user):
    return obtener_rol_cvb0003(user)


def _blocks(inspection):
    return [
        agregar_mediciones_campana_bloque({
            "polea": polea,
            "mediciones": list(polea.mediciones.order_by("orden", "punto")),
            "fotografias": list(polea.fotografias.order_by("creada_en", "id")),
        }, polea)
        for polea in inspection.poleas_inspeccionadas.order_by("orden", "numero")
    ]


def reporte_poleas_cvb0003(request, inspection):
    return render(
        request,
        "inspecciones/reporte_poleas_cvb0003.html",
        {
            "inspeccion": inspection,
            "bloques_poleas": _blocks(inspection),
            "rol": _role(request.user),
        },
    )


@login_required
def exportar_excel_poleas_cvb0003(request, inspeccion_id):
    inspection = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja",
            "inspector",
            "supervisor",
            "analista",
            "cliente",
        ).prefetch_related(
            "poleas_inspeccionadas__mediciones",
            "poleas_inspeccionadas__mediciones_campana",
            "poleas_inspeccionadas__fotografias",
        ),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.POLEAS,
    )
    if not puede_acceder_inspeccion_cvb0003(request.user, inspection, "ver"):
        return HttpResponseForbidden(
            "No tienes permiso para descargar este reporte."
        )
    tag = (inspection.faja.tag or "").upper().strip()
    if tag not in TAGS_CVB0003:
        return HttpResponseForbidden(
            "Esta exportación sólo corresponde al reporte de Poleas CVB0003."
        )

    output = generar_excel_poleas_cvb0003_master(
        inspection, _blocks(inspection)
    )
    response = HttpResponse(
        output.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
    response["Content-Disposition"] = (
        'attachment; filename="REPORTE_INSPECCION_CVB0003_POLEAS.xlsx"'
    )
    return response

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404

from inspecciones.models import CalibracionUTFajaCVB0003, Inspeccion

from .exporters.faja import generar_excel_faja_cvb0003_master
from .permissions import puede_acceder_inspeccion_cvb0003


TAGS_CVB0003 = {
    "CVB0003", "CVB003", "0220-CVB-0003", "0220-CVB0003", "0220-CVB-003",
}


@login_required
def exportar_excel_faja_cvb0003(request, inspeccion_id):
    inspection = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja", "inspector", "supervisor", "analista", "cliente",
        ).prefetch_related(
            "empalmes_cvb0003", "tramos_cvb0003", "fotografias_cvb0003",
            "calibraciones_ut_faja_cvb0003",
        ),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.FAJA,
    )
    if not puede_acceder_inspeccion_cvb0003(request.user, inspection, "ver"):
        return HttpResponseForbidden("No tienes permiso para descargar este reporte.")
    if (inspection.faja.tag or "").upper().strip() not in TAGS_CVB0003:
        return HttpResponseForbidden(
            "Esta exportación sólo corresponde al reporte Faja CVB0003."
        )
    CalibracionUTFajaCVB0003.crear_estructura(inspection)
    output = generar_excel_faja_cvb0003_master(inspection)
    response = HttpResponse(
        output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = (
        'attachment; filename="REPORTE_INSPECCION_CVB0003_FAJA.xlsx"'
    )
    return response

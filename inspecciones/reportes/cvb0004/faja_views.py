from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404

from inspecciones.models import Inspeccion

from .faja_excel import generar_excel_faja_cvb0004
from .utils import es_cvb0004


@login_required
def exportar_excel_faja_cvb0004(request, inspeccion_id):
    inspection = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja", "inspector", "supervisor", "analista", "cliente"
        ).prefetch_related("empalmes_cvb0003", "tramos_cvb0003", "fotografias_cvb0003"),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.FAJA,
    )
    if not es_cvb0004(inspection):
        return HttpResponseForbidden("Esta exportación corresponde únicamente a Faja CVB004.")

    output = generar_excel_faja_cvb0004(inspection)
    response = HttpResponse(
        output.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="REPORTE_INSPECCION_CVB0004_FAJA.xlsx"'
    return response

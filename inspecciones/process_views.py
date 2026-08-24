from __future__ import annotations

from django.contrib.auth.decorators import login_required
from django.db.models import Q
from django.shortcuts import render

from accounts.models import AccesoParada
from filtros_mm.models import ReporteFiltro

from .models import Inspeccion, Parada
from .presentation_scope import (
    INSPECCIONES_HISTORICAS_OFICIALES_IDS,
    PARADA_ACTUAL_ID,
)
from .views import (
    _agregar_contadores_dashboard,
    _codigo_equipo_dashboard,
    _ordenar_inspecciones_dashboard,
    obtener_rol,
    usuario_puede_abrir_inspeccion,
)


def _agrupar_chancado_por_equipo(inspecciones):
    resultado = []

    for codigo in ("CVB001", "CVB003", "CVB004"):
        items = [
            inspeccion
            for inspeccion in inspecciones
            if _codigo_equipo_dashboard(inspeccion) == codigo
        ]

        if items:
            resultado.append(
                {
                    "codigo": codigo,
                    "inspecciones": _ordenar_inspecciones_dashboard(items),
                }
            )

    return resultado


def _filtrar_chancado_para_usuario(usuario, queryset):
    rol = obtener_rol(usuario)

    if rol == "Administrador":
        return list(queryset)

    return [
        inspeccion
        for inspeccion in queryset
        if usuario_puede_abrir_inspeccion(usuario, inspeccion)
    ]


@login_required
def proceso_chancado(request):
    """
    Vista de presentación del proceso Chancado.

    No modifica estados, mediciones, fotos, Excel ni restricciones técnicas.
    Solo organiza reportes actuales e históricos que el usuario puede abrir.
    """
    rol = obtener_rol(request.user)

    actuales_qs = (
        Inspeccion.objects
        .select_related(
            "parada",
            "faja",
            "inspector",
            "supervisor",
            "analista",
            "cliente",
        )
        .prefetch_related("historial")
        .filter(parada_id=PARADA_ACTUAL_ID)
        .order_by("faja__tag", "tipo", "id")
    )

    actuales = _filtrar_chancado_para_usuario(
        request.user,
        actuales_qs,
    )

    for inspeccion in actuales:
        _agregar_contadores_dashboard(inspeccion)

    parada_actual = (
        actuales[0].parada
        if actuales and actuales[0].parada_id
        else Parada.objects.filter(id=PARADA_ACTUAL_ID).first()
    )

    historicos_qs = (
        Inspeccion.objects
        .select_related(
            "parada",
            "faja",
            "inspector",
            "supervisor",
            "analista",
            "cliente",
        )
        .prefetch_related("historial")
        .filter(id__in=INSPECCIONES_HISTORICAS_OFICIALES_IDS)
        .order_by("faja__tag", "tipo", "id")
    )

    historicos = _filtrar_chancado_para_usuario(
        request.user,
        historicos_qs,
    )

    for inspeccion in historicos:
        _agregar_contadores_dashboard(inspeccion)

    return render(
        request,
        "procesos/chancado.html",
        {
            "rol": rol,
            "parada_actual": parada_actual,
            "actuales_por_equipo": _agrupar_chancado_por_equipo(actuales),
            "historicos_por_equipo": _agrupar_chancado_por_equipo(historicos),
            "total_actuales": len(actuales),
            "total_historicos": len(historicos),
        },
    )


def _paradas_filtros_para_usuario(usuario):
    rol = obtener_rol(usuario)

    if rol == "Administrador":
        return list(
            Parada.objects
            .filter(planta__iexact="Filtros")
            .order_by("-fecha_inicio", "-id")
        )

    ids = (
        AccesoParada.objects
        .filter(
            usuario_id=usuario.id,
            rol=rol,
            parada__planta__iexact="Filtros",
        )
        .values_list("parada_id", flat=True)
        .distinct()
    )

    return list(
        Parada.objects
        .filter(id__in=ids)
        .order_by("-fecha_inicio", "-id")
    )


def _reportes_filtros_para_usuario(usuario, parada):
    rol = obtener_rol(usuario)

    qs = (
        ReporteFiltro.objects
        .filter(parada=parada)
        .order_by("tag", "codigo_catalogo", "id")
    )

    if rol == "Administrador":
        return list(qs)

    acceso = AccesoParada.objects.filter(
        parada=parada,
        usuario_id=usuario.id,
        rol=rol,
        activo=True,
    ).exists()

    if not acceso:
        return []

    if rol == "Cliente":
        qs = qs.filter(estado=ReporteFiltro.Estado.PUBLICADO)

    return list(qs)


def _agrupar_reportes_filtros(reportes):
    grupos = []

    orden_tags = []
    por_tag = {}

    for reporte in reportes:
        if reporte.tag not in por_tag:
            por_tag[reporte.tag] = []
            orden_tags.append(reporte.tag)

        por_tag[reporte.tag].append(reporte)

    for tag in orden_tags:
        grupos.append(
            {
                "tag": tag,
                "reportes": por_tag[tag],
            }
        )

    return grupos


@login_required
def proceso_filtros(request):
    """
    Vista de presentación de Filtros.

    La parada más reciente se considera actual.
    Las anteriores quedan en Histórico.
    """
    rol = obtener_rol(request.user)
    paradas = _paradas_filtros_para_usuario(request.user)

    parada_actual = paradas[0] if paradas else None
    paradas_historicas = paradas[1:] if len(paradas) > 1 else []

    reportes_actuales = (
        _reportes_filtros_para_usuario(request.user, parada_actual)
        if parada_actual
        else []
    )

    historicos = []

    for parada in paradas_historicas:
        reportes = _reportes_filtros_para_usuario(
            request.user,
            parada,
        )

        historicos.append(
            {
                "parada": parada,
                "total": len(reportes),
                "grupos": _agrupar_reportes_filtros(reportes),
            }
        )

    return render(
        request,
        "procesos/filtros.html",
        {
            "rol": rol,
            "parada_actual": parada_actual,
            "reportes_actuales": reportes_actuales,
            "actuales_por_tag": _agrupar_reportes_filtros(reportes_actuales),
            "total_actuales": len(reportes_actuales),
            "paradas_historicas": historicos,
        },
    )
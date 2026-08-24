import re
from decimal import Decimal

from django.contrib.auth.decorators import login_required
from django.http import HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, render

from inspecciones.models import FotoInspeccion, Inspeccion

from .excel_export import generar_reporte_faja_cvb0001_excel


TAGS_CVB0001 = {
    "CVB0001",
    "0220-CVB-0001",
    "0220-CVB0001",
}


def _es_cvb0001(inspeccion):
    return (inspeccion.faja.tag or "").upper().strip() in TAGS_CVB0001


def _obtener_rol(usuario):
    if usuario.is_superuser:
        return "Administrador"
    grupo = usuario.groups.first()
    return grupo.name if grupo else "Sin rol"


def _usuario_puede_abrir(usuario, inspeccion):
    rol = _obtener_rol(usuario)
    if rol in {"Administrador", "Inspector", "Supervisor", "Analista"}:
        return True
    if rol == "Cliente":
        return inspeccion.estado == Inspeccion.Estado.PUBLICADO
    return False


def _analizar_empalme(mediciones, nombre_empalme):
    candidatos = []
    for medicion in mediciones:
        for letra in "abcdefg":
            valor = getattr(medicion, letra)
            if valor is not None:
                candidatos.append((valor, letra.upper(), medicion))

    if not candidatos:
        return {
            "disponible": False,
            "texto": (
                f"El empalme {nombre_empalme} aún no cuenta con "
                "mediciones suficientes para generar el resumen automático."
            ),
        }

    minimo, letra, medicion = min(candidatos, key=lambda item: item[0])
    posicion_texto = {
        "-1 m": "a un metro antes del empalme",
        "+1 m": "a un metro después del empalme",
    }.get(medicion.posicion, f"en la posición {medicion.posicion}")
    zona = "zona de carga" if nombre_empalme == "E-01" else ""
    ubicacion = f"bastidor {medicion.bastidor}"
    if zona:
        ubicacion += f", {zona}"

    desgaste = None
    residual = None
    if medicion.espesor_nominal and medicion.espesor_nominal > 0:
        desgaste = medicion.espesor_nominal - minimo
        residual = minimo / medicion.espesor_nominal * Decimal("100")

    texto = (
        f"El empalme {nombre_empalme} se encontró en el {ubicacion}. "
        f"El espesor mínimo en {(medicion.lado or 'lado no indicado').lower()} "
        f"es de {minimo:.2f} mm en el punto {letra}, {posicion_texto}."
    )
    if desgaste is not None and residual is not None:
        texto += (
            f" El desgaste calculado es de {desgaste:.2f} mm y el "
            f"porcentaje residual es {residual:.2f}%."
        )

    return {
        "disponible": True,
        "minimo": minimo,
        "letra": letra,
        "posicion": medicion.posicion,
        "bastidor": medicion.bastidor,
        "lado": medicion.lado,
        "desgaste": desgaste,
        "porcentaje_residual": residual,
        "texto": texto,
    }


def _calcular_resumen_tramos(tramos):
    columnas = "abcdefg"
    resumen = {
        "minimos": {},
        "promedios": {},
        "nominal_minimo": None,
        "nominal_promedio": None,
    }
    nominales = [
        float(medicion.espesor_nominal)
        for medicion in tramos
        if medicion.espesor_nominal is not None
    ]
    if nominales:
        resumen["nominal_minimo"] = min(nominales)
        resumen["nominal_promedio"] = sum(nominales) / len(nominales)

    for columna in columnas:
        valores = [
            float(valor)
            for medicion in tramos
            if (valor := getattr(medicion, columna, None)) is not None
        ]
        resumen["minimos"][columna] = min(valores) if valores else None
        resumen["promedios"][columna] = (
            sum(valores) / len(valores) if valores else None
        )
    return resumen


def _contexto_reporte(inspeccion, usuario):
    mediciones = inspeccion.mediciones.order_by("orden", "id")
    empalme_e01 = list(
        mediciones.filter(seccion__iexact="EMPALME E-01")
    )
    empalme_e02 = list(
        mediciones.filter(seccion__iexact="EMPALME E-02")
    )
    tramos = list(mediciones.exclude(seccion__icontains="EMPALME"))
    return {
        "inspeccion": inspeccion,
        "empalme_e01": empalme_e01,
        "empalme_e02": empalme_e02,
        "tramos": tramos,
        "resumen_tramos": _calcular_resumen_tramos(tramos),
        "resumen_e01": _analizar_empalme(empalme_e01, "E-01"),
        "resumen_e02": _analizar_empalme(empalme_e02, "E-02"),
        "fotos_e01": list(
            inspeccion.fotografias.extra(
                where=["seccion = %s"], params=["EMPALME_E01"]
            ).order_by("creada_en", "id")
        ),
        "fotos_e02": list(
            inspeccion.fotografias.extra(
                where=["seccion = %s"], params=["EMPALME_E02"]
            ).order_by("creada_en", "id")
        ),
        "fotos_puntos": list(
            inspeccion.fotografias.extra(
                where=["seccion = %s"], params=["PUNTOS_MEDICION"]
            ).order_by("creada_en", "id")
        ),
        "rol": _obtener_rol(usuario),
    }


def reporte_faja_cvb0001(request, inspeccion):
    if not _es_cvb0001(inspeccion):
        return HttpResponseForbidden(
            "Este reporte exclusivo sólo corresponde a la faja CVB0001."
        )
    return render(
        request,
        "inspecciones/reporte_faja_cvb0001.html",
        _contexto_reporte(inspeccion, request.user),
    )


@login_required
def exportar_excel_cvb0001(request, inspeccion_id):
    inspeccion = get_object_or_404(
        Inspeccion.objects.select_related(
            "faja",
            "inspector",
            "supervisor",
            "analista",
            "cliente",
        ),
        id=inspeccion_id,
        tipo=Inspeccion.Tipo.FAJA,
    )
    if not _usuario_puede_abrir(request.user, inspeccion):
        return HttpResponseForbidden(
            "No tienes permiso para descargar este reporte."
        )
    if not _es_cvb0001(inspeccion):
        return HttpResponseForbidden(
            "La exportación solicitada sólo corresponde a la faja CVB0001."
        )

    contexto = _contexto_reporte(inspeccion, request.user)
    output = generar_reporte_faja_cvb0001_excel(
        inspeccion=inspeccion,
        empalme_e01=contexto["empalme_e01"],
        empalme_e02=contexto["empalme_e02"],
        tramos=contexto["tramos"],
        resumen_e01=contexto["resumen_e01"],
        resumen_e02=contexto["resumen_e02"],
        resumen_tramos=contexto["resumen_tramos"],
        fotos_e01=contexto["fotos_e01"],
        fotos_e02=contexto["fotos_e02"],
        fotos_puntos=contexto["fotos_puntos"],
    )
    codigo_seguro = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        inspeccion.codigo_reporte,
    ).strip("._") or str(inspeccion.id)
    response = HttpResponse(
        output.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
    response["Content-Disposition"] = (
        f'attachment; filename="REPORTE_CVB0001_{codigo_seguro}.xlsx"'
    )
    return response

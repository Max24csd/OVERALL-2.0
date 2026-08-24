"""Servicio transaccional de estados y trazabilidad para CVB003."""

from django.db import transaction

from inspecciones.models import HistorialEstado, Inspeccion

from .permissions import obtener_rol_cvb0003


TRANSICIONES_CVB0003 = {
    HistorialEstado.Accion.ENVIAR_A_REVISION: (
        {Inspeccion.Estado.BORRADOR, Inspeccion.Estado.DEVUELTO},
        Inspeccion.Estado.EN_REVISION,
    ),
    HistorialEstado.Accion.DEVOLVER_SUPERVISOR: (
        {Inspeccion.Estado.EN_REVISION}, Inspeccion.Estado.DEVUELTO,
    ),
    HistorialEstado.Accion.APROBAR_SUPERVISOR: (
        {Inspeccion.Estado.EN_REVISION}, Inspeccion.Estado.REVISADO,
    ),
    HistorialEstado.Accion.DEVOLVER_ANALISTA: (
        {Inspeccion.Estado.REVISADO}, Inspeccion.Estado.DEVUELTO,
    ),
    HistorialEstado.Accion.APROBAR_ANALISTA: (
        {Inspeccion.Estado.REVISADO}, Inspeccion.Estado.APROBADO,
    ),
    HistorialEstado.Accion.PUBLICAR: (
        {Inspeccion.Estado.APROBADO}, Inspeccion.Estado.PUBLICADO,
    ),
}


@transaction.atomic
def cambiar_estado_con_historial(
    inspeccion, usuario, nuevo_estado, accion, comentario=""
):
    comentario = (comentario or "").strip()
    if accion not in TRANSICIONES_CVB0003:
        raise ValueError("La acción de flujo no está permitida.")

    origenes, destino = TRANSICIONES_CVB0003[accion]
    bloqueada = Inspeccion.objects.select_for_update().get(pk=inspeccion.pk)
    anterior = bloqueada.estado
    if anterior not in origenes or nuevo_estado != destino:
        raise ValueError("La transición de estado solicitada no es válida.")
    if accion in {
        HistorialEstado.Accion.DEVOLVER_SUPERVISOR,
        HistorialEstado.Accion.DEVOLVER_ANALISTA,
    } and not comentario:
        raise ValueError("Escribe el motivo de devolución.")

    bloqueada.estado = nuevo_estado
    bloqueada.comentarios_revision = comentario
    bloqueada.save(
        update_fields=["estado", "comentarios_revision", "actualizada_en"]
    )
    evento = HistorialEstado.objects.create(
        inspeccion=bloqueada,
        estado_anterior=anterior,
        estado_nuevo=nuevo_estado,
        usuario=usuario,
        rol=obtener_rol_cvb0003(usuario),
        accion=accion,
        comentario=comentario,
    )
    inspeccion.estado = nuevo_estado
    inspeccion.comentarios_revision = comentario
    return evento

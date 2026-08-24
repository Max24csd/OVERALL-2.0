"""Autorización central por asignación para inspecciones CVB003."""

from __future__ import annotations

from inspecciones.models import Inspeccion
from inspecciones.presentation_scope import (
    INSPECCIONES_HISTORICAS_OFICIALES_IDS,
)

from .code_utils import es_tag_cvb0003


def obtener_rol_cvb0003(usuario) -> str:
    if usuario.is_superuser:
        return "Administrador"

    grupo = usuario.groups.first()
    return grupo.name if grupo else "Sin rol"


def es_inspeccion_cvb0003(inspeccion) -> bool:
    return es_tag_cvb0003(getattr(inspeccion.faja, "tag", ""))


def _cliente_actual_puede_ver_historico(usuario, inspeccion) -> bool:
    """
    Excepción estricta de SOLO LECTURA para históricos oficiales.

    El Cliente solamente puede consultar un histórico oficial cuando
    ese mismo reporte está asignado a su usuario mediante cliente_id.
    No concede permisos de edición ni altera estados o workflow.
    """
    return (
        inspeccion.id in INSPECCIONES_HISTORICAS_OFICIALES_IDS
        and inspeccion.cliente_id == usuario.id
    )


def puede_acceder_inspeccion_cvb0003(
    usuario,
    inspeccion,
    accion: str = "ver",
) -> bool:
    """Autoriza por rol, asignación, estado y acción solicitada."""
    if not usuario.is_authenticated or not es_inspeccion_cvb0003(inspeccion):
        return False

    rol = obtener_rol_cvb0003(usuario)
    estado = inspeccion.estado

    if rol == "Administrador":
        return True

    # Cliente: acceso de SOLO LECTURA a históricos oficiales
    # que están asignados a su propio usuario, incluso si son
    # registros antiguos conservados en estado BORRADOR.
    if (
        rol == "Cliente"
        and accion == "ver"
        and _cliente_actual_puede_ver_historico(
            usuario,
            inspeccion,
        )
    ):
        return True

    asignaciones = {
        "Inspector": inspeccion.inspector_id,
        "Supervisor": inspeccion.supervisor_id,
        "Analista": inspeccion.analista_id,
        "Cliente": inspeccion.cliente_id,
    }
    if asignaciones.get(rol) != usuario.id:
        return False

    if accion == "ver":
        if rol == "Cliente":
            return estado == Inspeccion.Estado.PUBLICADO
        return rol in {"Inspector", "Supervisor", "Analista"}

    if accion == "editar":
        estados_editables = {
            "Inspector": {
                Inspeccion.Estado.BORRADOR,
                Inspeccion.Estado.DEVUELTO,
            },
            "Supervisor": {Inspeccion.Estado.EN_REVISION},
            "Analista": {
                Inspeccion.Estado.REVISADO,
                Inspeccion.Estado.APROBADO,
            },
        }
        return estado in estados_editables.get(rol, set())

    if accion == "enviar_revision":
        return rol == "Inspector" and estado in {
            Inspeccion.Estado.BORRADOR,
            Inspeccion.Estado.DEVUELTO,
        }

    if accion in {"aprobar_supervisor", "devolver_supervisor"}:
        return rol == "Supervisor" and estado == Inspeccion.Estado.EN_REVISION

    if accion in {"aprobar_analista", "devolver_analista"}:
        return rol == "Analista" and estado == Inspeccion.Estado.REVISADO

    if accion == "publicar":
        return rol == "Analista" and estado == Inspeccion.Estado.APROBADO

    return False

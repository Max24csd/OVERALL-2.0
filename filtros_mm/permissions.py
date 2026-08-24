from __future__ import annotations

from accounts.models import AccesoParada, PerfilUsuario


ROLES_FILTROS = ("Inspector", "Supervisor", "Analista", "Cliente")


def obtener_rol(usuario) -> str:
    if usuario.is_superuser:
        return "Administrador"

    grupo = usuario.groups.first()
    return grupo.name if grupo else "Sin rol"


def usuario_vigente_en_sistema(usuario) -> bool:
    """
    Misma regla general usada en Chancado:
    - superusuario: vigente
    - sin PerfilUsuario: compatible con usuarios antiguos
    - PERMANENTE: vigente
    - CONTRATO: valida fechas
    - INTERMITENTE: la vigencia concreta se valida por AccesoParada
    """
    if usuario.is_superuser:
        return True

    try:
        perfil = usuario.perfil_sistema
    except PerfilUsuario.DoesNotExist:
        return True

    if perfil.tipo_vinculo == PerfilUsuario.TipoVinculo.PERMANENTE:
        return True

    if perfil.tipo_vinculo == PerfilUsuario.TipoVinculo.CONTRATO:
        return perfil.contrato_vigente()

    if perfil.tipo_vinculo == PerfilUsuario.TipoVinculo.INTERMITENTE:
        return True

    return False


def accesos_vigentes_parada(parada, rol: str | None = None):
    qs = (
        AccesoParada.objects
        .select_related("usuario")
        .filter(
            parada=parada,
            activo=True,
        )
        .order_by("rol", "usuario__first_name", "usuario__last_name", "usuario__username")
    )

    if rol:
        qs = qs.filter(rol=rol)

    # esta_vigente() ya contempla el rango temporal del acceso.
    return [
        acceso
        for acceso in qs
        if acceso.esta_vigente()
        and usuario_vigente_en_sistema(acceso.usuario)
    ]


def usuarios_vigentes_parada(parada, rol: str):
    return [
        acceso.usuario
        for acceso in accesos_vigentes_parada(parada, rol)
    ]


def usuario_asignado_a_parada(usuario, parada, rol: str) -> bool:
    return any(
        acceso.usuario_id == usuario.id
        for acceso in accesos_vigentes_parada(parada, rol)
    )


def nombre_usuario(usuario) -> str:
    return usuario.get_full_name().strip() or usuario.username


def nombres_asignados(parada, rol: str) -> str:
    return " / ".join(
        nombre_usuario(usuario)
        for usuario in usuarios_vigentes_parada(parada, rol)
    )


def contexto_responsables_parada(parada) -> dict:
    return {
        "inspectores_campo": nombres_asignados(parada, "Inspector"),
        "supervisores_campo": nombres_asignados(parada, "Supervisor"),
        "analistas": nombres_asignados(parada, "Analista"),
        "clientes": nombres_asignados(parada, "Cliente"),
    }


def puede_abrir_reporte_filtros(usuario, reporte) -> bool:
    rol = obtener_rol(usuario)

    if usuario.is_superuser or rol == "Administrador":
        return True

    if rol not in ROLES_FILTROS:
        return False

    if not reporte.parada_id:
        return False

    if not usuario_vigente_en_sistema(usuario):
        return False

    if not usuario_asignado_a_parada(usuario, reporte.parada, rol):
        return False

    if rol == "Cliente":
        return reporte.estado == reporte.Estado.PUBLICADO

    return True


def puede_editar_reporte_filtros(usuario, reporte) -> bool:
    rol = obtener_rol(usuario)

    if usuario.is_superuser or rol == "Administrador":
        return True

    if not puede_abrir_reporte_filtros(usuario, reporte):
        return False

    if rol == "Inspector":
        return reporte.estado in [
            reporte.Estado.BORRADOR,
            reporte.Estado.DEVUELTO,
        ]

    # Supervisor y Analista revisan/validan mediante acciones de workflow,
    # no modifican las mediciones técnicas capturadas por el inspector.
    return False

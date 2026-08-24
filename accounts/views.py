from django.contrib import messages
from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.contrib.auth.models import User
from django.http import HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from .forms import UsuarioSistemaForm
from .models import AccesoParada, PerfilUsuario


def _usuario_puede_iniciar_sesion(usuario):
    """
    Mantiene compatibilidad con usuarios antiguos sin PerfilUsuario.

    PERMANENTE:
        Puede iniciar sesión mientras la cuenta esté activa.

    CONTRATO:
        Puede iniciar sesión solamente mientras el contrato esté vigente.

    INTERMITENTE:
        Puede iniciar sesión solamente si existe al menos un acceso
        temporal activo a una parada dentro de su ventana autorizada.
    """
    if usuario.is_superuser:
        return True, ""

    try:
        perfil = usuario.perfil_sistema
    except PerfilUsuario.DoesNotExist:
        return True, ""

    if (
        perfil.tipo_vinculo
        == PerfilUsuario.TipoVinculo.PERMANENTE
    ):
        return True, ""

    if (
        perfil.tipo_vinculo
        == PerfilUsuario.TipoVinculo.CONTRATO
    ):
        if perfil.contrato_vigente():
            return True, ""

        return (
            False,
            "Tu acceso por contrato no se encuentra vigente.",
        )

    if (
        perfil.tipo_vinculo
        == PerfilUsuario.TipoVinculo.INTERMITENTE
    ):
        ahora = timezone.now()

        acceso_vigente = AccesoParada.objects.filter(
            usuario=usuario,
            activo=True,
            fecha_inicio__lte=ahora,
            fecha_fin__gte=ahora,
        ).exists()

        if acceso_vigente:
            return True, ""

        return (
            False,
            (
                "Tu acceso intermitente no está habilitado "
                "en este momento."
            ),
        )

    return False, "Tu cuenta no tiene un tipo de acceso válido."


def login_view(request):
    if request.user.is_authenticated:
        return redirect("dashboard")

    if request.method == "POST":
        username = request.POST.get("username", "").strip()
        password = request.POST.get("password", "")

        usuario = authenticate(
            request,
            username=username,
            password=password,
        )

        if usuario is not None:
            permitido, mensaje = _usuario_puede_iniciar_sesion(
                usuario
            )

            if permitido:
                login(request, usuario)
                return redirect("dashboard")

            messages.error(
                request,
                mensaje,
            )

        else:
            messages.error(
                request,
                "Usuario o contraseña incorrectos.",
            )

    return render(
        request,
        "accounts/login.html",
    )


def logout_view(request):
    if request.method == "POST":
        logout(request)
        return redirect("login")

    return redirect("dashboard")


def obtener_rol_usuario(usuario):
    if usuario.is_superuser:
        return "Administrador"

    grupo = usuario.groups.first()

    return grupo.name if grupo else "Sin rol"


def es_administrador(usuario):
    return obtener_rol_usuario(usuario) == "Administrador"


@login_required
def usuarios_lista(request):
    if not es_administrador(request.user):
        return HttpResponseForbidden(
            "Solo el administrador puede gestionar usuarios."
        )

    usuarios = (
        User.objects
        .prefetch_related("groups")
        .select_related("perfil_sistema")
        .order_by(
            "first_name",
            "last_name",
            "username",
        )
    )

    return render(
        request,
        "accounts/usuarios_lista.html",
        {
            "usuarios": usuarios,
        },
    )


@login_required
def usuario_crear(request):
    if not es_administrador(request.user):
        return HttpResponseForbidden(
            "Solo el administrador puede crear usuarios."
        )

    if request.method == "POST":
        formulario = UsuarioSistemaForm(
            request.POST
        )

        if formulario.is_valid():
            usuario = formulario.save()

            messages.success(
                request,
                (
                    f"El usuario {usuario.username} "
                    "fue creado correctamente."
                ),
            )

            return redirect("usuarios_lista")

        messages.error(
            request,
            "Revisa los campos marcados.",
        )

    else:
        formulario = UsuarioSistemaForm()

    return render(
        request,
        "accounts/usuario_formulario.html",
        {
            "formulario": formulario,
            "titulo": "Crear usuario",
            "es_edicion": False,
        },
    )


@login_required
def usuario_editar(request, usuario_id):
    if not es_administrador(request.user):
        return HttpResponseForbidden(
            "Solo el administrador puede editar usuarios."
        )

    usuario = get_object_or_404(
        User,
        id=usuario_id,
    )

    if request.method == "POST":
        formulario = UsuarioSistemaForm(
            request.POST,
            instance=usuario,
        )

        if formulario.is_valid():
            usuario = formulario.save()

            messages.success(
                request,
                (
                    f"El usuario {usuario.username} "
                    "fue actualizado correctamente."
                ),
            )

            return redirect("usuarios_lista")

        messages.error(
            request,
            "Revisa los campos marcados.",
        )

    else:
        formulario = UsuarioSistemaForm(
            instance=usuario
        )

    return render(
        request,
        "accounts/usuario_formulario.html",
        {
            "formulario": formulario,
            "titulo": "Editar usuario",
            "es_edicion": True,
            "usuario_editado": usuario,
        },
    )


@login_required
@require_POST
def usuario_cambiar_estado(request, usuario_id):
    if not es_administrador(request.user):
        return HttpResponseForbidden(
            "Solo el administrador puede cambiar usuarios."
        )

    usuario = get_object_or_404(
        User,
        id=usuario_id,
    )

    if usuario == request.user:
        messages.error(
            request,
            "No puedes desactivar tu propia cuenta.",
        )

        return redirect("usuarios_lista")

    usuario.is_active = not usuario.is_active

    usuario.save(
        update_fields=["is_active"]
    )

    estado = (
        "activado"
        if usuario.is_active
        else "desactivado"
    )

    messages.success(
        request,
        f"El usuario fue {estado}.",
    )

    return redirect("usuarios_lista")

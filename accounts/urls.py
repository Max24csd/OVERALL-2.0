from django.urls import path

from . import views


urlpatterns = [
    path(
        "login/",
        views.login_view,
        name="login",
    ),
    path(
        "logout/",
        views.logout_view,
        name="logout",
    ),

    path(
        "usuarios/",
        views.usuarios_lista,
        name="usuarios_lista",
    ),
    path(
        "usuarios/nuevo/",
        views.usuario_crear,
        name="usuario_crear",
    ),
    path(
        "usuarios/<int:usuario_id>/editar/",
        views.usuario_editar,
        name="usuario_editar",
    ),
    path(
        "usuarios/<int:usuario_id>/estado/",
        views.usuario_cambiar_estado,
        name="usuario_cambiar_estado",
    ),
]
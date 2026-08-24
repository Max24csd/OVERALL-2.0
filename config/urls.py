from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.urls import include, path


urlpatterns = [
    path(
        "admin/",
        admin.site.urls,
    ),

    # Usuarios / login / logout
    path(
        "",
        include("accounts.urls"),
    ),

    # Sistema actual de inspecciones de Chancado
    path(
        "",
        include("inspecciones.urls"),
    ),

    # Nuevo módulo de Filtros
    path(
        "filtros-mm/",
        include("filtros_mm.urls"),
    ),
]


if settings.DEBUG:
    urlpatterns += static(
        settings.MEDIA_URL,
        document_root=settings.MEDIA_ROOT,
    )
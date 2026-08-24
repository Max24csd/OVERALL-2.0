from django.urls import path

from . import views


app_name = "filtros_mm"


urlpatterns = [
    path(
        "",
        views.lista_paradas_filtros,
        name="lista_paradas",
    ),

    path(
        "paradas/nueva/",
        views.nueva_parada_filtros,
        name="nueva_parada",
    ),

    path(
        "parada/<int:parada_id>/",
        views.inicio_filtros,
        name="inicio",
    ),

    path(
        "parada/<int:parada_id>/reporte/<str:codigo_catalogo>/",
        views.formulario_reporte,
        name="formulario_reporte",
    ),

    path(
        "parada/<int:parada_id>/reporte/<str:codigo_catalogo>/excel/",
        views.exportar_excel_reporte,
        name="exportar_excel_reporte",
    ),
]

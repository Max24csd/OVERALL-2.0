from django.urls import path
from . import views
from .reportes.cvb0001 import views as views_cvb0001
from .reportes.cvb0001 import life_shaft_views as views_life_shaft_cvb0001
from .reportes.cvb0001 import poleas_views as views_poleas_cvb0001
from .reportes.cvb0003 import views as views_life_shaft_cvb0003
from .reportes.cvb0003 import poleas_views as views_poleas_cvb0003
from .reportes.cvb0003 import faja_views as views_faja_cvb0003
from .reportes.cvb0004 import life_shaft_views as views_life_shaft_cvb0004
from .reportes.cvb0004 import poleas_views as views_poleas_cvb0004
from .reportes.cvb0004 import faja_views as views_faja_cvb0004
from . import views
from . import process_views

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("inspecciones/<int:inspeccion_id>/faja/", views.formulario_faja, name="formulario_faja"),
    path("inspecciones/<int:inspeccion_id>/faja/reporte/", views.reporte_faja, name="reporte_faja"),
    path("inspecciones/<int:inspeccion_id>/faja/exportar-excel/", views.exportar_excel_faja_cvb0001, name="exportar_excel_faja_cvb0001"),
    path("inspecciones/<int:inspeccion_id>/molienda-cvb0006/exportar-excel/", views.exportar_excel_molienda_cvb0006, name="exportar_excel_molienda_cvb0006"),
    path("inspecciones/<int:inspeccion_id>/molienda-cvb0006/exportar-pdf/", views.exportar_pdf_molienda_cvb0006, name="exportar_pdf_molienda_cvb0006"),
    path("inspecciones/<int:inspeccion_id>/cvb0003/faja/exportar-excel/", views_faja_cvb0003.exportar_excel_faja_cvb0003, name="exportar_excel_faja_cvb0003"),
    path("inspecciones/<int:inspeccion_id>/cvb0004/faja/exportar-excel/", views_faja_cvb0004.exportar_excel_faja_cvb0004, name="exportar_excel_faja_cvb0004"),
    path("inspecciones/<int:inspeccion_id>/faja/cvb0001/exportar-excel/", views_cvb0001.exportar_excel_cvb0001, name="exportar_excel_cvb0001_exclusivo"),
    path("inspecciones/<int:inspeccion_id>/estado/<str:accion>/", views.cambiar_estado_inspeccion, name="cambiar_estado_inspeccion"),
    path("inspecciones/<int:inspeccion_id>/poleas/", views.formulario_poleas, name="formulario_poleas"),
    path("inspecciones/<int:inspeccion_id>/poleas/reporte/", views.reporte_poleas, name="reporte_poleas"),
    path("inspecciones/<int:inspeccion_id>/poleas/exportar-excel/", views_poleas_cvb0003.exportar_excel_poleas_cvb0003, name="exportar_excel_poleas_cvb0003"),
    path("inspecciones/<int:inspeccion_id>/cvb0001/poleas/exportar-excel/", views_poleas_cvb0001.exportar_excel_poleas_cvb0001, name="exportar_excel_poleas_cvb0001"),
    path("inspecciones/<int:inspeccion_id>/cvb0004/poleas/exportar-excel/", views_poleas_cvb0004.exportar_excel_poleas_cvb0004, name="exportar_excel_poleas_cvb0004"),
    path("inspecciones/<int:inspeccion_id>/life-shaft/", views.formulario_life_shaft, name="formulario_life_shaft"),
    path("inspecciones/<int:inspeccion_id>/life-shaft/reporte/", views.reporte_life_shaft, name="reporte_life_shaft"),
    path("inspecciones/<int:inspeccion_id>/life-shaft/exportar-excel/", views_life_shaft_cvb0003.exportar_excel_life_shaft_cvb0003, name="exportar_excel_life_shaft_cvb0003"),
    path("inspecciones/<int:inspeccion_id>/cvb0001/life-shaft/exportar-excel/", views_life_shaft_cvb0001.exportar_excel_life_shaft_cvb0001, name="exportar_excel_life_shaft_cvb0001"),
    path("inspecciones/<int:inspeccion_id>/cvb0004/life-shaft/exportar-excel/", views_life_shaft_cvb0004.exportar_excel_life_shaft_cvb0004, name="exportar_excel_life_shaft_cvb0004"),
    path("historial/", views.historial_inspecciones, name="historial_inspecciones"),
    path(
    "paradas/nueva/",
    views.nueva_parada,
    name="nueva_parada",
    ),
    path(
    "paradas/molienda/nueva/",
    views.nueva_parada_molienda,
    name="nueva_parada_molienda",
    ),
    path(
    "paradas/<int:parada_id>/asignaciones/",
    views.gestionar_asignaciones_parada,
    name="gestionar_asignaciones_parada",
    ),
    path(
    "procesos/chancado/",
    process_views.proceso_chancado,
    name="proceso_chancado",
    ),
    path(
    "procesos/molienda/",
    process_views.proceso_molienda,
    name="proceso_molienda",
    ),

    path(
    "procesos/filtros/",
    process_views.proceso_filtros,
    name="proceso_filtros",
    ),
]

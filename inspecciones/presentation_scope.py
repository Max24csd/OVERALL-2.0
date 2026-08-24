"""
Configuración de presentación de la versión aprobada.

IMPORTANTE:
- Este archivo NO modifica la lógica técnica.
- NO elimina inspecciones.
- NO interviene en históricos A-G.
- NO interviene en restricciones.
- NO interviene en Excel, fotografías o workflow.

Únicamente define qué registros se muestran en Dashboard e Historial.
"""


# ============================================================
# PARADA ACTUAL OFICIAL
# ============================================================

PARADA_ACTUAL_ID = 2


# ============================================================
# HISTÓRICO OFICIAL VISIBLE
# ============================================================
#
# Estos registros corresponden al conjunto histórico utilizado
# como referencia anterior a la parada actual.
#
# CVB001 -> 1, 2, 3
# CVB003 -> 7, 8, 9
# CVB004 -> 10, 11, 12
#

INSPECCIONES_HISTORICAS_OFICIALES_IDS = (
    1,
    2,
    3,
    7,
    8,
    9,
    10,
    11,
    12,
)


# ============================================================
# DEMOS OCULTAS EN PRESENTACIÓN
# ============================================================

INSPECCIONES_DEMO_IDS = (
    13,
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    22,
    23,
    24,
)


# ============================================================
# CONFIGURACIÓN DE CLIENTE
# ============================================================

MOSTRAR_HISTORICO_AL_CLIENTE = True

CANTIDAD_PARADAS_HISTORICAS_CLIENTE = 1


# ============================================================
# UTILIDADES
# ============================================================

def es_inspeccion_historica_oficial(inspeccion):
    return inspeccion.id in INSPECCIONES_HISTORICAS_OFICIALES_IDS


def es_inspeccion_demo(inspeccion):
    return inspeccion.id in INSPECCIONES_DEMO_IDS


def es_inspeccion_actual(inspeccion):
    return (
        getattr(inspeccion, "parada_id", None)
        == PARADA_ACTUAL_ID
    )
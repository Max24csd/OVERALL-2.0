"""Historial y validación técnica exclusiva de CVB003."""

from __future__ import annotations

from decimal import Decimal

from inspecciones.models import FaseCampana, Inspeccion, TipoMedicionComponente


CAMPOS_MEDICION = tuple("abcdefg")
MENSAJE_INVALIDO = (
    "Valor inválido. La nueva medición no puede ser mayor que la "
    "medición histórica equivalente: {valor} mm."
)


def _inspecciones_anteriores(inspeccion, fecha_tecnica):
    if not fecha_tecnica:
        return Inspeccion.objects.none()
    return (
        Inspeccion.objects.filter(
            faja=inspeccion.faja,
            tipo=inspeccion.tipo,
            fecha_inspeccion__isnull=False,
            fecha_inspeccion__lt=fecha_tecnica,
        )
        .exclude(pk=inspeccion.pk)
        .order_by("-fecha_inspeccion", "-id")
    )


def _valores(mediciones):
    return {
        medicion.punto: {
            campo: getattr(medicion, campo)
            for campo in CAMPOS_MEDICION
            if getattr(medicion, campo) is not None
        }
        for medicion in mediciones
    }


def _filas(mediciones):
    filas = []
    for medicion in mediciones:
        valores = [
            getattr(medicion, campo)
            for campo in CAMPOS_MEDICION
            if getattr(medicion, campo) is not None
        ]
        filas.append(
            {
                "punto": medicion.punto,
                "valores": [getattr(medicion, campo) for campo in CAMPOS_MEDICION],
                "minimo": min(valores) if valores else None,
                "promedio": (
                    round(sum(valores) / len(valores), 2) if valores else None
                ),
                "observacion": medicion.observacion,
            }
        )
    return filas


def _minimo_general(mediciones):
    valores = [
        getattr(medicion, campo)
        for medicion in mediciones
        for campo in CAMPOS_MEDICION
        if getattr(medicion, campo) is not None
    ]
    return min(valores) if valores else None


def _promedio_general(mediciones):
    valores = [
        getattr(medicion, campo)
        for medicion in mediciones
        for campo in CAMPOS_MEDICION
        if getattr(medicion, campo) is not None
    ]
    return round(sum(valores) / len(valores), 2) if valores else None


def _resumen_componente(inspeccion, componente, mediciones, tipo, es_actual=False):
    if not mediciones:
        return None
    condicion = componente.get_condicion_display()
    if componente.condicion == Inspeccion.Condicion.NO_MEDIDO:
        condicion = "TOLERABLE"
    return {
        "fecha": inspeccion.fecha_inspeccion,
        "codigo": inspeccion.codigo_reporte,
        "componente": componente.nombre,
        "tipo": tipo,
        "condicion": condicion,
        "observacion": componente.observacion_medicion,
        "filas": _filas(mediciones),
        "valores": _valores(mediciones),
        "minimo": _minimo_general(mediciones),
        "promedio": _promedio_general(mediciones),
        "fotografias": list(componente.fotografias.order_by("creada_en", "id")),
        "es_actual": es_actual,
    }


def historial_componente(inspeccion, componente, fecha_tecnica, relacion):
    """
    Devuelve el último valor histórico disponible por PUNTO y por coordenada A-G.

    Antes se tomaba una sola inspección anterior completa. Eso provocaba que,
    si la parada inmediatamente anterior tenía datos parciales (por ejemplo,
    sólo el punto 1 de la Polea 01), los puntos 2 y 3 existentes en una parada
    más antigua dejaran de aparecer y tampoco recibieran restricción.

    Ahora:
    - se recorren las inspecciones anteriores de la más reciente a la más antigua;
    - para cada punto y cada coordenada A-G se conserva el primer valor no nulo;
    - los huecos se completan con la siguiente inspección anterior disponible;
    - nunca se reemplaza un valor más reciente por uno más antiguo.
    """
    valores_por_punto = {}
    fuente_por_punto = {}

    inspeccion_referencia = None
    componente_referencia = None
    tipo_referencia = None

    for anterior in _inspecciones_anteriores(inspeccion, fecha_tecnica):
        previo = getattr(anterior, relacion).filter(
            numero=componente.numero
        ).first()

        if previo is None:
            continue

        if previo.tipo_medicion == TipoMedicionComponente.CAMPANA:
            mediciones = list(
                previo.mediciones_campana
                .filter(fase=FaseCampana.INICIO)
                .order_by("orden", "punto")
            )
            tipo = "INICIO DE CAMPAÑA"
        else:
            mediciones = list(
                previo.mediciones.order_by("orden", "punto")
            )
            tipo = "NORMAL"

        if not mediciones:
            continue

        aporto_en_esta_inspeccion = False

        for medicion in mediciones:
            punto = medicion.punto
            valores_punto = valores_por_punto.setdefault(punto, {})
            aporto_en_este_punto = False

            for campo in CAMPOS_MEDICION:
                valor = getattr(medicion, campo)

                # La inspección más reciente tiene prioridad.
                # Sólo se rellena un hueco que todavía no tenga histórico.
                if valor is not None and campo not in valores_punto:
                    valores_punto[campo] = valor
                    aporto_en_este_punto = True
                    aporto_en_esta_inspeccion = True

            if aporto_en_este_punto and punto not in fuente_por_punto:
                fuente_por_punto[punto] = {
                    "fecha": anterior.fecha_inspeccion,
                    "codigo": anterior.codigo_reporte,
                    "observacion": medicion.observacion,
                    "tipo": tipo,
                }

        if aporto_en_esta_inspeccion and inspeccion_referencia is None:
            inspeccion_referencia = anterior
            componente_referencia = previo
            tipo_referencia = tipo

    # Eliminar puntos que realmente no consiguieron ningún valor A-G.
    valores_por_punto = {
        punto: valores
        for punto, valores in valores_por_punto.items()
        if valores
    }

    if not valores_por_punto or inspeccion_referencia is None:
        return None

    filas = []
    todos_los_valores = []

    for punto in sorted(valores_por_punto):
        valores_punto = valores_por_punto[punto]
        fila_valores = [
            valores_punto.get(campo)
            for campo in CAMPOS_MEDICION
        ]
        disponibles = [
            valor
            for valor in fila_valores
            if valor is not None
        ]
        todos_los_valores.extend(disponibles)

        fuente = fuente_por_punto.get(punto, {})

        filas.append(
            {
                "punto": punto,
                "valores": fila_valores,
                "minimo": min(disponibles) if disponibles else None,
                "promedio": (
                    round(sum(disponibles) / len(disponibles), 2)
                    if disponibles
                    else None
                ),
                "observacion": fuente.get("observacion", ""),
                "fecha": fuente.get("fecha"),
                "codigo": fuente.get("codigo"),
            }
        )

    condicion = componente_referencia.get_condicion_display()
    if componente_referencia.condicion == Inspeccion.Condicion.NO_MEDIDO:
        condicion = "TOLERABLE"

    return {
        # La cabecera conserva como referencia la inspección anterior más
        # reciente que aportó al menos un dato.
        "fecha": inspeccion_referencia.fecha_inspeccion,
        "codigo": inspeccion_referencia.codigo_reporte,
        "componente": componente.nombre,
        "tipo": tipo_referencia,
        "condicion": condicion,
        "observacion": componente_referencia.observacion_medicion,
        "filas": filas,
        "valores": valores_por_punto,
        "minimo": min(todos_los_valores) if todos_los_valores else None,
        "promedio": (
            round(sum(todos_los_valores) / len(todos_los_valores), 2)
            if todos_los_valores
            else None
        ),
        "fotografias": list(
            componente_referencia.fotografias.order_by(
                "creada_en",
                "id",
            )
        ),
        "es_actual": False,
    }

def historial_componente_visible(inspeccion, componente, fecha_tecnica, relacion):
    """Devuelve exclusivamente una inspección cronológicamente anterior."""
    return historial_componente(inspeccion, componente, fecha_tecnica, relacion)


def _clave_empalme(medicion):
    """
    Identidad técnica estable del empalme.

    No usamos bastidor_lado porque ese texto ha cambiado entre campañas
    (por ejemplo, E-15 pasó de CARGA BC-221 a CARGA BC-244), aunque siga
    tratándose del mismo empalme y posición.
    """
    return (
        medicion.empalme,
        medicion.posicion,
    )


def _clave_tramo(medicion):
    """
    Identidad técnica estable del tramo.

    El nombre del tramo y el bastidor han cambiado entre campañas.
    Para comparar la misma secuencia de medición usamos tipo + número.
    """
    return (
        medicion.tipo,
        medicion.medicion,
    )


def historial_faja(inspeccion, fecha_tecnica, mediciones_actuales, clase):
    """
    Mapea cada fila actual contra la última medición histórica equivalente.

    Importante:
    - Empalmes: se compara por empalme + posición.
    - Tramos: se compara por tipo + número de medición.
    - Filas históricas vacías se ignoran y se sigue buscando más atrás.
    """
    es_empalme = clase == "empalme"
    clave = _clave_empalme if es_empalme else _clave_tramo
    relacion = "empalmes_cvb0003" if es_empalme else "tramos_cvb0003"

    pendientes = {
        clave(medicion): medicion
        for medicion in mediciones_actuales
    }
    resultados = {}

    for anterior in _inspecciones_anteriores(inspeccion, fecha_tecnica):
        anteriores = getattr(anterior, relacion).all()

        indice = {
            clave(medicion): medicion
            for medicion in anteriores
        }

        for identidad in list(pendientes):
            previo = indice.get(identidad)

            if previo is None:
                continue

            valores = [
                getattr(previo, campo)
                for campo in CAMPOS_MEDICION
                if getattr(previo, campo) is not None
            ]

            # Si la fila existe pero está vacía, no la usamos como histórico.
            if not valores:
                continue

            actual = pendientes.pop(identidad)

            resultados[actual.pk] = {
                "fecha": anterior.fecha_inspeccion,
                "codigo": anterior.codigo_reporte,
                "componente": " / ".join(
                    str(valor) for valor in identidad
                ),
                "tipo": "NORMAL",
                "condicion": anterior.get_condicion_general_display(),
                "observacion": previo.observacion,
                "filas": [
                    {
                        "punto": (
                            previo.posicion
                            if es_empalme
                            else previo.medicion
                        ),
                        "valores": [
                            getattr(previo, campo)
                            for campo in CAMPOS_MEDICION
                        ],
                        "minimo": min(valores),
                        "promedio": round(
                            sum(valores) / len(valores),
                            2,
                        ),
                        "observacion": previo.observacion,
                    }
                ],
                "valores": {
                    campo: getattr(previo, campo)
                    for campo in CAMPOS_MEDICION
                    if getattr(previo, campo) is not None
                },
                "minimo": min(valores),
                "fotografias": list(
                    anterior.fotografias_cvb0003
                    .order_by("creada_en", "id")
                ),
            }

        if not pendientes:
            break

    return resultados


def preparar_formset_historico(formset, historico, clave_form=None):
    """Inyecta los máximos históricos en los widgets sin validar todavía."""
    if not historico:
        return
    for form in formset.forms:
        clave = clave_form(form) if clave_form else _punto_form(form)
        valores = historico.get(clave, {})
        for campo, maximo in valores.items():
            if campo not in form.fields or maximo is None:
                continue
            form.fields[campo].widget.attrs.update(
                {
                    "data-historical-value": str(maximo),
                    "data-history-validation": "1",
                    "title": f"Máximo histórico permitido: {maximo} mm",
                }
            )


def validar_formset_historico(formset, historico, clave_form=None, restringir=True):
    """Valida cada coordenada luego de la validación normal del formset."""
    valido = formset.is_valid()
    if not valido or not historico or not restringir:
        return valido

    for form in formset.forms:
        if not hasattr(form, "cleaned_data") or form.cleaned_data.get("DELETE"):
            continue
        clave = clave_form(form) if clave_form else form.cleaned_data.get("punto")
        valores = historico.get(clave, {})
        for campo, maximo in valores.items():
            actual = form.cleaned_data.get(campo)
            if actual is not None and Decimal(actual) > Decimal(maximo):
                form.add_error(campo, MENSAJE_INVALIDO.format(valor=maximo))
                valido = False
    return valido


def _punto_form(form):
    if form.is_bound:
        valor = form.data.get(form.add_prefix("punto"))
        if valor not in (None, ""):
            try:
                return int(valor)
            except (TypeError, ValueError):
                pass
    return form.initial.get("punto") or getattr(form.instance, "punto", None)


def clave_pk_form(form):
    return getattr(form.instance, "pk", None)


def modo_campana_seleccionado(formulario_componente, componente):
    """Prioriza el tipo enviado y validado en el POST sobre el persistido."""
    if formulario_componente.is_bound:
        return (
            formulario_componente.data.get(
                formulario_componente.add_prefix("tipo_medicion")
            )
            == TipoMedicionComponente.CAMPANA
        )
    return componente.tipo_medicion == TipoMedicionComponente.CAMPANA

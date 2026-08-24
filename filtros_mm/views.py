from decimal import Decimal, InvalidOperation

from django.contrib import messages
from django.contrib.auth.decorators import login_required
from django.db import transaction
from django.http import Http404, HttpResponse, HttpResponseForbidden
from django.shortcuts import get_object_or_404, redirect, render

from accounts.models import AccesoParada, HistorialAsignacionParada
from inspecciones.models import Parada

from .config.catalogo import REPORTES_FILTROS
from .forms import ReporteFiltroCabeceraForm
from .parada_forms import NuevaParadaFiltrosForm
from .models import FotoReporteFiltro, ReporteFiltro
from .permissions import (
    contexto_responsables_parada,
    obtener_rol,
    puede_abrir_reporte_filtros,
    puede_editar_reporte_filtros,
    usuario_asignado_a_parada,
    usuarios_vigentes_parada,
)
from .services.excel_export import generar_excel_carrileria
from .services.faja_poleas_export import generar_excel_faja_poleas


PUNTOS_CARRILERIA = list(range(36, -33, -1))
LADOS_RUEDA = ("izquierdo", "derecho")
NUMEROS_RUEDA = (1, 2, 3)
PUNTOS_RUEDA = ("a1", "a2", "a3", "b1", "b2", "b3")
FOTO_SLOTS_CARRILERIA = [
    {
        "seccion": "zonas_soldadura",
        "titulo": "Registro fotografico de zonas de inspeccion",
        "input": f"foto_zonas_{orden}",
        "comentario": f"comentario_zonas_{orden}",
        "orden": orden,
    }
    for orden in range(1, 13)
] + [
    {
        "seccion": "ruedas",
        "titulo": "Registro fotografico de la inspeccion de las ruedas",
        "input": f"foto_ruedas_{orden}",
        "comentario": f"comentario_ruedas_{orden}",
        "orden": orden,
    }
    for orden in range(1, 13)
]
FOTO_SLOTS_FAJA_POLEAS = [
    {
        "seccion": seccion,
        "grupo": grupo,
        "titulo": titulo,
        "input": f"foto_{seccion}_{orden}",
        "comentario": f"comentario_{seccion}_{orden}",
        "orden": orden,
    }
    for seccion, grupo, titulo in [
        ("faja_empalme", "Top Cover y Empalme", "Registro fotografico de faja y empalme"),
        ("polea_cola", "Polea de Cola", "Registro fotografico de polea de cola"),
        ("polea_cabeza", "Polea de Cabeza", "Registro fotografico de polea de cabeza"),
    ]
    for orden in range(1, 6)
]
FAJA_POLEAS_BELTING = [
    ("item", "ITEM N°", "600"),
    ("qty", "QTY.", "52m"),
    ("manufacturer", "MANUFACTURER", "PHOENIX"),
    ("type", "TYPE", "EP 1250"),
    ("rated", "RATED N/mm", "1250"),
    ("belt_width", "BELT WIDTH", "2438mm"),
    ("plys", "PLYS", "4"),
    ("top_cover", "TOP COVER", "10"),
    ("bottom_cover", "BOTTOM COVER", "3"),
    ("cover_type", "COVER TYPE", "y"),
    ("splice_type", "SPLICE TYPE", "VULCANIZED"),
    ("remarks", "REMARKS", "-"),
]
FAJA_POLEAS_UT_FIELDS = [
    ("marca", "MARCA DEL EQUIPO", "OLYMPUS"),
    ("modelo", "MODELO", "EPOCH 6LT"),
    ("tipo_haz", "TIPO DE HAZ", "HAZ RECTO"),
    ("ganancia", "GANANCIA (dB)", "45"),
    ("frecuencia", "FRECUENCIA (MHz)", "1"),
    ("velocidad", "VELOCIDAD (m/s)", "1670"),
    ("ancho_banda", "ANCHO DE BANDA", "1 MHZ"),
    ("retardo", "RETARDO (us)", "3.28"),
    ("amortiguamiento", "AMORTIGUAMIENTO", "400"),
    ("diametro", "DIAMETRO (mm)", "13"),
]
FAJA_POLEAS_SECCIONES = [
    ("top_cover", "MEDICIÓN DE ESPESORES DEL TOP COVER", "TOP COVER DE LA FAJA"),
    ("empalme", "MEDICIÓN DE ESPESORES DEL EMPALME DE LA FAJA", "EMPALME DE LA FAJA"),
    ("polea_cola", "MEDICIÓN DE ESPESORES DEL LAGGING DE LA POLEA DE COLA", "POLEA DE COLA DE LA FAJA"),
    ("polea_cabeza", "MEDICIÓN DE ESPESORES DEL LAGGING DE LA POLEA DE CABEZA", "POLEA DE CABEZA DE LA FAJA"),
]
FAJA_POLEAS_COLUMNAS_MEDICION = ("A", "B", "C", "D", "E", "F", "G")


def _valor_decimal_o_vacio(valor):
    valor = (valor or "").strip()
    if not valor:
        return ""

    try:
        return str(Decimal(valor.replace(",", ".")))
    except (InvalidOperation, ValueError):
        return valor


def _faja_poleas_desde_post(post):
    datos = {
        "belting": {},
        "secciones": {},
    }

    for key, _label, _default in FAJA_POLEAS_BELTING:
        datos["belting"][key] = (post.get(f"belting_{key}") or "").strip()

    for key, _titulo, componente in FAJA_POLEAS_SECCIONES:
        seccion = {
            "procedimiento": (post.get(f"{key}_procedimiento") or "").strip(),
            "material": (post.get(f"{key}_material") or "").strip(),
            "componente": (post.get(f"{key}_componente") or componente).strip(),
            "comentarios": (post.get(f"{key}_comentarios") or "").strip(),
            "ut": {},
            "mediciones": [],
        }
        for campo, _label, _default in FAJA_POLEAS_UT_FIELDS:
            seccion["ut"][campo] = (post.get(f"{key}_ut_{campo}") or "").strip()

        filas = ("antes", "despues") if key == "empalme" else ("1", "2", "3")
        for fila in filas:
            valores = {
                columna.lower(): _valor_decimal_o_vacio(
                    post.get(f"{key}_{fila}_{columna.lower()}")
                )
                for columna in FAJA_POLEAS_COLUMNAS_MEDICION
            }
            seccion["mediciones"].append(
                {
                    "punto": fila,
                    "valores": valores,
                }
            )
        datos["secciones"][key] = seccion

    return datos


def _faja_poleas_para_template(reporte):
    datos = (reporte.datos or {}).get("faja_poleas") or {}
    belting_guardado = datos.get("belting") or {}
    secciones_guardadas = datos.get("secciones") or {}

    belting = [
        {
            "key": key,
            "label": label,
            "value": belting_guardado.get(key, default),
        }
        for key, label, default in FAJA_POLEAS_BELTING
    ]

    secciones = []
    for key, titulo, componente in FAJA_POLEAS_SECCIONES:
        guardada = secciones_guardadas.get(key) or {}
        ut_guardado = guardada.get("ut") or {}
        filas_guardadas = {
            str(item.get("punto")): item.get("valores", {})
            for item in guardada.get("mediciones", [])
        }
        filas = []
        for punto in (("antes", "despues") if key == "empalme" else ("1", "2", "3")):
            valores = filas_guardadas.get(punto, {})
            filas.append(
                {
                    "punto": punto,
                    "label": punto.upper() if key == "empalme" else punto,
                    "valores": [
                        {
                            "columna": columna,
                            "name": f"{key}_{punto}_{columna.lower()}",
                            "value": valores.get(columna.lower(), ""),
                        }
                        for columna in FAJA_POLEAS_COLUMNAS_MEDICION
                    ],
                }
            )

        secciones.append(
            {
                "key": key,
                "titulo": titulo,
                "procedimiento": guardada.get("procedimiento", ""),
                "material": guardada.get("material", "CAUCHO"),
                "componente": guardada.get("componente", f"{componente} {reporte.tag[-2:]}"),
                "comentarios": guardada.get("comentarios", ""),
                "ut": [
                    {
                        "key": campo,
                        "label": label,
                        "name": f"{key}_ut_{campo}",
                        "value": ut_guardado.get(campo, default),
                    }
                    for campo, label, default in FAJA_POLEAS_UT_FIELDS
                ],
                "filas": filas,
            }
        )

    return {
        "belting": belting,
        "secciones": secciones,
        "columnas": FAJA_POLEAS_COLUMNAS_MEDICION,
    }


def _mediciones_carrileria_desde_post(post):
    resultado = []

    for punto in PUNTOS_CARRILERIA:
        sufijo = str(punto).replace("-", "m")
        resultado.append(
            {
                "punto": punto,
                "espesor_a": _valor_decimal_o_vacio(post.get(f"espesor_a_{sufijo}")),
                "espesor_b": _valor_decimal_o_vacio(post.get(f"espesor_b_{sufijo}")),
                "espesor_c": _valor_decimal_o_vacio(post.get(f"espesor_c_{sufijo}")),
                "espesor_d": _valor_decimal_o_vacio(post.get(f"espesor_d_{sufijo}")),
                "ancho_izquierdo": _valor_decimal_o_vacio(
                    post.get(f"ancho_izquierdo_{sufijo}")
                ),
                "ancho_derecho": _valor_decimal_o_vacio(
                    post.get(f"ancho_derecho_{sufijo}")
                ),
            }
        )

    return resultado


def _mediciones_carrileria_para_template(reporte):
    guardadas = {
        int(item.get("punto")): item
        for item in (reporte.mediciones or [])
        if item.get("punto") is not None
    }

    filas = []

    for punto in PUNTOS_CARRILERIA:
        item = guardadas.get(punto, {})
        sufijo = str(punto).replace("-", "m")
        filas.append(
            {
                "punto": punto,
                "sufijo": sufijo,
                "espesor_a": item.get("espesor_a", ""),
                "espesor_b": item.get("espesor_b", ""),
                "espesor_c": item.get("espesor_c", ""),
                "espesor_d": item.get("espesor_d", ""),
                "ancho_izquierdo": item.get("ancho_izquierdo", ""),
                "ancho_derecho": item.get("ancho_derecho", ""),
            }
        )

    return filas


def _ruedas_desde_post(post):
    ruedas = {}

    for lado in LADOS_RUEDA:
        ruedas[lado] = []

        for numero in NUMEROS_RUEDA:
            item = {"numero": numero}

            for punto in PUNTOS_RUEDA:
                item[punto] = _valor_decimal_o_vacio(
                    post.get(f"rueda_{lado}_{numero}_{punto}")
                )

            ruedas[lado].append(item)

    return ruedas


def _ruedas_para_template(reporte):
    datos = reporte.datos or {}
    guardadas = datos.get("ruedas") or {}
    resultado = {"izquierdo": [], "derecho": []}

    for lado in LADOS_RUEDA:
        por_numero = {
            int(item.get("numero")): item
            for item in guardadas.get(lado, [])
            if item.get("numero") is not None
        }

        for numero in NUMEROS_RUEDA:
            item = por_numero.get(numero, {"numero": numero})
            resultado[lado].append(
                {
                    "numero": numero,
                    "a1": item.get("a1", ""),
                    "a2": item.get("a2", ""),
                    "a3": item.get("a3", ""),
                    "b1": item.get("b1", ""),
                    "b2": item.get("b2", ""),
                    "b3": item.get("b3", ""),
                }
            )

    return resultado


def _guardar_comentarios_fotos(datos, post):
    comentarios = dict((datos or {}).get("comentarios_fotos") or {})
    for slot in FOTO_SLOTS_CARRILERIA:
        comentarios[slot["input"]] = (post.get(slot["comentario"]) or "").strip()
    datos["comentarios_fotos"] = comentarios
    return datos


def _guardar_fotos_carrileria(reporte, files):
    for slot in FOTO_SLOTS_CARRILERIA:
        imagen = files.get(slot["input"])
        if not imagen:
            continue

        FotoReporteFiltro.objects.filter(
            reporte=reporte,
            seccion=slot["seccion"],
            orden=slot["orden"],
        ).delete()
        FotoReporteFiltro.objects.create(
            reporte=reporte,
            seccion=slot["seccion"],
            titulo=slot["titulo"],
            imagen=imagen,
            orden=slot["orden"],
        )


def _fotos_carrileria_para_template(reporte):
    fotos = {
        (foto.seccion, foto.orden): foto
        for foto in reporte.fotografias.filter(
            seccion__in=["zonas_soldadura", "ruedas"],
        )
    }
    comentarios = (reporte.datos or {}).get("comentarios_fotos") or {}

    return [
        {
            **slot,
            "foto": fotos.get((slot["seccion"], slot["orden"])),
            "comentario_valor": comentarios.get(slot["input"], ""),
        }
        for slot in FOTO_SLOTS_CARRILERIA
    ]


def _guardar_comentarios_faja_poleas(datos, post):
    comentarios = dict((datos or {}).get("comentarios_faja_poleas") or {})
    for slot in FOTO_SLOTS_FAJA_POLEAS:
        comentarios[slot["input"]] = (post.get(slot["comentario"]) or "").strip()
    datos["comentarios_faja_poleas"] = comentarios
    return datos


def _guardar_fotos_faja_poleas(reporte, files):
    for slot in FOTO_SLOTS_FAJA_POLEAS:
        imagen = files.get(slot["input"])
        if not imagen:
            continue

        FotoReporteFiltro.objects.filter(
            reporte=reporte,
            seccion=slot["seccion"],
            orden=slot["orden"],
        ).delete()
        FotoReporteFiltro.objects.create(
            reporte=reporte,
            seccion=slot["seccion"],
            titulo=slot["titulo"],
            imagen=imagen,
            orden=slot["orden"],
        )


def _fotos_faja_poleas_para_template(reporte):
    fotos = {
        (foto.seccion, foto.orden): foto
        for foto in reporte.fotografias.filter(
            seccion__in=["faja_empalme", "polea_cola", "polea_cabeza", "faja_poleas"]
        )
    }
    comentarios = (reporte.datos or {}).get("comentarios_faja_poleas") or {}
    grupos = []
    for grupo in ["Top Cover y Empalme", "Polea de Cola", "Polea de Cabeza"]:
        slots = []
        for slot in FOTO_SLOTS_FAJA_POLEAS:
            if slot["grupo"] != grupo:
                continue
            slots.append(
                {
                    **slot,
                    "foto": fotos.get((slot["seccion"], slot["orden"])),
                    "comentario_valor": comentarios.get(slot["input"], ""),
                }
            )
        grupos.append({"titulo": grupo, "slots": slots})
    return grupos


def _zonas_soldadura_para_template():
    positivos = list(range(36, 0, -1))
    negativos = [0] + list(range(-1, -33, -1))

    def puntos(letra, valores):
        return [
            {
                "label": f"{letra}.{valor}",
                "warning": f"{letra}.{valor}" in {"B.20", "C.25"},
            }
            for valor in valores
        ]

    return [
        {
            "titulo": "LH - LADO IZQUIERDO",
            "izquierda": puntos("A", positivos),
            "derecha": puntos("B", positivos),
        },
        {
            "titulo": "RH - LADO DERECHO",
            "izquierda": puntos("C", positivos),
            "derecha": puntos("D", positivos),
        },
        {
            "titulo": "LH - LADO IZQUIERDO",
            "izquierda": puntos("A", negativos),
            "derecha": puntos("B", negativos),
        },
        {
            "titulo": "RH - LADO DERECHO",
            "izquierda": puntos("C", negativos),
            "derecha": puntos("D", negativos),
        },
    ]


def _paradas_disponibles_usuario(usuario):
    rol = obtener_rol(usuario)

    if usuario.is_superuser or rol == "Administrador":
        return list(
            Parada.objects
            .filter(planta__iexact="Filtros")
            .order_by("-fecha_inicio", "-id")
        )

    parada_ids = (
        AccesoParada.objects
        .filter(
            usuario=usuario,
            rol=rol,
            activo=True,
        )
        .values_list("parada_id", flat=True)
        .distinct()
    )

    # La comprobación fina de fecha/contrato se realiza al abrir la parada.
    return list(
        Parada.objects
        .filter(
            id__in=parada_ids,
            planta__iexact="Filtros",
        )
        .order_by("-fecha_inicio", "-id")
    )


def _responsables_principales(parada):
    def primero(rol):
        usuarios = usuarios_vigentes_parada(parada, rol)
        return usuarios[0] if usuarios else None

    return {
        "inspector": primero("Inspector"),
        "supervisor": primero("Supervisor"),
        "analista": primero("Analista"),
        "cliente": primero("Cliente"),
    }


def _nombre_usuario(usuario):
    if usuario is None:
        return ""
    return usuario.get_full_name().strip() or usuario.username


def _responsables_formulario(parada, reporte, usuario):
    """
    Nombres visibles del reporte, tomados de AccesoParada.

    - Inspectores de campo: todos los inspectores vigentes asignados.
    - Supervisor: todos los supervisores vigentes asignados.
    - Analista: todos los analistas vigentes asignados.
    - Inspector que elabora: el Inspector autenticado que está trabajando
        el reporte; si abre un Administrador, se conserva el inspector principal.
    """
    responsables = dict(contexto_responsables_parada(parada))
    rol_usuario = obtener_rol(usuario)

    inspector_elabora = ""

    if (
        rol_usuario == "Inspector"
        and usuario_asignado_a_parada(usuario, parada, "Inspector")
    ):
        inspector_elabora = _nombre_usuario(usuario)
    elif reporte.inspector_id:
        inspector_elabora = _nombre_usuario(reporte.inspector)
    else:
        inspectores = usuarios_vigentes_parada(parada, "Inspector")
        if inspectores:
            inspector_elabora = _nombre_usuario(inspectores[0])

    responsables["inspector_elabora"] = inspector_elabora
    return responsables


def _crear_reporte_desde_catalogo(parada, codigo_catalogo, usuario):
    config = REPORTES_FILTROS.get(codigo_catalogo)

    if not config:
        raise Http404("El código solicitado no existe en el catálogo.")

    responsables = _responsables_principales(parada)

    reporte, creado = ReporteFiltro.objects.get_or_create(
        parada=parada,
        codigo_catalogo=codigo_catalogo,
        defaults={
            "area": config["area"],
            "tag": config["tag"],
            "componente": config["componente"],
            "familia": config["familia"],
            "tecnica": config["tecnica"],
            "codigo_reporte": config["codigo_reporte"],
            "inspector": responsables["inspector"],
            "supervisor": responsables["supervisor"],
            "analista": responsables["analista"],
            "cliente": responsables["cliente"],
            "creado_por": usuario,
        },
    )

    # Mientras no esté publicado, mantenemos los FK principales sincronizados
    # con la asignación actual de la parada. El listado completo se obtiene de
    # AccesoParada y se muestra en el formulario.
    if reporte.estado != ReporteFiltro.Estado.PUBLICADO:
        cambios = []

        for campo in ("inspector", "supervisor", "analista", "cliente"):
            nuevo = responsables[campo]
            actual_id = getattr(reporte, f"{campo}_id")

            if actual_id != getattr(nuevo, "id", None):
                setattr(reporte, campo, nuevo)
                cambios.append(campo)

        if cambios:
            reporte.save(update_fields=cambios)

    return reporte, config



@login_required
def nueva_parada_filtros(request):
    """
    Crea una parada exclusiva de Filtros.

    IMPORTANTE:
    - reutiliza inspecciones.Parada y accounts.AccesoParada;
    - NO crea Inspeccion de Chancado;
    - crea los ReporteFiltro definidos en catalogo.py;
    - registra el equipo autorizado de la parada.
    """
    rol = obtener_rol(request.user)

    if not (request.user.is_superuser or rol == "Administrador"):
        return HttpResponseForbidden(
            "Solo el Administrador puede crear una parada de Filtros."
        )

    if request.method == "POST":
        form = NuevaParadaFiltrosForm(request.POST)

        if form.is_valid():
            with transaction.atomic():
                parada = Parada.objects.create(
                    nombre=form.cleaned_data["nombre"].strip(),
                    planta="Filtros",
                    fecha_inicio=form.cleaned_data["fecha_inicio"],
                    fecha_fin=form.cleaned_data["fecha_fin"],
                    estado=Parada.Estado.PROGRAMADA,
                    observaciones=form.cleaned_data["observaciones"],
                    creado_por=request.user,
                )

                asignaciones = {
                    "Inspector": list(form.cleaned_data["inspectores"]),
                    "Supervisor": list(form.cleaned_data["supervisores"]),
                    "Analista": list(form.cleaned_data["analistas"]),
                    "Cliente": [form.cleaned_data["cliente"]],
                }

                for rol_acceso, usuarios in asignaciones.items():
                    for usuario in usuarios:
                        AccesoParada.objects.create(
                            parada=parada,
                            usuario=usuario,
                            rol=rol_acceso,
                            fecha_inicio=parada.fecha_inicio,
                            fecha_fin=parada.fecha_fin,
                            activo=True,
                            creado_por=request.user,
                        )

                        HistorialAsignacionParada.objects.create(
                            parada=parada,
                            rol=rol_acceso,
                            usuario_anterior=None,
                            usuario_nuevo=usuario,
                            motivo="Asignación inicial - Parada de Filtros",
                            cambiado_por=request.user,
                        )

                inspector_principal = asignaciones["Inspector"][0]
                supervisor_principal = asignaciones["Supervisor"][0]
                analista_principal = asignaciones["Analista"][0]
                cliente_principal = asignaciones["Cliente"][0]

                for codigo_catalogo, config in REPORTES_FILTROS.items():
                    ReporteFiltro.objects.create(
                        parada=parada,
                        codigo_catalogo=codigo_catalogo,
                        area=config["area"],
                        tag=config["tag"],
                        componente=config["componente"],
                        familia=config["familia"],
                        tecnica=config["tecnica"],
                        codigo_reporte=config["codigo_reporte"],
                        fecha_programada=parada.fecha_inicio,
                        inspector=inspector_principal,
                        supervisor=supervisor_principal,
                        analista=analista_principal,
                        cliente=cliente_principal,
                        estado=ReporteFiltro.Estado.BORRADOR,
                        creado_por=request.user,
                        datos={
                            "responsables": contexto_responsables_parada(parada),
                        },
                    )

            messages.success(
                request,
                (
                    "Parada de Filtros creada correctamente. "
                    f"Se generaron {len(REPORTES_FILTROS)} reportes."
                ),
            )

            return redirect("filtros_mm:inicio", parada_id=parada.id)
    else:
        form = NuevaParadaFiltrosForm()

    return render(
        request,
        "filtros_mm/nueva_parada.html",
        {
            "form": form,
            "rol": rol,
        },
    )

@login_required
def lista_paradas_filtros(request):
    paradas = _paradas_disponibles_usuario(request.user)

    return render(
        request,
        "filtros_mm/paradas.html",
        {
            "paradas": paradas,
            "rol": obtener_rol(request.user),
        },
    )


@login_required
def inicio_filtros(request, parada_id):
    parada = get_object_or_404(Parada, pk=parada_id)
    rol = obtener_rol(request.user)

    if not (
        request.user.is_superuser
        or rol == "Administrador"
        or usuario_asignado_a_parada(request.user, parada, rol)
    ):
        return HttpResponseForbidden(
            "No tienes una asignación vigente para esta parada."
        )

    reportes = [
        {
            "codigo": codigo,
            **config,
        }
        for codigo, config in REPORTES_FILTROS.items()
    ]

    return render(
        request,
        "filtros_mm/inicio.html",
        {
            "parada": parada,
            "reportes": reportes,
            "responsables": contexto_responsables_parada(parada),
            "rol": rol,
        },
    )


@login_required
def formulario_reporte(request, parada_id, codigo_catalogo):
    parada = get_object_or_404(Parada, pk=parada_id)

    reporte, config = _crear_reporte_desde_catalogo(
        parada,
        codigo_catalogo,
        request.user,
    )

    if not puede_abrir_reporte_filtros(request.user, reporte):
        return HttpResponseForbidden(
            "No tienes permiso para abrir este reporte."
        )

    if reporte.familia not in ("CARRILERIA", "FAJA_POLEAS"):
        return HttpResponseForbidden(
            "Esta familia todavía no está habilitada en el formulario genérico."
        )

    puede_editar = puede_editar_reporte_filtros(
        request.user,
        reporte,
    )

    responsables_formulario = _responsables_formulario(
        parada,
        reporte,
        request.user,
    )

    if request.method == "POST":
        if not puede_editar:
            return HttpResponseForbidden(
                "Este usuario puede revisar el reporte, pero no modificar "
                "las mediciones técnicas."
            )

        form = ReporteFiltroCabeceraForm(
            request.POST,
            request.FILES,
            instance=reporte,
        )

        if form.is_valid():
            reporte = form.save(commit=False)
            reporte = form.guardar_datos_texto(reporte)
            if reporte.familia == "CARRILERIA":
                reporte.mediciones = _mediciones_carrileria_desde_post(
                    request.POST
                )

            datos = dict(reporte.datos or {})
            if reporte.familia == "CARRILERIA":
                datos["ruedas"] = _ruedas_desde_post(request.POST)
                datos = _guardar_comentarios_fotos(datos, request.POST)
            elif reporte.familia == "FAJA_POLEAS":
                datos = _guardar_comentarios_faja_poleas(datos, request.POST)
                datos["faja_poleas"] = _faja_poleas_desde_post(request.POST)

            # Snapshot de responsables visibles para exportación/histórico.
            datos["responsables"] = responsables_formulario
            reporte.datos = datos

            reporte.save()
            if reporte.familia == "CARRILERIA":
                _guardar_fotos_carrileria(reporte, request.FILES)
            elif reporte.familia == "FAJA_POLEAS":
                _guardar_fotos_faja_poleas(reporte, request.FILES)

            messages.success(
                request,
                "Inspección guardada correctamente.",
            )

            return redirect(
                "filtros_mm:formulario_reporte",
                parada_id=parada.id,
                codigo_catalogo=codigo_catalogo,
            )
    else:
        form = ReporteFiltroCabeceraForm(instance=reporte)

    if reporte.familia == "FAJA_POLEAS":
        return render(
            request,
            "filtros_mm/formulario_faja_poleas.html",
            {
                "parada": parada,
                "reporte": reporte,
                "reporte_config": config,
                "form": form,
                "faja_poleas": _faja_poleas_para_template(reporte),
                "fotos_faja_poleas": _fotos_faja_poleas_para_template(reporte),
                "responsables": responsables_formulario,
                "rol": obtener_rol(request.user),
                "puede_editar": puede_editar,
            },
        )

    return render(
        request,
        "filtros_mm/formulario.html",
        {
            "parada": parada,
            "reporte": reporte,
            "reporte_config": config,
            "form": form,
            "filas_medicion": _mediciones_carrileria_para_template(reporte),
            "ruedas": _ruedas_para_template(reporte),
            "fotos_carrileria": _fotos_carrileria_para_template(reporte),
            "zonas_soldadura": _zonas_soldadura_para_template(),
            "responsables": responsables_formulario,
            "rol": obtener_rol(request.user),
            "puede_editar": puede_editar,
        },
    )


@login_required
def exportar_excel_reporte(request, parada_id, codigo_catalogo):
    parada = get_object_or_404(Parada, pk=parada_id)
    reporte, config = _crear_reporte_desde_catalogo(
        parada,
        codigo_catalogo,
        request.user,
    )

    if not puede_abrir_reporte_filtros(request.user, reporte):
        return HttpResponseForbidden(
            "No tienes permiso para descargar este reporte."
        )

    if reporte.familia not in ("CARRILERIA", "FAJA_POLEAS"):
        return HttpResponseForbidden(
            "La exportacion Excel por master solo esta habilitada para esta familia."
        )

    if reporte.familia == "CARRILERIA":
        output = generar_excel_carrileria(reporte, config)
    else:
        output = generar_excel_faja_poleas(reporte, config)
    filename = (
        f"REPORTE_{reporte.codigo_reporte or codigo_catalogo}_"
        f"{reporte.tag}_{reporte.familia}.xlsx"
    )

    response = HttpResponse(
        output.getvalue(),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response

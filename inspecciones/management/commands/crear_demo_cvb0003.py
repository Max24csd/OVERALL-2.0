from datetime import date

from django.core.management.base import BaseCommand, CommandError
from django.db import transaction

from inspecciones.models import (
    Faja,
    Inspeccion,
    LifeShaftInspeccion,
    MedicionEmpalmeCVB0003,
    MedicionLifeShaft,
    MedicionPolea,
    MedicionTramoCVB0003,
    PoleaInspeccion,
    TipoMedicionComponente,
)


FECHA_HISTORICA = date(2026, 8, 3)


class Command(BaseCommand):
    help = "Crea las tres inspecciones vacías de demostración para CVB003."

    def add_arguments(self, parser):
        parser.add_argument("--fecha", default="2026-08-10")

    @transaction.atomic
    def handle(self, *args, **options):
        try:
            fecha_demo = date.fromisoformat(options["fecha"])
        except ValueError as exc:
            raise CommandError("La fecha debe usar el formato YYYY-MM-DD.") from exc

        if fecha_demo <= FECHA_HISTORICA:
            raise CommandError("La fecha demo debe ser posterior al 03/08/2026.")

        try:
            faja = Faja.objects.get(tag="CVB0003")
        except Faja.DoesNotExist as exc:
            raise CommandError("No existe la faja CVB0003.") from exc

        creadas = 0
        for tipo in (
            Inspeccion.Tipo.FAJA,
            Inspeccion.Tipo.POLEAS,
            Inspeccion.Tipo.LIFE_SHAFT,
        ):
            historica = Inspeccion.objects.get(
                faja=faja,
                tipo=tipo,
                fecha_inspeccion=FECHA_HISTORICA,
            )
            demo, creada = Inspeccion.objects.get_or_create(
                faja=faja,
                tipo=tipo,
                fecha_programada=fecha_demo,
                defaults=self._cabecera_demo(historica, fecha_demo),
            )

            if not creada:
                self.stdout.write(
                    f"{tipo}: ya existe la inspección demo #{demo.pk}; sin cambios."
                )
                continue

            if tipo == Inspeccion.Tipo.FAJA:
                self._clonar_estructura_faja(historica, demo)
            elif tipo == Inspeccion.Tipo.POLEAS:
                self._clonar_estructura_poleas(historica, demo)
            else:
                self._clonar_estructura_life_shaft(historica, demo)

            creadas += 1
            self.stdout.write(
                self.style.SUCCESS(
                    f"{tipo}: creada inspección demo #{demo.pk} ({demo.codigo_reporte})."
                )
            )

        self.stdout.write(
            self.style.SUCCESS(f"Inspecciones nuevas creadas: {creadas}.")
        )

    @staticmethod
    def _cabecera_demo(historica, fecha_demo):
        return {
            "codigo_reporte": "PENDIENTE",
            "fecha_inspeccion": fecha_demo,
            "fecha_reporte": fecha_demo,
            "inspector": historica.inspector,
            "supervisor": historica.supervisor,
            "analista": historica.analista,
            "cliente": historica.cliente,
            "estado": Inspeccion.Estado.BORRADOR,
            "condicion_general": Inspeccion.Condicion.NORMAL,
            "planta": historica.planta,
            "proceso": historica.proceso,
            "etapa": historica.etapa,
            "condicion_equipo": historica.condicion_equipo,
            "inspector_campo_nombre": historica.inspector_campo_nombre,
            "supervisor_campo_nombre": historica.supervisor_campo_nombre,
            "analista_elabora_nombre": historica.analista_elabora_nombre,
            "analista_valida_nombre": historica.analista_valida_nombre,
            "circunstancias": "",
            "antecedentes": "",
            "observaciones": "",
            "recomendaciones": "",
            "comentarios_revision": "",
            "creado_por": historica.creado_por,
        }

    @staticmethod
    def _clonar_estructura_faja(historica, demo):
        MedicionEmpalmeCVB0003.objects.bulk_create(
            MedicionEmpalmeCVB0003(
                inspeccion=demo,
                zona=fila.zona,
                empalme=fila.empalme,
                bastidor_lado=fila.bastidor_lado,
                posicion=fila.posicion,
                espesor_nominal=fila.espesor_nominal,
                observacion="",
                orden=fila.orden,
            )
            for fila in historica.empalmes_cvb0003.order_by("orden", "id")
        )
        MedicionTramoCVB0003.objects.bulk_create(
            MedicionTramoCVB0003(
                inspeccion=demo,
                tipo=fila.tipo,
                tramo=fila.tramo,
                medicion=fila.medicion,
                bastidor=fila.bastidor,
                espesor_nominal=fila.espesor_nominal,
                observacion="",
                orden=fila.orden,
            )
            for fila in historica.tramos_cvb0003.order_by("tipo", "orden", "id")
        )

    @staticmethod
    def _clonar_estructura_poleas(historica, demo):
        campos_configuracion = (
            "nombre", "tag", "ubicacion", "componente", "material",
            "espesor_nominal", "marca_equipo", "modelo_equipo",
            "frecuencia_mhz", "rango_mm", "metodo_empleado",
            "componente_calibracion", "acoplante", "rectificacion",
            "velocidad_ms", "retardo_us", "tipo_scan", "orden",
        )
        for anterior in historica.poleas_inspeccionadas.order_by("orden", "numero"):
            datos = {campo: getattr(anterior, campo) for campo in campos_configuracion}
            actual = PoleaInspeccion.objects.create(
                inspeccion=demo,
                numero=anterior.numero,
                condicion=Inspeccion.Condicion.NORMAL,
                tipo_medicion=TipoMedicionComponente.NORMAL,
                observacion_visual="",
                observacion_medicion="",
                recomendaciones="",
                **datos,
            )
            MedicionPolea.objects.bulk_create(
                MedicionPolea(
                    polea=actual,
                    punto=fila.punto,
                    posicion=fila.posicion,
                    observacion="",
                    orden=fila.orden,
                )
                for fila in anterior.mediciones.order_by("orden", "punto")
            )

    @staticmethod
    def _clonar_estructura_life_shaft(historica, demo):
        campos_configuracion = (
            "nombre", "tag", "ubicacion", "marca_equipo", "tipo_haz",
            "frecuencia_mhz", "ancho_banda", "amortiguamiento",
            "velocidad_ms", "retardo_us", "orden",
        )
        for anterior in historica.life_shafts.order_by("orden", "numero"):
            datos = {campo: getattr(anterior, campo) for campo in campos_configuracion}
            actual = LifeShaftInspeccion.objects.create(
                inspeccion=demo,
                numero=anterior.numero,
                condicion=Inspeccion.Condicion.NORMAL,
                tipo_medicion=TipoMedicionComponente.NORMAL,
                observacion_visual="",
                observacion_medicion="",
                recomendaciones="",
                **datos,
            )
            MedicionLifeShaft.objects.bulk_create(
                MedicionLifeShaft(
                    life_shaft=actual,
                    punto=fila.punto,
                    ubicacion=fila.ubicacion,
                    observacion="",
                    orden=fila.orden,
                )
                for fila in anterior.mediciones.order_by("orden", "punto")
            )

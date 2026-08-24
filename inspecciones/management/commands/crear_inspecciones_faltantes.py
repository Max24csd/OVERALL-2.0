from datetime import date

from django.contrib.auth import get_user_model
from django.core.management.base import BaseCommand
from django.db import transaction

from inspecciones.models import Faja, Inspeccion


User = get_user_model()


class Command(BaseCommand):
    help = (
        "Crea CVB0003 y CVB0004 con sus inspecciones "
        "de Faja, Poleas y Life Shaft."
    )

    @transaction.atomic
    def handle(self, *args, **options):
        self.stdout.write(
            "Creando fajas e inspecciones faltantes..."
        )

        usuarios = self.obtener_usuarios()
        fajas = self.crear_fajas()
        self.crear_inspecciones(
            fajas=fajas,
            usuarios=usuarios,
        )

        self.stdout.write("")

        self.stdout.write(
            self.style.SUCCESS(
                "CVB0003 y CVB0004 fueron configuradas correctamente."
            )
        )

    def obtener_usuarios(self):
        nombres = {
            "inspector": "inspector",
            "supervisor": "supervisor",
            "analista": "analista",
            "cliente": "cliente",
            "administrador": "administrador",
        }

        usuarios = {}

        for clave, username in nombres.items():
            try:
                usuarios[clave] = User.objects.get(
                    username=username,
                )
            except User.DoesNotExist:
                raise RuntimeError(
                    (
                        f"No existe el usuario '{username}'. "
                        "Créalo antes de ejecutar este comando."
                    )
                )

        return usuarios

    def crear_fajas(self):
        configuraciones = [
            {
                "tag": "CVB0003",
                "nombre": "Faja Overland 03",
                "descripcion": (
                    "Faja CVB0003 del proceso de Chancado."
                ),
            },
            {
                "tag": "CVB0004",
                "nombre": "Faja Overland 04",
                "descripcion": (
                    "Faja CVB0004 del proceso de Chancado."
                ),
            },
        ]

        fajas = []

        for datos in configuraciones:
            faja, creada = Faja.objects.update_or_create(
                tag=datos["tag"],
                defaults={
                    "nombre": datos["nombre"],
                    "proceso": "Chancado",
                    "descripcion": datos["descripcion"],
                    "estado": Faja.Estado.ACTIVA,
                },
            )

            fajas.append(faja)

            accion = (
                "creada"
                if creada
                else "actualizada"
            )

            self.stdout.write(
                f"Faja {faja.tag}: {accion}"
            )

        return fajas

    def crear_inspecciones(
        self,
        fajas,
        usuarios,
    ):
        tipos = [
            Inspeccion.Tipo.FAJA,
            Inspeccion.Tipo.POLEAS,
            Inspeccion.Tipo.LIFE_SHAFT,
        ]

        fecha_programada = date.today()

        for faja in fajas:
            for tipo in tipos:
                codigo_reporte = self.generar_codigo(
                    faja=faja,
                    tipo=tipo,
                    fecha=fecha_programada,
                )

                inspeccion, creada = (
                    Inspeccion.objects.update_or_create(
                        faja=faja,
                        tipo=tipo,
                        fecha_programada=fecha_programada,
                        defaults={
                            "codigo_reporte": codigo_reporte,
                            "fecha_inspeccion": fecha_programada,
                            "fecha_reporte": fecha_programada,
                            "inspector": usuarios["inspector"],
                            "supervisor": usuarios["supervisor"],
                            "analista": usuarios["analista"],
                            "cliente": usuarios["cliente"],
                            "estado": Inspeccion.Estado.BORRADOR,
                            "condicion_general": (
                                Inspeccion.Condicion.NORMAL
                            ),
                            "planta": "Chancado",
                            "proceso": "Transporte de mineral",
                            "etapa": "Operaciones",
                            "condicion_equipo": "En uso",
                            "circunstancias": (
                                "Inspección programada para "
                                f"{faja.tag}."
                            ),
                            "antecedentes": "",
                            "observaciones": "",
                            "recomendaciones": "",
                            "creado_por": usuarios[
                                "administrador"
                            ],
                        },
                    )
                )

                accion = (
                    "creada"
                    if creada
                    else "actualizada"
                )

                self.stdout.write(
                    (
                        f"{faja.tag} - "
                        f"{inspeccion.get_tipo_display()}: "
                        f"{accion}"
                    )
                )

    @staticmethod
    def generar_codigo(
        faja,
        tipo,
        fecha,
    ):
        nombres_tipo = {
            Inspeccion.Tipo.FAJA: "FAJA",
            Inspeccion.Tipo.POLEAS: "POLEAS",
            Inspeccion.Tipo.LIFE_SHAFT: "LIFE-SHAFT",
        }

        return (
            f"{fecha:%Y%m%d}-VTUT-"
            f"{faja.tag}-"
            f"{nombres_tipo[tipo]}"
        )
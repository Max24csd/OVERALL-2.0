from datetime import date, timedelta

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group, Permission
from django.core.management.base import BaseCommand
from django.db import transaction

from inspecciones.models import Faja, Inspeccion


User = get_user_model()


class Command(BaseCommand):
    help = (
        "Crea los roles, usuarios de prueba, las tres fajas "
        "de Chancado y sus nueve inspecciones iniciales."
    )

    @transaction.atomic
    def handle(self, *args, **options):
        self.stdout.write("Configurando Overall Chancado...")

        grupos = self.crear_grupos()
        usuarios = self.crear_usuarios(grupos)
        fajas = self.crear_fajas()
        self.crear_inspecciones(fajas, usuarios)

        self.stdout.write("")
        self.stdout.write(
            self.style.SUCCESS(
                "Configuración terminada correctamente."
            )
        )

        self.mostrar_credenciales()

    def crear_grupos(self):
        nombres = [
            "Administrador",
            "Analista",
            "Supervisor",
            "Inspector",
            "Cliente",
        ]

        grupos = {}

        for nombre in nombres:
            grupo, creado = Group.objects.get_or_create(
                name=nombre
            )

            grupos[nombre] = grupo

            texto = "creado" if creado else "actualizado"
            self.stdout.write(
                f"Grupo {nombre}: {texto}"
            )

        self.asignar_permisos(grupos)

        return grupos

    def asignar_permisos(self, grupos):
        permisos_inspecciones = Permission.objects.filter(
            content_type__app_label="inspecciones"
        )

        permisos_usuarios = Permission.objects.filter(
            content_type__app_label="auth",
            content_type__model__in=["user", "group"],
        )

        # Administrador:
        # todos los permisos operativos y administración de usuarios.
        grupos["Administrador"].permissions.set(
            list(permisos_inspecciones) +
            list(permisos_usuarios)
        )

        # Analista:
        # todos los permisos del módulo de inspecciones,
        # pero no puede administrar usuarios ni grupos.
        grupos["Analista"].permissions.set(
            permisos_inspecciones
        )

        supervisor_codigos = [
            "view_faja",
            "view_inspeccion",
            "change_inspeccion",
            "view_medicion",
            "change_medicion",
            "view_fotoinspeccion",
            "view_historialestado",
            "devolver_inspector",
            "revisar_inspeccion",
        ]

        grupos["Supervisor"].permissions.set(
            permisos_inspecciones.filter(
                codename__in=supervisor_codigos
            )
        )

        inspector_codigos = [
            "view_faja",
            "view_inspeccion",
            "change_inspeccion",
            "view_medicion",
            "add_medicion",
            "change_medicion",
            "delete_medicion",
            "view_fotoinspeccion",
            "add_fotoinspeccion",
            "change_fotoinspeccion",
            "delete_fotoinspeccion",
            "view_historialestado",
            "add_historialestado",
            "enviar_revision",
        ]

        grupos["Inspector"].permissions.set(
            permisos_inspecciones.filter(
                codename__in=inspector_codigos
            )
        )

        cliente_codigos = [
            "view_faja",
            "view_inspeccion",
            "view_medicion",
            "view_fotoinspeccion",
        ]

        grupos["Cliente"].permissions.set(
            permisos_inspecciones.filter(
                codename__in=cliente_codigos
            )
        )

        self.stdout.write(
            self.style.SUCCESS(
                "Permisos asignados a los cinco roles."
            )
        )

    def crear_usuarios(self, grupos):
        configuraciones = [
            {
                "username": "administrador",
                "email": "administrador@overall.local",
                "first_name": "Administrador",
                "last_name": "Overall",
                "password": "Admin2026*",
                "grupo": "Administrador",
                "is_staff": True,
            },
            {
                "username": "analista",
                "email": "analista@overall.local",
                "first_name": "Analista",
                "last_name": "Técnico",
                "password": "Analista2026*",
                "grupo": "Analista",
                "is_staff": True,
            },
            {
                "username": "supervisor",
                "email": "supervisor@overall.local",
                "first_name": "Supervisor",
                "last_name": "Chancado",
                "password": "Supervisor2026*",
                "grupo": "Supervisor",
                "is_staff": True,
            },
            {
                "username": "inspector",
                "email": "inspector@overall.local",
                "first_name": "Inspector",
                "last_name": "Campo",
                "password": "Inspector2026*",
                "grupo": "Inspector",
                "is_staff": False,
            },
            {
                "username": "cliente",
                "email": "cliente@overall.local",
                "first_name": "Cliente",
                "last_name": "Las Bambas",
                "password": "Cliente2026*",
                "grupo": "Cliente",
                "is_staff": False,
            },
        ]

        usuarios = {}

        for datos in configuraciones:
            usuario, creado = User.objects.get_or_create(
                username=datos["username"],
                defaults={
                    "email": datos["email"],
                    "first_name": datos["first_name"],
                    "last_name": datos["last_name"],
                    "is_active": True,
                    "is_staff": datos["is_staff"],
                },
            )

            usuario.email = datos["email"]
            usuario.first_name = datos["first_name"]
            usuario.last_name = datos["last_name"]
            usuario.is_active = True
            usuario.is_staff = datos["is_staff"]

            # La contraseña se establece incluso si el usuario ya existía,
            # para que las credenciales de prueba sean conocidas.
            usuario.set_password(datos["password"])
            usuario.save()

            usuario.groups.clear()
            usuario.groups.add(
                grupos[datos["grupo"]]
            )

            usuarios[datos["grupo"]] = usuario

            texto = "creado" if creado else "actualizado"

            self.stdout.write(
                f"Usuario {usuario.username}: {texto}"
            )

        return usuarios

    def crear_fajas(self):
        configuraciones = [
            {
                "tag": "CVB0001",
                "nombre": "Faja Overland 01",
                "descripcion": (
                    "Faja principal del proceso de Chancado."
                ),
            },
            {
                "tag": "CVB0002",
                "nombre": "Faja Overland 02",
                "descripcion": (
                    "Segunda faja del proceso de Chancado."
                ),
            },
            {
                "tag": "CVB0003",
                "nombre": "Faja Overland 03",
                "descripcion": (
                    "Tercera faja del proceso de Chancado."
                ),
            },
        ]

        fajas = {}

        for datos in configuraciones:
            faja, creado = Faja.objects.update_or_create(
                tag=datos["tag"],
                defaults={
                    "nombre": datos["nombre"],
                    "proceso": "Chancado",
                    "descripcion": datos["descripcion"],
                    "estado": Faja.Estado.ACTIVA,
                },
            )

            fajas[datos["tag"]] = faja

            texto = "creada" if creado else "actualizada"

            self.stdout.write(
                f"Faja {faja.tag}: {texto}"
            )

        return fajas

    def crear_inspecciones(self, fajas, usuarios):
        tipos = [
            Inspeccion.Tipo.FAJA,
            Inspeccion.Tipo.POLEAS,
            Inspeccion.Tipo.LIFE_SHAFT,
        ]

        fecha_programada = date.today() + timedelta(days=1)

        for faja in fajas.values():
            for tipo in tipos:
                codigo = self.generar_codigo(
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
                            "codigo_reporte": codigo,
                            "fecha_inspeccion": fecha_programada,
                            "fecha_reporte": fecha_programada,
                            "inspector": usuarios["Inspector"],
                            "supervisor": usuarios["Supervisor"],
                            "analista": usuarios["Analista"],
                            "cliente": usuarios["Cliente"],
                            "estado": Inspeccion.Estado.BORRADOR,
                            "condicion_general": (
                                Inspeccion.Condicion.NORMAL
                            ),
                            "planta": "Chancado",
                            "proceso": "Transporte de mineral",
                            "etapa": "Operaciones",
                            "condicion_equipo": "En uso",
                            "circunstancias": (
                                "Inspección programada para el "
                                "proceso de Chancado."
                            ),
                            "antecedentes": "",
                            "observaciones": "",
                            "recomendaciones": "",
                            "creado_por": (
                                usuarios["Administrador"]
                            ),
                        },
                    )
                )

                texto = "creada" if creada else "actualizada"

                self.stdout.write(
                    f"Inspección {codigo}: {texto}"
                )

    @staticmethod
    def generar_codigo(faja, tipo, fecha):
        codigos_tipo = {
            Inspeccion.Tipo.FAJA: "FAJA",
            Inspeccion.Tipo.POLEAS: "POLEAS",
            Inspeccion.Tipo.LIFE_SHAFT: "LIFE-SHAFT",
        }

        tipo_codigo = codigos_tipo[tipo]

        return (
            f"{fecha:%Y%m%d}-VTUT-"
            f"{faja.tag}-{tipo_codigo}"
        )

    def mostrar_credenciales(self):
        self.stdout.write("")
        self.stdout.write(
            self.style.WARNING(
                "CREDENCIALES TEMPORALES"
            )
        )

        credenciales = [
            ("administrador", "Admin2026*"),
            ("analista", "Analista2026*"),
            ("supervisor", "Supervisor2026*"),
            ("inspector", "Inspector2026*"),
            ("cliente", "Cliente2026*"),
        ]

        for usuario, clave in credenciales:
            self.stdout.write(
                f"{usuario:<15} {clave}"
            )

        self.stdout.write("")
        self.stdout.write(
            self.style.WARNING(
                "Cambia estas contraseñas antes de publicar."
            )
        )
import os

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand
from django.db import transaction

from inspecciones.models import Faja


ROLES = (
    "Administrador",
    "Inspector",
    "Supervisor",
    "Analista",
    "Cliente",
)


class Command(BaseCommand):
    help = "Crea roles, equipos base y, si hay variables de entorno, el admin inicial."

    @transaction.atomic
    def handle(self, *args, **options):
        grupos = {}
        for role in ROLES:
            grupos[role], created = Group.objects.get_or_create(name=role)
            status = "creado" if created else "existente"
            self.stdout.write(f"Rol {role}: {status}")

        self.crear_equipos_chancado()
        self.crear_equipos_molienda()

        username = os.environ.get("ADMIN_USERNAME")
        password = os.environ.get("ADMIN_PASSWORD")
        email = os.environ.get("ADMIN_EMAIL", "")

        if not username or not password:
            self.stdout.write(
                self.style.WARNING(
                    "ADMIN_USERNAME y ADMIN_PASSWORD no definidos; "
                    "solo se crearon/verificaron los roles."
                )
            )
            return

        User = get_user_model()
        user, created = User.objects.get_or_create(
            username=username,
            defaults={
                "email": email,
                "first_name": "Administrador",
                "last_name": "Overall",
                "is_active": True,
                "is_staff": True,
                "is_superuser": True,
            },
        )

        user.email = email
        user.is_active = True
        user.is_staff = True
        user.is_superuser = True
        user.set_password(password)
        user.save()
        user.groups.add(grupos["Administrador"])

        status = "creado" if created else "actualizado"
        self.stdout.write(
            self.style.SUCCESS(
                f"Administrador {username}: {status}"
            )
        )

    def crear_equipos_chancado(self):
        equipos = [
            {
                "tag": "CVB0001",
                "nombre": "Faja Overland 01",
                "descripcion": "Faja principal del proceso de Chancado.",
            },
            {
                "tag": "CVB0003",
                "nombre": "Faja Overland 03",
                "descripcion": "Faja del proceso de Chancado.",
            },
            {
                "tag": "CVB0004",
                "nombre": "Faja Overland 04",
                "descripcion": "Faja del proceso de Chancado.",
            },
        ]

        for equipo in equipos:
            _faja, created = Faja.objects.update_or_create(
                tag=equipo["tag"],
                defaults={
                    "nombre": equipo["nombre"],
                    "proceso": "Chancado",
                    "descripcion": equipo["descripcion"],
                    "estado": Faja.Estado.ACTIVA,
                },
            )
            status = "creado" if created else "actualizado"
            self.stdout.write(f"Equipo {equipo['tag']}: {status}")

    def crear_equipos_molienda(self):
        equipos = [
            {
                "tag": "CVB0006",
                "nombre": "Faja 06",
                "descripcion": "Faja 6 del proceso de Molienda.",
            },
            {
                "tag": "CVB0007",
                "nombre": "Faja 07",
                "descripcion": "Faja 7 del proceso de Molienda.",
            },
            {
                "tag": "CVB0010",
                "nombre": "Faja 10",
                "descripcion": "Faja 10 del proceso de Molienda.",
            },
            {
                "tag": "CVB0010-ENTRANTE",
                "nombre": "Faja 10 Entrante",
                "descripcion": "Top Cover entrante de Faja 10 del proceso de Molienda.",
            },
            {
                "tag": "CVB0010-SALIENTE",
                "nombre": "Faja 10 Saliente",
                "descripcion": "Top Cover saliente de Faja 10 del proceso de Molienda.",
            },
            {
                "tag": "CVB0011",
                "nombre": "Faja 11",
                "descripcion": "Faja 11 del proceso de Molienda.",
            },
            {
                "tag": "CVB0015",
                "nombre": "Faja 15",
                "descripcion": "Faja 15 del proceso de Molienda.",
            },
            {
                "tag": "CVB0017",
                "nombre": "Faja 17",
                "descripcion": "Faja 17 del proceso de Molienda.",
            },
            {
                "tag": "CVB0018",
                "nombre": "Faja 18",
                "descripcion": "Faja 18 del proceso de Molienda.",
            },
        ]

        for equipo in equipos:
            _faja, created = Faja.objects.update_or_create(
                tag=equipo["tag"],
                defaults={
                    "nombre": equipo["nombre"],
                    "proceso": "Molienda",
                    "descripcion": equipo["descripcion"],
                    "estado": Faja.Estado.ACTIVA,
                },
            )
            status = "creado" if created else "actualizado"
            self.stdout.write(f"Equipo {equipo['tag']}: {status}")

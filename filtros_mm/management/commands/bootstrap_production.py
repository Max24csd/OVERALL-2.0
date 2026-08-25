import os

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.management.base import BaseCommand
from django.db import transaction


ROLES = (
    "Administrador",
    "Inspector",
    "Supervisor",
    "Analista",
    "Cliente",
)


class Command(BaseCommand):
    help = "Crea roles base y, si hay variables de entorno, el administrador inicial."

    @transaction.atomic
    def handle(self, *args, **options):
        grupos = {}
        for role in ROLES:
            grupos[role], created = Group.objects.get_or_create(name=role)
            status = "creado" if created else "existente"
            self.stdout.write(f"Rol {role}: {status}")

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

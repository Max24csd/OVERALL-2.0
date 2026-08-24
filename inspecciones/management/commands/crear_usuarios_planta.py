import secrets
import string
import unicodedata

from django.contrib.auth.models import Group, User
from django.core.management.base import BaseCommand
from django.db import transaction


USUARIOS_PLANTA = [
    {
        "nombres": "Freddy Jesus",
        "apellidos": "Carpio Ayaqui",
        "cargo": "Supervisor Operativo",
        "grupo": "Supervisor",
        "username": "freddy.carpio",
    },
    {
        "nombres": "Maximo",
        "apellidos": "Sardon Lozano",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "maximo.sardon",
    },
    {
        "nombres": "Rolny Remy",
        "apellidos": "Gomez Hilasaca",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "rolny.gomez",
    },
    {
        "nombres": "Marcelo Darwin",
        "apellidos": "Espinoza Corrales",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "marcelo.espinoza",
    },
    {
        "nombres": "Vladimiro",
        "apellidos": "Fajardo Albertis",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "vladimiro.fajardo",
    },
    {
        "nombres": "Juan David",
        "apellidos": "Aguilar Colquehuanca",
        "cargo": "Supervisor Operativo",
        "grupo": "Supervisor",
        "username": "juan.aguilar",
    },
    {
        "nombres": "Rall Fredy",
        "apellidos": "Llanos Aguilar",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "rall.llanos",
    },
    {
        "nombres": "Yerinine Vivian",
        "apellidos": "Santi Caliente",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "yerinine.santi",
    },
    {
        "nombres": "Jesus",
        "apellidos": "Rodriguez Delgado",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "jesus.rodriguez",
    },
    {
        "nombres": "Raul Steven",
        "apellidos": "Carpio Ayaqui",
        "cargo": "Inspector NDT",
        "grupo": "Inspector",
        "username": "raul.carpio",
    },
    {
        "nombres": "John Lehnon",
        "apellidos": "Alvarez Farje",
        "cargo": "Inspector NDT / Conductor",
        "grupo": "Inspector",
        "username": "john.alvarez",
    },
]


def quitar_tildes(texto):
    texto_normalizado = unicodedata.normalize(
        "NFKD",
        texto,
    )

    return "".join(
        caracter
        for caracter in texto_normalizado
        if not unicodedata.combining(caracter)
    )


def generar_password(longitud=12):
    letras = string.ascii_letters
    numeros = string.digits
    especiales = "!@#$%"

    password = [
        secrets.choice(string.ascii_uppercase),
        secrets.choice(string.ascii_lowercase),
        secrets.choice(numeros),
        secrets.choice(especiales),
    ]

    caracteres = letras + numeros + especiales

    password.extend(
        secrets.choice(caracteres)
        for _ in range(longitud - len(password))
    )

    secrets.SystemRandom().shuffle(password)

    return "".join(password)


class Command(BaseCommand):
    help = (
        "Crea grupos y usuarios de planta con "
        "contraseñas temporales."
    )

    @transaction.atomic
    def handle(self, *args, **options):
        nombres_grupos = [
            "Administrador",
            "Inspector",
            "Supervisor",
            "Analista",
            "Cliente",
        ]

        grupos = {}

        self.stdout.write("")
        self.stdout.write(
            self.style.MIGRATE_HEADING(
                "CREANDO GRUPOS"
            )
        )

        for nombre_grupo in nombres_grupos:
            grupo, creado = Group.objects.get_or_create(
                name=nombre_grupo
            )

            grupos[nombre_grupo] = grupo

            estado = (
                "CREADO"
                if creado
                else "YA EXISTÍA"
            )

            self.stdout.write(
                f"  {nombre_grupo}: {estado}"
            )

        credenciales_nuevas = []

        self.stdout.write("")
        self.stdout.write(
            self.style.MIGRATE_HEADING(
                "CREANDO USUARIOS DE PLANTA"
            )
        )

        for datos in USUARIOS_PLANTA:
            username = quitar_tildes(
                datos["username"].lower()
            )

            usuario, creado = User.objects.get_or_create(
                username=username,
                defaults={
                    "first_name": datos["nombres"],
                    "last_name": datos["apellidos"],
                    "is_active": True,
                    "is_staff": False,
                    "is_superuser": False,
                },
            )

            usuario.first_name = datos["nombres"]
            usuario.last_name = datos["apellidos"]
            usuario.is_active = True
            usuario.is_staff = False
            usuario.is_superuser = False

            # Cada trabajador tendrá solamente su rol operativo.
            usuario.groups.clear()
            usuario.groups.add(
                grupos[datos["grupo"]]
            )

            password_temporal = None

            if creado:
                password_temporal = generar_password()
                usuario.set_password(
                    password_temporal
                )

            usuario.save()

            if creado:
                credenciales_nuevas.append(
                    {
                        "nombre": (
                            f"{datos['nombres']} "
                            f"{datos['apellidos']}"
                        ),
                        "username": username,
                        "password": password_temporal,
                        "grupo": datos["grupo"],
                        "cargo": datos["cargo"],
                    }
                )

                self.stdout.write(
                    self.style.SUCCESS(
                        f"  CREADO: {username} "
                        f"({datos['grupo']})"
                    )
                )
            else:
                self.stdout.write(
                    self.style.WARNING(
                        f"  ACTUALIZADO: {username} "
                        f"({datos['grupo']})"
                    )
                )

        self.stdout.write("")
        self.stdout.write(
            self.style.MIGRATE_HEADING(
                "CREDENCIALES NUEVAS"
            )
        )

        if not credenciales_nuevas:
            self.stdout.write(
                self.style.WARNING(
                    "No se generaron contraseñas porque "
                    "todos los usuarios ya existían."
                )
            )
        else:
            self.stdout.write(
                "Guarda estas credenciales en un lugar seguro:"
            )
            self.stdout.write("")

            for credencial in credenciales_nuevas:
                self.stdout.write(
                    "----------------------------------------"
                )
                self.stdout.write(
                    f"Nombre:   {credencial['nombre']}"
                )
                self.stdout.write(
                    f"Cargo:    {credencial['cargo']}"
                )
                self.stdout.write(
                    f"Grupo:    {credencial['grupo']}"
                )
                self.stdout.write(
                    f"Usuario:  {credencial['username']}"
                )
                self.stdout.write(
                    f"Clave:    {credencial['password']}"
                )

            self.stdout.write(
                "----------------------------------------"
            )

        self.stdout.write("")
        self.stdout.write(
            self.style.SUCCESS(
                "Usuarios y grupos configurados correctamente."
            )
        )
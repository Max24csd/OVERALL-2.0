import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("accounts", "0001_initial"),
        ("inspecciones", "0011_parada"),
        migrations.swappable_dependency(
            settings.AUTH_USER_MODEL
        ),
    ]

    operations = [
        migrations.CreateModel(
            name="HistorialAsignacionParada",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "rol",
                    models.CharField(
                        choices=[
                            ("Inspector", "Inspector"),
                            ("Supervisor", "Supervisor"),
                            ("Analista", "Analista"),
                            ("Cliente", "Cliente"),
                        ],
                        max_length=20,
                    ),
                ),
                (
                    "motivo",
                    models.TextField(
                        blank=True,
                        default="",
                    ),
                ),
                (
                    "fecha",
                    models.DateTimeField(
                        auto_now_add=True,
                    ),
                ),
                (
                    "cambiado_por",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.PROTECT,
                        related_name="cambios_asignacion_parada",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
                (
                    "parada",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="historial_asignaciones",
                        to="inspecciones.parada",
                    ),
                ),
                (
                    "usuario_anterior",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.PROTECT,
                        related_name="asignaciones_parada_reemplazadas",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
                (
                    "usuario_nuevo",
                    models.ForeignKey(
                        blank=True,
                        null=True,
                        on_delete=django.db.models.deletion.PROTECT,
                        related_name="asignaciones_parada_recibidas",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
            ],
            options={
                "verbose_name": "Historial de asignación de parada",
                "verbose_name_plural": "Historial de asignaciones de parada",
                "ordering": ["-fecha"],
            },
        ),
        migrations.CreateModel(
            name="AccesoParada",
            fields=[
                (
                    "id",
                    models.BigAutoField(
                        auto_created=True,
                        primary_key=True,
                        serialize=False,
                        verbose_name="ID",
                    ),
                ),
                (
                    "rol",
                    models.CharField(
                        choices=[
                            ("Inspector", "Inspector"),
                            ("Supervisor", "Supervisor"),
                            ("Analista", "Analista"),
                            ("Cliente", "Cliente"),
                        ],
                        max_length=20,
                    ),
                ),
                (
                    "fecha_inicio",
                    models.DateTimeField(),
                ),
                (
                    "fecha_fin",
                    models.DateTimeField(),
                ),
                (
                    "activo",
                    models.BooleanField(
                        default=True,
                    ),
                ),
                (
                    "creado_en",
                    models.DateTimeField(
                        auto_now_add=True,
                    ),
                ),
                (
                    "actualizado_en",
                    models.DateTimeField(
                        auto_now=True,
                    ),
                ),
                (
                    "creado_por",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.PROTECT,
                        related_name="accesos_parada_creados",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
                (
                    "parada",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="accesos_temporales",
                        to="inspecciones.parada",
                    ),
                ),
                (
                    "usuario",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="accesos_paradas",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
            ],
            options={
                "verbose_name": "Acceso temporal a parada",
                "verbose_name_plural": "Accesos temporales a paradas",
                "ordering": ["-creado_en"],
            },
        ),
        migrations.AddConstraint(
            model_name="accesoparada",
            constraint=models.UniqueConstraint(
                fields=(
                    "parada",
                    "usuario",
                    "rol",
                ),
                name="acceso_unico_usuario_parada_rol",
            ),
        ),
    ]
import django.db.models.deletion
from django.conf import settings
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("inspecciones", "0010_historialestado_rol_accion"),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
    ]

    operations = [
        migrations.CreateModel(
            name="Parada",
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
                    "nombre",
                    models.CharField(
                        blank=True,
                        default="",
                        max_length=150,
                    ),
                ),
                (
                    "planta",
                    models.CharField(
                        default="Chancado",
                        max_length=100,
                    ),
                ),
                (
                    "fecha_inicio",
                    models.DateField(),
                ),
                (
                    "fecha_fin",
                    models.DateField(
                        blank=True,
                        null=True,
                    ),
                ),
                (
                    "estado",
                    models.CharField(
                        choices=[
                            ("PROGRAMADA", "Programada"),
                            ("EN_CURSO", "En curso"),
                            ("FINALIZADA", "Finalizada"),
                            ("CANCELADA", "Cancelada"),
                        ],
                        default="PROGRAMADA",
                        max_length=20,
                    ),
                ),
                (
                    "observaciones",
                    models.TextField(
                        blank=True,
                        default="",
                    ),
                ),
                (
                    "creada_en",
                    models.DateTimeField(
                        auto_now_add=True,
                    ),
                ),
                (
                    "actualizada_en",
                    models.DateTimeField(
                        auto_now=True,
                    ),
                ),
                (
                    "creado_por",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.PROTECT,
                        related_name="paradas_creadas",
                        to=settings.AUTH_USER_MODEL,
                    ),
                ),
            ],
            options={
                "verbose_name": "Parada",
                "verbose_name_plural": "Paradas",
                "ordering": ["-fecha_inicio", "-id"],
            },
        ),

        migrations.AddField(
            model_name="inspeccion",
            name="parada",
            field=models.ForeignKey(
                blank=True,
                null=True,
                on_delete=django.db.models.deletion.PROTECT,
                related_name="inspecciones",
                to="inspecciones.parada",
            ),
        ),
    ]
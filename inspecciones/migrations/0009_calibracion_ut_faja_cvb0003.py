import django.db.models.deletion
from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("inspecciones", "0008_condiciones_cvb0003"),
    ]

    operations = [
        migrations.CreateModel(
            name="CalibracionUTFajaCVB0003",
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
                ("numero", models.PositiveSmallIntegerField()),
                ("marca_equipo", models.CharField(default="Olympus", max_length=80)),
                ("modelo_equipo", models.CharField(default="Epoch 6Lt", max_length=80)),
                ("frecuencia_mhz", models.CharField(default="1", max_length=30)),
                ("rango_mm", models.CharField(default="30-90", max_length=30)),
                ("metodo_empleado", models.CharField(default="Pulso - eco", max_length=80)),
                ("acoplante", models.CharField(default="Echo gel", max_length=60)),
                ("rectificacion", models.CharField(default="Full", max_length=50)),
                ("velocidad_ms", models.CharField(blank=True, max_length=30)),
                ("retardo_us", models.CharField(blank=True, max_length=30)),
                ("tipo_scan", models.CharField(default="A Scan", max_length=50)),
                (
                    "inspeccion",
                    models.ForeignKey(
                        on_delete=django.db.models.deletion.CASCADE,
                        related_name="calibraciones_ut_faja_cvb0003",
                        to="inspecciones.inspeccion",
                    ),
                ),
            ],
            options={
                "ordering": ["numero"],
                "constraints": [
                    models.UniqueConstraint(
                        fields=("inspeccion", "numero"),
                        name="calibracion_ut_faja_cvb0003_unica",
                    )
                ],
            },
        ),
    ]
